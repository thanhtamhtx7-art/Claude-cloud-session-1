# -*- coding: utf-8 -*-
"""
Bộ công cụ hóa đơn - chạy giao diện trên máy của bạn.

Cách dùng: đặt file này và Bo_cong_cu_hoa_don.html vào CÙNG THƯ MỤC với main.py
(chương trình tải hóa đơn Gmail), rồi chạy:

    python app.py

Trình duyệt sẽ tự mở http://127.0.0.1:8765. Bấm "Quét & tải" trên giao diện để chạy main.py.
Chỉ dùng thư viện có sẵn của Python, không cần cài thêm gì.

Tùy chọn:
    python app.py --port 8800          đổi cổng
    python app.py --no-browser         không tự mở trình duyệt
    python app.py --main D:\\tool\\main.py   main.py nằm ở thư mục khác
"""
import argparse
import datetime as dt
import json
import os
import re
import secrets
import signal
import sqlite3
import subprocess
import sys
import threading
import time
import unicodedata
import importlib.util
import xml.etree.ElementTree as ET
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

APP_VERSION = "2.0.2"
APP_DIR = Path(__file__).resolve().parent
HTML_NAME = "Bo_cong_cu_hoa_don.html"
UI_CONFIG = APP_DIR / "app_ui.json"          # profile bạn thêm từ giao diện (không đụng tới app_config.json)
DEFAULT_PROFILES = ["lh", "S7", "Tam_LH", "THP"]
PROFILE_RE = re.compile(r"^[A-Za-z0-9_.-]{1,40}$")
ID_RE = re.compile(r"^[A-Za-z0-9_\-]{1,80}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
MAX_LINES = 30000

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ------------------------------------------------------------------ cấu hình
class Settings:
    main_py = APP_DIR / "main.py"
    python = sys.executable
    token = secrets.token_urlsafe(24)
    port = 8765


def tool_dir():
    return Settings.main_py.parent


def main_features():
    """Đọc main.py để biết có các tùy chọn tự giải CAPTCHA (bản mới) hay chưa."""
    try:
        src = Settings.main_py.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        src = ""
    return {
        "auto_captcha": "--petro-captcha-attempts" in src,
        "petro_xml": "--petro-xml" in src,
        "manual_captcha": "--petro-manual-captcha" in src,
        "keep_xml": "--keep-xml" in src,
    }


V25_NAMES = ("v25.py", "scan_pdf_v26.py", "scan_pdf_v25.py", "v26.py", "scan_pdf.py")
# Bước sửa MST bên mua sau khi v25 chạy xong (không sửa v25.py).
# v25 hay để trống MST, hoặc lấy nhầm MST bên BÁN, ở một số mẫu hóa đơn (Fast, 1C26MWX, K26T...).
# Tên đơn vị mua vẫn đọc đúng, nên suy ra MST từ tên chính thức của các công ty mình,
# hoặc từ MST nằm cuối tên file (vd 1C26MAA_02078107_0313887686). Ô sửa được tô vàng + ghi chú.
V25_FIX = r"""
import glob, json, re, time, unicodedata

FIX_KNOWN = {"0309868627": "THP", "0313887686": "LH", "0319020723": "LH-THP",
             "0313655050": "S7", "0313073567": "HN", "0301450059": "Q3"}
FIX_NAMES = [
    ("CONG TY TNHH THUONG MAI VAN TAI XAY DUNG THIEN HOANG", "0309868627"),
    ("LIEN HIEP HOP TAC XA VAN TAI THIEN HOANG PHAT", "0319020723"),
    ("LIEN HIEP HOP TAC XA VAN TAI CO GIOI THANH PHO HO CHI MINH", "0313887686"),
    ("HOP TAC XA VAN TAI CO GIOI SO 7", "0313655050"),
]

def fix_load_extra(extra):
    for k, v in (extra or {}).items():
        d = re.sub(r"\D", "", str(k))
        if len(d) == 10 and str(v).strip():
            FIX_KNOWN[d] = str(v).strip()

def fix_norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).replace("đ", "d").replace("Đ", "D")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s).strip()
    s = re.sub(r"^(?:HO TEN NGUOI MUA HANG|HO TEN NGUOI MUA|TEN NGUOI MUA HANG|TEN NGUOI MUA|TEN DON VI MUA HANG|TEN DON VI MUA|TEN DON VI|DON VI MUA HANG)\s*", "", s)
    return s

def fix_mst_from_name(name):
    n = fix_norm(name)
    if not n:
        return ""
    for key, mst in FIX_NAMES:
        if n == key or n.startswith(key + " "):
            return mst
    return ""

def fix_mst_from_file(fname):
    found = {t for t in re.findall(r"(?<!\d)(\d{10})(?!\d)", str(fname or "")) if t in FIX_KNOWN}
    return found.pop() if len(found) == 1 else ""

def fix_row(ten_file, name, mst):
    # -> (mst_moi, ly_do) hoặc ("", "")
    cur = re.sub(r"\D", "", str(mst or ""))[:10]
    if cur in FIX_KNOWN:
        return "", ""
    by_name = fix_mst_from_name(name)
    by_file = fix_mst_from_file(ten_file)
    name_blank = not str(name or "").strip() or str(name).strip().lower() == "nan"
    if by_name and by_file and by_name != by_file:
        return "", ""
    new = by_name or (by_file if name_blank else "")
    if not new:
        return "", ""
    why = ("theo tên đơn vị mua" if by_name else "theo MST trong tên file")
    why += (", thay MST " + str(mst).strip() + " (thường là MST bên bán)") if cur else ", ô MST đang trống"
    return new, why

def fix_workbook(path, log=print):
    from openpyxl import load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    head = [str(c.value or "").strip() for c in ws[1]]
    try:
        c_file, c_name, c_mst = head.index("Ten file"), head.index("Tên đơn vị mua"), head.index("Mã số thuế bên mua")
    except ValueError:
        log("[UI] Ket_Qua.xlsx không có cột MST/Tên đơn vị mua, bỏ qua bước sửa MST.")
        return 0
    yellow = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    fixed, blank_left = [], 0
    for r in range(2, ws.max_row + 1):
        f, nm, ms = ws.cell(r, c_file + 1).value, ws.cell(r, c_name + 1).value, ws.cell(r, c_mst + 1).value
        new, why = fix_row(f, nm, ms)
        if not new:
            if ms is None or not str(ms).strip():
                blank_left += 1
            continue
        mc, nc = ws.cell(r, c_mst + 1), ws.cell(r, c_name + 1)
        mc.value = new
        mc.number_format = "@"
        nc.value = FIX_KNOWN[new]
        for c in (mc, nc):
            c.fill = yellow
        mc.comment = Comment("app.py sửa " + why + ". Tên trên hóa đơn: " + str(nm or "(trống)"), "app.py")
        fixed.append((f, ms, new, why))
    if fixed:
        wb.save(path)
    n_blank = sum(1 for x in fixed if not str(x[1] or "").strip())
    log("[UI] Da sua MST ben mua: %d dong (%d o trong, %d o lay nham MST ben ban). Con trong: %d dong." % (len(fixed), n_blank, len(fixed) - n_blank, blank_left))
    for f, old, new, why in fixed[:400]:
        log("[UI]   %s: %s -> %s %s" % (f, str(old).strip() if old not in (None, "") else "(trong)", new, FIX_KNOWN[new]))
    return len(fixed)
"""

V25_RUNNER = r"""
import importlib.util, os, sys, json, glob, time
p, folder, out, skip = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] == "1"
do_fix = len(sys.argv) > 5 and sys.argv[5] == "1"
sys.path.insert(0, os.path.dirname(p))
spec = importlib.util.spec_from_file_location("v25_ui", p)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
m.IMPORT_FOLDER = folder
m.OUTPUT_PATH = out
if skip:
    _orig = m._collect_import_inputs
    def _filtered(f):
        items = _orig(f)
        keep = [it for it in items if os.path.basename(os.path.dirname(os.path.abspath(it[0]))).lower() != "hoadon_petrolimex"]
        if len(items) != len(keep):
            print("[UI] Bo qua %d file trong thu muc HoaDon_Petrolimex (da co so lieu tu XML)." % (len(items) - len(keep)), flush=True)
        return keep
    m._collect_import_inputs = _filtered
_orig0 = m._collect_import_inputs
def _no_copies(f):
    items = _orig0(f)
    keep = [it for it in items if os.path.basename(os.path.dirname(os.path.abspath(it[0]))).lower() not in ("da_rap_xml", "cac_hang_khac", "trung_lap")]
    if len(items) != len(keep):
        print("[UI] Bo qua %d file trong thu muc Cac_hang_khac / Trung_lap (PDF da co XML hoac trung)." % (len(items) - len(keep)), flush=True)
    return keep
m._collect_import_inputs = _no_copies
_skip_file = os.environ.get("UI_SKIP_STEMS_FILE")
if _skip_file and os.path.isfile(_skip_file):
    _stems = set(json.load(open(_skip_file, encoding="utf-8")))
    _orig2 = m._collect_import_inputs
    _root = os.path.normcase(os.path.abspath(folder))
    def _skip_xml(f):
        items = _orig2(f)
        keep = [it for it in items if not (os.path.normcase(os.path.dirname(os.path.abspath(it[0]))) == _root
                and it[0].lower().endswith(".pdf") and os.path.splitext(os.path.basename(it[0]))[0] in _stems)]
        if len(items) != len(keep):
            print("[UI] Bo qua %d file PDF da co so lieu tu XML (khong quet lai)." % (len(items) - len(keep)), flush=True)
        return keep
    m._collect_import_inputs = _skip_xml
_t0 = time.time() - 2
m.main()
if do_fix:
    base, ext = os.path.splitext(out)
    cands = [f for f in [out] + glob.glob(base + "_*" + ext) if os.path.isfile(f) and os.path.getmtime(f) >= _t0]
    if cands:
        target = max(cands, key=os.path.getmtime)
        try:
            fix_load_extra(getattr(m, "_MST_TO_DON_VI_MUA", {}))
            fix_load_extra(json.loads(os.environ.get("UI_MST_TABLE") or "{}"))
            fix_workbook(target, lambda t: print(t, flush=True))
        except PermissionError:
            print("[UI] Ket_Qua.xlsx dang mo trong Excel nen chua sua MST duoc. Dong file roi quet lai.", flush=True)
        except Exception as e:
            print("[UI] Loi khi sua MST: %r" % (e,), flush=True)
"""


# Tải lại hóa đơn lỗi: chạy lại ĐÚNG cách xử lý của main.py (trình duyệt ẩn, chờ nút, tải lại trang, lấy XML),
# lấy lại link đúng từ email trên Gmail, rồi mới tới nút PDF trên trang xem hóa đơn / tải thẳng HTTP.
# Import main.py như thư viện (không sửa main.py). Có file log riêng: logs/run_<profile>_tai_lai_<giờ>.log
RETRY_CODE = r"""
# -*- coding: utf-8 -*-
# Tải lại hóa đơn lỗi: DÙNG LẠI ĐÚNG CÁCH XỬ LÝ CỦA main.py (mở trình duyệt ẩn, chờ nút, tải lại trang, lấy XML...),
# lấy lại link đúng từ chính email trên Gmail, không được thì mới tải thẳng qua HTTP. Không sửa main.py.
# Tham số: <items.json> <thư mục main.py> <tham số main.py: --profile X [--download-dir D] [--keep-xml]>
import asyncio, hashlib, importlib.util, io, json, os, re, sys, time, urllib.request, zipfile
from urllib.parse import unquote

items_path, tool = sys.argv[1], sys.argv[2]
main_args = sys.argv[3:]
items = json.load(open(items_path, encoding="utf-8"))
os.chdir(tool)
sys.path.insert(0, tool)
spec = importlib.util.spec_from_file_location("main_tool", os.path.join(tool, "main.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

sys.argv = ["main.py"] + main_args
args = m.parse_args()
prof = m.safe_profile_name(args.profile) if args.profile else ""
pc = m.find_profile_config(prof) or {}
m.setup_logging(m._resolve_path(args.log_dir or pc.get("log_dir") or "logs"), (prof or "default") + "_tai_lai")
m.configure_runtime(args)
P = m.print
import quet_email as qe

PETRO = ("petrolimex", "petro")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

def link_from_content(content, supplier):
    # giống bước phân loại link trong main.py
    s = supplier or ""
    if "easyinvoice" in s:
        return m.extract_easyinvoice_link(content), "easyinvoice"
    if "meinvoice.vn" in s:
        return m.extract_meinvoice_link(content), "meinvoice"
    if "pvoil" in s:
        return m.extract_pvoil_link(content), "pvoil"
    if "smartsign.com.vn" in s:
        return m.extract_smartsign_link(content), "smartsign"
    if "fast" in s or "fastsoftware" in s:
        return m.extract_fast_link(content), "fast"
    if "ehoadon.vn" in s:
        return m.extract_ehoadon_link(content), "ehoadon"
    if "vnpt-invoice.com.vn" in s or "vnpt" in s:
        return m.extract_vnpt_link(content), "vnpt"
    return None, None

def key_from_name(sup, url):
    s = (sup or "").lower() + " " + (url or "").lower()
    for k, keys in (("easyinvoice", ["easyinvoice"]), ("meinvoice", ["meinvoice"]), ("pvoil", ["pvoil"]),
                    ("smartsign", ["smartsign"]), ("fast", ["fast"]), ("ehoadon", ["ehoadon", "bkav"]), ("vnpt", ["vnpt"])):
        if any(x in s for x in keys):
            return k
    return None

EXTRACT = {"easyinvoice": m.extract_easyinvoice_link, "meinvoice": m.extract_meinvoice_link, "pvoil": m.extract_pvoil_link,
           "smartsign": m.extract_smartsign_link, "fast": m.extract_fast_link, "ehoadon": m.extract_ehoadon_link,
           "vnpt": m.extract_vnpt_link}

PROC = {"easyinvoice": m.process_easyinvoice, "meinvoice": m.process_meinvoice, "pvoil": m.process_pvoil,
        "smartsign": m.process_smartsign, "fast": m.process_fast, "ehoadon": m.process_ehoadon, "vnpt": m.process_vnpt}

# ---------- tải thẳng qua HTTP (cách cũ, chỉ dùng khi trình duyệt không tải được) ----------
def http_pdf(url, eid, subject):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/pdf,*/*"})
    with urllib.request.urlopen(req, timeout=120) as r:
        data, ctype, cd = r.read(), r.headers.get("Content-Type", ""), r.headers.get("Content-Disposition", "")
    fname = ""
    mm = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", cd or "", re.I) or re.search(r'filename\s*=\s*"?([^";]+)"?', cd or "", re.I)
    if mm:
        fname = unquote(mm.group(1).strip().strip('"'))
    if data[:4] == b"%PDF" or data[:2] == b"PK":
        if not fname:
            so = re.search(r"s[oố]\s*(\d{3,})", subject or "", re.I)
            fname = (so.group(1) if so else eid) + (".pdf" if data[:4] == b"%PDF" else ".zip")
        fname = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", fname).strip(" ._")[:150] or (eid + ".pdf")
        name = m.get_unique_filename(m.DOWNLOAD_DIR, fname)
        path = os.path.join(m.DOWNLOAD_DIR, name)
        with open(path, "wb") as f:
            f.write(data)
        return m._finalize_download(path, name, "TAI_LAI_HTTP", eid)
    head = data[:3000].decode("utf-8", "ignore")
    t = re.search(r"<title>(.*?)</title>", head, re.I | re.S)
    raise Exception("Link trả về trang web, không phải PDF (%s)" % ((t.group(1).strip()[:80] if t else ctype) or "không rõ"))

_PDF_BTN_RE = re.compile(r"(t[aả]i|download).{0,20}pdf", re.I)

async def page_pdf(ctx, key, url, eid, content, subject):
    # Mở trang xem / tra cứu hóa đơn (cùng trang dùng lấy XML) rồi bấm nút tải PDF như người dùng
    views = [u for how, u in m.web_xml_plan({"supplier": key, "url": url, "content": content}) if how == "page"]
    if views:
        P("  Thử mở trang xem hóa đơn và bấm nút tải PDF...")
    for view in views:
        page = await ctx.new_page()
        try:
            try:
                await page.goto(view, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                if "Download is starting" not in str(e):
                    raise
            try:
                await page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            btn, opened = None, 0
            for _ in range(8):
                btn = await m._find_visible_text(page, _PDF_BTN_RE)
                if btn:
                    break
                if opened < 2:
                    menu = await m._find_visible_text(page, m._XML_MENU_RE)
                    if menu:
                        try:
                            await menu.click(timeout=5000); opened += 1
                            await page.wait_for_timeout(800)
                            continue
                        except Exception:
                            pass
                await page.wait_for_timeout(2000)
            if not btn:
                P(f"  Trang xem hóa đơn không có nút tải PDF: {view}")
                continue
            data = await m._catch_file(page, ctx, lambda: btn.click(timeout=10000))
            if data[:2] == b"PK" or data[:4] == b"%PDF":
                so = re.search(r"s[oố]\s*:?\s*(\d{3,})", subject or "", re.I)
                name = f"HoaDon_{key}_{so.group(1) if so else eid}" + (".pdf" if data[:4] == b"%PDF" else ".zip")
                name = m.get_unique_filename(m.DOWNLOAD_DIR, name)
                path = os.path.join(m.DOWNLOAD_DIR, name)
                with open(path, "wb") as f:
                    f.write(data)
                return m._finalize_download(path, name, key, eid)
            P(f"  File tải từ nút PDF không phải PDF ({len(data)} byte)")
        except Exception as e:
            P(f"  Bấm nút PDF trên trang xem hóa đơn chưa được: {repr(e)[:200]}")
        finally:
            await page.close()
    return 0

def is_petro(it):
    return (it.get("supplier") or "").lower() in PETRO or "PETRO" in (it.get("type") or "").upper() \
        or "petrolimex.com.vn" in (it.get("url") or "").lower()

class OnlyTheseState:
    # Sổ theo dõi nhưng chỉ đưa ra các hóa đơn Petro được chọn để tải lại
    def __init__(self, state, ids):
        self._s, self._ids = state, set(ids)
    def petro_pending_downloads(self, limit=None):
        rows = [r for r in self._s.petro_pending_downloads() if r.get("message_id") in self._ids]
        return rows[:limit] if limit else rows
    def __getattr__(self, name):
        return getattr(self._s, name)

async def main():
    petro_items = [it for it in items if is_petro(it)]
    link_items = [it for it in items if not is_petro(it)]
    P(f"[RETRY] Tải lại {len(items)} hóa đơn lỗi (dùng cách xử lý của main.py) | Thư mục tải: {m.DOWNLOAD_DIR}")
    if petro_items:
        P(f"[RETRY] Gồm {len(link_items)} hóa đơn tải qua link + {len(petro_items)} hóa đơn Petrolimex (tự giải CAPTCHA)")
    if m.KEEP_XML:
        P(f"[RETRY] Có lấy thêm XML vào: {m._xml_dir()}")
    state = m._open_state()
    service = creds = None
    try:
        creds = m.get_creds()
        service = m.build("gmail", "v1", credentials=creds)
    except Exception as e:
        P(f"[WARN] Không mở được Gmail ({e!r}) -> dùng link trong báo cáo lỗi")
    ok = fail = skip = 0
    sem = asyncio.Semaphore(1)
    from playwright.async_api import async_playwright
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        for i, it in enumerate(link_items, 1):
            eid, url, sup_name = it["id"], it["url"], (it.get("supplier") or "").lower()
            subject, mdate = it.get("subject") or None, it.get("date") or None
            P(f"\n[{i}/{len(items)}] {sup_name or '?'} | EMAIL={eid}")
            key, has_pdf, content = key_from_name(sup_name, url), False, ""
            # 1) Đọc lại email trên Gmail: lấy link đúng (báo cáo có thể ghi link trang đăng nhập) / tải PDF đính kèm
            if service is not None:
                try:
                    content, has_pdf, meta = qe.get_message_content_and_pdf(service, eid)
                    subject = subject or meta.get("subject"); mdate = mdate or meta.get("mail_date")
                    supplier = qe.extract_supplier(content) if not has_pdf else None
                    link, k2 = link_from_content(content, supplier) if supplier else (None, None)
                    if not link and key in EXTRACT and content:
                        link, k2 = EXTRACT[key](content), key
                    if link and link != url:
                        P(f"  Link lấy lại từ email: {link}")
                        url = link
                    key = k2 or key
                except Exception as e:
                    P(f"  [WARN] Không đọc lại được email ({e!r}) -> dùng link trong báo cáo")
            if not has_pdf and re.search(r"/account/logon", url or "", re.I):
                # Email EasyInvoice "Tài khoản tra cứu hóa đơn điện tử": chỉ báo tài khoản đăng nhập, không kèm hóa đơn
                skip += 1
                state.record(eid, m.SKIPPED, supplier=key or sup_name or None, files=[],
                             reason="KHONG_PHAI_HOA_DON: email thông báo tài khoản tra cứu, link chỉ là trang đăng nhập",
                             subject=subject, mail_date=mdate)
                P("  -> BỎ QUA: không phải email hóa đơn (thông báo tài khoản tra cứu, link chỉ là trang đăng nhập). "
                  "Sổ ghi SKIPPED. Mở email nếu cần: https://mail.google.com/mail/u/0/#all/" + eid)
                continue
            saved, how = 0, ""
            if has_pdf:
                saved = m.download_single_pdf(eid, creds)
                how = "PDF đính kèm email"
            ctx = await browser.new_context(accept_downloads=True)
            try:
                # 2) Cách xử lý của main.py (trình duyệt ẩn, chờ nút, tải lại trang, SmartSign/Bkav/EasyInvoice...)
                if not saved and key in PROC and url:
                    P(f"  Tải bằng cách xử lý {key} của main.py: {url}")
                    try:
                        saved = await PROC[key](ctx, url, sem, eid, key)
                        how = f"trình duyệt ({key})"
                    except Exception as e:
                        P(f"  -> Cách của main.py chưa được: {repr(e)[:300]}")
                # 2b) Mở trang xem hóa đơn, bấm nút "Tải tệp PDF" / "Tải hóa đơn dạng PDF"
                if not saved and key in m.WEB_XML_SUPPLIERS and url:
                    saved = await page_pdf(ctx, key, url, eid, content, subject)
                    if saved:
                        how = "nút tải PDF trên trang xem hóa đơn"
                # 3) Tải thẳng qua HTTP
                if not saved and url:
                    try:
                        P("  Thử tải thẳng qua HTTP...")
                        saved = http_pdf(url, eid, subject)
                        how = "HTTP"
                    except Exception as e:
                        P(f"  -> Tải thẳng cũng chưa được: {e}")
                # 4) XML (khi bật Lưu XML), kể cả khi PDF vẫn lỗi
                if m.KEEP_XML and key in m.WEB_XML_SUPPLIERS and not has_pdf:
                    await m.fetch_web_xml(ctx, {"supplier": key, "url": url, "email_id": eid, "content": content}, sem)
            finally:
                await ctx.close()
            files = m.pop_tracked_files(eid)
            if saved:
                ok += 1
                state.record(eid, m.DONE, supplier=key or sup_name or None, saved_count=saved, files=files,
                             reason="TAI_LAI_QUA_LINK", subject=subject, mail_date=mdate)
                P(f"  -> ĐÃ TẢI XONG ({how}): {', '.join(files) or saved}")
            else:
                fail += 1
                P("  -> VẪN LỖI. Mở email để tải tay: https://mail.google.com/mail/u/0/#all/" + eid)
        await browser.close()
    # ---- Petrolimex: tải lại bằng đúng bước tải PDF Petro của main.py (tự giải CAPTCHA, chạy ẩn) ----
    if petro_items:
        P(f"\n[RETRY] Petrolimex: tải lại {len(petro_items)} hóa đơn bằng bước tải PDF Petro của main.py")
        before = {r.get("message_id") for r in state.petro_pending_downloads()}
        ids = [it["id"] for it in petro_items if it["id"] in before]
        for it in petro_items:
            if it["id"] not in before:
                skip += 1
                P(f"  -> BỎ QUA: EMAIL={it['id']} không còn trong danh sách Petro chờ tải (đã tải hoặc trùng hóa đơn đã có)")
        if ids:
            st = await m.run_petro_downloads(OnlyTheseState(state, ids), prompt=False)
            ok += st.get("done", 0)
            fail += len(ids) - st.get("done", 0)
            if st.get("done"):
                try:
                    m._export_petro(state)
                except Exception as e:
                    P(f"[WARN] Không cập nhật được file Excel Petro: {e!r}")
            try:
                m._petro_sync_master(state)
            except Exception as e:
                P(f"[WARN] Không cập nhật được file tổng Petro: {e!r}")
    try:
        state.close()
    except Exception:
        pass
    if m.KEEP_XML:
        P(f"[KEEP-XML] Lần này đã lưu {len(m._XML_SAVED)} file XML vào: {m._xml_dir()}")
    P(f"\n[RETRY] Xong: {ok} thành công, {fail} lỗi, {skip} bỏ qua. File lưu tại: {m.DOWNLOAD_DIR}")
    P(f"[LOG] Log lần tải lại này: {m.LOG_PATH}")

asyncio.run(main())
"""


REPORT_RE = re.compile(r"^loi_can_xu_ly_(.+)_(\d{8}_\d{6})\.xlsx$", re.I)


def read_xlsx_rows(path):
    """Đọc sheet đầu của file .xlsx bằng thư viện chuẩn (không cần openpyxl)."""
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % ns["m"])))
        sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        if not sheets:
            return []
        root = ET.fromstring(z.read(sheets[0]))
    rows = []
    for row in root.iter("{%s}row" % ns["m"]):
        vals = {}
        for c in row.findall("m:c", ns):
            ref = c.get("r", "")
            col = 0
            for ch in re.match(r"[A-Z]+", ref).group(0) if re.match(r"[A-Z]+", ref) else "":
                col = col * 26 + ord(ch) - 64
            t = c.get("t")
            if t == "inlineStr":
                v = "".join(x.text or "" for x in c.iter("{%s}t" % ns["m"]))
            else:
                ve = c.find("m:v", ns)
                v = ve.text if ve is not None else ""
                if t == "s" and v != "":
                    v = shared[int(v)]
            vals[col - 1] = v or ""
        if vals:
            rows.append([vals.get(i, "") for i in range(max(vals) + 1)])
    return rows


def error_report_items(rows):
    """Lấy các dòng có link tải + Email ID từ báo cáo loi_can_xu_ly_*.xlsx."""
    if not rows:
        return []
    head = [str(h).strip() for h in rows[0]]
    def col(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return -1
    c_link, c_id, c_sup = col("Link tải / tra cứu", "Link"), col("Email ID"), col("Nhà cung cấp")
    c_sub, c_date, c_err, c_type = col("Tiêu đề email"), col("Ngày mail"), col("Lý do / chi tiết lỗi"), col("Loại")
    c_code, c_so = col("Mã tra cứu"), col("Số hóa đơn")
    out = []
    for r in rows[1:]:
        g = lambda i: (str(r[i]).strip() if 0 <= i < len(r) else "")
        link, eid = g(c_link), g(c_id)
        if not re.match(r"^https?://", link, re.I) or not re.match(r"^[A-Za-z0-9_-]{6,80}$", eid):
            continue
        err = g(c_err)
        petro = "petrolimex" in g(c_sup).lower() or "PETRO" in g(c_type).upper()
        out.append({"id": eid, "url": link, "supplier": g(c_sup)[:30], "subject": g(c_sub)[:300], "date": g(c_date)[:30],
                    "type": g(c_type)[:30], "error": (re.split(r"\\n|\n", err)[0] if err else "")[:300],
                    "petro": petro, "code": g(c_code)[:40], "so": g(c_so)[:30]})
    return out


def latest_error_report(pid):
    paths = profile_paths(pid)
    d = paths["error_dir"]
    if not d or not Path(d).is_dir():
        return None
    safe = paths["safe"].lower()
    cands = [f for f in Path(d).glob("loi_can_xu_ly_*.xlsx")
             if REPORT_RE.match(f.name) and REPORT_RE.match(f.name).group(1).lower() == safe]
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None


def v25_runner_code():
    # phần sửa MST phải định nghĩa trước khi runner gọi
    return V25_FIX + "\n" + V25_RUNNER


def find_v25():
    """Tìm file v25 (quét hóa đơn -> Excel): đường dẫn đã lưu, hoặc cạnh main.py / app.py."""
    saved = load_ui().get("v25_path")
    if saved and Path(saved).is_file():
        return Path(saved)
    for d in (tool_dir(), APP_DIR):
        for n in V25_NAMES:
            if (d / n).is_file():
                return d / n
    return None


def ddddocr_installed():
    try:
        r = subprocess.run([Settings.python, "-c", "import ddddocr"], capture_output=True, timeout=60,
                           cwd=str(tool_dir()), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0)
        return r.returncode == 0
    except Exception:
        return None


def safe_profile_name(profile):
    """Giống main.py: ký tự lạ đổi thành _."""
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(profile or "").strip())


def read_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def load_ui():
    data = read_json(UI_CONFIG, {})
    if not isinstance(data, dict):
        data = {}
    data.setdefault("profiles", [])
    return data


_UI_LOCK = threading.RLock()


def save_ui(data):
    with _UI_LOCK:                                   # nhiều luồng cùng ghi (vd bắt đầu chạy + ghi mốc bố cục)
        tmp = UI_CONFIG.with_suffix(".%d.tmp" % threading.get_ident())
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, UI_CONFIG)


def app_config_profiles():
    """Các profile có sẵn trong app_config.json của main.py (chỉ đọc)."""
    data = read_json(tool_dir() / "app_config.json", {})
    items = data.get("profiles", []) if isinstance(data, dict) else []
    return [p for p in items if isinstance(p, dict) and p.get("id")]


def resolve(path):
    if not path:
        return None
    p = Path(path)
    return p if p.is_absolute() else tool_dir() / p


def profile_paths(pid):
    """Tính đường dẫn giống configure_runtime() của main.py."""
    safe = safe_profile_name(pid)
    pc = None
    for item in app_config_profiles():
        if str(item.get("id", "")).lower() == safe.lower():
            pc = item
            break
    ui = next((p for p in load_ui()["profiles"] if str(p.get("id", "")).lower() == safe.lower()), {}) or {}
    pcd = pc or {}
    download_dir = ui.get("download_dir") or (pcd.get("download_dir") if pc else os.path.join("downloads", safe))
    token = pcd.get("token") if pc else os.path.join("tokens", "token_%s.json" % safe)
    credentials = (pcd.get("credentials") if pc else None) or "credentials.json"
    state = pcd.get("state") or os.path.join("state", "processed_%s.db" % safe)
    petro_output = pcd.get("petro_output") or os.path.join("outputs", safe, "hoa_don_petrolimex.xlsx")
    log_dir = pcd.get("log_dir") or "logs"
    error_dir = pcd.get("error_dir") or os.path.join("outputs", safe)
    dl = resolve(download_dir)
    petro_dl = resolve(pcd.get("petro_download_dir")) if pcd.get("petro_download_dir") else (dl / "HoaDon_Petrolimex" if dl else None)
    return {
        "safe": safe, "config": pc, "ui": ui,
        "download_dir": dl, "token": resolve(token), "credentials": resolve(credentials),
        "state": resolve(state), "petro_output": resolve(petro_output), "log_dir": resolve(log_dir),
        "error_dir": resolve(error_dir), "petro_download_dir": petro_dl,
    }


def all_profiles():
    seen, out = set(), []
    ui = load_ui()

    def add(pid, **extra):
        key = pid.lower()
        if key in seen:
            for o in out:
                if o["id"].lower() == key:
                    for k, v in extra.items():
                        if v:
                            o[k] = v
            return
        seen.add(key)
        item = {"id": pid, "builtin": False, "custom": False, "in_config": False, "email": "", "download_dir": ""}
        item.update(extra)
        out.append(item)

    for pid in DEFAULT_PROFILES:
        add(pid, builtin=True)
    for pc in app_config_profiles():
        add(str(pc["id"]), in_config=True)
    for p in ui["profiles"]:
        if PROFILE_RE.match(str(p.get("id", ""))):
            add(p["id"], custom=True, email=p.get("email", ""), download_dir=p.get("download_dir", ""))
    return out


def db_connect(path):
    uri = Path(path).resolve().as_uri() + "?mode=ro"
    try:
        return sqlite3.connect(uri, uri=True, timeout=5)
    except sqlite3.Error:
        return sqlite3.connect(str(path), timeout=5)


def table_cols(conn, table):
    try:
        return {r[1] for r in conn.execute("PRAGMA table_info(%s)" % table).fetchall()}
    except sqlite3.Error:
        return set()


def profile_status(p):
    paths = profile_paths(p["id"])
    st = dict(p)
    st.update({
        "logged_in": bool(paths["token"] and paths["token"].exists()),
        "credentials_found": bool(paths["credentials"] and paths["credentials"].exists()),
        "download_dir": str(paths["download_dir"] or ""),
        "state_path": str(paths["state"] or ""),
        "counts": {}, "petro": {}, "petro_todo": 0, "last_run": None,
    })
    if paths["state"] and paths["state"].exists():
        try:
            conn = db_connect(paths["state"])
            try:
                if "status" in table_cols(conn, "emails"):
                    st["counts"] = {r[0]: r[1] for r in conn.execute("SELECT status, COUNT(*) FROM emails GROUP BY status")}
                pcols = table_cols(conn, "petro_invoices")
                if "pdf_status" in pcols:
                    st["petro"] = {r[0]: r[1] for r in conn.execute(
                        "SELECT COALESCE(pdf_status,'PENDING'), COUNT(*) FROM petro_invoices GROUP BY 1")}
                    if "ma_tra_cuu" in pcols:      # hóa đơn có mã tra cứu, chưa tải được PDF (giống nút "Tải PDF còn thiếu")
                        st["petro_todo"] = conn.execute(
                            "SELECT COUNT(*) FROM petro_invoices WHERE COALESCE(ma_tra_cuu,'') <> '' "
                            "AND UPPER(COALESCE(pdf_status,'PENDING')) IN ('PENDING','FAILED')").fetchone()[0]
            finally:
                conn.close()
        except Exception as e:
            st["db_error"] = str(e)
    try:
        logs = sorted(paths["log_dir"].glob("run_%s_*.log" % paths["safe"]), key=lambda f: f.stat().st_mtime) if paths["log_dir"] else []
        if logs:
            st["last_run"] = dt.datetime.fromtimestamp(logs[-1].stat().st_mtime).strftime("%d/%m/%Y %H:%M")
    except Exception:
        pass
    st["folders"] = folder_stats(paths)
    return st


def _pdf_mtimes(folder):
    """Thời điểm sửa của các file PDF nằm ngay trong thư mục (không có thư mục -> [])."""
    try:
        return [e.stat().st_mtime for e in os.scandir(folder) if e.is_file() and e.name.lower().endswith(".pdf")]
    except OSError:
        return []


_TONG_CACHE = {}


def tong_hop_counts(path):
    """Số dòng hóa đơn trong Tong_hop_ket_qua.xlsx theo nhóm (cột đầu 'Nhóm'). Chỉ đọc lại khi file đổi."""
    path = Path(path)
    try:
        st = path.stat()
    except OSError:
        return None
    sig = (st.st_mtime, st.st_size)
    hit = _TONG_CACHE.get(str(path))
    if hit and hit[0] == sig:
        return hit[1]
    import openpyxl
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        counts = {"petro": 0, "cac": 0, "kh": 0}
        rows = wb.worksheets[0].iter_rows(values_only=True)
        next(rows, None)                                  # dòng tiêu đề
        for r in rows:
            if not r or not any(v not in (None, "") for v in r):
                continue
            g = str(r[0] or "").lower()
            counts["petro" if g.startswith("petro") else "kh" if g.startswith("không") else "cac"] += 1
    finally:
        wb.close()
    _TONG_CACHE[str(path)] = (sig, counts)
    return counts


def folder_stats(paths):
    """Số PDF trong từng thư mục con của thư mục tải (cho trang Kết quả & Tổng hợp). Chỉ đọc, không sửa gì."""
    out = {"layout": False, "petro_pdf": 0, "cac_pdf": 0, "kh_pdf": 0, "kh_new": 0, "xml_wait": 0,
           "trung_pdf": 0, "tong_hop": "", "tong_rows": None, "kh_ket_qua": ""}
    dl = paths.get("download_dir")
    try:
        if not dl or not Path(dl).is_dir():
            return out
        dl = Path(dl)
        fmt = lambda ts: dt.datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M")
        out["layout"] = layout_since(dl) is not None
        if paths.get("petro_download_dir"):
            out["petro_pdf"] = len(_pdf_mtimes(paths["petro_download_dir"]))
        out["cac_pdf"] = len(_pdf_mtimes(dl / CAC_DIR))
        kh = _pdf_mtimes(dl / KHONG_DIR)
        out["kh_pdf"] = len(kh)
        kq = dl / KHONG_DIR / "Ket_Qua.xlsx"
        kq_time = kq.stat().st_mtime if kq.exists() else 0
        out["kh_new"] = sum(1 for t in kh if t > kq_time)         # PDF tải về sau lần quét v25 gần nhất
        if kq_time:
            out["kh_ket_qua"] = fmt(kq_time)
        xml_dir = dl / XML_SUBDIR
        if xml_dir.is_dir():
            out["xml_wait"] = sum(1 for e in os.scandir(xml_dir) if e.is_file() and e.name.lower().endswith(".xml"))
        out["trung_pdf"] = len(_pdf_mtimes(dl / TRUNG_DIR))
        th = dl / TONG_HOP
        if th.exists():
            out["tong_hop"] = fmt(th.stat().st_mtime)
            try:
                out["tong_rows"] = tong_hop_counts(th)          # {"petro", "cac", "kh"}: số hóa đơn trong Tổng hợp
            except Exception:
                pass                                            # file đang mở/ghi dở -> chỉ hiện số PDF
    except Exception:
        pass
    return out


def ledger_rows(pid, limit=5000):
    paths = profile_paths(pid)
    if not paths["state"] or not paths["state"].exists():
        return []
    conn = db_connect(paths["state"])
    try:
        cols = table_cols(conn, "emails")
        if not cols:
            return []
        want = [c for c in ("message_id", "status", "supplier", "subject", "mail_date", "reason", "attempts",
                            "saved_count", "files", "updated_at") if c in cols]
        rows = conn.execute("SELECT %s FROM emails ORDER BY updated_at DESC, message_id LIMIT ?" % ",".join(want), (int(limit),)).fetchall()
        # Kèm trạng thái tải PDF Petrolimex của email (bảng petro_invoices) để sổ email thấy được mã tải lỗi
        petro = {}
        pcols = table_cols(conn, "petro_invoices")
        if "message_id" in pcols:
            pw = [c for c in ("message_id", "so_hoa_don", "ma_tra_cuu", "pdf_status", "pdf_error", "pdf_attempts", "pdf_file", "dup_of") if c in pcols]
            for pr in conn.execute("SELECT %s FROM petro_invoices" % ",".join(pw)).fetchall():
                pd_ = dict(zip(pw, pr))
                petro[pd_["message_id"]] = pd_
        out = []
        for r in rows:
            d = dict(zip(want, r))
            if d.get("message_id") in petro:
                d["petro"] = petro[d["message_id"]]
            try:
                d["files"] = "; ".join(json.loads(d.get("files") or "[]"))
            except Exception:
                d["files"] = str(d.get("files") or "")
            out.append(d)
        return out
    finally:
        conn.close()


def petro_rows(pid):
    paths = profile_paths(pid)
    if not paths["state"] or not paths["state"].exists():
        return []
    conn = db_connect(paths["state"])
    try:
        cols = table_cols(conn, "petro_invoices")
        if not cols:
            return []
        want = [c for c in ("message_id", "mail_date", "header_date", "sender", "subject", "so_hoa_don", "ma_tra_cuu",
                            "pdf_status", "pdf_file", "pdf_error", "pdf_attempts", "dup_of") if c in cols]
        rows = conn.execute("SELECT %s FROM petro_invoices ORDER BY COALESCE(mail_date,'') DESC, message_id" % ",".join(want)).fetchall()
        return [dict(zip(want, r)) for r in rows]
    finally:
        conn.close()


# ------------------------------------------------------------------ chạy main.py
class Job:
    def __init__(self, jid, profile, action, title, commands):
        self.id = jid
        self.profile = profile
        self.action = action
        self.title = title
        self.commands = commands          # danh sách [tên bước, [tham số], chạy trong cửa sổ riêng?]
        self.lines = []
        self.status = "running"
        self.returncode = None
        self.started = time.time()
        self.ended = None
        self.proc = None
        self.part = 0
        self.stop_requested = False
        self.extra = {}
        self.result = None
        self.lock = threading.Lock()

    def add(self, text, kind="out", replace=False, t=None):
        with self.lock:
            item = {"t": t or time.strftime("%H:%M:%S"), "text": text, "kind": kind}
            if replace and self.lines and self.lines[-1].get("transient"):
                self.lines[-1] = dict(item, transient=True)
            else:
                if replace:
                    item["transient"] = True
                self.lines.append(item)
            if len(self.lines) > MAX_LINES:
                del self.lines[: len(self.lines) - MAX_LINES]
                self.lines[0] = {"t": "", "text": "... (đã lược bớt các dòng cũ)", "kind": "info"}

    def summary(self):
        return {
            "id": self.id, "profile": self.profile, "action": self.action, "title": self.title,
            "status": self.status, "returncode": self.returncode, "started": self.started, "ended": self.ended,
            "elapsed": round((self.ended or time.time()) - self.started, 1),
            "part": self.part, "parts": [c[0] for c in self.commands], "lines": len(self.lines),
            "result": self.result,
        }


DDDDOCR = {"ok": None}
JOBS = {}
JOB_ORDER = []
CURRENT = {"job": None}
JOB_LOCK = threading.Lock()


def kill_tree(proc):
    if proc is None or proc.poll() is not None:
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True, timeout=15)
        else:
            os.killpg(os.getpgid(proc.pid), signal.SIGTERM)
            for _ in range(30):
                if proc.poll() is not None:
                    return
                time.sleep(0.1)
            os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def pump_output(job, proc):
    buf = b""
    stream = proc.stdout
    while True:
        chunk = stream.read1(4096) if hasattr(stream, "read1") else stream.read(1)
        if not chunk:
            break
        buf += chunk
        while True:
            m = re.search(rb"\r\n|\n|\r", buf)
            if not m:
                break
            line, sep, buf = buf[: m.start()], m.group(0), buf[m.end():]
            text = line.decode("utf-8", errors="replace").rstrip()
            if text.strip():
                job.add(text, replace=(sep == b"\r"))
    if buf.strip():
        job.add(buf.decode("utf-8", errors="replace").rstrip())


LOG_PREFIX_RE = re.compile(r"^\[(\d{2}:\d{2}:\d{2})\] ")


def tail_log(job, proc, log_dir, safe, started, known):
    """Chế độ cửa sổ riêng: main.py in ra cửa sổ của nó, còn app.py đọc file log main.py đang ghi."""
    path, pos, buf = None, 0, ""
    while True:
        alive = proc.poll() is None
        if path is None:
            try:
                cands = [f for f in Path(log_dir).glob("run_%s_*.log" % safe)
                         if str(f) not in known and f.stat().st_mtime >= started - 2]
                if cands:
                    path = max(cands, key=lambda f: f.stat().st_mtime)
            except Exception:
                pass
        if path is not None:
            try:
                with open(path, "rb") as f:
                    f.seek(pos)
                    data = f.read()
                    pos += len(data)
                buf += data.decode("utf-8", errors="replace")
                *done, buf = buf.split("\n")
                for line in done:
                    line = line.rstrip("\r")
                    m = LOG_PREFIX_RE.match(line)
                    t = m.group(1) if m else None
                    text = line[m.end():] if m else line
                    if text.strip():
                        job.add(text.rstrip(), t=t)
            except Exception:
                pass
        if not alive:
            if buf.strip():
                job.add(buf.rstrip())
            if path is None:
                job.add("Không thấy file log của main.py (có thể main.py dừng ngay khi mở). Xem thông báo trong cửa sổ dòng lệnh vừa mở.", kind="err")
            break
        time.sleep(0.4)


def run_job(job):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1", PYTHONUTF8="1")
    rc = 0
    if job.action in ("scan", "retry_links"):
        ensure_layout_marker(job.profile)             # PDF tải từ lần chạy này được chia thư mục theo bố cục mới
    try:
        for idx, cmd_item in enumerate(job.commands):
            name, args = cmd_item[0], cmd_item[1]
            console = len(cmd_item) > 2 and cmd_item[2]
            if job.stop_requested:
                break
            job.part = idx
            program = cmd_item[3] if len(cmd_item) > 3 else "main"
            if program == "v25":
                v25p, folder, out, skip, workers = args[:5]
                fix = args[5] if len(args) > 5 else True
                cmd = [Settings.python, "-u", "-c", v25_runner_code(), v25p, folder, out, "1" if skip else "0", "1" if fix else "0"]
                env["UI_MST_TABLE"] = json.dumps(load_ui().get("mst_table") or {}, ensure_ascii=False)
                job.add("$ python %s  (quét thư mục %s, %s luồng)" % (Path(v25p).name, folder, workers), kind="cmd")
                env["INVOICE_WORKERS"] = str(workers)
                try:
                    xlog = lambda text, kind="out": job.add(text, kind=kind)
                    merge_invoice_xml(folder, xlog)
                    stems = sorted({str(r.get("Ten file") or "") for r in xml_store_rows(folder)} - {""})
                    if stems:
                        import tempfile
                        fd, tmp_skip = tempfile.mkstemp(prefix="skipxml_", suffix=".json")
                        with os.fdopen(fd, "w", encoding="utf-8") as fh:
                            json.dump(stems, fh, ensure_ascii=False)
                        job.extra.setdefault("tmp", []).append(tmp_skip)
                        env["UI_SKIP_STEMS_FILE"] = tmp_skip
                except Exception as e:
                    job.add("[XML-HD] Lỗi khi ráp XML trước khi quét: %r" % e, kind="err")
            elif program == "retry":
                cmd = [Settings.python, "-u", "-c", RETRY_CODE] + [str(a) for a in args]
                job.add("$ tải lại %s hóa đơn lỗi bằng cách xử lý của main.py (%s)" % (job.extra.get("count", "?"), " ".join(quote_arg(a) for a in args[2:])), kind="cmd")
            else:
                cmd = [Settings.python, "-u", str(Settings.main_py)] + args
                job.add("$ python main.py " + " ".join(quote_arg(a) for a in args), kind="cmd")
            if program in ("v25", "retry"):
                kwargs = dict(cwd=str(Path(args[0]).parent) if program == "v25" else str(tool_dir()), stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT, env=env)
                if os.name == "nt":
                    kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                else:
                    kwargs["start_new_session"] = True
            elif console:
                # Mở cửa sổ dòng lệnh riêng để main.py nhận là "có người ngồi máy" -> được nhập tay CAPTCHA
                paths = profile_paths(job.profile)
                log_dir = paths["log_dir"] or (tool_dir() / "logs")
                try:
                    known = {str(f) for f in Path(log_dir).glob("run_%s_*.log" % paths["safe"])}
                except Exception:
                    known = set()
                kwargs = dict(cwd=str(tool_dir()), env=env)
                if os.name == "nt":
                    kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_CONSOLE", 0x10)
                else:
                    kwargs.update(stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
                job.add("Đã mở cửa sổ dòng lệnh riêng cho main.py (để bạn nhập tay CAPTCHA khi cần). Đừng tắt cửa sổ đó.", kind="info")
            else:
                kwargs = dict(cwd=str(tool_dir()), stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT, env=env)
                if os.name == "nt":
                    kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                else:
                    kwargs["start_new_session"] = True
            started = time.time()
            try:
                proc = subprocess.Popen(cmd, **kwargs)
            except Exception as e:
                job.add("Không chạy được main.py: %r" % e, kind="err")
                rc = -1
                break
            job.proc = proc
            if console:
                tail_log(job, proc, log_dir, paths["safe"], started, known)
            else:
                pump_output(job, proc)
            rc = proc.wait()
            job.proc = None
            if job.stop_requested:
                break
            if rc != 0:
                job.add("main.py kết thúc với mã lỗi %s" % rc, kind="err")
                break
        if job.action == "scan_v25":
            out = None
            with job.lock:
                for ln in job.lines:
                    mm = re.match(r"^DONE\s*→\s*(.+)$", ln["text"].strip())
                    if mm:
                        out = mm.group(1).strip()
                    mm2 = re.search(r"File đang mở — đã lưu vào:\s*(.+)$", ln["text"])
                    if mm2:
                        out = mm2.group(1).strip()
            v_folder = Path(job.commands[0][1][1])
            if out:
                outp = Path(out)
                if not outp.is_absolute():
                    outp = Path(job.commands[0][1][0]).parent / outp
                job.result = str(outp)
            if not job.stop_requested:
                try:
                    xrows = xml_store_rows(v_folder)
                    if xrows:
                        target = Path(job.result) if job.result else v_folder / "Ket_Qua.xlsx"
                        added, replaced = sync_xml_rows(target, xrows, None)
                        job.result = str(target)
                        job.add("[XML-HD] Đã ghi %d hóa đơn lấy từ XML vào %s%s." % (
                            added, target.name, (" (thay %d dòng quét PDF)" % replaced) if replaced else ""), kind="out")
                except PermissionError:
                    job.add("[XML-HD] Không ghi được hóa đơn XML vào Ket_Qua.xlsx (file đang mở?). Đóng file rồi quét lại.", kind="err")
                except Exception as e:
                    job.add("[XML-HD] Lỗi khi ghi hóa đơn XML vào Ket_Qua.xlsx: %r" % e, kind="err")
        if job.action in ("scan", "retry_links"):
            try:
                merge_invoice_xml(profile_paths(job.profile)["download_dir"], lambda text, kind="out": job.add(text, kind=kind))
            except Exception as e:
                job.add("[XML-HD] Lỗi khi ráp XML: %r" % e, kind="err")
        if job.action in ("scan", "download_petro", "retry_links"):
            try:
                merge_petro_xml(job.profile, lambda text, kind="out": job.add(text, kind=kind))
            except Exception as e:
                job.add("[XML] Lỗi khi ráp XML: %r" % e, kind="err")
            try:
                organize_download(job.profile, lambda text, kind="out": job.add(text, kind=kind))
            except Exception as e:
                job.add("[XML-HD] Lỗi khi chia thư mục / tổng hợp: %r" % e, kind="err")
        if job.action == "scan_v25" and not job.stop_requested:
            try:
                vf = Path(job.commands[0][1][1])
                if vf.name.lower() == KHONG_DIR.lower():
                    pid = _pid_for_folder(vf.parent)
                    if pid:
                        build_tong_hop(pid, lambda text, kind="out": job.add(text, kind=kind))
            except Exception as e:
                job.add("[XML-HD] Lỗi khi ghi Tổng hợp kết quả: %r" % e, kind="err")
    finally:
        for tmp in job.extra.get("tmp", []):
            try:
                os.remove(tmp)
            except Exception:
                pass
        job.returncode = rc
        job.ended = time.time()
        job.status = "stopped" if job.stop_requested else ("done" if rc == 0 else "error")
        job.add({"done": "Đã chạy xong.", "stopped": "Đã dừng theo yêu cầu.", "error": "Chạy không thành công."}[job.status], kind="end")
        with JOB_LOCK:
            if CURRENT["job"] is job:
                CURRENT["job"] = None


def quote_arg(a):
    a = str(a)
    return '"%s"' % a.replace('"', '\\"') if re.search(r'[\s"&|<>^()]', a) or a == "" else a


def clamp_int(v, lo, hi, default=None):
    try:
        n = int(v)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, n))


def build_commands(pid, action, o):
    """Dựng tham số cho main.py từ lựa chọn trên giao diện. Chỉ các lệnh trong danh sách này được chạy."""
    if action == "retry_links":
        items = o.get("items") or []
        if not isinstance(items, list) or not items:
            raise ValueError("Chưa có hóa đơn nào để tải lại")
        clean = []
        for it in items[:300]:
            if not isinstance(it, dict):
                continue
            eid, url = str(it.get("id") or "").strip(), str(it.get("url") or "").strip()
            if re.match(r"^[A-Za-z0-9_-]{6,80}$", eid) and re.match(r"^https?://[^\s]{4,2000}$", url, re.I):
                clean.append({"id": eid, "url": url, "supplier": str(it.get("supplier") or "")[:30],
                              "subject": str(it.get("subject") or "")[:300], "date": str(it.get("date") or "")[:30],
                              "type": str(it.get("type") or "")[:30]})
        if not clean:
            raise ValueError("Không có dòng nào có link tải và Email ID hợp lệ")
        paths = profile_paths(pid)
        if not paths["download_dir"]:
            raise ValueError("Không biết thư mục tải về của profile này")
        import tempfile
        fd, tmp = tempfile.mkstemp(prefix="retry_", suffix=".json")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(clean, f, ensure_ascii=False)
        o["_tmp"] = [tmp]
        o["_count"] = len(clean)
        margs = ["--profile", pid]
        if paths["ui"].get("download_dir") and not paths["config"]:
            margs += ["--download-dir", str(paths["ui"]["download_dir"])]
        feats = main_features()
        if o.get("keep_xml") and feats["keep_xml"]:
            margs.append("--keep-xml")
        margs.append("--petro-headless")               # Petro: chỉ tự giải CAPTCHA, không chờ nhập tay
        if o.get("petro_xml") and feats["petro_xml"]:
            margs.append("--petro-xml")
        return "Tải lại hóa đơn lỗi", [["Tải lại hóa đơn lỗi", [tmp, str(tool_dir())] + margs, False, "retry"]]
    if action == "scan_v25":
        v25 = find_v25()
        if not v25:
            raise ValueError("Chưa tìm thấy file v25.py. Nhập đường dẫn file v25 ở trang Quét hóa đơn.")
        folder = Path(str(o.get("folder") or "").strip().strip('"'))
        if o.get("khong_co_xml"):
            dl = profile_paths(pid)["download_dir"]
            if not dl:
                raise ValueError("Không biết thư mục tải về của profile này")
            folder = dl / KHONG_DIR
            if not folder.is_dir() or not any(folder.glob("*.pdf")):
                raise ValueError("Thư mục %s chưa có PDF nào cần quét" % folder)
        if not str(folder) or not folder.is_absolute():
            folder = tool_dir() / folder
        if not folder.is_dir():
            raise ValueError("Không thấy thư mục: %s" % folder)
        workers = clamp_int(o.get("workers"), 1, 32, max(1, min(8, os.cpu_count() or 1)))
        out = folder / "Ket_Qua.xlsx"
        return "Quét hóa đơn → Excel (v25)", [["Quét bằng v25", [str(v25), str(folder), str(out), bool(o.get("skip_petro")), workers, o.get("fix_mst", True) is not False], False, "v25"]]
    base = ["--profile", pid]
    paths = profile_paths(pid)
    if paths["ui"].get("download_dir") and not paths["config"]:
        base += ["--download-dir", str(paths["ui"]["download_dir"])]

    feats = main_features()
    mode = str(o.get("captcha_mode") or "auto")
    if mode not in ("auto", "auto_manual", "manual"):
        raise ValueError("Cách giải CAPTCHA không hợp lệ")
    if not feats["auto_captcha"]:
        mode = "legacy"                       # main.py bản cũ: chỉ nhập tay

    def petro_opts():
        a = []
        lim = clamp_int(o.get("petro_limit"), 0, 100000, 0)
        if lim:
            a += ["--petro-limit", str(lim)]
        if mode in ("auto_manual", "manual", "legacy"):
            cap = clamp_int(o.get("captcha_timeout"), 10, 3600, None)
            if cap and cap != 90:
                a += ["--petro-captcha-timeout", str(cap)]
        if mode in ("auto", "auto_manual"):
            att = clamp_int(o.get("captcha_attempts"), 1, 50, None)
            if att and att != 5:
                a += ["--petro-captcha-attempts", str(att)]
        if mode == "auto":
            a.append("--petro-headless")
        if mode == "manual":
            a.append("--petro-manual-captcha")
        if o.get("petro_xml") and feats["petro_xml"]:
            a.append("--petro-xml")
        return a

    # Nhập tay CAPTCHA cần main.py thấy "có người ngồi máy" -> chạy trong cửa sổ dòng lệnh riêng
    console = mode in ("auto_manual", "manual")

    if action == "scan":
        a = list(base)
        src = o.get("source") or "inbox"
        if src not in ("inbox", "all", "label"):
            raise ValueError("Nguồn email không hợp lệ")
        if src != "inbox":
            a += ["--source", src]
        if src == "label":
            label = str(o.get("label") or "").strip()
            if not label or len(label) > 100:
                raise ValueError("Hãy nhập tên label Gmail")
            a += ["--label", label]
        for key, flag in (("since", "--since"), ("until", "--until")):
            v = str(o.get(key) or "").strip()
            if v:
                if not DATE_RE.match(v):
                    raise ValueError("Ngày phải có dạng YYYY-MM-DD")
                a += [flag, v]
        if o.get("since") and o.get("until") and str(o["since"]) > str(o["until"]):
            raise ValueError("Ngày bắt đầu đang sau ngày kết thúc")
        q = str(o.get("query") or "").strip()
        if q:
            if len(q) > 300:
                raise ValueError("Điều kiện Gmail quá dài")
            a += ["--query", q]
        att = clamp_int(o.get("max_attempts"), 1, 20, None)
        if att and att != 3:
            a += ["--max-attempts", str(att)]
        if o.get("retry_failed"):
            a.append("--retry-failed")
        if o.get("reprocess_skipped"):
            a.append("--reprocess-skipped")
        if o.get("force"):
            a.append("--force")
        if o.get("keep_xml") and feats["keep_xml"]:
            a.append("--keep-xml")
        if not o.get("petro_download"):
            a.append("--no-petro-download")
            return "Quét & tải", [["Quét & tải", a, False]]
        if mode == "legacy":
            # main.py bản cũ tự bỏ bước 6 khi không có bàn phím -> tách thành 2 lượt như trước
            return "Quét & tải", [["Quét & tải", a + ["--no-petro-download"], False],
                                   ["Tải PDF Petrolimex", base + ["--download-petro"] + petro_opts(), False]]
        return "Quét & tải", [["Quét & tải", a + petro_opts(), console]]
    if action == "download_petro":
        return "Tải PDF Petrolimex còn thiếu", [["Tải PDF Petrolimex", base + ["--download-petro"] + petro_opts(), console]]
    if action == "report":
        return "Báo cáo & xuất CSV", [["Báo cáo", base + ["--report", "--limit", "0"]]]
    if action == "verify":
        return ("Kiểm tra và đánh dấu email thiếu file" if o.get("fix") else "Kiểm tra file đã tải"), \
               [["Kiểm tra", base + ["--verify"] + (["--fix"] if o.get("fix") else [])]]
    if action == "export_petro":
        return "Xuất lại Excel Petrolimex", [["Xuất Excel Petro", base + ["--export-petro"]]]
    if action == "mark_done" and o.get("petro_ids"):
        # Chọn lẫn email thường và email Petrolimex tải PDF lỗi: đánh dấu email xong + hóa đơn Petro đã tải
        cmds, n = [], 0
        for key, flag in (("ids", "--mark-done"), ("petro_ids", "--mark-petro-downloaded")):
            vals = list(dict.fromkeys(str(x).strip().rstrip("*") for x in (o.get(key) or []) if str(x).strip()))
            for x in vals:
                if not ID_RE.match(x):
                    raise ValueError("Mã không hợp lệ: %s" % x)
            if vals:
                n += len(vals)
                cmds.append([("Đánh dấu %d email đã xong" if key == "ids" else "Đánh dấu %d hóa đơn Petro đã tải") % len(vals), base + [flag] + vals, False])
        if len(o.get("ids") or []) + len(o.get("petro_ids") or []) > 500:
            raise ValueError("Chọn tối đa 500 dòng mỗi lần")
        return "Đánh dấu %d mục đã xong" % n, cmds
    if action in ("mark_done", "mark_pending", "mark_petro"):
        ids = [str(x).strip().rstrip("*") for x in (o.get("ids") or []) if str(x).strip()]
        ids = list(dict.fromkeys(ids))
        if not ids:
            raise ValueError("Chưa chọn email nào")
        if len(ids) > 500:
            raise ValueError("Chọn tối đa 500 dòng mỗi lần")
        for x in ids:
            if not ID_RE.match(x):
                raise ValueError("Mã không hợp lệ: %s" % x)
        flag = {"mark_done": "--mark-done", "mark_pending": "--mark-pending", "mark_petro": "--mark-petro-downloaded"}[action]
        title = {"mark_done": "Đánh dấu %d email đã xong", "mark_pending": "Đánh dấu %d email xử lý lại",
                 "mark_petro": "Đánh dấu %d hóa đơn Petro đã tải"}[action] % len(ids)
        return title, [[title, base + [flag] + ids]]
    raise ValueError("Lệnh không được hỗ trợ: %s" % action)


def start_job(pid, action, options):
    if not PROFILE_RE.match(pid or ""):
        raise ValueError("Tên profile không hợp lệ")
    if action != "scan_v25" and not Settings.main_py.exists():
        raise ValueError("Không tìm thấy main.py tại %s" % Settings.main_py)
    options = options or {}
    title, cmds = build_commands(pid, action, options)
    with JOB_LOCK:
        cur = CURRENT["job"]
        if cur is not None and cur.status == "running":
            for tmp in options.get("_tmp") or []:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            raise RuntimeError("Đang chạy \"%s\" cho profile %s. Hãy chờ xong hoặc bấm Dừng." % (cur.title, cur.profile))
        jid = "%d-%s" % (int(time.time() * 1000), secrets.token_hex(3))
        job = Job(jid, pid, action, title, cmds)
        job.extra = {"tmp": options.get("_tmp") or [], "count": options.get("_count")}
        JOBS[jid] = job
        JOB_ORDER.append(jid)
        while len(JOB_ORDER) > 30:
            JOBS.pop(JOB_ORDER.pop(0), None)
        CURRENT["job"] = job
    threading.Thread(target=run_job, args=(job,), daemon=True).start()
    return job


def open_folder(path):
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(path))  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ------------------------------------------------------------------ Petro: XML -> Ket_Qua.xlsx
# "MST bên bán" (cuối bảng): chỉ có khi lấy từ XML; dùng để không gộp nhầm 2 người bán trùng ký hiệu + số
KQ_COLUMNS = ["Ten file", "Nha phat hanh", "Loai PDF", "Tên đơn vị mua", "Mã số thuế bên mua", "Ký hiệu",
              "Số hóa đơn", "Ngày hóa đơn", "Cộng tiền hàng", "Biển số xe", "Tên hàng hóa, dịch vụ", "Nguon",
              "MST bên bán"]
KQ_WIDTHS = [32, 14, 10, 34, 16, 12, 12, 13, 16, 14, 22, 10, 16]
KQ_TEXT_COLS = ("Mã số thuế bên mua", "Số hóa đơn", "MST bên bán")   # giữ dạng chữ (số 0 đầu)
MST_TO_DON_VI = {  # giống bảng _MST_TO_DON_VI_MUA của v25
    "0309868627": "THP", "0313887686": "LH", "0319020723": "LH-THP",
    "0313655050": "S7", "0313073567": "HN", "0301450059": "Q3",
}
PLATE_RE = re.compile(r"\b(\d{2}[A-Z]{1,2})[\s.\-]*(\d{3})[\s.\-]*(\d{2,3})(?!\d)")


def no_accent(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def _ln(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _find(node, name):
    if node is None:
        return None
    for el in node.iter():
        if _ln(el.tag) == name:
            return el
    return None


def _party(root, name):
    """Lấy khối người bán/người mua (NBan/NMua) chứa dữ liệu thật.
    Bkav (ehoadon.vn) có thêm <DSCKS><NBan/><NMua/> rỗng (chữ ký số) đứng TRƯỚC phần hóa đơn,
    nên ưu tiên khối nằm trong NDHDon, rồi tới khối đầu tiên có Ten hoặc MST."""
    ndh = _find(root, "NDHDon")
    if ndh is not None:
        el = _find(ndh, name)
        if el is not None:
            return el
    first = None
    for el in root.iter():
        if _ln(el.tag) == name:
            if first is None:
                first = el
            if _text(el, "Ten") or _text(el, "MST"):
                return el
    return first


def _text(node, name):
    el = _find(node, name)
    return (el.text or "").strip() if el is not None and el.text else ""


def _plate(text):
    norm = re.sub(r"[­‐-―]", "-", no_accent(text).upper())
    m = PLATE_RE.search(norm)
    return "%s-%s.%s" % (m.group(1), m.group(2), m.group(3)) if m else ""


def _hang_hoa(name):
    """Giống clean_hang_hoa của v25."""
    n = no_accent(name).lower()
    if not n.strip():
        return "Cần kiểm tra"
    if "sua chua" in n or "bao duong" in n:
        return "Sửa chữa"
    if "diezen" in n or "diesel" in n:
        return "Dầu"
    if "xang" in n or "ron" in n or "e5" in n or "95" in n:
        return "Xăng"
    if "dau" in n or re.search(r"\bdo\b", n):
        return "Dầu"
    return "Cần kiểm tra"


def _to_int(v):
    try:
        return int(round(float(str(v).replace(",", ""))))
    except (TypeError, ValueError):
        return None


def parse_petro_xml(path):
    return parse_invoice_xml(path, "petrolimex")


def parse_invoice_xml(path, issuer=None):
    """Đọc hóa đơn điện tử XML (chuẩn TT78: TTChung, NDHDon/NMua, DSHHDVu, TToan) -> 1 dòng Ket_Qua.
    issuer=None: tự nhận 'petrolimex' nếu tên người bán có chữ PETROLIMEX, còn lại 'standard' (giống cột của v25)."""
    root = ET.parse(str(path)).getroot()
    ttc = _find(root, "TTChung")
    so = _text(ttc, "SHDon")
    khhd = _text(ttc, "KHHDon")
    if not so or not khhd:
        raise ValueError("không thấy Số hóa đơn (SHDon) hoặc Ký hiệu (KHHDon)")
    ky_hieu = (_text(ttc, "KHMSHDon") + khhd).upper()
    so_n = so.lstrip("0") or so
    nlap = _text(ttc, "NLap")
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", nlap)
    ngay = "%s/%s/%s" % (m.group(3), m.group(2), m.group(1)) if m else nlap
    nmua = _party(root, "NMua")
    mst = re.sub(r"\s", "", _text(nmua, "MST"))
    ten_mua = _text(nmua, "Ten")
    don_vi = MST_TO_DON_VI.get(re.sub(r"\D", "", mst), ten_mua)
    # Biển số: Petrolimex ghi ở trường SoPhuongTien (phần người bán); mẫu khác ghi ở họ tên người mua
    # Một số nhà cung cấp (vd 1C26MDV, 1C26MXD) ghi biển số ở ô "số điện thoại" của người mua
    cands = [_text(nmua, "HVTNMHang"), _text(nmua, "SDThoai")]
    for tt in root.iter():
        if _ln(tt.tag) == "TTin":
            key = no_accent(_text(tt, "TTruong")).lower()
            val = _text(tt, "DLieu")
            if val and any(k in key.replace(" ", "") for k in ("sophuongtien", "phuongtien", "bienso", "soxe", "bsx", "plate")):
                cands.insert(1, val)
            elif val:
                cands.append(val)
    cands += [ten_mua, _text(nmua, "DChi")]
    bien_so = ""
    for c in cands:
        bien_so = _plate(c)
        if bien_so:
            break
    # Hàng hóa: dòng hàng đầu tiên (bỏ dòng ghi chú/chiết khấu)
    ten_hang = ""
    lines = [el for el in root.iter() if _ln(el.tag) == "HHDVu"]
    for el in lines:
        tchat = _text(el, "TChat")
        name = _text(el, "THHDVu")
        if name and tchat in ("", "1"):
            ten_hang = name
            break
    tien = _to_int(_text(_find(root, "TToan"), "TgTCThue"))
    if tien is None:
        tien = sum(_to_int(_text(el, "ThTien")) or 0 for el in lines if _text(el, "TChat") in ("", "1")) or None
    nban = _party(root, "NBan")
    if not issuer:
        seller = no_accent(_text(nban, "Ten")).upper()
        issuer = "petrolimex" if "PETROLIMEX" in seller else "standard"
    return {
        "Ten file": "", "Nha phat hanh": issuer, "Loai PDF": "xml",
        "Tên đơn vị mua": don_vi, "Mã số thuế bên mua": mst, "Ký hiệu": ky_hieu,
        "Số hóa đơn": so_n, "Ngày hóa đơn": ngay, "Cộng tiền hàng": tien if tien is not None else "Cần kiểm tra",
        "Biển số xe": bien_so or "Không biển số", "Tên hàng hóa, dịch vụ": _hang_hoa(ten_hang), "Nguon": "xml",
        "MST bên bán": re.sub(r"\s", "", _text(nban, "MST")),
        "_khhd": khhd.upper(),
    }


def _pdf_stem_for(row, pdf_stems, xml_stem):
    """Tên file PDF tương ứng (để công cụ xếp theo biển số tìm được file)."""
    if xml_stem in pdf_stems:
        return xml_stem
    so = row["Số hóa đơn"]
    hits = [st for st in pdf_stems if re.search(r"(?<!\d)0*%s(?!\d)" % re.escape(so), st)]
    both = [st for st in hits if row["_khhd"] in st.upper()]
    pick = both or hits
    if not pick:   # Bkav: XML "C26MTP-00949638-MXP5QBYA1VD-DPH" <-> PDF "HoaDon_Bkav_MXP5QBYA1VD"
        codes = [t.upper() for t in re.split(r"[_\-\s]+", xml_stem) if len(t) >= 8 and re.search(r"[A-Za-z]", t) and re.search(r"\d", t)]
        pick = [st for st in pdf_stems if any(st.upper().endswith(c) or ("_" + c + "_") in st.upper() + "_" for c in codes)]
    return sorted(pick, key=len)[0] if pick else xml_stem


def _ascii_key(v):
    return no_accent(v).lower().strip()


# ---- Lọc trùng hóa đơn: Tên đơn vị mua + Ký hiệu + Số hóa đơn (+ MST bên bán nếu có) ----
# Ô trống ở 1 bên (vd dòng v25 không đọc được tên đơn vị / không có MST bên bán) được coi là khớp,
# để số liệu XML vẫn thay được dòng quét PDF của cùng hóa đơn.
_BLANK_NAMES = {"", "NAN", "NONE", "CAN KIEM TRA"}


def _buyer_key(row):
    mst = re.sub(r"\D", "", str(row.get("Mã số thuế bên mua") or ""))[:10]
    name = MST_TO_DON_VI.get(mst) or row.get("Tên đơn vị mua")
    key = re.sub(r"[^A-Z0-9]+", " ", no_accent(name).upper()).strip()
    return "" if key in _BLANK_NAMES else key


def _seller_key(row):
    return re.sub(r"\D", "", str(row.get("MST bên bán") or ""))


class _KqIndex:
    """Tìm hóa đơn đã có theo Tên đơn vị mua + Ký hiệu + Số (+ MST bên bán)."""

    def __init__(self):
        self._by_key = {}

    @staticmethod
    def _ident(row):
        return _buyer_key(row), _seller_key(row)

    def find(self, row):
        key = _kq_key(row.get("Ký hiệu"), row.get("Số hóa đơn"))
        if not key[1]:
            return None
        mine = self._ident(row)
        for other, ref in self._by_key.get(key, []):
            if all(not a or not b or a == b for a, b in zip(mine, other)):
                return ref
        return None

    def add(self, row, ref):
        key = _kq_key(row.get("Ký hiệu"), row.get("Số hóa đơn"))
        if key[1]:
            self._by_key.setdefault(key, []).append((self._ident(row), ref))


def _kq_row_at(ws, r, cols):
    return {name: ws.cell(r, cols[name]).value for name in KQ_COLUMNS}


def merge_petro_xml(pid, log):
    """Ráp các file XML trong thư mục PDF Petro vào Ket_Qua.xlsx cùng thư mục, rồi xóa XML đã ráp."""
    paths = profile_paths(pid)
    folder = paths["petro_download_dir"]
    res = {"found": 0, "added": 0, "existing": 0, "failed": 0, "deleted": 0, "file": ""}
    if not folder or not Path(folder).is_dir():
        return res
    folder = Path(folder)
    xmls = sorted(f for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".xml")
    res["found"] = len(xmls)
    if not xmls:
        return res
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except Exception:
        log("[XML] Chưa có thư viện openpyxl nên chưa ráp được XML vào Ket_Qua.xlsx (cài: python -m pip install openpyxl).", "err")
        return res
    kq = folder / "Ket_Qua.xlsx"
    res["file"] = str(kq)
    pdf_stems = [f.stem for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"]
    parsed, bad = [], []
    for f in xmls:
        try:
            row = parse_petro_xml(f)
            row["Ten file"] = _pdf_stem_for(row, pdf_stems, f.stem)
            parsed.append((f, row))
        except Exception as e:
            bad.append(f)
            log("[XML] Không đọc được %s: %s (giữ lại file XML)" % (f.name, e), "err")
    res["failed"] = len(bad)
    if not parsed:
        return res
    try:
        if kq.exists():
            wb = openpyxl.load_workbook(kq)
            ws = wb.worksheets[0]
            hdr, cols = None, {}
            for r in range(1, min(ws.max_row or 1, 10) + 1):
                heads = {_ascii_key(ws.cell(r, c).value): c for c in range(1, (ws.max_column or 1) + 1) if ws.cell(r, c).value}
                if "ten file" in heads and "so hoa don" in heads:
                    hdr = r
                    cols = {name: heads.get(_ascii_key(name)) for name in KQ_COLUMNS}
                    break
            if hdr is None:
                raise ValueError("Ket_Qua.xlsx không có dòng tiêu đề (Ten file, Số hóa đơn)")
            nxt = max(ws.max_column or 1, 1)
            for name in KQ_COLUMNS:          # thêm cột còn thiếu vào cuối
                if not cols.get(name):
                    nxt += 1
                    ws.cell(hdr, nxt, name)
                    cols[name] = nxt
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            hdr = 1
            cols = {name: i + 1 for i, name in enumerate(KQ_COLUMNS)}
            for name, c in cols.items():
                cell = ws.cell(1, c, name)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", start_color="1F4E79")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions[cell.column_letter].width = KQ_WIDTHS[c - 1]
            ws.freeze_panes = "A2"
        have = _KqIndex()
        last = hdr
        for r in range(hdr + 1, (ws.max_row or hdr) + 1):
            old = _kq_row_at(ws, r, cols)
            if old["Ký hiệu"] or old["Số hóa đơn"] or old["Ten file"]:
                last = r
            have.add(old, r)
        for f, row in parsed:
            if have.find(row):
                res["existing"] += 1
                continue
            last += 1
            have.add(row, last)
            _kq_write_row(ws, last, cols, row)
            res["added"] += 1
        wb.save(kq)
    except PermissionError:
        log("[XML] Không ghi được %s (file đang mở trong Excel?). Đóng file rồi bấm \"Ráp XML vào Ket_Qua.xlsx\" ở trang Petrolimex. File XML vẫn được giữ lại." % kq, "err")
        return res
    except Exception as e:
        log("[XML] Lỗi khi ghi Ket_Qua.xlsx: %r. File XML vẫn được giữ lại." % e, "err")
        return res
    for f, _ in parsed:
        try:
            f.unlink()
            res["deleted"] += 1
        except Exception as e:
            log("[XML] Không xóa được %s: %r" % (f.name, e), "err")
    log("[XML] Đã ráp %d file XML vào %s: thêm %d hóa đơn mới%s, đã xóa %d file XML."
        % (len(parsed), kq, res["added"], (", %d hóa đơn đã có sẵn" % res["existing"]) if res["existing"] else "", res["deleted"]), "out")
    return res


# ------------------------------------------------------------------ XML hóa đơn thường -> Ket_Qua.xlsx (giai đoạn 2)
# main.py --keep-xml lưu XML vào <thư mục tải>/XML. app.py đọc các XML đó, lưu số liệu vào kho
# XML/_xml_da_rap.xlsx (để không mất dữ liệu khi v25 ghi đè Ket_Qua.xlsx), xóa file XML đã đọc,
# rồi ghi các hóa đơn đó vào Ket_Qua.xlsx của thư mục tải. Không đụng tới file PDF.
XML_SUBDIR = "XML"
XML_STORE = "_xml_da_rap.xlsx"


def _kq_key(ky, so):
    return (str(ky or "").strip().upper(), re.sub(r"\D", "", str(so or "")).lstrip("0"))


def _kq_open(path, create=True):
    """Mở (hoặc tạo) file dạng Ket_Qua -> (wb, ws, hdr, cols). cols: tên cột -> số cột (1-based)."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    path = Path(path)
    if path.exists():
        wb = openpyxl.load_workbook(path)
        ws = wb.worksheets[0]
        for r in range(1, min(ws.max_row or 1, 10) + 1):
            heads = {_ascii_key(ws.cell(r, c).value): c for c in range(1, (ws.max_column or 1) + 1) if ws.cell(r, c).value}
            if "ten file" in heads and "so hoa don" in heads:
                cols = {name: heads.get(_ascii_key(name)) for name in KQ_COLUMNS}
                nxt = max(ws.max_column or 1, 1)
                for name in KQ_COLUMNS:
                    if not cols.get(name):
                        nxt += 1
                        ws.cell(r, nxt, name)
                        cols[name] = nxt
                return wb, ws, r, cols
        raise ValueError("%s không có dòng tiêu đề (Ten file, Số hóa đơn)" % path.name)
    if not create:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    cols = {name: i + 1 for i, name in enumerate(KQ_COLUMNS)}
    for name, c in cols.items():
        cell = ws.cell(1, c, name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = KQ_WIDTHS[c - 1]
    ws.freeze_panes = "A2"
    return wb, ws, 1, cols


def _kq_write_row(ws, r, cols, row, only=None):
    for name in (only or KQ_COLUMNS):
        cell = ws.cell(r, cols[name], row.get(name, ""))
        if name in KQ_TEXT_COLS:
            cell.number_format = "@"
        elif name == "Cộng tiền hàng" and isinstance(row.get(name), int):
            cell.number_format = "#,##0"


def _is_blank_cell(v):
    return v is None or str(v).strip() in ("", "Không biển số", "Cần kiểm tra")


def _kq_fill_blanks(ws, r, cols, row):
    """Điền vào các ô đang trống của dòng r bằng số liệu XML (không ghi đè ô đã có). Trả về số ô đã điền."""
    names = [n for n in KQ_COLUMNS if n not in ("Ten file",)
             and _is_blank_cell(ws.cell(r, cols[n]).value) and not _is_blank_cell(row.get(n))]
    if names:
        _kq_write_row(ws, r, cols, row, only=names)
    return len(names)


def xml_store_rows(folder):
    """Các hóa đơn đã ráp từ XML của thư mục (đọc từ kho XML/_xml_da_rap.xlsx)."""
    store = Path(folder) / XML_SUBDIR / XML_STORE
    if not store.exists():
        return []
    try:
        got = _kq_open(store, create=False)
    except Exception:
        return []
    if not got:
        return []
    wb, ws, hdr, cols = got
    out = []
    for r in range(hdr + 1, (ws.max_row or hdr) + 1):
        row = {name: ws.cell(r, cols[name]).value for name in KQ_COLUMNS}
        if row.get("Số hóa đơn") not in (None, ""):
            out.append({k: ("" if v is None else v) for k, v in row.items()})
    return out


def sync_xml_rows(kq, rows, log):
    """Ghi các hóa đơn lấy từ XML vào Ket_Qua.xlsx:
      - chưa có -> thêm dòng mới
      - đã có dòng do quét PDF (v25) -> thay bằng số liệu XML (giữ biển số của v25 nếu XML không ghi biển số)
      - đã có dòng từ XML -> bỏ qua. Trả về (thêm, thay)."""
    if not rows:
        return 0, 0
    kq = Path(kq)
    wb, ws, hdr, cols = _kq_open(kq)
    index, last = _KqIndex(), hdr
    for r in range(hdr + 1, (ws.max_row or hdr) + 1):
        old = _kq_row_at(ws, r, cols)
        if old["Ký hiệu"] or old["Số hóa đơn"] or old["Ten file"]:
            last = r
        index.add(old, r)
    added = replaced = filled = 0
    for row in rows:
        r = index.find(row)
        if r:
            if str(ws.cell(r, cols["Nguon"]).value or "").strip().lower() == "xml":
                filled += 1 if _kq_fill_blanks(ws, r, cols, row) else 0
                continue
            new = dict(row)
            old_file = str(ws.cell(r, cols["Ten file"]).value or "").strip()
            if old_file:
                new["Ten file"] = old_file
            old_plate = str(ws.cell(r, cols["Biển số xe"]).value or "").strip()
            if new.get("Biển số xe") in ("", "Không biển số") and old_plate and old_plate != "Không biển số":
                new["Biển số xe"] = old_plate
            _kq_write_row(ws, r, cols, new)
            replaced += 1
        else:
            last += 1
            _kq_write_row(ws, last, cols, row)
            index.add(row, last)
            added += 1
    if added or replaced or filled:
        wb.save(kq)
    if filled and log:
        log("[XML-HD] Đã bổ sung ô còn trống cho %d hóa đơn XML trong %s." % (filled, kq.name), "out")
    return added, replaced

# ---- Bố cục thư mục tải (áp dụng cho PDF tải từ lúc bật bố cục mới; file cũ để nguyên chỗ cũ) ----
#   HoaDon_Petrolimex/  : PDF + Ket_Qua.xlsx Petrolimex (đã có sẵn)
#   Cac_hang_khac/      : PDF các hãng khác ĐÃ có XML + Ket_Qua.xlsx ráp từ XML
#   Khong_co_XML/       : PDF tải được nhưng không có XML (quét v25 -> Ket_Qua.xlsx trong đó)
#   Tong_hop_ket_qua.xlsx : Petrolimex + Cac_hang_khac + Khong_co_XML (v25)
CAC_DIR = "Cac_hang_khac"
KHONG_DIR = "Khong_co_XML"
TRUNG_DIR = "Trung_lap"      # PDF trùng ký hiệu + số với hóa đơn đã có XML (không đưa vào Tổng hợp, v25 không quét)


def _dup_row_for(stem, rows, cac_stems):
    """PDF chưa có XML (tên `stem`) có phải bản tải lặp của 1 hóa đơn đã có XML trong Cac_hang_khac không?
    Coi là trùng khi PDF của hóa đơn đó đang nằm trong Cac_hang_khac và tên file:
      - chỉ khác đuôi _1, _2... (tải 2 lần cùng tên), hoặc
      - chứa cả ký hiệu (vd C26MAQ) lẫn số hóa đơn (vd 554173) của hóa đơn đó.
    Trả về dòng XML của hóa đơn bị trùng, không trùng -> None."""
    base = re.sub(r"_\d+$", "", stem)
    up = stem.upper()
    for row in rows:
        tf = str(row.get("Ten file") or "").strip()
        if not tf or tf == stem or tf not in cac_stems:
            continue
        if base == tf:
            return row
        so = re.sub(r"\D", "", str(row.get("Số hóa đơn") or "")).lstrip("0")
        ky = str(row.get("Ký hiệu") or "").strip().upper()
        code = ky[1:] if len(ky) == 7 and ky[0].isdigit() else ky          # "1C26MAQ" -> "C26MAQ"
        if len(so) >= 3 and len(code) >= 5 and code in up and re.search(r"(?<!\d)0*%s(?!\d)" % re.escape(so), stem):
            return row
    return None
TONG_HOP = "Tong_hop_ket_qua.xlsx"

def _email_by_file(folder):
    """Tên file PDF -> thông tin email (lấy từ sổ theo dõi của profile có thư mục tải này)."""
    out = {}
    try:
        target = os.path.normcase(str(Path(folder).resolve()))
        for p in all_profiles():
            paths = profile_paths(p["id"])
            dl = paths.get("download_dir")
            if not dl or os.path.normcase(str(Path(dl).resolve())) != target or not paths["state"] or not Path(paths["state"]).exists():
                continue
            conn = sqlite3.connect("file:%s?mode=ro" % Path(paths["state"]).as_posix(), uri=True)
            try:
                for mid, sup, subj, md, files in conn.execute("SELECT message_id, supplier, subject, mail_date, files FROM emails"):
                    try:
                        names = json.loads(files or "[]")
                    except Exception:
                        names = []
                    for n in names if isinstance(names, list) else []:
                        out.setdefault(str(n), {"id": mid, "supplier": sup or "", "subject": subj or "", "date": md or ""})
            finally:
                conn.close()
    except Exception:
        pass
    return out


def _same_file(a, b):
    try:
        if a.stat().st_size != b.stat().st_size:
            return False
        import hashlib
        return hashlib.sha1(a.read_bytes()).digest() == hashlib.sha1(b.read_bytes()).digest()
    except Exception:
        return False


def _folder_key(folder):
    return os.path.normcase(str(Path(folder).resolve()))


def layout_since(folder):
    """Mốc thời gian bật bố cục thư mục mới của thư mục tải (None = chưa bật)."""
    v = (load_ui().get("layout_since") or {}).get(_folder_key(folder))
    return float(v) if isinstance(v, (int, float)) else None


def ensure_layout_marker(pid):
    """Bật bố cục mới cho thư mục tải của profile (chỉ lần đầu): PDF tải từ lúc này mới được chia thư mục."""
    try:
        dl = profile_paths(pid)["download_dir"]
        if not dl:
            return
        with _UI_LOCK:
            ui = load_ui()
            marks = ui.setdefault("layout_since", {})
            key = _folder_key(dl)
            if key not in marks:
                marks[key] = time.time() - 2
                save_ui(ui)
    except Exception:
        pass


def _kq_read_rows(path):
    """Đọc các dòng hóa đơn của 1 file dạng Ket_Qua (không có file -> [])."""
    path = Path(path)
    if not path.exists():
        return []
    got = _kq_open(path, create=False)
    if not got:
        return []
    wb, ws, hdr, cols = got
    out = []
    for r in range(hdr + 1, (ws.max_row or hdr) + 1):
        row = {name: ws.cell(r, cols[name]).value for name in KQ_COLUMNS}
        if any(v not in (None, "") for v in row.values()):
            out.append({k: ("" if v is None else v) for k, v in row.items()})
    return out


def _move_pdf(f, dest_dir):
    """Chuyển PDF vào thư mục con. Trùng tên: giống hệt -> bỏ bản thừa; khác nội dung -> thêm hậu tố."""
    dest_dir.mkdir(exist_ok=True)
    t = dest_dir / f.name
    if t.exists():
        if _same_file(f, t):
            f.unlink()
            return t
        k = 2
        while t.exists():
            t = dest_dir / ("%s_%d%s" % (f.stem, k, f.suffix))
            k += 1
    os.replace(str(f), str(t))
    return t


def _pid_for_folder(folder):
    key = _folder_key(folder)
    for p in all_profiles():
        dl = profile_paths(p["id"]).get("download_dir")
        if dl and _folder_key(dl) == key:
            return p["id"]
    return None


def build_tong_hop(pid, log):
    """Ghi Tong_hop_ket_qua.xlsx = Ket_Qua Petrolimex + Ket_Qua Cac_hang_khac + Ket_Qua Khong_co_XML (v25).
    Trùng hóa đơn (cùng ký hiệu + số): số liệu XML được ưu tiên hơn dòng quét PDF."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    paths = profile_paths(pid)
    folder = paths["download_dir"]
    if not folder or layout_since(folder) is None:
        return None
    kh_stems = {f.stem for f in (folder / KHONG_DIR).glob("*.pdf")} if (folder / KHONG_DIR).is_dir() else set()
    sources = [("Petrolimex", paths["petro_download_dir"] / "Ket_Qua.xlsx" if paths["petro_download_dir"] else None),
               ("Các hãng khác", folder / CAC_DIR / "Ket_Qua.xlsx"),
               ("Không có XML (v25)", folder / KHONG_DIR / "Ket_Qua.xlsx")]
    seen, out, counts = _KqIndex(), [], {}
    for group, src in sources:
        if not src:
            continue
        try:
            rows = _kq_read_rows(src)
        except Exception as e:
            log("[XML-HD] Không đọc được %s: %r (bỏ qua khi tổng hợp)" % (src, e), "err")
            continue
        for row in rows:
            if group.startswith("Không có XML") and row.get("Ten file") and str(row["Ten file"]) not in kh_stems:
                continue                                   # PDF này đã có XML và chuyển sang Cac_hang_khac
            if seen.find(row):
                continue                                   # trùng Tên đơn vị + Ký hiệu + Số với dòng đã lấy
            seen.add(row, True)
            out.append((group, row))
            counts[group] = counts.get(group, 0) + 1
    target = folder / TONG_HOP
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    heads = ["Nhóm"] + KQ_COLUMNS
    ws.append(heads)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, w in enumerate([18] + KQ_WIDTHS, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for group, row in out:
        ws.append([group] + [row.get(n, "") for n in KQ_COLUMNS])
        rr = ws.max_row
        for name in KQ_TEXT_COLS:
            ws.cell(rr, heads.index(name) + 1).number_format = "@"
        if isinstance(row.get("Cộng tiền hàng"), (int, float)):
            ws.cell(rr, heads.index("Cộng tiền hàng") + 1).number_format = "#,##0"
    ws.freeze_panes = "A2"
    if out:
        ws.auto_filter.ref = ws.dimensions
    try:
        wb.save(str(target))
    except PermissionError:
        log("[XML-HD] Không ghi được %s (file đang mở trong Excel?). Đóng file rồi bấm Ráp XML lại." % target, "err")
        return None
    log("[XML-HD] Tổng hợp kết quả: %d hóa đơn (%s) -> %s" % (
        len(out), ", ".join("%s %d" % (g, n) for g, n in counts.items()) or "chưa có dòng nào", target), "out")
    return {"rows": len(out), "counts": counts, "file": str(target)}


def organize_download(pid, log):
    """Chia PDF mới tải vào Cac_hang_khac (có XML) / Khong_co_XML (không có XML), cập nhật Ket_Qua của
    Cac_hang_khac rồi dựng lại file Tổng hợp. PDF tải trước khi bật bố cục mới để nguyên chỗ cũ."""
    folder = profile_paths(pid)["download_dir"]
    if not folder or not Path(folder).is_dir():
        return None
    folder = Path(folder)
    since = layout_since(folder)
    if since is None:
        return None
    rows = xml_store_rows(folder)
    stems = {str(r.get("Ten file") or "").strip() for r in rows} - {""}
    cac, kh = folder / CAC_DIR, folder / KHONG_DIR
    to_cac, to_kh = [], []
    for f in sorted(folder.iterdir()):
        if f.is_file() and f.suffix.lower() == ".pdf" and f.stat().st_mtime >= since:
            (to_cac if f.stem in stems else to_kh).append(f)
    later = [f for f in sorted(kh.glob("*.pdf")) if f.stem in stems] if kh.is_dir() else []
    for f in to_cac + later:
        _move_pdf(f, cac)
    for f in to_kh:
        _move_pdf(f, kh)
    kh.mkdir(exist_ok=True)
    cac.mkdir(exist_ok=True)
    cac_stems = {f.stem for f in cac.glob("*.pdf")}
    # PDF chưa có XML nhưng trùng ký hiệu + số với hóa đơn đã có XML (vd. cùng hóa đơn gửi trong 2 email)
    # -> chuyển vào Trung_lap: không đưa vào Tổng hợp, v25 không phải quét thừa. Xét cả PDF đã nằm sẵn trong Khong_co_XML.
    dups = []
    for f in sorted(kh.glob("*.pdf")):
        row = _dup_row_for(f.stem, rows, cac_stems)
        if row:
            _move_pdf(f, folder / TRUNG_DIR)
            dups.append("%s (trùng %s số %s)" % (f.name, row.get("Ký hiệu"), row.get("Số hóa đơn")))
    added = replaced = 0
    try:
        added, replaced = sync_xml_rows(cac / "Ket_Qua.xlsx", [r for r in rows if str(r.get("Ten file") or "") in cac_stems], None)
    except PermissionError:
        log("[XML-HD] Không ghi được %s (file đang mở trong Excel?)." % (cac / "Ket_Qua.xlsx"), "err")
    except Exception as e:
        log("[XML-HD] Lỗi khi ghi Ket_Qua.xlsx của %s: %r" % (CAC_DIR, e), "err")
    kh_all = sorted(f.name for f in kh.glob("*.pdf"))
    msg = "Chia thư mục: %d PDF có XML -> %s%s, %d PDF không có XML -> %s" % (
        len(to_cac) + len(later), CAC_DIR, (" (thêm %d hóa đơn vào Ket_Qua.xlsx)" % added) if added else "", len(to_kh), KHONG_DIR)
    if dups:
        msg += ", %d PDF trùng hóa đơn đã có -> %s: %s" % (len(dups), TRUNG_DIR, "; ".join(dups[:10]) + (" …" if len(dups) > 10 else ""))
    if kh_all:
        names = ", ".join(kh_all[:15]) + (" … (+%d)" % (len(kh_all) - 15) if len(kh_all) > 15 else "")
        log("[XML-HD] %s. Thư mục %s đang có %d PDF CHƯA CÓ XML: %s. Bấm 'Quét v25 PDF không có XML' để lấy số liệu." % (
            msg, KHONG_DIR, len(kh_all), names), "err")
    else:
        log("[XML-HD] %s. Không còn PDF nào thiếu XML." % msg, "out")
    th = build_tong_hop(pid, log)
    return {"to_cac": len(to_cac) + len(later), "to_kh": len(to_kh), "kh_pdf": kh_all, "dups": dups, "tong_hop": th}


def merge_invoice_xml(folder, log, kq_name="Ket_Qua.xlsx"):
    """Ráp XML trong <folder>/XML vào kho + Ket_Qua.xlsx của <folder>, rồi xóa file XML đã ráp."""
    res = {"found": 0, "parsed": 0, "failed": 0, "deleted": 0, "added": 0, "replaced": 0, "file": ""}
    if not folder:
        return res
    folder = Path(folder)
    xml_dir = folder / XML_SUBDIR
    xmls = sorted(f for f in xml_dir.iterdir() if f.is_file() and f.suffix.lower() == ".xml") if xml_dir.is_dir() else []
    res["found"] = len(xmls)
    store = xml_dir / XML_STORE
    if not xmls and not store.exists():
        return res
    try:
        import openpyxl  # noqa: F401
    except Exception:
        log("[XML-HD] Chưa có thư viện openpyxl nên chưa ráp được XML (cài: python -m pip install openpyxl).", "err")
        return res
    kq = folder / kq_name
    res["file"] = str(kq)
    new_layout = layout_since(folder) is not None
    parsed = []
    if xmls:
        pdf_stems = [f.stem for d in (folder, folder / CAC_DIR, folder / KHONG_DIR) if d.is_dir()
                     for f in d.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"]
        for f in xmls:
            try:
                row = parse_invoice_xml(f)
                row["Ten file"] = _pdf_stem_for(row, pdf_stems, f.stem)
                parsed.append((f, row))
            except Exception as e:
                res["failed"] += 1
                log("[XML-HD] Không đọc được %s: %s (giữ lại file XML)" % (f.name, e), "err")
        res["parsed"] = len(parsed)
        if parsed:
            # 1) lưu vào kho trước -> an toàn để xóa XML
            try:
                wb, ws, hdr, cols = _kq_open(store)
                have, last = _KqIndex(), hdr
                for r in range(hdr + 1, (ws.max_row or hdr) + 1):
                    old = _kq_row_at(ws, r, cols)
                    if any(v not in (None, "") for v in old.values()):
                        last = r
                    have.add(old, r)
                for f, row in parsed:
                    r = have.find(row)
                    if r:
                        _kq_fill_blanks(ws, r, cols, row)
                        continue
                    last += 1
                    have.add(row, last)
                    _kq_write_row(ws, last, cols, row)
                wb.save(store)
            except Exception as e:
                log("[XML-HD] Không ghi được kho %s: %r. File XML vẫn được giữ lại." % (store, e), "err")
                return res
            for f, _ in parsed:
                try:
                    f.unlink()
                    res["deleted"] += 1
                except Exception as e:
                    log("[XML-HD] Không xóa được %s: %r" % (f.name, e), "err")
    if new_layout:                                        # bố cục mới: Ket_Qua nằm trong Cac_hang_khac (organize_download)
        if parsed:
            log("[XML-HD] Đã ráp %d file XML hóa đơn thường (đã xóa %d file XML, PDF giữ nguyên)." % (len(parsed), res["deleted"]), "out")
        return res
    # 2) ghi toàn bộ hóa đơn trong kho vào Ket_Qua.xlsx (thêm mới / thay dòng quét PDF)
    rows = xml_store_rows(folder)
    try:
        res["added"], res["replaced"] = sync_xml_rows(kq, rows, log)
    except PermissionError:
        log("[XML-HD] Không ghi được %s (file đang mở trong Excel?). Số liệu XML vẫn giữ trong %s, lần chạy sau sẽ tự ghi vào." % (kq, store), "err")
        return res
    except Exception as e:
        log("[XML-HD] Lỗi khi ghi Ket_Qua.xlsx: %r. Số liệu XML vẫn giữ trong %s." % (e, store), "err")
        return res
    if parsed:
        log("[XML-HD] Đã ráp %d file XML hóa đơn thường vào %s: thêm %d hóa đơn%s, đã xóa %d file XML (PDF giữ nguyên)."
            % (len(parsed), kq, res["added"], (", thay %d dòng quét PDF bằng số liệu XML" % res["replaced"]) if res["replaced"] else "", res["deleted"]), "out")
    elif res["added"] or res["replaced"]:
        log("[XML-HD] Đã ghi lại %d hóa đơn từ XML (đã ráp trước đó) vào %s%s."
            % (res["added"], kq, (", thay %d dòng quét PDF" % res["replaced"]) if res["replaced"] else ""), "out")
    return res


# ------------------------------------------------------------------ Petro: nhập mã tra cứu từ file
CODE_RE = re.compile(r"^[A-Z0-9]{5,24}$")


def load_email_state():
    path = tool_dir() / "email_state.py"
    if not path.exists():
        raise ValueError("Không thấy email_state.py cạnh main.py")
    spec = importlib.util.spec_from_file_location("email_state_for_app", str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def import_petro_codes(pid, items, source):
    es = load_email_state()
    paths = profile_paths(pid)
    state = es.EmailState(str(paths["state"]))
    added, existed, invalid = [], [], []
    try:
        known = {}
        for r in state.petro_rows(include_duplicates=True):
            known.setdefault(es.norm_code(r.get("ma_tra_cuu")), set()).add(es.norm_so(r.get("so_hoa_don")))
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        subject = ("Nhập từ " + str(source or "danh sách"))[:200]
        for it in items[:2000]:
            code = re.sub(r"\s", "", str(it.get("ma") or "")).upper().rstrip("*")
            so = re.sub(r"\D", "", str(it.get("so") or ""))
            if not CODE_RE.match(code):
                invalid.append(code)
                continue
            so_n = so.lstrip("0")
            if code in known and any(es.same_so(so_n, k) for k in known[code]):
                existed.append(code)
                continue
            known.setdefault(code, set()).add(so_n)
            state.save_petro("nhap-tay-" + code + ("-" + so_n if so_n else ""), now, "", "Nhập tay", subject,
                             so, code + "*", commit=False)
            added.append(code)
        state.commit()
    finally:
        state.close()
    return {"added": added, "existed": existed, "invalid": invalid}


# ------------------------------------------------------------------ HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "HoaDonApp/" + APP_VERSION

    def log_message(self, fmt, *args):  # yên lặng, chỉ in lỗi
        pass

    def _host_ok(self):
        host = (self.headers.get("Host") or "").split(":")[0].lower()
        return host in ("127.0.0.1", "localhost")

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _auth(self):
        return self._host_ok() and secrets.compare_digest(self.headers.get("X-Local-Token", ""), Settings.token)

    def _json_body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 1_000_000:
            raise ValueError("Dữ liệu quá lớn")
        raw = self.rfile.read(n) if n else b"{}"
        return json.loads(raw.decode("utf-8") or "{}")

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, {"error": "Chỉ mở bằng 127.0.0.1"})
        u = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(u.query)
        q = lambda k, d="": (qs.get(k) or [d])[0]
        if u.path in ("/", "/index.html"):
            return self.serve_html()
        if u.path == "/api/ping":
            return self._send(200, {"app": "hoadon", "version": APP_VERSION})
        if not u.path.startswith("/api/"):
            return self._send(404, {"error": "Không có trang này"})
        if not self._auth():
            return self._send(403, {"error": "Phiên làm việc đã hết hạn. Hãy tải lại trang."})
        try:
            if u.path == "/api/info":
                return self._send(200, self.info())
            if u.path == "/api/ledger":
                pid = q("profile")
                if not PROFILE_RE.match(pid):
                    raise ValueError("Tên profile không hợp lệ")
                return self._send(200, {"profile": pid, "rows": ledger_rows(pid, clamp_int(q("limit", "5000"), 1, 100000, 5000))})
            if u.path == "/api/petro":
                pid = q("profile")
                if not PROFILE_RE.match(pid):
                    raise ValueError("Tên profile không hợp lệ")
                return self._send(200, {"profile": pid, "rows": petro_rows(pid)})
            if u.path == "/api/job":
                job = JOBS.get(q("id")) if q("id") else CURRENT["job"]
                if not job:
                    return self._send(200, {"job": None, "lines": [], "next": 0})
                start = clamp_int(q("from", "0"), 0, 10**9, 0)
                with job.lock:
                    lines = job.lines[start:]
                    nxt = len(job.lines)
                return self._send(200, {"job": job.summary(), "lines": lines, "next": nxt})
            if u.path == "/api/error-report":
                pid = q("profile")
                if not PROFILE_RE.match(pid or ""):
                    raise ValueError("Tên profile không hợp lệ")
                f = latest_error_report(pid)
                if not f:
                    return self._send(200, {"file": "", "items": []})
                return self._send(200, {"file": f.name, "path": str(f), "mtime": f.stat().st_mtime,
                                        "items": error_report_items(read_xlsx_rows(f))})
            if u.path == "/api/v25-result":
                job = JOBS.get(q("id"))
                if not job or job.action != "scan_v25" or not job.result or not Path(job.result).is_file():
                    return self._send(404, {"error": "Chưa có file kết quả"})
                data = Path(job.result).read_bytes()
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if u.path == "/api/jobs":
                return self._send(200, {"jobs": [JOBS[j].summary() for j in reversed(JOB_ORDER) if j in JOBS]})
            return self._send(404, {"error": "Không có API này"})
        except Exception as e:
            return self._send(400, {"error": str(e)})

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        if not self._auth():
            return self._send(403, {"error": "Phiên làm việc đã hết hạn. Hãy tải lại trang."})
        try:
            body = self._json_body()
            if u.path == "/api/run":
                job = start_job(str(body.get("profile") or ""), str(body.get("action") or ""), body.get("options") or {})
                with _UI_LOCK:
                    ui = load_ui()
                    ui["last_profile"] = job.profile
                    save_ui(ui)
                return self._send(200, {"job": job.summary()})
            if u.path == "/api/stop":
                job = JOBS.get(str(body.get("id") or "")) or CURRENT["job"]
                if job and job.status == "running":
                    job.stop_requested = True
                    job.add("Đang dừng…", kind="info")
                    kill_tree(job.proc)
                return self._send(200, {"ok": True})
            if u.path == "/api/profiles/add":
                pid = str(body.get("id") or "").strip()
                if not PROFILE_RE.match(pid):
                    raise ValueError("Tên profile chỉ gồm chữ không dấu, số, dấu gạch dưới, gạch ngang hoặc dấu chấm (tối đa 40 ký tự)")
                email = str(body.get("email") or "").strip()[:120]
                ddir = str(body.get("download_dir") or "").strip()[:500]
                ui = load_ui()
                existing = [p for p in ui["profiles"] if str(p.get("id", "")).lower() == pid.lower()]
                if existing:
                    existing[0].update({"email": email or existing[0].get("email", ""), "download_dir": ddir})
                else:
                    ui["profiles"].append({"id": pid, "email": email, "download_dir": ddir})
                ui["last_profile"] = pid
                save_ui(ui)
                return self._send(200, self.info())
            if u.path == "/api/profiles/remove":
                pid = str(body.get("id") or "")
                ui = load_ui()
                ui["profiles"] = [p for p in ui["profiles"] if str(p.get("id", "")).lower() != pid.lower()]
                save_ui(ui)
                return self._send(200, self.info())
            if u.path in ("/api/petro/import", "/api/petro/merge-xml"):
                pid = str(body.get("profile") or "")
                if not PROFILE_RE.match(pid):
                    raise ValueError("Tên profile không hợp lệ")
                cur = CURRENT["job"]
                if cur is not None and cur.status == "running":
                    raise RuntimeError("Đang chạy \"%s\". Chờ chạy xong rồi thử lại." % cur.title)
                if u.path == "/api/petro/import":
                    items = body.get("items") or []
                    if not isinstance(items, list) or not items:
                        raise ValueError("Danh sách mã trống")
                    return self._send(200, import_petro_codes(pid, items, str(body.get("source") or "")[:120]))
                msgs = []
                res = merge_petro_xml(pid, lambda text, kind="out": msgs.append({"text": text, "kind": kind}))
                res["messages"] = msgs
                return self._send(200, res)
            if u.path == "/api/error-report/read":
                import base64, io
                raw = base64.b64decode(str(body.get("data") or ""), validate=False)
                if not raw or len(raw) > 700_000:
                    raise ValueError("File trống hoặc quá lớn")
                try:
                    rows = read_xlsx_rows(io.BytesIO(raw))
                except Exception:
                    raise ValueError("Không đọc được file. Hãy chọn đúng file loi_can_xu_ly_*.xlsx")
                items = error_report_items(rows)
                return self._send(200, {"file": str(body.get("name") or "")[:200], "items": items})
            if u.path == "/api/xml/merge":
                pid = str(body.get("profile") or "")
                if not PROFILE_RE.match(pid):
                    raise ValueError("Tên profile không hợp lệ")
                cur = CURRENT["job"]
                if cur is not None and cur.status == "running":
                    raise RuntimeError("Đang chạy \"%s\". Chờ chạy xong rồi thử lại." % cur.title)
                msgs = []
                res = merge_invoice_xml(profile_paths(pid)["download_dir"], lambda text, kind="out": msgs.append({"text": text, "kind": kind}))
                organize_download(pid, lambda text, kind="out": msgs.append({"text": text, "kind": kind}))
                res["messages"] = msgs
                return self._send(200, res)
            if u.path == "/api/v25/config":
                path = str(body.get("path") or "").strip().strip('"')
                ui = load_ui()
                if path:
                    pth = Path(path)
                    if pth.is_dir():
                        cands = [pth / n for n in V25_NAMES if (pth / n).is_file()]
                        if not cands:
                            raise ValueError("Trong thư mục này không có file v25.py / scan_pdf_v26.py")
                        pth = cands[0]
                    if not pth.is_file() or pth.suffix.lower() != ".py":
                        raise ValueError("Không thấy file .py: %s" % path)
                    ui["v25_path"] = str(pth)
                else:
                    ui.pop("v25_path", None)
                save_ui(ui)
                return self._send(200, self.info())
            if u.path == "/api/prefs":
                ui = load_ui()
                pid = str(body.get("last_profile") or "")
                if PROFILE_RE.match(pid):
                    ui["last_profile"] = pid
                    save_ui(ui)
                return self._send(200, {"ok": True})
            if u.path == "/api/open":
                pid = str(body.get("profile") or "")
                kind = str(body.get("kind") or "")
                if kind == "tool":
                    target = tool_dir()
                elif kind == "job_result":
                    job = JOBS.get(str(body.get("id") or ""))
                    if not job or not job.result or not Path(job.result).is_file():
                        raise ValueError("Chưa có file kết quả")
                    if os.name == "nt":
                        os.startfile(job.result)  # noqa
                    else:
                        subprocess.Popen(["xdg-open", job.result], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    return self._send(200, {"ok": True, "path": job.result})
                else:
                    if not PROFILE_RE.match(pid):
                        raise ValueError("Tên profile không hợp lệ")
                    paths = profile_paths(pid)
                    if kind == "ket_qua_dl":
                        kq = paths["download_dir"] / "Ket_Qua.xlsx" if paths["download_dir"] else None
                        if not kq or not kq.exists():
                            raise ValueError("Chưa có Ket_Qua.xlsx trong thư mục tải về")
                        if os.name == "nt":
                            os.startfile(str(kq))  # noqa
                        else:
                            subprocess.Popen(["xdg-open" if sys.platform != "darwin" else "open", str(kq)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return self._send(200, {"ok": True, "path": str(kq)})
                    if kind == "ket_qua":
                        kq = paths["petro_download_dir"] / "Ket_Qua.xlsx" if paths["petro_download_dir"] else None
                        if not kq or not kq.exists():
                            raise ValueError("Chưa có Ket_Qua.xlsx trong thư mục PDF Petro")
                        if os.name == "nt":
                            os.startfile(str(kq))  # noqa
                        else:
                            subprocess.Popen(["xdg-open" if sys.platform != "darwin" else "open", str(kq)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return self._send(200, {"ok": True, "path": str(kq)})
                    target = {"downloads": paths["download_dir"], "petro": paths["petro_download_dir"],
                              "outputs": paths["error_dir"], "logs": paths["log_dir"],
                              "state": paths["state"].parent if paths["state"] else None,
                              "xml": (paths["download_dir"] / "XML") if paths["download_dir"] else None,
                              "cac_hang_khac": (paths["download_dir"] / CAC_DIR) if paths["download_dir"] else None,
                              "khong_co_xml": (paths["download_dir"] / KHONG_DIR) if paths["download_dir"] else None,
                              "trung_lap": (paths["download_dir"] / TRUNG_DIR) if paths["download_dir"] else None}.get(kind)
                    if kind == "tong_hop":
                        lf = paths["download_dir"] / TONG_HOP if paths["download_dir"] else None
                        if not lf or not lf.exists():
                            raise ValueError("Chưa có file Tổng hợp kết quả (tạo sau lần tải / ráp XML đầu tiên theo bố cục mới)")
                        if os.name == "nt":
                            os.startfile(str(lf))  # noqa
                        else:
                            subprocess.Popen(["xdg-open" if sys.platform != "darwin" else "open", str(lf)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return self._send(200, {"ok": True, "path": str(lf)})
                if not target:
                    raise ValueError("Không biết mở thư mục nào")
                open_folder(target)
                return self._send(200, {"ok": True, "path": str(target)})
            return self._send(404, {"error": "Không có API này"})
        except (ValueError, RuntimeError) as e:
            return self._send(400, {"error": str(e)})
        except Exception as e:
            return self._send(500, {"error": "Lỗi: %r" % e})

    def info(self):
        ui = load_ui()
        cur = CURRENT["job"]
        return {
            "app": {"version": APP_VERSION, "folder": str(tool_dir()), "main_found": Settings.main_py.exists(),
                    "main_path": str(Settings.main_py), "python": Settings.python,
                    "credentials_found": (tool_dir() / "credentials.json").exists(),
                    "features": main_features(), "ddddocr": DDDDOCR["ok"], "windows": os.name == "nt",
                    "v25": str(find_v25() or ""), "cpu": os.cpu_count() or 1},
            "profiles": [profile_status(p) for p in all_profiles()],
            "last_profile": ui.get("last_profile") or DEFAULT_PROFILES[0],
            "job": cur.summary() if cur else None,
        }

    def serve_html(self):
        path = APP_DIR / HTML_NAME
        if not path.exists():
            return self._send(404, "<meta charset=utf-8><p>Không thấy file %s cạnh app.py.</p>" % HTML_NAME, "text/html; charset=utf-8")
        html = path.read_text(encoding="utf-8")
        inject = "<script>window.LOCAL_API=%s;</script>" % json.dumps({"token": Settings.token, "version": APP_VERSION})
        marker = '<div id="root"></div>'
        html = html.replace(marker, marker + "\n" + inject, 1) if marker in html else inject + html
        return self._send(200, html, "text/html; charset=utf-8",
                          {"Content-Security-Policy": "frame-ancestors 'none'", "Referrer-Policy": "no-referrer"})


def already_running(port):
    try:
        with urllib.request.urlopen("http://127.0.0.1:%d/api/ping" % port, timeout=1.5) as r:
            return json.loads(r.read().decode("utf-8")).get("app") == "hoadon"
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser(description="Giao diện Bộ công cụ hóa đơn chạy trên máy của bạn")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--main", help="Đường dẫn main.py (mặc định: cùng thư mục với app.py)")
    ap.add_argument("--python", help="Python dùng để chạy main.py (mặc định: Python đang chạy app.py)")
    args = ap.parse_args()
    if args.main:
        Settings.main_py = Path(args.main).expanduser().resolve()
    if args.python:
        Settings.python = args.python

    port = args.port
    if already_running(port):
        print("Giao diện đang chạy sẵn, mở lại trình duyệt: http://127.0.0.1:%d" % port)
        if not args.no_browser:
            webbrowser.open("http://127.0.0.1:%d/" % port)
        return
    server = None
    for p in range(port, port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            port = p
            break
        except OSError:
            continue
    if server is None:
        print("Không mở được cổng nào từ %d đến %d." % (args.port, args.port + 19))
        sys.exit(1)
    server.daemon_threads = True
    threading.Thread(target=lambda: DDDDOCR.update(ok=ddddocr_installed()), daemon=True).start()
    Settings.port = port
    url = "http://127.0.0.1:%d/" % port
    print("=" * 60)
    print(" Bộ công cụ hóa đơn  -  phiên bản %s" % APP_VERSION)
    print(" Mở trình duyệt: %s" % url)
    print(" main.py: %s%s" % (Settings.main_py, "" if Settings.main_py.exists() else "   (CHƯA THẤY!)"))
    print(" Giữ cửa sổ này mở trong lúc dùng. Nhấn Ctrl+C để tắt.")
    print("=" * 60)
    if not (APP_DIR / HTML_NAME).exists():
        print("[CẢNH BÁO] Không thấy %s cạnh app.py." % HTML_NAME)
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        cur = CURRENT["job"]
        if cur and cur.status == "running":
            cur.stop_requested = True
            kill_tree(cur.proc)
        server.server_close()
        print("Đã tắt.")


if __name__ == "__main__":
    main()
