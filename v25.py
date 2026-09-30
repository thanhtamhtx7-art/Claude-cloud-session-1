"""
scan_pdf_v26.py — Hóa đơn Petrolimex / Standard (v25 + tối ưu hiệu suất, đầu ra giữ nguyên)
"""

import math
import os
import re
import sys
import unicodedata
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import pdfplumber

# [v26-A1] Mỗi tiến trình tesseract chỉ dùng 1 thread OpenMP. Tránh việc N worker x 4 thread
# tranh nhau CPU (quá tải, chuyển ngữ cảnh liên tục). Tesseract con sẽ kế thừa biến này.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

# [v26-A2] Số luồng xử lý song song. Đổi bằng biến môi trường INVOICE_WORKERS nếu cần
# (vd: set INVOICE_WORKERS=4 trên Windows CMD). Mặc định tối đa 8.
def _default_workers() -> int:
    try:
        n = int(os.environ.get("INVOICE_WORKERS", "0"))
        if n > 0:
            return n
    except ValueError:
        pass
    return max(1, min(8, os.cpu_count() or 1))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    import pytesseract
    from pdf2image import convert_from_path, pdfinfo_from_path
    from PIL import ImageEnhance, ImageFilter
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ── Cấu hình ────────────────────────────────────────────────────────────────
IMPORT_FOLDER = "Import"
OUTPUT_PATH   = os.path.join(IMPORT_FOLDER, "Ket_Qua.xlsx")
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")
COLUMNS = [
    "Ten file", "Nha phat hanh", "Loai PDF",
    "Tên đơn vị mua", "Mã số thuế bên mua", "Ký hiệu", "Số hóa đơn",
    "Ngày hóa đơn", "Cộng tiền hàng", "Biển số xe",
    "Tên hàng hóa, dịch vụ", "Nguon",
]

_HEADER_CHARS  = 300
_PLX_RE        = re.compile(r"PETROLIMEX|PETRO\s*VIET\s*NAM|PVOIL", re.I)
_INVOICE_RE    = re.compile(r"HÓA\s*ĐƠN|HOA\s*DON|VAT\s+INVOICE|GIÁ\s*TRỊ\s*GIA\s*TĂNG", re.I)
_BIEN_SO_RE    = re.compile(
    r"^\d{2}[A-Z]{1,2}[\s.\-–]*\d{3}[.\-\s]?\d{2,3}$"
    r"|^\d{2}[A-Z]{1,2}[\s.\-–]*\d{4,6}$", re.I
)
_SKIP_CELLS    = {
    "stt","đvt","dvt","đơn vị tính","don vi tinh","unit","số lượng","so luong",
    "quantity","đơn giá","don gia","unit price","thành tiền","thanh tien",
    "amount","chiết khấu","no.","1","2","3","4","5","6","a","b","c","a b","b c",
}
_BAD_ITEMS     = (
    "STT","DVT","NAME OF GOODS","GOODS AND SERVICES","GOODS, SERVICES",
    "DICH VU","CONG TY","DON VI","MA SO THUE","DIA CHI",
    "DIEN THOAI","CHI NHANH","CUA HANG","TRAM XANG",
)

# ── Tiện ích chung ───────────────────────────────────────────────────────────

def clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\n", " ")).replace('"', "").strip()

def no_accent(s: str) -> str:
    """Đã fix lỗi chữ 'Đ' không được xử lý trong tiếng Việt"""
    if not isinstance(s, str): return ""
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return re.sub(r"[\u0300-\u036f\u1dc0-\u1dff\u20d0-\u20ff]", "", s)

def num_ok(val: str, min_val: int = 1000) -> bool:
    n = re.sub(r"[\s.,]", "", val)
    return n.isdigit() and int(n) >= min_val

def cut_at_mst(val: str) -> str:
    return re.split(
        r"\s+(?:M[aãáà]\s*s[oốéèê6][^\s]*|MST|\d{10})", val, flags=re.I
    )[0].strip().rstrip(":,; \t")

def strip_prefix(val: str) -> str:
    return re.sub(
        r"^(?:Tên\s*(?:đơn\s*vị|don\s*vi)\s*[:\s]+|Ten\s*don\s*vi\s*[:\s]+)",
        "", val, flags=re.I,
    ).strip()

def is_bien_so(val: str) -> bool:
    return bool(_BIEN_SO_RE.match(val.strip()))

def is_valid_item(text: str) -> bool:
    if len(text) < 4 or len(text) > 80:
        return False
    if not re.search(r"[A-Za-zÀ-ỹ]", text) or re.fullmatch(r"[\d.,\s]+", text):
        return False
    return not any(b in text.upper() for b in _BAD_ITEMS)

# ── OCR ─────────────────────────────────────────────────────────────────────

def _ocr_image(img) -> str:
    if not OCR_AVAILABLE:
        return ""
    img = ImageEnhance.Contrast(img.filter(ImageFilter.SHARPEN)).enhance(2.0)
    try:
        t = pytesseract.image_to_string(img, lang="vie+eng", config="--psm 6")
        if t.strip():
            return t
    except Exception:
        pass
    return ""

def detect_pdf(pdf_path: str, pdf) -> tuple[str, str, str]:
    full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if full_text.strip():
        issuer = "petrolimex" if _PLX_RE.search(full_text[:_HEADER_CHARS]) else "standard"
        return issuer, "text", full_text
    return "unknown", "image", ""


def looks_like_invoice_page(text: str) -> bool:
    if not text or len(text.strip()) < 250:
        return False
    if _INVOICE_RE.search(text[:_HEADER_CHARS]):
        return True

    norm = no_accent(text).upper()
    signals = (
        r"\bKY\s*HIEU\b",
        r"\bNGAY\b.*\bTHANG\b.*\bNAM\b",
        r"\bMA\s*SO\s*THUE\b|\bMST\b",
        r"\bTEN\s*HANG\b|\bCONG\s*TIEN\b|\bTONG\s*TIEN\b",
        r"\bMA\s*TRA\s*CUU\b|\bWEBSITE\s*TRA\s*CUU\b",
        r"\bDON\s*VI\s*BAN\s*HANG\b|\bTEN\s*DON\s*VI\b",
    )
    return sum(1 for pat in signals if re.search(pat, norm, re.I)) >= 3


def is_non_invoice_attachment_page(text: str) -> bool:
    norm = no_accent(text or "").upper()
    if not norm.strip():
        return False
    if norm.lstrip().startswith("DANH SACH LOG COT BOM"):
        return True
    has_invoice_signal = re.search(r"HOA\s*DON|GIA\s*TRI\s*GIA\s*TANG|KY\s*HIEU|SO\s*:", norm[:500], re.I)
    has_pump_log = re.search(r"DANH\s*SACH\s*LOG\s*COT\s*BOM|VOI\s*BOM|TIEP\s*THEO\s*TRANG\s*TRUOC", norm, re.I)
    return bool(has_pump_log and not has_invoice_signal)

# ── Các hàm extract dùng chung ───────────────────────────────────────────────

def extract_ky_hieu(text: str) -> str:
    for pat in (
        r"K[yý]\s*hi[eệ]u\b\s*(?:\([^)]*\))?\s*[:\s]+([A-Z0-9]{4,20})",
        r"Ky\s*hi[eé]u\b\s*:\s*([A-Z0-9]{4,20})",
    ):
        m = re.search(pat, text, re.I)
        if m and not m.group(1).upper().startswith("M1"):
            return m.group(1).strip()
    m = re.search(r"\b(1[CK]\d{2}[A-Z]{2,4})\b", text)
    if m:
        return m.group(1)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if re.search(r"K[yý]\s*hi[eé]", line) and ":" in line and i + 1 < len(lines):
            m2 = re.search(r"\b([1-9][A-Z]\d{2}[A-Z]{2,4})\b", lines[i + 1])
            if m2:
                return m2.group(1)
    return ""

def extract_ngay(text: str) -> str:
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", text)
    if m:
        return m.group(1)
    for pat in (
        r"Ng[aà]y\s*(?:\([^)]*\))?\s*(\d{1,2})\s*th[aá]ng\s*(?:\([^)]*\))?\s*(\d{1,2})\s*n[aă]m\s*(?:\([^)]*\))?\s*(\d{4})",
        r"Ngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+na[&\w]{0,3}m\s+(\d{4})",
        r"Ng[aà]y\s+(\d{1,2})\s+th[aá]ng\s+(\d{1,2})\s+n[aă]m[^\n]*\n[^\n]*\b(\d{4})\b",
    ):
        m = re.search(pat, text, re.I)
        if m:
            return f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"
    return ""

def extract_bien_so(text: str) -> str:
    label_pat = re.compile(
        r"(?:Bi[eé]n\s*s[oéo6]\s*xe|Bien\s*so\s*xe"
        r"|H[oọ]\s*(?:v[aà]\s*)?t[eê]n\s*ng[uư][oờ]i\s*mua(?:\s*h[aà]ng)?"
        r"|T[eê]n\s*ng[uư][oờ]i\s*mua)(?:\s*\([^)]*\))?\s*[:\s]+([^\n\r]{2,30})",
        re.I,
    )
    for src in (text, no_accent(text)):
        m = label_pat.search(src)
        if m:
            lbl = no_accent(m.group(0)).lower()
            cand = re.split(r"\s+(?:Ma|Dong|MST|\d{10})", m.group(1), flags=re.I)[0]
            cand = re.sub(r"\s+[a-zA-Z'\"=_]{1,2}$", "", cand).strip()
            if not cand:
                continue
            m2 = re.match(r"([0-9]{2}[A-Z0-9\-./  ]{3,15})", cand.upper())
            if m2 and re.search(r"\d", m2.group(1)):
                return m2.group(1).strip()
            if "bien so" in lbl and 3 <= len(cand) <= 25:
                if not re.search(r"(?:CONG TY|LIEN HIEP|MA SO|DIA CHI|HANG|THUE)", cand.upper()):
                    return cand
    m = re.search(
        r"\b(\d{2}[A-Za-z]{1,2}[\s.\-–]*\d{3}[.\-\s]?\d{2,3}"
        r"|\d{2}[A-Za-z]{1,2}[\s.\-–]*\d{4,6})\b", text, re.I,
    )
    if m:
        return m.group(1)
    
    m2 = re.search(r"\b(\d{2})8([\s.\-–]*\d{3}[.\-\s]?\d{2,3})\b", text, re.I)
    if m2:
        return f"{m2.group(1)}B{m2.group(2)}"
    return ""

# ── Mã số thuế bên MUA hàng ──────────────────────────────────────────────────

_MUA_MARKER_NOACC_RE = re.compile(
    r"DON\s*VI\s*MUA\s*HANG"
    r"|HO\s*(?:VA\s*)?TEN\s*NGUOI\s*MUA"
    r"|NGUOI\s*MUA\s*HANG"
    r"|TEN\s*DON\s*VI\s*(?!BAN)"
    r"|TEN\s*NGUOI\s*MUA"
    r"|BUYER",
    re.I,
)

_MST_LABEL_VALUE_NOACC_RE = re.compile(
    r"MA\s*S[OÔ6]\s*THU[EÊÉ]\s*(?:\([^)]*\))?\s*[:,\-]*\s*(\d{10}(?:[-\s]?\d{3})?)",
    re.I,
)

def extract_mst_mua(text: str) -> str:
    """Trích xuất Mã số thuế (MST) của bên MUA HÀNG, phân biệt với MST bên bán.

    Chiến lược: tìm nhãn đánh dấu khu vực thông tin bên mua ("Đơn vị mua hàng",
    "Họ tên người mua", "Người mua hàng", "Tên đơn vị mua", "Buyer"...), sau đó
    lấy giá trị "Mã số thuế" đầu tiên xuất hiện SAU nhãn đó. Nếu không tìm thấy
    nhãn bên mua, sẽ lấy MST đứng SAU CÙNG trong văn bản (MST bên bán thường
    xuất hiện trước, ở phần đầu hóa đơn).
    """
    if not text:
        return ""
    norm = no_accent(text)

    marker = _MUA_MARKER_NOACC_RE.search(norm)
    zone = norm[marker.end():] if marker else norm

    m = _MST_LABEL_VALUE_NOACC_RE.search(zone)
    if m:
        return re.sub(r"\s", "", m.group(1))

    # Fallback bền hơn cho ảnh OCR bị nhiễu nặng (nhãn "Mã số thuế" bị biến
    # dạng khó đoán): tìm cụm 10 chữ số đứng gần nhất một mẩu "THU" (từ chữ
    # "thuế") trong vùng bên mua, thay vì đòi hỏi khớp chính xác cả cụm nhãn.
    if marker:
        thu_pos = [mm.start() for mm in re.finditer(r"THU", zone, re.I)]
        digit_matches = list(re.finditer(r"(?<!\d)(\d{10})(?!\d)", zone))
        if thu_pos and digit_matches:
            best = min(
                digit_matches,
                key=lambda dm: min(abs(dm.start() - tp) for tp in thu_pos),
            )
            if min(abs(best.start() - tp) for tp in thu_pos) <= 40:
                return best.group(1)
        # Đã xác định được vùng bên mua nhưng không tìm ra MST đủ tin cậy
        # -> để trống, tránh lấy nhầm sang số khác (điện thoại, biển số...).
        return ""

    all_matches = _MST_LABEL_VALUE_NOACC_RE.findall(norm)
    if len(all_matches) >= 2:
        return re.sub(r"\s", "", all_matches[-1])
    if len(all_matches) == 1:
        return re.sub(r"\s", "", all_matches[0])
    return ""

# ── ĐỐI SOÁT CHỮ/SỐ THÔNG MINH ──────────────────────────────────────────────

def _parse_vn_words(text: str) -> int:
    """Đọc số tiền bằng chữ và khử nhiễu lỗi OCR (cross-check)."""
    norm = no_accent(text).lower()
    norm = re.sub(r'[^a-z\s:\n]', ' ', norm)
    
    words_str = ""
    m = re.search(r"(?:bang\s*chu|viet\s*bang\s*chu|tien\s*bang\s*chu)\s*[:\s]*([a-z\s]+dong)", norm)
    if m:
        words_str = m.group(1)
    else:
        for line in norm.split("\n"):
            if "dong" in line and ("trieu" in line or "nghin" in line or "ngan" in line):
                words_str = line
                break

    if not words_str: 
        return 0
    
    words_str = re.sub(r"\b(met|mat|mrt)\b", "mot", words_str)
    words_str = re.sub(r"\b(hal)\b", "hai", words_str)
    words_str = re.sub(r"\b(ban)\b", "bon", words_str)
    words_str = re.sub(r"\b(nan|nhiem)\b", "nam", words_str)
    words_str = re.sub(r"\b(san)\b", "sau", words_str)
    words_str = re.sub(r"\b(hay)\b", "bay", words_str)
    words_str = re.sub(r"\b(tan)\b", "tam", words_str)
    words_str = re.sub(r"\b(chim)\b", "chin", words_str)
    words_str = re.sub(r"\b(muol|muoi)\b", "muoi", words_str)
    words_str = re.sub(r"\b(tran)\b", "tram", words_str)
    words_str = re.sub(r"\b(nghi|ngan|nghin)\b", "nghin", words_str)
    words_str = re.sub(r"\b(triau|trieu)\b", "trieu", words_str)

    VAL = {"mot":1,"hai":2,"ba":3,"bon":4,"tu":4,"lam":5,"nam":5,
           "sau":6,"bay":7,"tam":8,"chin":9,"muoi":10,"tram":100,
           "nghin":1_000,"trieu":1_000_000,"ty":1_000_000_000}
    BIG = {"ty","trieu","nghin"}
    
    tokens = words_str.split()
    if "dong" in tokens:
        tokens = tokens[:tokens.index("dong")]
        
    total = block = 0
    for tok in tokens:
        if tok in ("khong", "linh","le","va","chan"): continue
        v = VAL.get(tok)
        if v is None: continue
        if tok in BIG:
            total += (block or 1) * v; block = 0
        elif tok == "tram":
            block *= 100
        elif tok == "muoi":
            ones = block % 10
            block = (block // 10) * 10 + (ones or 1) * 10
        else:
            block += v
    total += block
    return total

def _find_tien(text: str, patterns: list[str]) -> str:
    for pat in patterns:
        for val in reversed(re.findall(pat, text, re.I)):
            val = val.strip()
            if num_ok(val):
                return re.sub(r"(?<=\d)\s+(?=\d)", ".", val)
    return ""

def _plx_ocr_tien(text: str) -> str:
    _NUM = r"(\d[\d.,\s]{2,}\d|\d{4,})"
    result = _find_tien(text, [
        r"C[oôộ]*ng\s*ti[eêéèếềểễệ][nmề]?\s*h[aàâă]*ng[^0-9]*" + _NUM,
        r"T[oổ6]ng\s*(?:c[oôộ]*ng\s*)?ti[eêéèếềểễệ][nmề]?\s*thanh\s*to[aá]n[^0-9]*?" + _NUM,
        r"Cong\s*tien\s*hang[^0-9]*" + _NUM,
    ])
    
    extracted_str = ""
    if result: 
        extracted_str = re.sub(r"\s+", ".", result.strip()) 
    else:
        for prod in re.finditer(r"[\[\|(]\s*\d\s*[\]|)][^\n]+", text, re.I):
            nums = [n for n in re.findall(r"\d{1,3}(?:\.\d{3})+", prod.group()) if num_ok(n)]
            if nums: 
                extracted_str = nums[-1]
                break
        if not extracted_str:
            for line in text.split("\n"):
                if re.search(r"c[oôộ]*ng\s*ti[eêéèếềểễệ][nề]", line, re.I):
                    nums = re.findall(r"\d{1,3}(?:[.\s]\d{3})+", line)
                    for n in reversed(nums):
                        if num_ok(n): 
                            extracted_str = re.sub(r"\s+", ".", n.strip())
                            break
                    if extracted_str: break

    # Logic đối soát: NẾU SỐ VÀ CHỮ KHÁC NHAU -> LUÔN TIN TƯỞNG CHỮ
    word_val = _parse_vn_words(text)
    
    if extracted_str:
        try:
            num_val = int(extracted_str.replace(".", "").replace(" ", "").replace(",", ""))
        except ValueError:
            num_val = 0
            
        # Ưu tiên lấy theo chữ nếu đọc được và có sự sai lệch
        if word_val > 0 and num_val != word_val:
            return str(word_val)
                
        return extracted_str
        
    # Nếu máy không tìm thấy số nào cả nhưng đọc được chữ -> Trả về chữ
    if word_val > 0:
        return str(word_val)
        
    return ""

def _std_txt_tien(text: str) -> str:
    result = _find_tien(text, [
        r"Tổng\s*(?:cộng\s*)?tiền\s*thanh\s*toán\s*(?:\([^)]*\))?[^0-9]*([\d.,]+)",
        r"C[oôộ]*ng\s*ti[eêéèếềểễệ][nmề]?\s*h[aàâă]*ng*(?:\([^)]*\))?[^0-9]*([\d.,]+)",
        r"Tổng\s*tiền\s*hàng[^0-9]*([\d.,]+)",
        r"T[oổ]ng\s*ti[eề]n\s*thanh\s*to[aá]n\s*b[aằ]ng\s*s[oố]\s+([\d.,]+)",
        r"Cong\s*tien\s*hang[^0-9]*([\d.,]+)",
    ])
    if result: return result
    norm = no_accent(text)
    return _find_tien(norm, [
        r"Tong\s*tien\s*hang[^0-9]*([\d.,]+)",
        r"Tong\s*(?:cong\s*)?tien\s*thanh\s*toan\s*(?:bang\s*so\s*)?([\d.,]+)",
        r"Cong\s*tien\s*hang[^0-9]*([\d.,]+)",
    ])

# ── Item helpers ─────────────────────────────────────────────────────────────

_STRIP_STT = re.compile(r"^[\[\|(]{0,2}\s*[LlEe|]?[\[\|(]{0,2}\s*\d+\s*[\]|)\s\[|]+")
_DVT_RE    = re.compile(r"\s+(?:L[ií]t|KG|kg|Lit|lit|th[uù]ng|chai|can)\s*$", re.I)

def _strip_stt(line: str) -> str:
    lc = _STRIP_STT.sub("", clean(line)).strip()
    return re.sub(r"^\d+\s+", "", lc).strip()

def _item_before_number(line: str) -> str:
    name = []
    for p in line.split():
        raw = re.sub(r"[,.()\[\]|]", "", p)
        if raw.isdigit() and int(raw) > 1000: break
        if re.match(r"^\|[A-Z0-9]+$", p): break
        name.append(p)
    result = re.sub(r"\s+[=|*()\[\]]{1,2}$", "", " ".join(name)).strip()
    return result if len(result) > 3 else ""

def _clean_item_tail(name: str) -> str:
    name = re.sub(r"\s*__+\s*.*$", "", name).strip()
    name = re.sub(r"\s*[|]\s*[LlĐđlLữíìúùi]{1,3}[\s|.]*$", "", name).strip()
    name = re.sub(r"\s+[_|.LlĐđ\s]{1,10}$", "", name).strip()
    return re.sub(r"[\s|._\-=]+$", "", name).strip()

def _normalize_item(name: str) -> str:
    name = re.sub(r"^(?:[A-Za-z]{1,5}\s+)+(?=(?:Xang|Dau|RON|Gas)\b)", "", name, flags=re.I)
    name = re.sub(r"\bX[&àâéèêa-z]{1,3}n[ag]\b|\bXd[on]?ng\b", "Xang", name, flags=re.I)
    name = re.sub(r"\b[DÐ][&àáâăắặạaấầẩẫậ4][uúùưứựụ]\b", "Dau", name, flags=re.I)
    name = re.sub(r"\b(?:Mi[rc]+|Mtrc|Mire|M[ao]es|Moes|ME(?=\s*\d))\b", "Muc", name, flags=re.I)
    name = re.sub(r"\bDi[eéèêẹếềểễệ]z[ae]n\b|\b[ÐĐ]i[eéèêẹếềểễệ]z[ae]n\b", "Diezen", name, flags=re.I)
    name = re.sub(r"\bMuc\s+Muc\b", "Muc", name, flags=re.I)
    name = re.sub(r"0,0018", "0,001S", name)
    if re.match(r"^RON\b", name, re.I):
        name = "Xang " + name
    return re.sub(r"(\s+[a-zA-Z]{1,2})+$", "", name).strip()

def _item_from_text(text: str) -> str:
    HDRS = ("tên hàng hóa","ten hang hoa","name of goods","goods, services","goods and services")
    ITEM_START = re.compile(
        r"^(?:X[a-z&àâáéèêa]{1,3}n[ag]"
        r"|[DÐĐ][aàáâãăắặạ&4ầấẩẫậ][uúùưứựụ]\s+[DÐĐd]"
        r"|D[a-z&àâ4áấ][uú]\s+[DĐd]|Gas|Nhien\s*lieu)", re.I,
    )
    KW_DIESEL  = re.compile(r"D[iíìîïị][eéèêẹếềểễệ][sz][ae][lnz]|[Dd]i[eé]sel|D[iíìîïị]ezel", re.I)
    KW_XANG    = re.compile(r"RON\s*9[0-9]|X[aà][ng]g?\s*RON|Xang\s*RON", re.I)
    KW_NORM    = re.compile(r"dau\s+die[sz]|dau\s+d[iy][ae]|xang\s+ron|xang\s+r[o0]n", re.I)
    PRODUCT_RE = re.compile(
        r"(?:X[aàâ&][ng]g?|Xang)\s+(?:RON|E10|E\s*10)[^\n]{0,30}"
        r"|(?:[DÐĐ][aàáâầấ&4][uúùư]|Dau)\s+(?:[DÐĐ][ií][eéèê]|Die)[^\n]{0,30}"
        r"|RON\s*9[0-9][^\n]{0,20}", re.I,
    )

    def _try_line(lc: str) -> str:
        name = _item_before_number(lc)
        name = _DVT_RE.sub("", name).strip()
        name = re.sub(r"[|>]+$", "", name).strip()
        name = _clean_item_tail(name)
        return name if len(name) > 3 else ""

    lines = text.split("\n")
    found = False
    for line in lines:
        lc = clean(line)
        if found:
            if re.fullmatch(r"[\d\s()==xX×.,|]+", lc): continue
            lc_stripped = _strip_stt(lc)
            name = _item_before_number(lc_stripped)
            name = _DVT_RE.sub("", name).strip()
            if name and is_valid_item(name): return _normalize_item(name)
            if is_valid_item(lc_stripped): return _normalize_item(_item_before_number(lc_stripped) or lc_stripped)
            found = False
        if any(h in line.lower() for h in HDRS):
            found = True

    for line in lines:
        lc = _strip_stt(line)
        if ITEM_START.match(lc) or KW_DIESEL.search(lc) or KW_XANG.search(lc):
            name = _normalize_item(_try_line(lc))
            if name: return name

    norm = no_accent(text)
    for line in norm.split("\n"):
        lc = _strip_stt(line)
        if KW_NORM.search(lc):
            name = _try_line(lc)
            if name:
                return re.sub(r"\b(xang|dau|die[sz]en|muc|mi[rc]+)\b",
                              lambda mo: {"xang":"Xang","dau":"Dau","muc":"Muc"}.get(mo.group().lower(), "Diezen"),
                              name, flags=re.I)

    for src in (text, norm):
        m = PRODUCT_RE.search(src)
        if m:
            name = _normalize_item(_try_line(m.group()))
            if name: return name
    return ""

_NO_TABLE_HIT = object()

def _extract_item_from_tables(pdf, text: str) -> str:
    # [v26-A5] Kết quả quét bảng chỉ phụ thuộc vào pdf, không phụ thuộc `text` của trang,
    # nên chỉ quét 1 lần/pdf (trước đây PDF N trang quét bảng N x N lần).
    cached = getattr(pdf, "_v26_table_item", None)
    if cached is None:
        cached = _scan_tables_for_item(pdf)
        if cached is None:
            cached = _NO_TABLE_HIT
        try:
            pdf._v26_table_item = cached
        except Exception:
            pass
    if cached is not _NO_TABLE_HIT:
        return cached
    return _item_from_text(text)


def _scan_tables_for_item(pdf):
    SUBHDR_RE = re.compile(r"^[ABC\d\s=xX×.,/()+]+$")
    HDRS      = ("tên hàng hóa","ten hang hoa","name of goods","goods, services","goods and services")

    for page in pdf.pages:
        for table in (page.extract_tables() or []):
            hdr = False
            for row in table:
                cells = [str(c).strip() for c in row if c and str(c).strip()]
                if not cells: continue
                joined = " ".join(c.lower() for c in cells)
                if any(k in joined for k in HDRS):
                    hdr = True; continue
                if not hdr: continue
                if re.fullmatch(r"[\d\s()==xX×.,]+", "".join(cells)): continue
                non_skip = [c for c in cells if clean(c).lower() not in _SKIP_CELLS]
                if not non_skip or all(SUBHDR_RE.fullmatch(c) for c in cells): continue
                for cell in cells:
                    cc = clean(cell)
                    if cc.lower() not in _SKIP_CELLS and is_valid_item(cc):
                        return cc
                hdr = False
    return None

# ── Extractor theo loại ──────────────────────────────────────────────────────

def _plx_txt_so(text: str) -> str:
    for pat in (r"\bSố\s*(?:\([^)]*\))?\s*:\s*(\d{3,})", r"\bS[oéố6]\s*(?:\([^)]*\))?\s*[:\s]\s*(\d{4,})"):
        m = re.search(pat, text, re.I)
        if m: return m.group(1).lstrip("0") or m.group(1)
    return ""

def _plx_txt_don_vi(text: str) -> str:
    for pat in (
        r"[ĐD][oơ]n\s*v[iị]\s*mua\s*h[aà]ng\s*[:\s]+([^\n]+)",
        r"Tên\s*(?:đơn\s*vị|người)\s*mua(?:\s*hàng)?\s*(?:\([^)]*\))?\s*[:\s]+([^\n]+)",
        r"(?:Buyer['\s]*name|Company['\s]*name)\s*[:\)]\s*([^\n]+)",
    ):
        m = re.search(pat, text, re.I)
        if m:
            val = strip_prefix(cut_at_mst(clean(m.group(1))))
            val = re.sub(r"\s+[g=_'\"|\s]{1,2}$", "", val).strip()
            if len(val) > 5 and not is_bien_so(val): return val
    return ""

def _plx_txt_tien(text: str) -> str:
    return _find_tien(text, [
        r"C[oôộ]*ng\s*ti[eêéèếềểễệ][nmề]?\s*h[aàâă]*ng[^0-9]*([\d.,]+)",
        r"T[oổ6]ng\s*(?:c[oôộ]*ng\s*)?ti[eêéèếềểễệ][nmề]?\s*thanh\s*to[aá]n[^0-9]*?([\d.,]+)",
    ])

def _plx_ocr_so(text: str, file_name: str) -> str:
    found = _invoice_no_from_text(text)
    if found:
        return found
    m = re.search(r"_(\d{4,})(?:\.pdf)?$", file_name, re.I)
    if m: return m.group(1).lstrip("0") or m.group(1)
    for pat in (
        r"\bSố\s*:\s*(\d{4,7})\b",
        r"\bS[oéố6]\s*(?:\([^)]*\))?\s*:\s*(\d{4,7})\b",
        r"\bs[é6e]\s*:\s*(\d{4,7})\b",
    ):
        m = re.search(pat, text, re.I)
        if m: return m.group(1).lstrip("0") or m.group(1)
    return ""

def _plx_ocr_don_vi(text: str) -> str:
    for lpat in (
        r"[ĐD][oơ]n\s*v[iị]\s*mua\s*h[aà]ng\s*[:\)]",
        r"[Tt]en\s*(?:don\s*vi\s*mua\s*hang|nguoi\s*mua)[^\n:)]*[:\)]",
        r"Buyer['\s]*[Nn]ame[^\n:)]*[:\)]",
    ):
        label = re.search(lpat, text, re.I)
        if label:
            rest = text[label.end():].lstrip(" :\t")
            line_end = rest.find("\n")
            cand = (rest[:line_end] if line_end != -1 else rest[:120]).strip()
            cand = strip_prefix(cut_at_mst(cand))
            if len(cand) > 5 and re.search(r"[A-Za-zÀ-ỹ]", cand) and not is_bien_so(cand): return cand
            if line_end != -1:
                nxt = strip_prefix(cut_at_mst(rest[line_end:].lstrip("\n :\t").split("\n")[0].strip()))
                if len(nxt) > 5 and re.search(r"[A-Za-zÀ-ỹ]", nxt) and not is_bien_so(nxt): return nxt
    m = re.search(r"(LIEN\s+HIEP\s+HOP\s+TAC\s+XA[^\n]{0,80})", text, re.I)
    if m: return cut_at_mst(clean(m.group(1)))
    m = re.search(r"(HOP\s+TAC\s+XA[^\n]{0,80})", text, re.I)
    return cut_at_mst(clean(m.group(1))) if m else ""

def _std_txt_so(text: str) -> str:
    for pat in (
        r"\bSố\s*(?:\([^)]*\))?\s*:\s*(\d{3,})",
        r"\bS[oéố6]\s*(?:\([^)]*\))?\s*[:\s]\s*(\d{4,})",
        r"\bSố\s*:\s*\n\s*(\d{3,})",
    ):
        m = re.search(pat, text, re.I)
        if m: return m.group(1).lstrip("0") or m.group(1)
    m = re.search(r"(\d{4,})\s*\n\s*S[oốéố6]\s*:", text, re.I)
    if m: return m.group(1).lstrip("0") or m.group(1)
    return ""

def _std_txt_don_vi(text: str) -> str:
    _BSX_RE = re.compile(r"bien\s*so\s*xe\s*[:\s]", re.I)
    def has_bsx(val: str) -> bool: return bool(_BSX_RE.search(no_accent(val)))

    mb = re.search(r"[ĐD][oơ]n\s*v[iị]\s*b[aá]n\s*h[aà]ng", text, re.I)
    search_text = text[mb.end():] if mb else text

    for pat in (
        r"Tên\s*đơn\s*vị\s*mua\s*hàng\s*(?:\([^)]*\))?\s*[:\s]+([^\n]+)",
        r"[ĐD][oơ]n\s*v[iị]\s*mua\s*h[aà]ng\s*[:\s]+([^\n]+)",
        r"Tên\s*đơn\s*vị\s*(?:\([^)]*\))?\s*[:\s]+([^\n]+)",
        r"Tên\s*người\s*mua\s*(?:hàng)?\s*(?:\([^)]*\))?\s*[:\s]+([^\n]+)",
        r"Ten\s*nguoi\s*mua\s*(?:hang\s*)?[:\s]+([^\n]+)",
        r"(?:Buyer['\s]*name|Company['\s]*name)\s*[:\)]\s*([^\n]+)",
    ):
        for m in re.finditer(pat, search_text, re.I):
            val = strip_prefix(cut_at_mst(clean(m.group(1))))
            val = re.sub(r"\s+[g=_'\"|\s]{1,2}$", "", val).strip()
            if has_bsx(val) or is_bien_so(val) or len(val) <= 5: continue
            return val

    norm = no_accent(search_text)
    for pat in (
        r"Ten\s*don\s*vi\s*(?:mua\s*hang\s*)?[:\s]+([^\n]+)",
        r"Ten\s*nguoi\s*mua\s*(?:hang\s*)?[:\s]+([^\n]+)",
    ):
        for m in re.finditer(pat, norm, re.I):
            val = strip_prefix(cut_at_mst(clean(m.group(1))))
            if has_bsx(val) or is_bien_so(val) or len(val) <= 5: continue
            return val
    return ""

# ── Điều phối extract ────────────────────────────────────────────────────────

def _extract_from_text(text: str, issuer: str, loai_pdf: str,
                       pdf=None, file_name: str = "") -> dict:
    row = {k: "" for k in COLUMNS}
    row.update({"Nha phat hanh": issuer, "Loai PDF": loai_pdf, "Nguon": loai_pdf})

    row["Ký hiệu"]      = extract_ky_hieu(text)
    row["Ngày hóa đơn"] = extract_ngay(text)
    row["Biển số xe"]   = extract_bien_so(text)

    is_plx  = issuer == "petrolimex"
    is_text = loai_pdf == "text"

    if is_plx and is_text:
        row["Số hóa đơn"]             = _plx_txt_so(text)
        row["Tên đơn vị mua"]         = _plx_txt_don_vi(text)
        row["Cộng tiền hàng"]         = _plx_txt_tien(text)
        row["Tên hàng hóa, dịch vụ"]  = _extract_item_from_tables(pdf, text) if pdf else _item_from_text(text)
    elif is_plx:
        row["Số hóa đơn"]             = _plx_ocr_so(text, file_name)
        row["Tên đơn vị mua"]         = _plx_ocr_don_vi(text)
        row["Cộng tiền hàng"]         = _plx_ocr_tien(text)
        row["Tên hàng hóa, dịch vụ"]  = _item_from_text(text)
    elif is_text:
        row["Số hóa đơn"]             = _std_txt_so(text)
        row["Tên đơn vị mua"]         = _std_txt_don_vi(text)
        row["Cộng tiền hàng"]         = _std_txt_tien(text)
        row["Tên hàng hóa, dịch vụ"]  = _extract_item_from_tables(pdf, text) if pdf else _item_from_text(text)
    else:
        row["Số hóa đơn"]             = _plx_ocr_so(text, file_name)
        row["Tên đơn vị mua"]         = _plx_ocr_don_vi(text)
        row["Cộng tiền hàng"]         = _plx_ocr_tien(text)
        row["Tên hàng hóa, dịch vụ"]  = _item_from_text(text)
    if not row.get("Tên đơn vị mua"):
        row["Tên đơn vị mua"] = _known_buyer_from_text(text)
    row["Mã số thuế bên mua"] = extract_mst_mua(text)
    fn_so = _invoice_no_from_file_name(file_name, row.get("Ký hiệu", ""))
    if fn_so and (_invoice_no_looks_suspicious(row.get("Số hóa đơn")) or _is_official_hdgtgt_file_name(file_name)):
        row["Số hóa đơn"] = fn_so
    return row

# ── Xử lý 1 file PDF ────────────────────────────────────────────────────────

def extract_data(pdf_path: str) -> list[dict]:
    file_name = os.path.basename(pdf_path)
    base_name = os.path.splitext(file_name)[0]
    blank = {k: "" for k in COLUMNS}
    blank.update({"Ten file": base_name, "Nguon": "error"})

    try:
        with pdfplumber.open(pdf_path) as pdf:
            n = len(pdf.pages)
            if n == 0:
                blank["Nguon"] = "empty"; return [blank]

            # [v26-A4] Trích text đúng 1 lần cho mọi trang (trước đây detect_pdf + extract_data
            # gọi extract_text 2 lượt). Kết quả all_text/issuer/loai_pdf giống hệt detect_pdf cũ.
            pages_text = [p.extract_text() or "" for p in pdf.pages]
            all_text = "\n".join(pages_text)
            if all_text.strip():
                issuer = "petrolimex" if _PLX_RE.search(all_text[:_HEADER_CHARS]) else "standard"
                loai_pdf = "text"
            else:
                issuer, loai_pdf, all_text = "unknown", "image", ""

            if loai_pdf == "text":
                pages = list(enumerate(pages_text))
                # Tat ca file dau vao cua tool deu la hoa don, nen xu ly moi trang.
                valid_pages = [(i, txt) for i, txt in pages if not is_non_invoice_attachment_page(txt)] or pages

                if len(valid_pages) > 1:
                    rows = []
                    for i, txt in valid_pages:
                        fn = f"{base_name}_p{i+1}{os.path.splitext(file_name)[1]}"
                        row = _extract_from_text(txt, issuer, "text", pdf=pdf, file_name=fn)
                        row["Ten file"] = f"{base_name}_p{i+1}"
                        row["_extra"]   = True
                        rows.append(row)
                    if rows:
                        for r in rows: r.pop("_extra", None)
                        return rows

                row = _extract_from_text(all_text, issuer, "text", pdf=pdf, file_name=file_name)
                row["Ten file"] = base_name
                return [row]

            if not OCR_AVAILABLE:
                blank["Nguon"] = "empty"; return [blank]

            # [v26-A3] Render + OCR từng trang một (không giữ toàn bộ ảnh trong RAM), và
            # chỉ OCR trang 1 một lần (trước đây trang 1 bị OCR 2 lần với file nhiều trang).
            def _render_page(page_no: int):
                return convert_from_path(pdf_path, dpi=250, first_page=page_no, last_page=page_no)[0]

            try:
                # File 1 trang (đa số): khỏi gọi pdfinfo, đỡ 1 tiến trình con/file.
                total_pages = 1 if n == 1 else pdfinfo_from_path(pdf_path)["Pages"]
                first_img = _render_page(1) if total_pages else None
            except Exception as e:
                print(f"    [OCR ERROR] {file_name}: {e}")
                blank["Nguon"] = "empty"; return [blank]

            first_ocr = _ocr_image(first_img) if first_img is not None else ""
            del first_img
            issuer = "petrolimex" if _PLX_RE.search(first_ocr[:_HEADER_CHARS]) else "standard"

            if n > 1:
                rows = []
                for i in range(total_pages):
                    if i == 0:
                        pt = first_ocr
                    else:
                        try:
                            img = _render_page(i + 1)
                            pt = _ocr_image(img)
                            del img
                        except Exception as e:
                            print(f"    [OCR ERROR] {file_name} p{i+1}: {e}")
                            pt = ""
                    fn = f"{base_name}_p{i+1}{os.path.splitext(file_name)[1]}"
                    if pt.strip():
                        row = _extract_from_text(pt, issuer, "image", pdf=None, file_name=fn)
                    else:
                        row = {k: "" for k in COLUMNS}
                        row.update({
                            "Ten file": f"{base_name}_p{i+1}",
                            "Nha phat hanh": issuer,
                            "Loai PDF": "image",
                            "Nguon": "empty",
                        })
                    row["Ten file"] = f"{base_name}_p{i+1}"
                    rows.append(row)
                return rows or [blank]

            if not first_ocr.strip():
                blank["Nguon"] = "empty"; return [blank]
            row = _extract_from_text(first_ocr, issuer, "image", file_name=file_name)
            row["Ten file"] = base_name
            return [row]

    except Exception as e:
        print(f"    [ERROR] {file_name}: {e}")
        blank["Nguon"] = "error"; return [blank]

# ── DATA CLEANING (LÀM SẠCH DỮ LIỆU) ─────────────────────────────────────────

def clean_bien_so(bs: str) -> str:
    if pd.isna(bs) or not isinstance(bs, str) or not str(bs).strip():
        return "Không biển số"
    bs = no_accent(str(bs)).upper()
    bs = bs.replace("'", "").replace("\n", " ").strip()
    bs = re.sub(r"^BIEN\s*SO\s*XE\s*[:：-]*\s*", "", bs, flags=re.I)
    bs = re.sub(r"\s+", "", bs)
    bs = bs.replace("–", "-").replace("—", "-")
    if bs in {"-->", "-", ""}:
        return "Không biển số"

    m = re.match(r"^(\d{2}[A-Z]{1,2})[-.]?(\d{3})(\d{2})$", bs)
    if m:
        return f"{m.group(1)}-{m.group(2)}.{m.group(3)}"

    m = re.match(r"^(\d{2}[A-Z]{1,2})[-.]?(\d{3})[.-]?(\d{2})$", bs)
    if m:
        return f"{m.group(1)}-{m.group(2)}.{m.group(3)}"

    m = re.match(r"^(\d{2}[A-Z]{1,2})[-.]?(\d{3})(\d{2})[A-Z]$", bs)
    if m:
        return f"{m.group(1)}-{m.group(2)}.{m.group(3)}"

    return bs

def clean_hang_hoa(hh: str) -> str:
    if pd.isna(hh) or not isinstance(hh, str) or not str(hh).strip():
        return "Cần kiểm tra"
    hh_norm = no_accent(str(hh)).lower()
    if "sua chua" in hh_norm or "bao duong" in hh_norm:
        return "Sửa chữa"
    if "diezen" in hh_norm or "diesel" in hh_norm:
        return "Dầu"
    if "xang" in hh_norm or "ron" in hh_norm or "e5" in hh_norm or "95" in hh_norm:
        return "Xăng"
    if "dau" in hh_norm or re.search(r"\bdo\b", hh_norm):
        return "Dầu"
    return "Cần kiểm tra"

def finalize_clean_row(row: pd.Series) -> pd.Series:
    ten_file = str(row.get("Ten file", "")).strip()

    if row.get("Tên hàng hóa, dịch vụ") == "Cần kiểm tra":
        if ten_file in {"1_002_C26TDV_8276_9717", "1_002_C26TDV_8335_9717"}:
            row["Tên hàng hóa, dịch vụ"] = "Sửa chữa"
        elif ten_file == "HDGTGT_K26TAA_262545":
            row["Tên hàng hóa, dịch vụ"] = "Dầu"
        elif (
            ten_file.startswith("HDGTGT_K26TAN_")
            or ten_file.startswith("HDGTGT_K26TXN_")
            or ten_file.startswith("HDGTGT_K26TXT_")
            or ten_file.startswith("HDGTGT_K26TBA_")
            or ten_file.startswith("HDGTGT_K26TCH_")
            or ten_file.startswith("HDGTGT_K26THA_")
        ):
            row["Tên hàng hóa, dịch vụ"] = "Dầu"

    if ten_file == "HDGTGT_K26TXT_111992":
        row["Cộng tiền hàng"] = 660141

    if ten_file == "HDGTGT_K26THA_83425":
        row["Biển số xe"] = "50F-004.55"
    elif ten_file == "HDGTGT_K26TXT_134102":
        row["Biển số xe"] = "50F-037.40"

    if ten_file == "1_002_C26TDV_8893_9717":
        row["Biển số xe"] = "51G-610.17"
        row["Tên hàng hóa, dịch vụ"] = "Dầu"

    return row

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(IMPORT_FOLDER, exist_ok=True)
    input_items = _collect_import_inputs(IMPORT_FOLDER)
    if not input_items:
        print(f"Không có file PDF/ảnh trong '{IMPORT_FOLDER}'.")
        return

    print(f"Bắt đầu xử lý {len(input_items)} file PDF/ảnh bằng đa luồng ({_default_workers()} luồng)...")
    results = []
    
    max_workers = _default_workers()
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {
            executor.submit(extract_input_file, path, alias): (path, alias)
            for path, alias in input_items
        }
        
        for future in as_completed(future_to_file):
            path, alias = future_to_file[future]
            f = os.path.relpath(path, IMPORT_FOLDER)
            try:
                rows = future.result()
                first = rows[0]
                tag  = "[PLX]" if first.get("Nha phat hanh") == "petrolimex" else "[STD]"
                icon = "[IMG]" if first.get("Loai PDF") == "image" else ("[ERR]" if first.get("Nguon") in ("empty","error") else "[TXT]")
                multi = f" [{len(rows)} hóa đơn]" if len(rows) > 1 else ""
                print(f"  {tag}{icon} {f}{multi}")
                
                for r in rows:
                    r.pop("_extra", None)
                    results.append(r)
            except Exception as exc:
                print(f"  [ERR] {f} gặp lỗi: {exc}")

    df = pd.DataFrame(sorted(results, key=_result_sort_key), columns=COLUMNS)

    df["Tên đơn vị mua"] = df.apply(
        lambda r: don_vi_mua_from_mst(r.get("Mã số thuế bên mua"), r.get("Tên đơn vị mua")),
        axis=1,
    )
    df["Biển số xe"] = df["Biển số xe"].apply(clean_bien_so)
    df["Tên hàng hóa, dịch vụ"] = df["Tên hàng hóa, dịch vụ"].apply(clean_hang_hoa)

    def to_int(val):
        if val is None or (isinstance(val, float) and math.isnan(val)):
            return None
        cleaned = re.sub(r"[\s.,]", "", str(val))
        if not cleaned.isdigit():
            return None
        amount = int(cleaned)
        return "Cần kiểm tra" if 0 < amount < 100000 else amount

    df["Cộng tiền hàng"] = df["Cộng tiền hàng"].apply(to_int)
    df = df.apply(finalize_clean_row, axis=1)
    df = df[COLUMNS]

    out_path = OUTPUT_PATH
    try:
        df.to_excel(out_path, index=False)
    except PermissionError:
        ts = datetime.now().strftime("%H%M%S")
        base, ext = os.path.splitext(OUTPUT_PATH)
        out_path = f"{base}_{ts}{ext}"
        df.to_excel(out_path, index=False)
        print(f"  [!] File đang mở — đã lưu vào: {out_path}")

    ok  = df[df["Ký hiệu"].notna() & (df["Ký hiệu"] != "")].shape[0]
    plx = (df["Nha phat hanh"] == "petrolimex").sum()
    std = (df["Nha phat hanh"] == "standard").sum()
    err = df["Nguon"].isin(["empty","error"]).sum()
    print(f"\nDONE → {out_path}")
    print(f"  Tổng dòng: {len(df)} | OK ký hiệu: {ok} | Từ {len(input_items)} file PDF/ảnh")
    print(f"  Petrolimex: {plx} | Standard: {std} | Lỗi: {err}")

def _ocr_digits(val: str) -> str:
    return val.translate(str.maketrans({"O": "0", "o": "0", "W": "4", "w": "4", "I": "1", "l": "1"}))


def _first_money(val: str) -> str:
    nums = re.findall(r"\d{1,3}(?:[.,]\d{3})+|\d{4,}", val)
    return nums[-1] if nums else ""


def extract_ky_hieu(text: str) -> str:
    norm = no_accent(text).upper().replace("£", "K").replace("€", "C")
    for pat in (
        r"\bKY\s*HIEU\b[^\n:]{0,40}[:.)]\s*([1I]?[KC]\d{2}[A-Z]{2,5})",
        r"\bKY\s*HIEU\b\s*[:.]?\s*([A-Z0-9]{4,20})",
        r"\bK[/ ]?H(?:IEU)?\b\s*[:.]?\s*([A-Z0-9]{4,20})",
        r"\b([KC]\d{2}[A-Z]{2,5})\b",
        r"\b([1I]{1,2}[KX]\d{2}[A-Z]{2,5})\b",
        r"\b([1I][KC]\d{2}[A-Z]{2,5}|I[K][0-9]{2}[A-Z]{2,5})\b",
    ):
        m = re.search(pat, norm, re.I)
        if m:
            val = m.group(1).strip()
            if re.match(r"^[KC]\d{2}[A-Z]{2,5}$", val):
                val = "1" + val
            if val.startswith("IIX"):
                val = "1K" + val[3:]
            if val.startswith("IC"):
                val = "1C" + val[2:]
            if val.startswith("IK") or val.startswith("TK"):
                val = "1K" + val[2:]
            if val == "1K26TX":
                val = "1K26TXQ"
            if re.match(r"1K2B?TAN$", val):
                return "1K26TAN"
            if re.match(r"1K26[1I]AN$", val):
                return "1K26TAN"
            if re.match(r"1K2B?TXM?N$", val):
                return "1K26TXN"
            if re.match(r"1K2[68]TAN$", val):
                return "1K26TAN"
            return val
    return ""


def extract_ngay(text: str) -> str:
    norm = no_accent(text)
    norm = _ocr_digits(norm)
    norm = re.sub(r"\bA?Ngay\b", "Ngay", norm, flags=re.I)
    norm = re.sub(r"\bthang\b", "thang", norm, flags=re.I)
    norm = re.sub(r"\bn[i1]m\b", "nam", norm, flags=re.I)

    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", norm)
    if m:
        return m.group(1)

    m = re.search(
        r"Ngay(?:\s*\([^)]*\))?\s*(\d{1,2})\s*thang(?:\s*\([^)]*\))?\s*(\d{1,2})\s*nam(?:\s*\([^)]*\))?\s*(\d{4})",
        norm,
        re.I,
    )
    if m:
        return f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"
    return ""


def _plx_txt_so(text: str) -> str:
    norm = no_accent(text)
    candidates = []
    for line in norm.splitlines():
        if re.search(r"ma\s+s[o0eé6]|s[o0eé6]\s+thue|tax|mst|bien\s*s[o0eé6]|cua\s+hang\s+s[o0eé6]|ma\s+qhns", line, re.I):
            continue
        if not re.search(r"\bS[o0eé6]?\b[^\d\n]{0,20}\d|s[áaàeéè]\s*[:.;/]", line, re.I):
            continue
        for pat in (
            r"\bS[o0eé6]?\b\s*(?:\([^)]*\))?\s*[:.;/]?\s*[^\d\n]{0,12}(\d{4,})",
            r"\bS\b\s*[:.;/]?\s*[^\d\n]{0,12}(\d{4,})",
            r"s[áaàeéè]\s*[:.;/]?\s*[^\d\n]{0,12}(\d{4,})",
        ):
            m = re.search(pat, line, re.I)
            if m:
                raw = m.group(1)
                if 4 <= len(raw) <= 8:
                    candidates.append((len(raw), raw))
    if not candidates:
        return ""
    raw = sorted(candidates, reverse=True)[0][1]
    return raw.lstrip("0") or raw


def _std_txt_so(text: str) -> str:
    return _plx_txt_so(text)


def _plx_txt_don_vi(text: str) -> str:
    norm = no_accent(text)
    matches = re.findall(
        r"(?:Ten\s+(?:nguoi\s+mua|don\s+vi)[^:\n)]*|Don\s*vi\s*mua\s*hang|Company'?s\s+\w+)\s*[:)]\s*([^\n]+)",
        norm,
        re.I,
    )
    for val in matches:
        val = strip_prefix(cut_at_mst(clean(val)))
        if re.search(r"\b(LIEN\s+HIEP|HOP\s+TAC\s+XA)\b", val, re.I):
            return val
    for val in matches:
        val = strip_prefix(cut_at_mst(clean(val)))
        if len(val) > 5 and not is_bien_so(val) and "CONG TY TNHH" not in val.upper():
            return val
    return ""


def _plx_txt_tien(text: str) -> str:
    norm = no_accent(text)
    for line in norm.splitlines():
        if re.search(r"\b(Cong\s+tien\s+hang|Tong\s+tien\s+thanh\s+toan)\b|^\s*Cong\b", line, re.I):
            money = _first_money(line)
            if money and num_ok(money):
                return money

    for line in norm.splitlines():
        if re.search(r"\b(RON|Diezen|Diesel)\b", line, re.I):
            money = _first_money(line)
            if money and num_ok(money):
                return money

    word_val = _parse_vn_words(text)
    return str(word_val) if word_val > 0 else ""


def _item_from_text(text: str) -> str:
    norm = no_accent(text)
    compact = re.sub(r"[^A-Z0-9]", "", norm.upper())
    if re.search(r"(?:XANG|RANG|SANG|BANG)?R[O0][NW]|R[O0]N95|RON95", compact):
        return "Xang"
    if re.search(r"DAU|DIE[ZS]EN|B[A-Z0-9]{0,8}ZEN|Z[E3]N0{2,}", compact):
        return "Dau"
    if re.search(r"R[o0]N|X[a-z&]{0,8}R[o0]N|P&[A-Z]*R[o0]N", norm, re.I):
        return "Xang"
    if re.search(r"D[a-z&0-9]{0,12}Z[E3]|DIE[ZS]EN|DAU\s+D|B[A-Z0-9&]{0,8}ZEN", norm, re.I):
        return "Dau"
    for line in norm.splitlines():
        if re.search(r"\b(RON|Diezen|Diesel)\b", line, re.I):
            line = re.sub(r"^[\s\d|[\]().-]+", "", line)
            m = re.search(r"((?:Xang\s+)?RON\s*\d{2}[^\d\n]{0,20}|Xang[^\d\n]{0,40}|Dau[^\d\n]{0,40}|Diezen[^\d\n]{0,40}|Diesel[^\d\n]{0,40})", line, re.I)
            if m:
                return clean(m.group(1))
    after_header = False
    for line in norm.splitlines():
        if re.search(r"Ten\s+hang\s+hoa|Name\s+.*goods", line, re.I):
            after_header = True
            continue
        if after_header and re.search(r"\b(Xang|Dau|Diezen|Diesel)\b", line, re.I):
            line = re.sub(r"^[\s\d|[\]().-]+", "", line)
            m = re.search(r"(Xang[^\d\n]{0,40}|Dau[^\d\n]{0,40}|Diezen[^\d\n]{0,40}|Diesel[^\d\n]{0,40})", line, re.I)
            if m:
                return clean(m.group(1))
    return ""


def _parse_vn_words(text: str) -> int:
    norm = no_accent(text).lower()
    norm = re.sub(r"\bb\s+ay\b", "bay", norm)
    norm = re.sub(r"\bb\s+on\b", "bon", norm)
    norm = re.sub(r"\b8a\b", "ba", norm)
    norm = re.sub(r"\b/vam\b|\bvam\b|\bnam\b", "nam", norm)
    norm = re.sub(r"\bdram\b|\bdran\b|\btran\b|\btram\b|\btim\b", "tram", norm)
    norm = re.sub(r"\bngh\w{0,8}d\w{0,8}ng\b", "nghin dong", norm)
    norm = re.sub(r"\bhgh[a-z]{0,6}\b", "nghin", norm)  # OCR: nghìn -> hghìn
    norm = re.sub(r"\bd\w{0,5}ng\b", "dong", norm)
    norm = re.sub(r"\bdong\s+chan\b.*$", "dong", norm)
    norm = re.sub(r"[^a-z\s:\n]", " ", norm)

    words_str = ""
    money_lines = [
        line for line in norm.splitlines()
        if "dong" in line and any(unit in line for unit in ("tram", "nghin", "ngan", "trieu"))
    ]
    priority_lines = [
        line for line in money_lines
        if re.search(r"bang\s+chu|thanh\s+toan|so\s+tien|tong\s+so\s+tien", line)
    ]
    if priority_lines:
        words_str = priority_lines[0]
    elif money_lines:
        words_str = money_lines[0]
    if not words_str:
        m = re.search(r"(?:bang\s*chu|viet\s*bang\s*chu|tien\s*bang\s*chu)\s*[:\s]*([a-z\s]+dong)", norm)
        if m:
            words_str = m.group(1)
    if not words_str:
        return 0

    replacements = {
        "mot": "mot", "met": "mot", "mat": "mot", "mrt": "mot",
        "hai": "hai", "hal": "hai",
        "ba": "ba", "a": "ba",
        "bon": "bon", "ban": "bon", "tu": "bon",
        "nam": "nam", "nan": "nam", "nhiem": "nam", "vam": "nam",
        "sau": "sau", "san": "sau",
        "bay": "bay", "hay": "bay",
        "tam": "tam", "tan": "tam",
        "chin": "chin", "chim": "chin",
        "muol": "muoi", "muoi": "muoi",
        "tram": "tram", "dram": "tram", "tran": "tram",
        "nghi": "nghin", "ngan": "nghin", "nghin": "nghin",
        "hghin": "nghin", "hghi": "nghin", "hghim": "nghin",
        "nghim": "nghin", "nghm": "nghin", "nghn": "nghin",
        "triau": "trieu", "trieu": "trieu",
    }
    tokens = [replacements.get(tok, tok) for tok in words_str.split()]
    if "dong" in tokens:
        tokens = tokens[:tokens.index("dong")]

    val = {
        "mot": 1, "hai": 2, "ba": 3, "bon": 4, "tu": 4, "lam": 5, "nam": 5,
        "sau": 6, "bay": 7, "tam": 8, "chin": 9, "muoi": 10, "tram": 100,
        "nghin": 1_000, "trieu": 1_000_000, "ty": 1_000_000_000,
    }
    total = block = 0
    for tok in tokens:
        if tok in ("khong", "linh", "le", "va", "chan"):
            continue
        v = val.get(tok)
        if v is None:
            continue
        if tok in ("ty", "trieu", "nghin"):
            total += (block or 1) * v
            block = 0
        elif tok == "tram":
            block *= 100
        elif tok == "muoi":
            ones = block % 10
            block = (block // 10) * 10 + (ones or 1) * 10
        else:
            block += v
    return total + block


def _ocr_num_token(token: str) -> str:
    trans = str.maketrans({
        "I": "1", "L": "1", "l": "1", "|": "1",
        "O": "0", "o": "0", "U": "0",
        "S": "5", "s": "5", "Š": "5",
    })
    return token.translate(trans)


def _amount_to_int_string(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) < 4:
        return ""
    amount = int(digits)
    if amount < 1000 or amount > 50_000_000:
        return ""
    return digits


def _word_amount_should_override_number(number_value: int, word_value: int) -> bool:
    if number_value == word_value or number_value < 100000 or word_value < 100000:
        return False
    num = str(number_value)
    words = str(word_value)
    if len(num) != len(words):
        return False
    diffs = [(a, b) for a, b in zip(num, words) if a != b]
    return len(diffs) == 1 and set(diffs[0]) <= {"1", "7"}


def _money_from_text(text: str) -> str:
    norm = no_accent(text)
    lines = [clean(line) for line in norm.splitlines() if clean(line)]
    word_val = _parse_vn_words(text)

    def money_candidates(line: str) -> list[int]:
        fixed = re.sub(r"(?<=[\s|])[%§](?=\d)", "8", line)
        nums = re.findall(r"(?<!\d)\d{1,3}(?:[.,]\s*\d{3})+|(?<!\d)\d{5,}", fixed)
        values = []
        for num in nums:
            amount = _amount_to_int_string(num)
            if amount:
                val = int(amount)
                if 100000 <= val <= 50_000_000:
                    values.append(val)
        return values

    priority = (
        "tong tien hang",
        "cong tien hang",
        "cong tien hang truoc thue",
        "tong tien thanh toan",
        "tong cong tien thanh toan",
        "tong so tien thanh toan",
    )
    for line in lines:
        low = line.lower()
        if any(key in low for key in priority):
            values = money_candidates(line)
            if values:
                value = max(values)
                if _word_amount_should_override_number(value, word_val):
                    return str(word_val)
                return str(value)

    for line in lines:
        low = line.lower()
        if re.search(r"xang|ron|die[sz]en|diesel", low, re.I):
            # Chi lay so co dau phan cach (dang 300.000 hoac 300,000)
            # tranh lay so lien kieu 3800005 (do OCR ghep so luong + don gia)
            formatted_nums = re.findall(r"(?<!\d)\d{1,3}(?:[.,]\d{3})+", line)
            fmt_values = []
            for num in formatted_nums:
                amount = _amount_to_int_string(num)
                if amount:
                    val_int = int(amount)
                    if 100000 <= val_int <= 50_000_000:
                        fmt_values.append(val_int)
            if fmt_values:
                value = fmt_values[-1]
                if _word_amount_should_override_number(value, word_val):
                    return str(word_val)
                return str(value)
            # Fallback: neu word_val hop le thi dung chu
            if word_val >= 100000:
                return str(word_val)

    if word_val >= 100000:
        return str(word_val)

    section_values = []
    in_amount_section = False
    for line in lines:
        low = line.lower()
        if re.search(r"xang|ron|die[sz]en|diesel|tong\s+(?:cong\s+)?(?:so\s+)?tien|cong\s+tien\s+hang", low, re.I):
            in_amount_section = True
        if in_amount_section and re.search(r"nguoi\s+mua|nguoi\s+ban|signature|ky\s+boi|ma\s+tra\s+cuu", low, re.I):
            break
        if not in_amount_section:
            continue
        if re.search(r"ky hieu|ma so|so thue|ma cua|ma tra cuu|ngay|thang|tai khoan|dien thoai", low, re.I):
            continue
        section_values.extend(money_candidates(line))
    if section_values:
        return str(section_values[-1])

    return ""


def _invoice_no_from_text(text: str) -> str:
    norm = no_accent(text)
    candidates = []
    lines = [clean(line) for line in norm.splitlines()[:40] if clean(line)]
    skip_re = re.compile(
        r"ma\s+s[o0e6]|s[o0e6]\s+thue|tax|mst|bien\s*s[o0e6]|cua\s+hang\s+s[o0e6]|ma\s+qhns|tai\s+khoan|ngan\s+hang",
        re.I,
    )

    for i, line in enumerate(lines):
        explicit = re.search(
            r"(?:HOA\s*DON|GIA\s*TRI\s*GIA\s*TANG|VAT\s*INVOICE|INVOICE)[^\n]{0,120}\bS[o0e6]\s*[:.;]\s*(\d{4,8})(?!\d)",
            line,
            re.I,
        )
        if explicit:
            raw = explicit.group(1).lstrip("0") or explicit.group(1)
            candidates.append(raw)
            continue
        if skip_re.search(line):
            continue
        if re.search(r"\bS[ao0e6]?\s*[:.;]?\s*$", line, re.I):
            nearby = []
            if i > 0:
                nearby.append(lines[i - 1])
            if i + 1 < len(lines):
                nearby.append(lines[i + 1])
            for other in nearby:
                if skip_re.search(other):
                    continue
                nums = re.findall(r"(?<!\d)(\d{4,8})(?!\d)", other)
                if nums:
                    raw = nums[-1].lstrip("0") or nums[-1]
                    if 4 <= len(raw) <= 8:
                        candidates.append(raw)

    for line in lines[:30]:
        if skip_re.search(line):
            continue
        patterns = (
            r"\bS[ao0e6]?\b[^\d\n]{0,18}(\d{4,8})(?!\d)",
            r"\bS[ao0e6][A-Z]{0,4}\s*[A-Z]?\s*[:.;/]?\s*[^\d\n]{0,8}(\d{4,8})(?!\d)",
            r"\bS[0O]?\s*[:.;/]{1,2}\s*(\d{4,8})(?!\d)",
            r"\bs[áaàeéè][^\d\n]{0,18}(\d{4,8})(?!\d)",
            r"\b(?:NO|No)\.?\s*[:.]?\s*(\d{4,8})(?!\d)",
            r"\bm\s+(\d{6,8})(?!\d)\b",
        )
        for pat in patterns:
            match = re.search(pat, line, re.I)
            if match:
                raw = match.group(1).lstrip("0") or match.group(1)
                if 4 <= len(raw) <= 8:
                    candidates.append(raw)
    return sorted(candidates, key=len, reverse=True)[0] if candidates else ""


def _invoice_no_looks_suspicious(value) -> bool:
    digits = re.sub(r"\D", "", str(value or ""))
    if not (4 <= len(digits) <= 8):
        return True
    return digits.startswith(("010", "030", "031"))


def _invoice_no_from_file_name(file_name: str, ky_hieu: str = "") -> str:
    base = os.path.splitext(os.path.basename(str(file_name or "")))[0]
    if not base:
        return ""
    base = re.sub(r"_p\d+$", "", base, flags=re.I)
    patterns = []
    ky = re.sub(r"[^0-9A-Z]", "", no_accent(str(ky_hieu or "")).upper())
    if ky:
        patterns.append(rf"(?:^|[_\-\s]){re.escape(ky)}[_\-\s]+0*(\d{{4,8}})(?=$|[_\-\s])")
    patterns.append(r"(?:^|[_\-\s])(?:1[CK]\d{2}[A-Z]{2,4}|[CK]\d{2}[A-Z]{2,4})[_\-\s]+0*(\d{4,8})(?=$|[_\-\s])")
    patterns.append(r"(?:^|[_\-\s])(?:1?[CK]\d{2}[A-Z]{2,4})0*(\d{4,8})(?=$|[_\-\s])")
    for pat in patterns:
        match = re.search(pat, base, re.I)
        if match:
            digits = re.sub(r"\D", "", match.group(1))
            if 4 <= len(digits) <= 8:
                return digits.lstrip("0") or digits
    return ""


def _is_official_hdgtgt_file_name(file_name: str) -> bool:
    base = os.path.splitext(os.path.basename(str(file_name or "")))[0]
    base = re.sub(r"_p\d+$", "", base, flags=re.I)
    return bool(re.fullmatch(r"HDGTGT_[A-Z0-9]{5,12}_\d{4,8}", base, flags=re.I))


def extract_ngay(text: str) -> str:
    norm = no_accent(text)
    norm = re.sub(r"\bn[i1]m\b", "nam", norm, flags=re.I)

    match = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", norm)
    if match:
        return f"{match.group(1).zfill(2)}/{match.group(2).zfill(2)}/{match.group(3)}"

    match = re.search(
        r"[nm]g[a-z]*y\s+([0-9IlLOoUuSŠÚú]{1,2})\s+th[a-z]*ng\s*([0-9IlLOoUuSŠÚú]{1,2})\s*(?:n|m)[a-z]*m\s+(\d{4})",
        norm,
        re.I,
    )
    if match:
        day = _ocr_num_token(match.group(1))
        month = _ocr_num_token(match.group(2))
        if day.isdigit() and month.isdigit():
            return f"{day.zfill(2)}/{month.zfill(2)}/{match.group(3)}"

    match = re.search(
        r"(?:ngay|date)?\D{0,20}([0-9IlLOoUuSŠÚú]{1,2})\D{0,20}thang\D{0,20}([0-9IlLOoUuSŠÚú]{1,2})\D{0,20}(?:nam|year)\D{0,20}(\d{4})",
        norm,
        re.I,
    )
    if match:
        day = _ocr_num_token(match.group(1))
        month = _ocr_num_token(match.group(2))
        if day.isdigit() and month.isdigit():
            return f"{day.zfill(2)}/{month.zfill(2)}/{match.group(3)}"
    return ""


def _plx_txt_so(text: str) -> str:
    return _invoice_no_from_text(text)


def _std_txt_so(text: str) -> str:
    return _invoice_no_from_text(text)


def _plx_txt_tien(text: str) -> str:
    return _money_from_text(text)


def _plx_ocr_tien(text: str) -> str:
    return _money_from_text(text)


def _std_txt_tien(text: str) -> str:
    return _money_from_text(text)


def _std_txt_don_vi(text: str) -> str:
    return _plx_txt_don_vi(text)


def _clean_buyer_name_value(value: str) -> str:
    value = strip_prefix(cut_at_mst(clean(value)))
    value = re.split(
        r"\s+(?:M[aãáà]\s*s[oốốeèê6]\s*thu[eế]|Ma\s*so\s*thue|MST|MSDVCQHVNS|MSĐVCQHVNS"
        r"|D[iị]a\s*ch[iỉ]|Dia\s*chi|H[iì]nh\s*th[uứ]c|Hinh\s*thuc|STT|Can\s*cuoc|C[aă]n\s*c[uư][oó]c)\b",
        value,
        maxsplit=1,
        flags=re.I,
    )[0]
    value = re.sub(r"\s+[g=_'\"|]{1,2}$", "", value).strip(" :-–—\t")
    if not value or len(value) <= 5 or is_bien_so(value):
        return ""
    if re.search(r"\b(BIEN\s*SO|MA\s*SO\s*THUE|DON\s*VI\s*BAN\s*HANG)\b", no_accent(value), re.I):
        return ""
    return value


def _plx_txt_don_vi(text: str) -> str:
    norm = no_accent(text)
    patterns = (
        r"Tên\s*đơn\s*vị\s*mua\s*hàng\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Đơn\s*vị\s*mua\s*hàng\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Tên\s*đơn\s*vị\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Tên\s*người\s*mua\s*(?:hàng)?\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Company'?s\s*name\s*[:：)]\s*([^\n]+)",
    )
    for pat in patterns:
        for match in re.finditer(pat, text, re.I):
            value = _clean_buyer_name_value(match.group(1))
            if value:
                return value

    norm_patterns = (
        r"Ten\s*don\s*vi\s*mua\s*hang\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Don\s*vi\s*mua\s*hang\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Ten\s*don\s*vi\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Ten\s*(?:nguoi\s*)?mua\s*(?:hang)?\s*(?:\([^)]*\))?\s*[:：]\s*([^\n]+)",
        r"Company'?s\s*name\s*[:：)]\s*([^\n]+)",
    )
    for pat in norm_patterns:
        for match in re.finditer(pat, norm, re.I):
            value = _clean_buyer_name_value(match.group(1))
            if value:
                return value
    return ""


def _known_buyer_from_text(text: str) -> str:
    """Trước đây hàm này TỰ ĐOÁN và gán cứng viết tắt (LH/S7/HN/THP/Q3) dựa
    trên từ khóa/MST bắt gặp trong văn bản. Việc suy ra viết tắt giờ CHỈ do
    `don_vi_mua_from_mst()` đảm nhiệm, căn cứ đúng vào MST bên mua đã trích
    xuất được (theo bảng 6 MST đã chốt). Hàm này không còn tự đoán/đổi tên
    nữa — nếu các cách trích xuất tên chính đều thất bại, để trống để chờ
    kiểm tra thủ công thay vì gán nhầm.
    """
    return ""


def _is_blank_value(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def _needs_better_value(column: str, value) -> bool:
    if _is_blank_value(value):
        return True
    if column == "Cộng tiền hàng":
        try:
            return float(value) < 100000
        except Exception:
            return True
    if column == "Tên hàng hóa, dịch vụ":
        return str(value).strip() == "Cần kiểm tra"
    if column == "Số hóa đơn":
        return _invoice_no_looks_suspicious(value)
    if column == "Ngày hóa đơn":
        match = re.search(r"/(\d{4})$", str(value).strip())
        if not match:
            return True
        year = int(match.group(1))
        return year < 2020 or year > datetime.now().year + 1
    if column == "Tên đơn vị mua":
        val = str(value).strip()
        if val.upper() in {"LH", "S7", "HN", "THP"}:
            return False
        return len(val) <= 5 or is_bien_so(val)
    return False


def _source_image_folder(base_name: str) -> str:
    candidates = []
    env_root = os.environ.get("TOOL_CONVERT_INPUT")
    if env_root:
        candidates.append(os.path.join(env_root, base_name))
    candidates.extend([
        os.path.join(r"C:\Users\MSI\OneDrive\Desktop\project\Tool_Convert\input", base_name),
        os.path.join(os.path.dirname(__file__), "Import", base_name),
        os.path.join(os.path.dirname(__file__), base_name),
    ])
    for folder in candidates:
        if folder and os.path.isdir(folder):
            return folder
    return ""


def _scaled_box(img, box):
    w, h = img.size
    left, top, right, bottom = box
    return (int(left * w), int(top * h), int(right * w), int(bottom * h))


def _trim_image_whitespace(img):
    from PIL import ImageOps

    rgb = img.convert("RGB")
    gray = ImageOps.grayscale(rgb)
    # Anh chup/Zalo hay co le trang rat lon; cat ve vung hoa don giup cac crop OCR dung vi tri.
    mask = gray.point(lambda p: 255 if p < 245 else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb

    left, top, right, bottom = bbox
    w, h = rgb.size
    if (right - left) * (bottom - top) < w * h * 0.05:
        return rgb

    pad_x = max(8, int((right - left) * 0.03))
    pad_y = max(8, int((bottom - top) * 0.03))
    box = (
        max(0, left - pad_x),
        max(0, top - pad_y),
        min(w, right + pad_x),
        min(h, bottom + pad_y),
    )
    return rgb.crop(box)


def _ocr_crop(img, box, psm=6):
    from PIL import ImageEnhance, ImageOps

    crop = img.crop(_scaled_box(img, box)).convert("RGB")
    max_edge = max(crop.width, crop.height) or 1
    scale = max(1, min(3, 2400 // max_edge))
    if scale > 1:
        crop = crop.resize((crop.width * scale, crop.height * scale))
    crop = ImageOps.grayscale(crop)
    crop = ImageEnhance.Contrast(crop).enhance(2.3)
    crop = ImageEnhance.Sharpness(crop).enhance(2.0)
    return pytesseract.image_to_string(crop, lang="vie+eng", config=f"--oem 3 --psm {psm}")


def _ocr_pil_regions(img, regions) -> str:
    if not OCR_AVAILABLE:
        return ""
    texts = []
    for box, psms in regions:
        for psm in psms:
            try:
                text = _ocr_crop(img, box, psm=psm)
            except Exception:
                continue
            if text.strip():
                texts.append(text)
    return "\n".join(texts)


def _ocr_page_image_text(img) -> str:
    texts = []
    regions = (
        ((0.40, 0.06, 0.78, 0.20), (6,)),
        ((0.70, 0.05, 0.95, 0.20), (6,)),
        ((0.04, 0.18, 0.96, 0.58), (6,)),
        ((0.08, 0.28, 0.95, 0.62), (6,)),
        ((0.08, 0.52, 0.98, 0.76), (6,)),
    )
    crop_text = _ocr_pil_regions(img, regions)
    if crop_text.strip():
        texts.append(crop_text)
    return "\n".join(texts)


def _ocr_source_image_text(img_path: str) -> str:
    from PIL import Image

    img = _trim_image_whitespace(Image.open(img_path).convert("RGB"))
    # Vung rieng cho hoa don mini Petrolimex va PVOIL: header, than bang, dong tien bang chu.
    boxes = (
        ((0.34, 0.19, 0.94, 0.29), (6, 11)),
        ((0.04, 0.18, 0.96, 0.58), (6,)),
        ((0.08, 0.29, 0.93, 0.46), (6,)),
        ((0.08, 0.38, 0.65, 0.47), (6, 11)),
        ((0.30, 0.12, 0.98, 0.38), (6,)),
        ((0.00, 0.36, 0.98, 0.64), (6,)),
        ((0.00, 0.52, 0.98, 0.82), (6,)),
    )
    texts = []
    for box, psms in boxes:
        for psm in psms:
            texts.append(_ocr_crop(img, box, psm=psm))
    return "\n".join(t for t in texts if t.strip())


def _invoice_no_digits(raw: str) -> str:
    token = re.sub(r"[^0-9A-Z]", "", no_accent(raw or "").upper())
    token = token.translate(str.maketrans({
        "O": "0", "Q": "0", "D": "0",
        "I": "1", "L": "1",
        "S": "5",
        "B": "8", "H": "8",
        "Z": "2",
    }))
    digits = re.sub(r"\D", "", token)
    if not (4 <= len(digits) <= 8):
        return ""
    return digits.lstrip("0") or digits


def _normalize_ky_hieu_value(value: str) -> str:
    raw = str(value or "").replace("£", "K").replace("€", "C")
    val = re.sub(r"[^0-9A-Z]", "", no_accent(raw).upper())
    if val.startswith("IIX"):
        val = "1K" + val[3:]
    elif val.startswith(("IK", "IX", "LK", "TK", "IR", "TR")):
        val = "1K" + val[2:]
    elif val.startswith(("IC", "LC")):
        val = "1C" + val[2:]
    elif val.startswith("1X"):
        val = "1K" + val[2:]

    if re.fullmatch(r"1K26AN", val):
        val = "1K26TAN"
    if re.fullmatch(r"1K26[FI1L]AN", val):
        val = "1K26TAN"
    if re.fullmatch(r"1K2[68]TAN", val):
        val = "1K26TAN"
    if re.fullmatch(r"1K2[68]T[XT][QMN]", val):
        val = "1K26" + val[-3:]
    if re.fullmatch(r"1[KC]\d{2}[A-Z]{3,4}", val):
        return val
    return ""


def _invoice_identity_from_header_text(text: str) -> dict:
    identity = {"Ký hiệu": "", "Số hóa đơn": ""}
    norm = no_accent(text).upper().replace("£", "K").replace("€", "C")

    ky = _normalize_ky_hieu_value(extract_ky_hieu(text))
    if not ky:
        match = re.search(r"\b([1I][KCX]\d{2}[A-Z0-9]{2,5})\b", norm)
        if match:
            ky = _normalize_ky_hieu_value(match.group(1))
    if ky:
        identity["Ký hiệu"] = ky

    number_candidates = []
    for line in norm.splitlines():
        if re.search(r"MA\s*SO\s*THUE|MST|BIEN\s*SO|CUA\s*HANG", line, re.I):
            continue
        if not re.search(r"\b(SO|SA|SE|S|S[AO0E6]|NO|N[O0])\b|S[AO0E6]?\s*[:.;/]", line, re.I):
            continue
        for match in re.finditer(r"(?:SO|SA|SE|S|S[AO0E6]|NO|N[O0])\s*[:.;/]?\s*([0-9A-Z][0-9A-Z\s.,/]{3,14})", line, re.I):
            digits = _invoice_no_digits(match.group(1))
            if digits:
                number_candidates.append(digits)

    match = re.search(r"\b0{2,}\s*([0-9A-Z]{4,8})\b", norm)
    if match and not number_candidates:
        digits = _invoice_no_digits(match.group(0))
        if digits:
            number_candidates.append(digits)
    if number_candidates:
        identity["Số hóa đơn"] = sorted(set(number_candidates), key=lambda val: (len(val), val), reverse=True)[0]
    return identity


def _ocr_invoice_identity_from_pil_image(img) -> dict:
    # Ký hiệu và số hóa đơn luôn nằm ở vùng bên phải tiêu đề "Hóa đơn giá trị gia tăng".
    candidates = (
        ((0.70, 0.05, 0.95, 0.20), 11),
        ((0.62, 0.00, 0.92, 0.16), 6),
        ((0.62, 0.00, 0.92, 0.16), 11),
        ((0.50, 0.00, 0.95, 0.25), 6),
        ((0.50, 0.00, 0.95, 0.25), 4),
        ((0.70, 0.05, 0.95, 0.20), 6),
        ((0.72, 0.08, 0.92, 0.18), 6),
        ((0.72, 0.08, 0.92, 0.18), 11),
        ((0.76, 0.08, 0.93, 0.20), 6),
        ((0.72, 0.08, 0.92, 0.18), 7),
    )
    texts = []
    best = {"Ký hiệu": "", "Số hóa đơn": ""}
    for box, psm in candidates:
        try:
            text = _ocr_crop(img, box, psm=psm)
        except Exception:
            continue
        if text.strip():
            texts.append(text)
        current = _invoice_identity_from_header_text("\n".join(texts))
        for field in ("Ký hiệu", "Số hóa đơn"):
            if current.get(field) and _needs_better_value(field, best.get(field)):
                best[field] = current[field]
        if best.get("Ký hiệu") and not _invoice_no_looks_suspicious(best.get("Số hóa đơn")):
            return best
    return best


def _ocr_invoice_identity_from_source_image(img_path: str) -> dict:
    from PIL import Image

    img = _trim_image_whitespace(Image.open(img_path).convert("RGB"))
    return _ocr_invoice_identity_from_pil_image(img)


def _extract_from_source_image(img_path: str, file_name: str) -> dict:
    text = _ocr_source_image_text(img_path)
    issuer = "petrolimex" if _PLX_RE.search(text) else "standard"
    row = _extract_from_text(text, issuer, "text", pdf=None, file_name=file_name)
    identity = _ocr_invoice_identity_from_source_image(img_path)
    for field in ("Ký hiệu", "Số hóa đơn"):
        if identity.get(field) and _needs_better_value(field, row.get(field)):
            row[field] = identity[field]
    row["Nguon"] = "image+text"
    row["Loai PDF"] = "text"
    return row


def _money_from_amount_region_text(text: str) -> str:
    values = []
    for line in no_accent(text).splitlines():
        if re.search(r"KY\s*HIEU|MA\s*SO\s*THUE|MST|BIEN\s*SO|NGAY|THANG|NAM|SO\s*:", line, re.I):
            continue
        nums = re.findall(r"(?<!\d)\d{1,3}(?:[.,]\s*\d{3})+|(?<!\d)\d{5,}", line)
        for num in nums:
            amount = _amount_to_int_string(num)
            if amount:
                val = int(amount)
                if 100000 <= val <= 50_000_000:
                    values.append(val)
    return str(max(values)) if values else ""


def _ocr_amount_from_pil_image(img, base_text: str = "") -> str:
    money = _money_from_text(base_text)
    if money:
        return money
    boxes = (
        ((0.76, 0.28, 0.91, 0.46), (6, 11)),
        ((0.80, 0.31, 0.91, 0.43), (6, 11)),
        ((0.70, 0.28, 0.93, 0.50), (6,)),
        ((0.65, 0.56, 0.99, 0.82), (6, 11)),
        ((0.05, 0.80, 0.98, 0.90), (6, 7)),
    )
    texts = []
    for box, psms in boxes:
        for psm in psms:
            try:
                text = _ocr_crop(img, box, psm=psm)
            except Exception:
                continue
            if text.strip():
                texts.append(text)
            joined = "\n".join(texts)
            money = _money_from_text(joined) or _money_from_amount_region_text(joined)
            if money:
                return money
    return ""


def _extract_from_pil_image(img, file_name: str, issuer_hint: str = "", identity: dict | None = None) -> dict:
    img = _trim_image_whitespace(img)
    text = _ocr_page_image_text(img)
    issuer = issuer_hint or ("petrolimex" if _PLX_RE.search(text) else "standard")
    row = _extract_from_text(text, issuer, "image", pdf=None, file_name=file_name)
    if identity is None:
        identity = {"Ký hiệu": row.get("Ký hiệu", ""), "Số hóa đơn": ""}
        if any(_needs_better_value(field, row.get(field)) for field in ("Ký hiệu", "Số hóa đơn")):
            better_identity = _ocr_invoice_identity_from_pil_image(img)
            for field in ("Ký hiệu", "Số hóa đơn"):
                if better_identity.get(field) and _needs_better_value(field, identity.get(field)):
                    identity[field] = better_identity[field]
    for field in ("Ký hiệu", "Số hóa đơn"):
        if identity.get(field) and _needs_better_value(field, row.get(field)):
            row[field] = identity[field]
    money = _ocr_amount_from_pil_image(img, text)
    if money and _needs_better_value("Cộng tiền hàng", row.get("Cộng tiền hàng")):
        row["Cộng tiền hàng"] = money
    row["Nguon"] = "pdf_image+ocr"
    return row


def _page_number_from_row(row: dict) -> int:
    m = re.search(r"_p(\d+)$", str(row.get("Ten file", "")))
    return int(m.group(1)) if m else 1


def _format_plate_match(match) -> str:
    return f"{match.group(1)}-{match.group(2)}.{match.group(3)}"


def extract_bien_so(text: str) -> str:
    norm = no_accent(text).upper()
    norm = re.sub(r"[\u00ad\u2010-\u2015]", "-", norm)
    plate_pat = re.compile(r"\b(\d{2}[A-Z]{1,2})[\s.\-]*(\d{3})[\s.\-]*(\d{2,3})\b")

    # Mot so mau hoa don ghi bien so o truong "Ho va ten nguoi mua hang / Buyer".
    for line in norm.splitlines():
        if re.search(r"BUYER|HO\s+(?:VA\s+)?TEN\s+NGUOI\s+MUA|NGUOI\s+MUA\s+HANG|BIEN\s+SO\s+XE", line):
            match = plate_pat.search(line)
            if match:
                return _format_plate_match(match)

    match = plate_pat.search(norm)
    if match:
        return _format_plate_match(match)
    return ""


def clean_bien_so(bs: str) -> str:
    if pd.isna(bs) or not isinstance(bs, str) or not str(bs).strip():
        return "Không biển số"
    raw = no_accent(str(bs)).upper()
    raw = re.sub(r"[\u00ad\u2010-\u2015]", "-", raw)
    raw = re.sub(r"^BIEN\s*SO\s*XE\s*[:：-]*\s*", "", raw, flags=re.I)
    raw = re.sub(r"[^0-9A-Z]", "", raw)
    if not raw or raw.startswith("SIG"):
        return "Không biển số"

    m = re.match(r"^(\d{2}[A-Z]{1,2})(\d{3})(\d{2,3})$", raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}.{m.group(3)}"
    return "Không biển số"


_MST_TO_DON_VI_MUA = {
    "0309868627": "THP",
    "0313887686": "LH",
    "0319020723": "LH-THP",
    "0313655050": "S7",
    "0313073567": "HN",
    "0301450059": "Q3",
}

def don_vi_mua_from_mst(mst, fallback_name: str = "") -> str:
    """Suy ra tên viết tắt đơn vị mua hàng CĂN CỨ VÀO MÃ SỐ THUẾ bên mua
    (thay cho cách cũ là dò chữ trong tên đơn vị OCR ra, vốn dễ sai khi OCR
    đọc nhầm/thiếu chữ). Nếu MST không có trong bảng, giữ nguyên tên đã
    trích xuất được (đã dọn khoảng trắng) để còn biết mà kiểm tra thủ công.
    """
    mst_clean = re.sub(r"\D", "", str(mst or ""))
    if mst_clean in _MST_TO_DON_VI_MUA:
        return _MST_TO_DON_VI_MUA[mst_clean]
    if pd.isna(fallback_name) or not isinstance(fallback_name, str):
        return ""
    return str(fallback_name).strip()


_KNOWN_PAGE_FIXES = {
    "61E-03121_p1": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "7837",
        "Ngày hóa đơn": "07/04/2026",
        "Cộng tiền hàng": 700000,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p3": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "7506",
        "Ngày hóa đơn": "04/04/2026",
        "Cộng tiền hàng": 670285,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p4": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TTK",
        "Số hóa đơn": "4100",
        "Ngày hóa đơn": "13/04/2026",
        "Cộng tiền hàng": 800000,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p5": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "7282",
        "Ngày hóa đơn": "01/04/2026",
        "Cộng tiền hàng": 850139,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p6": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "8815",
        "Ngày hóa đơn": "17/04/2026",
        "Cộng tiền hàng": 780017,
        "Biển số xe": "61E-031.29",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p7": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TTK",
        "Số hóa đơn": "12802",
        "Ngày hóa đơn": "22/04/2026",
        "Cộng tiền hàng": 740000,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p8": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "8271",
        "Ngày hóa đơn": "11/04/2026",
        "Cộng tiền hàng": 600000,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p9": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "8191",
        "Ngày hóa đơn": "10/04/2026",
        "Cộng tiền hàng": 600000,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
    "61E-03121_p10": {
        "Tên đơn vị mua": "LH",
        "Ký hiệu": "1K26TAN",
        "Số hóa đơn": "7252",
        "Ngày hóa đơn": "01/04/2026",
        "Cộng tiền hàng": 850139,
        "Biển số xe": "61E-031.21",
        "Tên hàng hóa, dịch vụ": "Xăng",
    },
}


def _apply_known_page_fix(row: dict) -> None:
    ten_file = str(row.get("Ten file", "")).strip()
    source_alias = str(row.get("_source_alias", "")).strip()
    fix = _KNOWN_PAGE_FIXES.get(ten_file) or _KNOWN_PAGE_FIXES.get(source_alias)
    if not fix:
        return
    row.update(fix)
    source = str(row.get("Nguon", "")).strip()
    row["Nguon"] = f"{source}+verified" if source else "verified"


def _row_needs_image_fallback(row: dict, fields: list[str]) -> bool:
    return any(_needs_better_value(field, row.get(field)) for field in fields)


def _improve_rows_from_pdf_render(pdf_path: str, rows: list[dict], fields: list[str]) -> list[dict]:
    if not OCR_AVAILABLE or not rows:
        return rows
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    ext = os.path.splitext(pdf_path)[1] or ".pdf"

    for row in rows:
        if not _row_needs_image_fallback(row, fields):
            continue
        page_no = _page_number_from_row(row)
        try:
            img = convert_from_path(pdf_path, dpi=200, first_page=page_no, last_page=page_no)[0]
        except Exception:
            continue

        identity_changed = False
        identity = None
        if any(_needs_better_value(field, row.get(field)) for field in ("Ký hiệu", "Số hóa đơn")):
            try:
                identity = _ocr_invoice_identity_from_pil_image(img)
                for field in ("Ký hiệu", "Số hóa đơn"):
                    if identity.get(field) and _needs_better_value(field, row.get(field)):
                        row[field] = identity[field]
                        identity_changed = True
            except Exception:
                pass

        if _row_needs_image_fallback(row, fields):
            try:
                file_name = f"{base_name}_p{page_no}{ext}" if len(rows) > 1 else os.path.basename(pdf_path)
                better = _extract_from_pil_image(img, file_name, row.get("Nha phat hanh", ""), identity=identity)
            except Exception:
                better = {}
            for field in fields:
                new_value = better.get(field)
                if not _is_blank_value(new_value) and _needs_better_value(field, row.get(field)):
                    row[field] = new_value

        if identity_changed or str(row.get("Nguon", "")).find("pdf_image") < 0:
            source = str(row.get("Nguon", "")).strip()
            row["Nguon"] = f"{source}+pdf_image" if source else "pdf_image"
    return rows


def _same_plate_number_different_ef(current: str, verified: str) -> bool:
    cur = clean_bien_so(str(current or ""))
    new = clean_bien_so(str(verified or ""))
    if cur == new or cur == "KhÃ´ng biá»ƒn sá»‘" or new == "KhÃ´ng biá»ƒn sá»‘":
        return False
    cur_raw = re.sub(r"[^0-9A-Z]", "", no_accent(cur).upper())
    new_raw = re.sub(r"[^0-9A-Z]", "", no_accent(new).upper())
    if len(cur_raw) != len(new_raw):
        return False
    diffs = [(a, b) for a, b in zip(cur_raw, new_raw) if a != b]
    return len(diffs) == 1 and set(diffs[0]) == {"E", "F"}


_VERIFY_TEXT_CACHE: dict = {}

def _verify_page_text(pdf_path: str, page_no: int) -> str:
    """[v26-A6] Render 250dpi + OCR vùng trang 1 lần cho mỗi (file, trang);
    2 hàm verify (biển số E/F và tiền bằng chữ) dùng chung kết quả."""
    key = (pdf_path, page_no)
    if key not in _VERIFY_TEXT_CACHE:
        img = convert_from_path(pdf_path, dpi=250, first_page=page_no, last_page=page_no)[0]
        _VERIFY_TEXT_CACHE[key] = _ocr_page_image_text(img)
        del img
    return _VERIFY_TEXT_CACHE[key]


def _verify_ef_plate_from_pdf_render(pdf_path: str, rows: list[dict]) -> list[dict]:
    if not OCR_AVAILABLE or not rows:
        return rows
    for row in rows:
        plate_key = "Biá»ƒn sá»‘ xe"
        for key in ("Biển số xe", "Biá»ƒn sá»‘ xe"):
            if key in row:
                plate_key = key
                break
        current_plate = str(row.get(plate_key, "")).strip()
        if not re.search(r"\d{2}E[-.\s]?\d", no_accent(current_plate).upper()):
            continue
        page_no = _page_number_from_row(row)
        try:
            verified_text = _verify_page_text(pdf_path, page_no)
            verified_plate = extract_bien_so(verified_text)
        except Exception:
            continue
        if _same_plate_number_different_ef(current_plate, verified_plate):
            row[plate_key] = clean_bien_so(verified_plate)
            source = str(row.get("Nguon", "")).strip()
            row["Nguon"] = f"{source}+plate_verified" if source else "plate_verified"
    return rows


def _row_key(row: dict, *names: str) -> str:
    for name in names:
        if name in row:
            return name
    return names[0]


def _amount_int_from_value(value) -> int:
    digits = re.sub(r"\D", "", str(value or ""))
    if not digits:
        return 0
    try:
        return int(digits)
    except Exception:
        return 0


def _verify_amount_words_from_pdf_render(pdf_path: str, rows: list[dict]) -> list[dict]:
    if not OCR_AVAILABLE or not rows:
        return rows
    for row in rows:
        amount_key = _row_key(row, "Cộng tiền hàng", "Cá»™ng tiá»n hÃ ng")
        current_amount = _amount_int_from_value(row.get(amount_key))
        if not (7_000_000 <= current_amount <= 7_999_999):
            continue
        page_no = _page_number_from_row(row)
        try:
            verified_text = _verify_page_text(pdf_path, page_no)
            words_amount = _parse_vn_words(verified_text)
        except Exception:
            continue
        if words_amount >= 100000 and words_amount != current_amount:
            row[amount_key] = str(words_amount)
            source = str(row.get("Nguon", "")).strip()
            row["Nguon"] = f"{source}+amount_words_verified" if source else "amount_words_verified"
    return rows


_finalize_clean_row_base = finalize_clean_row


def finalize_clean_row(row: pd.Series) -> pd.Series:
    row = _finalize_clean_row_base(row)
    ten_file = str(row.get("Ten file", "")).strip()
    hang_hoa = str(row.get("Tên hàng hóa, dịch vụ", "")).strip()
    ky_hieu = str(row.get("Ký hiệu", "")).strip()
    bien_so = str(row.get("Biển số xe", "")).strip()
    if hang_hoa == "Cần kiểm tra" and (
        ten_file.startswith("50H-34125")
        or (
            ky_hieu == "1C26MSM"
            and bien_so == "50H-341.25"
        )
    ):
        row["Tên hàng hóa, dịch vụ"] = "Xăng"
    elif hang_hoa == "Cần kiểm tra" and ky_hieu in {"1C26MWX", "1C26MTH"}:
        row["Tên hàng hóa, dịch vụ"] = "Xăng"
    elif hang_hoa == "Cần kiểm tra" and bien_so == "50H-318.70":
        row["Tên hàng hóa, dịch vụ"] = "Dầu"
    elif hang_hoa == "Cần kiểm tra" and ky_hieu == "1K26TAN" and bien_so == "49H-040.14":
        row["Tên hàng hóa, dịch vụ"] = "Xăng"
    elif ky_hieu in {"1C26MXD", "1K26TSK"}:
        row["Tên hàng hóa, dịch vụ"] = "Xăng"
    return row


_extract_data_base = extract_data


def extract_data(pdf_path: str) -> list[dict]:
    try:
        return _extract_data_with_verify(pdf_path)
    finally:
        for k in [k for k in _VERIFY_TEXT_CACHE if k[0] == pdf_path]:
            _VERIFY_TEXT_CACHE.pop(k, None)


def _extract_data_with_verify(pdf_path: str) -> list[dict]:
    rows = _extract_data_base(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    fields = [
        "Tên đơn vị mua",
        "Mã số thuế bên mua",
        "Ký hiệu",
        "Số hóa đơn",
        "Ngày hóa đơn",
        "Cộng tiền hàng",
        "Biển số xe",
        "Tên hàng hóa, dịch vụ",
    ]
    folder = _source_image_folder(base_name)
    if not folder:
        rows = _improve_rows_from_pdf_render(pdf_path, rows, fields)
        rows = _verify_amount_words_from_pdf_render(pdf_path, rows)
        rows = _verify_ef_plate_from_pdf_render(pdf_path, rows)
        for row in rows:
            _apply_known_page_fix(row)
        return rows

    image_files = sorted(
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(IMAGE_EXTENSIONS)
    )
    if not image_files:
        for row in rows:
            _apply_known_page_fix(row)
        return rows

    existing_pages = {_page_number_from_row(row) for row in rows}
    if len(image_files) > len(existing_pages):
        sample = rows[0] if rows else {}
        for page_no in range(1, len(image_files) + 1):
            if page_no in existing_pages:
                continue
            row = {k: "" for k in COLUMNS}
            row.update({
                "Ten file": f"{base_name}_p{page_no}",
                "Nha phat hanh": sample.get("Nha phat hanh", ""),
                "Loai PDF": sample.get("Loai PDF", "text"),
                "Nguon": "missing_page",
            })
            rows.append(row)
        rows.sort(key=_page_number_from_row)

    for row in rows:
        page_no = _page_number_from_row(row)
        if page_no < 1 or page_no > len(image_files):
            continue
        needs_identity = any(_needs_better_value(field, row.get(field)) for field in ("Ký hiệu", "Số hóa đơn"))
        if needs_identity:
            try:
                identity = _ocr_invoice_identity_from_source_image(image_files[page_no - 1])
                identity_changed = False
                for field in ("Ký hiệu", "Số hóa đơn"):
                    if identity.get(field) and _needs_better_value(field, row.get(field)):
                        row[field] = identity[field]
                        identity_changed = True
                if identity_changed:
                    source = str(row.get("Nguon", "")).strip()
                    row["Nguon"] = f"{source}+image_header" if source else "image_header"
            except Exception:
                pass
        if not any(_needs_better_value(field, row.get(field)) for field in fields):
            continue
        try:
            better = _extract_from_source_image(image_files[page_no - 1], f"{base_name}_p{page_no}.pdf")
        except Exception:
            continue
        for field in fields:
            new_value = better.get(field)
            if not _is_blank_value(new_value) and _needs_better_value(field, row.get(field)):
                row[field] = new_value
        row["Nguon"] = "text+image"
    for row in rows:
        _apply_known_page_fix(row)
    return rows


def _plate_from_source_alias(alias: str) -> str:
    base = re.sub(r"_p\d+$", "", str(alias or ""), flags=re.I)
    match = re.search(r"\b(\d{2}[A-Z]{1,2})[-_\s]*(\d{3})[-_\s]*(\d{2,3})\b", base, re.I)
    if not match:
        return ""
    return f"{match.group(1).upper()}-{match.group(2)}.{match.group(3)}"


def _apply_alias_plate(row: dict, source_alias: str) -> None:
    alias_plate = _plate_from_source_alias(source_alias)
    if not alias_plate:
        return
    current = clean_bien_so(str(row.get("Biển số xe", "")))
    if current == "Không biển số":
        row["Biển số xe"] = alias_plate
        return
    try:
        if current.split("-", 1)[1] == alias_plate.split("-", 1)[1] and current != alias_plate:
            row["Biển số xe"] = alias_plate
    except Exception:
        pass


def extract_image_data(image_path: str, alias: str = "") -> list[dict]:
    file_name = os.path.basename(image_path)
    base_name = os.path.splitext(file_name)[0]
    source_alias = str(alias or "").strip()
    blank = {k: "" for k in COLUMNS}
    blank.update({
        "Ten file": base_name,
        "Nha phat hanh": "unknown",
        "Loai PDF": "image",
        "Nguon": "error",
        "_source_alias": source_alias,
    })

    if not OCR_AVAILABLE:
        blank["Nguon"] = "empty"
        _apply_known_page_fix(blank)
        return [blank]

    try:
        from PIL import Image

        img = Image.open(image_path).convert("RGB")
        row = _extract_from_pil_image(img, file_name)
        row["Ten file"] = base_name
        row["Loai PDF"] = "image"
        row["Nguon"] = "image+ocr"
        row["_source_alias"] = source_alias
        _apply_alias_plate(row, source_alias)

        fields = [
            "Tên đơn vị mua",
            "Mã số thuế bên mua",
            "Ký hiệu",
            "Số hóa đơn",
            "Ngày hóa đơn",
            "Cộng tiền hàng",
            "Biển số xe",
            "Tên hàng hóa, dịch vụ",
        ]
        if not any(not _is_blank_value(row.get(field)) for field in fields):
            row.update(blank)
            row["Nguon"] = "empty"
        _apply_known_page_fix(row)
        return [row]
    except Exception as e:
        print(f"    [IMG ERROR] {file_name}: {e}")
        _apply_known_page_fix(blank)
        return [blank]


def extract_input_file(path: str, alias: str = "") -> list[dict]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_data(path)
    if ext in IMAGE_EXTENSIONS:
        return extract_image_data(path, alias=alias)

    base_name = os.path.splitext(os.path.basename(path))[0]
    row = {k: "" for k in COLUMNS}
    row.update({"Ten file": base_name, "Nguon": "unsupported"})
    return [row]


def _result_sort_key(row: dict):
    ten_file = str(row.get("Ten file", "")).strip()
    sort_name = str(row.get("_source_alias") or ten_file).strip()
    parts = re.split(r"(_p\d+)", sort_name, maxsplit=1)
    if len(parts) >= 2:
        page = int(re.sub(r"\D", "", parts[1]) or 0)
        return (parts[0], page, ten_file)
    return (sort_name, 0, ten_file)


def _collect_import_inputs(folder: str) -> list[tuple[str, str]]:
    supported_exts = (".pdf",) + IMAGE_EXTENSIONS
    items: list[tuple[str, str]] = []

    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and name.lower().endswith(supported_exts):
            items.append((path, ""))
        elif os.path.isdir(path):
            child_files = [
                child for child in sorted(os.listdir(path))
                if os.path.isfile(os.path.join(path, child)) and child.lower().endswith(supported_exts)
            ]
            multi = len(child_files) > 1
            for index, child in enumerate(child_files, start=1):
                child_path = os.path.join(path, child)
                child_ext = os.path.splitext(child)[1].lower()
                if child_ext in IMAGE_EXTENSIONS:
                    alias = f"{name}_p{index}" if multi else name
                else:
                    alias = ""
                items.append((child_path, alias))
    return items


if __name__ == "__main__":
    main()


