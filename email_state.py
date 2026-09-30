"""
Sổ theo dõi email đã quét / đã xử lý (SQLite) - mỗi profile một file .db.

Trạng thái (status):
  DONE     : đã tải đủ file, các lần chạy sau bỏ qua.
  FAILED   : tải lỗi, lần sau tự thử lại cho tới khi hết max_attempts.
  SKIPPED  : đã quét nhưng không có PDF / supplier chưa hỗ trợ / không có link.
  PENDING  : được đánh dấu thủ công để xử lý lại ở lần chạy sau.
  (không có trong sổ) : chưa quét.
"""
import csv
import json
import os
import sqlite3
import threading
from datetime import datetime

DONE = "DONE"
FAILED = "FAILED"
SKIPPED = "SKIPPED"
PENDING = "PENDING"
ALL_STATUSES = (DONE, FAILED, SKIPPED, PENDING)

GMAIL_LINK = "https://mail.google.com/mail/u/0/#all/{}"

CSV_COLUMNS = [
    "updated_at", "status", "supplier", "mail_date", "subject", "reason",
    "attempts", "saved_count", "files", "gmail_link", "message_id",
]


# ---------- Chuẩn hóa & so khớp hóa đơn Petro (dùng cho lọc trùng) ----------
def norm_code(value):
    """Mã tra cứu: bỏ khoảng trắng và dấu * cuối, không phân biệt hoa/thường."""
    return str(value if value is not None else "").strip().rstrip("*").strip().upper()

def norm_so(value):
    """Số hóa đơn: bỏ số 0 đầu (Excel hay biến '00012345' thành 12345)."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value if value is not None else "").strip().lstrip("0").upper()

def same_so(a, b):
    """Hai số hóa đơn được coi là khớp nếu bằng nhau, hoặc một bên không có dữ liệu."""
    return (not a) or (not b) or a == b

def in_master(master_keys, code_n, so_n):
    """master_keys: {mã tra cứu chuẩn hóa: {các số hóa đơn chuẩn hóa}}"""
    return bool(master_keys) and code_n in master_keys and any(same_so(so_n, m) for m in master_keys[code_n])


def gmail_link(message_id):
    return GMAIL_LINK.format(message_id)


def _now():
    return datetime.now().isoformat(timespec="seconds")


class EmailState:
    def __init__(self, db_path, max_attempts=3):
        self.db_path = db_path
        self.max_attempts = max_attempts
        folder = os.path.dirname(db_path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA synchronous=NORMAL")
        self._init_schema()

    # ---------- schema ----------
    def _init_schema(self):
        with self._lock:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS emails (
                    message_id  TEXT PRIMARY KEY,
                    status      TEXT NOT NULL,
                    supplier    TEXT,
                    subject     TEXT,
                    mail_date   TEXT,
                    reason      TEXT,
                    attempts    INTEGER NOT NULL DEFAULT 0,
                    saved_count INTEGER NOT NULL DEFAULT 0,
                    files       TEXT,
                    first_seen  TEXT,
                    updated_at  TEXT
                )
                """
            )
            self.conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_status ON emails(status)")
            # Dữ liệu hóa đơn Petrolimex (Số hóa đơn / Mã tra cứu) - dùng để dựng lại file Excel đầy đủ
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS petro_invoices (
                    message_id  TEXT PRIMARY KEY,
                    mail_date   TEXT,
                    header_date TEXT,
                    sender      TEXT,
                    subject     TEXT,
                    so_hoa_don  TEXT,
                    ma_tra_cuu  TEXT,
                    updated_at  TEXT
                )
                """
            )
            # Nâng cấp DB cũ: thêm cột theo dõi việc tải PDF Petro (mặc định = chưa tải)
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(petro_invoices)").fetchall()}
            for name, ddl in (
                ("pdf_status", "TEXT DEFAULT 'PENDING'"),
                ("pdf_file", "TEXT"),
                ("pdf_error", "TEXT"),
                ("pdf_attempts", "INTEGER DEFAULT 0"),
                ("pdf_updated_at", "TEXT"),
                ("dup_of", "TEXT"),                        # trùng với email nào / 'MASTER' (file tổng)
                ("master_synced", "INTEGER DEFAULT 0"),    # đã ghi vào file tổng chưa
                ("force_pdf", "INTEGER DEFAULT 0"),        # người dùng chủ động yêu cầu tải lại -> bỏ qua lọc trùng
            ):
                if name not in cols:
                    self.conn.execute(f"ALTER TABLE petro_invoices ADD COLUMN {name} {ddl}")
            self.conn.commit()

    def close(self):
        with self._lock:
            try:
                self.conn.commit()
            finally:
                self.conn.close()

    def commit(self):
        with self._lock:
            self.conn.commit()

    # ---------- ghi ----------
    def record(self, message_id, status, supplier=None, saved_count=0, files=None,
               reason=None, subject=None, mail_date=None, commit=True):
        """Ghi kết quả xử lý một email. Mỗi lần gọi tăng `attempts` thêm 1."""
        now = _now()
        with self._lock:
            old = self.conn.execute(
                "SELECT attempts, first_seen, subject, mail_date, supplier, files, saved_count "
                "FROM emails WHERE message_id=?",
                (message_id,),
            ).fetchone()
            attempts = (old["attempts"] if old else 0) + 1
            first_seen = old["first_seen"] if old else now
            subject = subject or (old["subject"] if old else None)
            mail_date = mail_date or (old["mail_date"] if old else None)
            supplier = supplier or (old["supplier"] if old else None)
            # files=None -> giữ nguyên danh sách file đã lưu trước đó (vd. PDF Petro đã tải)
            if files is None and old:
                files_json = old["files"] or "[]"
                saved_count = old["saved_count"]
            else:
                files_json = json.dumps(files or [], ensure_ascii=False)
            self.conn.execute(
                """
                INSERT OR REPLACE INTO emails
                (message_id, status, supplier, subject, mail_date, reason, attempts,
                 saved_count, files, first_seen, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    message_id, status, supplier, subject, mail_date, reason, attempts,
                    int(saved_count or 0), files_json,
                    first_seen, now,
                ),
            )
            if commit:
                self.conn.commit()

    def mark(self, message_id, status):
        """Đổi trạng thái thủ công (DONE / PENDING). Tạo dòng mới nếu chưa có."""
        now = _now()
        reason = "MANUAL_DONE" if status == DONE else "MANUAL_PENDING"
        with self._lock:
            old = self.conn.execute(
                "SELECT message_id FROM emails WHERE message_id=?", (message_id,)
            ).fetchone()
            if old:
                if status == PENDING:
                    self.conn.execute(
                        "UPDATE emails SET status=?, reason=?, attempts=0, updated_at=? WHERE message_id=?",
                        (status, reason, now, message_id),
                    )
                else:
                    self.conn.execute(
                        "UPDATE emails SET status=?, reason=?, updated_at=? WHERE message_id=?",
                        (status, reason, now, message_id),
                    )
            else:
                self.conn.execute(
                    """
                    INSERT INTO emails (message_id, status, reason, attempts, saved_count, files, first_seen, updated_at)
                    VALUES (?,?,?,0,0,'[]',?,?)
                    """,
                    (message_id, status, reason, now, now),
                )
            if status == PENDING:
                # Xử lý lại email Petro -> đồng thời cho phép tải lại PDF của nó
                self.conn.execute(
                    "UPDATE petro_invoices SET pdf_status='PENDING', pdf_file=NULL, pdf_error=NULL, "
                    "dup_of=NULL, master_synced=0, force_pdf=1 WHERE message_id=?",
                    (message_id,),
                )
            self.conn.commit()
            return bool(old)

    def seed(self, message_ids, reason):
        """Đánh dấu hàng loạt là DONE (bỏ qua id đã có trong sổ). Trả về số dòng mới thêm."""
        now = _now()
        with self._lock:
            before = self.conn.total_changes
            self.conn.executemany(
                """
                INSERT OR IGNORE INTO emails
                (message_id, status, reason, attempts, saved_count, files, first_seen, updated_at)
                VALUES (?,?,?,0,0,'[]',?,?)
                """,
                [(mid, DONE, reason, now, now) for mid in message_ids],
            )
            self.conn.commit()
            return self.conn.total_changes - before

    # ---------- đọc ----------
    def partition_ids(self, retry_failed=False, reprocess_skipped=False):
        """
        Chia các email trong sổ thành 2 nhóm:
          skip    : không cần đụng tới nữa (DONE, SKIPPED, FAILED đã hết lượt thử)
          include : cần xử lý lại (PENDING, FAILED còn lượt thử, ...) - kể cả khi mail đã rời Inbox
        """
        skip, include = set(), set()
        with self._lock:
            rows = self.conn.execute("SELECT message_id, status, attempts FROM emails").fetchall()
        for r in rows:
            mid, status, attempts = r["message_id"], r["status"], r["attempts"]
            if status == DONE:
                skip.add(mid)
            elif status == SKIPPED:
                (include if reprocess_skipped else skip).add(mid)
            elif status == FAILED:
                if attempts < self.max_attempts or retry_failed:
                    include.add(mid)
                else:
                    skip.add(mid)
            else:  # PENDING
                include.add(mid)
        return skip, include

    def counts(self):
        with self._lock:
            rows = self.conn.execute("SELECT status, COUNT(*) AS n FROM emails GROUP BY status").fetchall()
        return {r["status"]: r["n"] for r in rows}

    def rows(self, status=None, limit=None):
        sql = "SELECT * FROM emails"
        params = []
        if status:
            sql += " WHERE status=?"
            params.append(status)
        sql += " ORDER BY updated_at DESC, message_id"
        if limit:
            sql += " LIMIT ?"
            params.append(int(limit))
        with self._lock:
            return [dict(r) for r in self.conn.execute(sql, params).fetchall()]

    def reason_breakdown(self, status):
        with self._lock:
            rows = self.conn.execute(
                "SELECT COALESCE(reason,'') AS reason, COUNT(*) AS n FROM emails WHERE status=? "
                "GROUP BY reason ORDER BY n DESC",
                (status,),
            ).fetchall()
        return [(r["reason"], r["n"]) for r in rows]

    @staticmethod
    def files_of(row):
        try:
            return json.loads(row.get("files") or "[]")
        except Exception:
            return []

    # ---------- Petrolimex ----------
    def save_petro(self, message_id, mail_date, header_date, sender, subject,
                   so_hoa_don, ma_tra_cuu, commit=True):
        """Lưu/cập nhật dữ liệu hóa đơn Petro. Giữ nguyên trạng thái tải PDF, trừ khi mã tra cứu đổi."""
        with self._lock:
            self.conn.execute(
                """
                INSERT INTO petro_invoices
                (message_id, mail_date, header_date, sender, subject, so_hoa_don, ma_tra_cuu,
                 updated_at, pdf_status, pdf_attempts)
                VALUES (?,?,?,?,?,?,?,?, 'PENDING', 0)
                ON CONFLICT(message_id) DO UPDATE SET
                    mail_date   = COALESCE(excluded.mail_date, petro_invoices.mail_date),
                    header_date = excluded.header_date,
                    sender      = excluded.sender,
                    subject     = excluded.subject,
                    so_hoa_don  = excluded.so_hoa_don,
                    ma_tra_cuu  = excluded.ma_tra_cuu,
                    updated_at  = excluded.updated_at,
                    pdf_status  = CASE WHEN petro_invoices.ma_tra_cuu <> excluded.ma_tra_cuu
                                       THEN 'PENDING' ELSE petro_invoices.pdf_status END,
                    pdf_file    = CASE WHEN petro_invoices.ma_tra_cuu <> excluded.ma_tra_cuu
                                       THEN NULL ELSE petro_invoices.pdf_file END
                """,
                (message_id, mail_date, header_date, sender, subject,
                 so_hoa_don or "", ma_tra_cuu or "", _now()),
            )
            if commit:
                self.conn.commit()

    def petro_rows(self, include_duplicates=True):
        """Hóa đơn Petro trong sổ, mới nhất trước. include_duplicates=False: bỏ các email trùng."""
        sql = "SELECT * FROM petro_invoices"
        if not include_duplicates:
            sql += " WHERE COALESCE(pdf_status,'PENDING') <> 'DUPLICATE'"
        sql += " ORDER BY COALESCE(mail_date,'') DESC, message_id"
        with self._lock:
            rows = self.conn.execute(sql).fetchall()
        return [dict(r) for r in rows]

    def petro_duplicate_count(self):
        with self._lock:
            return self.conn.execute(
                "SELECT COUNT(*) FROM petro_invoices WHERE pdf_status='DUPLICATE'"
            ).fetchone()[0]

    def petro_count(self):
        with self._lock:
            return self.conn.execute("SELECT COUNT(*) FROM petro_invoices").fetchone()[0]

    def petro_pending_downloads(self, limit=None):
        """Hóa đơn Petro có mã tra cứu nhưng chưa tải PDF (gồm cả lần trước tải lỗi)."""
        sql = (
            "SELECT * FROM petro_invoices WHERE ma_tra_cuu <> '' "
            "AND COALESCE(pdf_status,'PENDING') NOT IN ('DONE','DUPLICATE') "
            "ORDER BY COALESCE(mail_date,'') DESC, message_id"
        )
        params = []
        if limit:
            sql += " LIMIT ?"
            params.append(int(limit))
        with self._lock:
            return [dict(r) for r in self.conn.execute(sql, params).fetchall()]

    def set_petro_pdf(self, message_id, status, pdf_file=None, error=None):
        """Ghi kết quả tải PDF Petro. Khi DONE, ghi luôn tên file vào sổ email (cột files trong CSV)."""
        with self._lock:
            self.conn.execute(
                """
                UPDATE petro_invoices
                SET pdf_status=?, pdf_file=?, pdf_error=?,
                    pdf_attempts=COALESCE(pdf_attempts,0)+1, pdf_updated_at=?,
                    force_pdf=CASE WHEN ?='DONE' THEN 0 ELSE force_pdf END
                WHERE message_id=?
                """,
                (status, pdf_file, error, _now(), status, message_id),
            )
            if status == DONE and pdf_file:
                self.conn.execute(
                    "UPDATE emails SET files=?, saved_count=1 WHERE message_id=?",
                    (json.dumps([pdf_file], ensure_ascii=False), message_id),
                )
            self.conn.commit()

    def petro_mark_duplicates(self, master_keys=None):
        """
        Lọc trùng hóa đơn Petro theo (mã tra cứu + số hóa đơn). Chạy lại nhiều lần vẫn cho cùng kết quả.
          - Trùng = cùng mã tra cứu (không phân biệt hoa/thường, bỏ dấu *) VÀ số hóa đơn khớp
            (nếu cả hai bên đều có số hóa đơn thì phải giống nhau).
          - Trùng với hóa đơn đã có trong FILE TỔNG (master_keys), hoặc với email khác trong sổ
            (ưu tiên giữ email đã tải xong, sau đó là email cũ nhất) -> đánh dấu DUPLICATE, không tải.
          - Chỉ đánh dấu các hóa đơn CHƯA tải; hóa đơn đã tải xong (DONE) không bị đụng tới.
          - Hóa đơn bạn chủ động yêu cầu tải lại (--mark-pending / --verify --fix) không bị coi là trùng.
          - Bản trùng không còn trùng nữa (vd. đã xoá khỏi file tổng) -> trả về trạng thái chờ tải.
        Trả về {"total": số email trùng hiện có, "new": số email mới bị đánh dấu trùng}.
        """
        master_keys = master_keys or {}
        with self._lock:
            rows = [dict(r) for r in self.conn.execute(
                "SELECT message_id, mail_date, so_hoa_don, ma_tra_cuu, pdf_status, dup_of, "
                "COALESCE(force_pdf,0) AS force_pdf FROM petro_invoices WHERE ma_tra_cuu <> ''"
            ).fetchall()]

        def order(r):
            done = (r["pdf_status"] or "PENDING") == DONE
            return (0 if done else 1, r["mail_date"] or "", r["message_id"])

        rows.sort(key=order)
        reps, result = {}, {}
        for r in rows:
            code_n, so_n = norm_code(r["ma_tra_cuu"]), norm_so(r["so_hoa_don"])
            status = r["pdf_status"] or "PENDING"
            dup_of = None
            if status != DONE and not r["force_pdf"]:      # hóa đơn bị ép tải lại thì không lọc trùng
                if in_master(master_keys, code_n, so_n):
                    dup_of = "MASTER"
                else:
                    for rep_so, rep_id in reps.get(code_n, []):
                        if same_so(so_n, rep_so):
                            dup_of = rep_id
                            break
            if dup_of is None:
                reps.setdefault(code_n, []).append((so_n, r["message_id"]))
            result[r["message_id"]] = dup_of

        new = 0
        with self._lock:
            for r in rows:
                mid, dup_of, status = r["message_id"], result[r["message_id"]], (r["pdf_status"] or "PENDING")
                if dup_of is not None:
                    if status != "DUPLICATE" or r["dup_of"] != dup_of:
                        self.conn.execute(
                            "UPDATE petro_invoices SET pdf_status='DUPLICATE', dup_of=?, pdf_file=NULL, pdf_error=NULL "
                            "WHERE message_id=?", (dup_of, mid))
                        if status != "DUPLICATE":
                            new += 1
                    where = "file tổng" if dup_of == "MASTER" else f"email {dup_of}"
                    self.conn.execute(
                        "UPDATE emails SET reason=? WHERE message_id=?",
                        (f"PETRO_DUPLICATE: trùng mã tra cứu/số hóa đơn với {where} - không tải lại", mid))
                elif status == "DUPLICATE":
                    self.conn.execute(
                        "UPDATE petro_invoices SET pdf_status='PENDING', dup_of=NULL WHERE message_id=?", (mid,))
                    self.conn.execute(
                        "UPDATE emails SET reason='PETRO_EXTRACTED' WHERE message_id=?", (mid,))
            self.conn.commit()
        return {"total": self.petro_duplicate_count(), "new": new}

    def petro_unsynced_master_rows(self):
        """Hóa đơn đã tải xong nhưng chưa được ghi vào file tổng (cũ nhất trước)."""
        with self._lock:
            rows = self.conn.execute(
                "SELECT * FROM petro_invoices WHERE COALESCE(pdf_status,'PENDING')='DONE' "
                "AND COALESCE(master_synced,0)=0 AND ma_tra_cuu <> '' "
                "ORDER BY COALESCE(mail_date,'') ASC, message_id"
            ).fetchall()
        return [dict(r) for r in rows]

    def set_master_synced(self, message_ids):
        with self._lock:
            self.conn.executemany(
                "UPDATE petro_invoices SET master_synced=1 WHERE message_id=?", [(m,) for m in message_ids])
            self.conn.commit()

    def petro_pdf_counts(self):
        with self._lock:
            rows = self.conn.execute(
                """
                SELECT CASE WHEN ma_tra_cuu = '' THEN 'NO_CODE'
                            ELSE COALESCE(pdf_status,'PENDING') END AS s, COUNT(*) AS n
                FROM petro_invoices GROUP BY s
                """
            ).fetchall()
        return {r["s"]: r["n"] for r in rows}

    def mark_petro_downloaded(self, keys=None):
        """
        Đánh dấu hóa đơn Petro là ĐÃ TẢI PDF (bạn đã tải tay / bằng down_petro.py trước đây).
        keys: danh sách message_id hoặc mã tra cứu (có hay không có dấu * cuối đều được).
        Không truyền keys = đánh dấu TẤT CẢ hóa đơn đang chờ tải.
        """
        keys = [k.strip() for k in (keys or []) if k and k.strip()]
        params = [_now()]
        where = "1=1"
        if keys:
            norm = [k.rstrip("*") for k in keys]
            where = (f"(message_id IN ({','.join('?' * len(keys))}) "
                     f"OR RTRIM(ma_tra_cuu,'*') IN ({','.join('?' * len(norm))}))")
            params += keys + norm
        with self._lock:
            cur = self.conn.execute(
                "UPDATE petro_invoices SET pdf_status='DONE', pdf_error='MANUAL', pdf_updated_at=? "
                "WHERE ma_tra_cuu <> '' AND COALESCE(pdf_status,'PENDING') NOT IN ('DONE','DUPLICATE') AND " + where,
                params,
            )
            self.conn.commit()
            return cur.rowcount

    # ---------- kiểm tra chéo với thư mục tải ----------
    def verify(self, download_dir):
        """
        Với mỗi email DONE có ghi tên file: kiểm tra file còn trong download_dir không.
        Trả về (số email đã kiểm tra, số email DONE không có thông tin file, danh sách dòng bị thiếu file).
        """
        checked, no_info, problems = 0, 0, []
        for row in self.rows(DONE):
            files = self.files_of(row)
            if not files:
                no_info += 1
                continue
            checked += 1
            missing = [f for f in files if not os.path.exists(os.path.join(download_dir, f))]
            if missing:
                row["missing"] = missing
                problems.append(row)
        return checked, no_info, problems

    # ---------- xuất CSV để mở bằng Excel ----------
    def export_csv(self, csv_path):
        folder = os.path.dirname(csv_path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        rows = self.rows()
        # utf-8-sig: Excel đọc tiếng Việt không bị lỗi font
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(CSV_COLUMNS)
            for r in rows:
                writer.writerow([
                    r["updated_at"], r["status"], r["supplier"] or "", r["mail_date"] or "",
                    r["subject"] or "", r["reason"] or "", r["attempts"], r["saved_count"],
                    "; ".join(self.files_of(r)), gmail_link(r["message_id"]), r["message_id"],
                ])
        return len(rows)
