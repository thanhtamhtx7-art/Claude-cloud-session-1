from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.header import decode_header
from datetime import datetime
import os.path
import os
import base64
import re
from collections import Counter
import concurrent.futures
import threading

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# Gmail giới hạn tốc độ gọi API (lỗi 429 / 403 rateLimitExceeded) và thỉnh thoảng lỗi 5xx / rớt mạng.
# execute(num_retries=...) tự thử lại các lỗi này, chờ tăng dần (1s, 2s, 4s...) giữa các lần.
GMAIL_RETRIES = 5
SCAN_WORKERS = 8          # số luồng đọc email song song (15 luồng trước đây dễ vượt giới hạn tốc độ)

# Fix #1: Vẫn giữ thread_local nhưng cache theo token để tránh stale credentials
thread_local = threading.local()


def get_creds(token_path="token.json", credentials_path="credentials.json"):
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES
            )
            creds = flow.run_local_server(port=0)
        token_dir = os.path.dirname(token_path)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        with open(token_path, "w") as token:
            token.write(creds.to_json())
    return creds


# Fix #1: Cache service theo token thay vì cache mãi mãi
# Nếu creds được refresh (token đổi), service sẽ được tạo lại tự động
def get_local_service(creds):
    current_token = creds.token
    if (
        not hasattr(thread_local, "service")
        or not hasattr(thread_local, "token")
        or thread_local.token != current_token
    ):
        thread_local.service = build("gmail", "v1", credentials=creds)
        thread_local.token = current_token
    return thread_local.service


def clean_text(text):
    urls = re.findall(r'https?://\S+', text)
    for u in urls:
        if "gov.vn" in u:
            text = text.replace(u, "")
    return text


def normalize_domain(domain):
    parts = domain.split(".")
    if len(parts) >= 3:
        return ".".join(parts[-3:])
    return domain


# Dấu hiệu tên miền của các nhà cung cấp hóa đơn mà main.py tải được (giống thứ tự kiểm tra trong main.py)
SUPPORTED_SUPPLIER_HINTS = ("easyinvoice", "meinvoice.vn", "pvoil", "smartsign.com.vn", "fast",
                            "ehoadon.vn", "vnpt-invoice.com.vn", "vnpt")


def is_supported_supplier(domain):
    d = (domain or "").lower()
    return any(h in d for h in SUPPORTED_SUPPLIER_HINTS)


def extract_supplier(text):
    """Tên miền nhà cung cấp hóa đơn trong nội dung email.
    Ưu tiên tên miền của nhà cung cấp đã hỗ trợ (EasyInvoice, VNPT...) dù nó không đứng đầu:
    email hay có website của bên bán đứng trước link hóa đơn. Không có -> tên miền .vn đầu tiên như cũ."""
    domains = re.findall(r"[a-zA-Z0-9.-]+\.(?:com\.vn|vn)", text)
    normalized = [normalize_domain(d) for d in domains]
    for d in normalized:
        if is_supported_supplier(d):
            return d
    return normalized[0] if normalized else None


def _decode_mime_header(value):
    """Giải mã tiêu đề mail dạng =?UTF-8?B?...?= thành chuỗi đọc được."""
    if not value:
        return ""
    try:
        out = ""
        for part, enc in decode_header(value):
            if isinstance(part, bytes):
                out += part.decode(enc or "utf-8", errors="ignore")
            else:
                out += part
        return out.strip()
    except Exception:
        return value.strip()


def attachment_parts(payload):
    """Danh sách file đính kèm của email (phần lá có tên file hoặc là PDF), chỉ giữ thông tin cần để tải:
    filename, mimeType, body={attachmentId, size}. Rất nhẹ (vài trăm byte/email) -> main.py tải PDF/XML
    bằng attachments.get() luôn, không phải gọi messages.get(format="full") lần nữa."""
    out = []

    def walk(p):
        if p.get("parts"):
            for c in p["parts"]:
                walk(c)
            return
        fname = p.get("filename") or ""
        mime = p.get("mimeType") or ""
        if not fname and mime != "application/pdf":
            return
        body = p.get("body") or {}
        slim = {"size": body.get("size", 0)}
        if body.get("attachmentId"):
            slim["attachmentId"] = body["attachmentId"]
        elif body.get("data"):
            slim["data"] = body["data"]      # file nhỏ nằm sẵn trong thư (hiếm)
        out.append({"filename": fname, "mimeType": mime, "body": slim})

    walk(payload or {})
    return out


def get_message_content_and_pdf(service, msg_id):
    """Trả về (nội dung text, có PDF đính kèm không, meta={subject, mail_date, sender, attachments})."""
    msg = service.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute(num_retries=GMAIL_RETRIES)

    text = ""
    has_pdf = False

    def extract_parts(parts):
        nonlocal text, has_pdf
        for part in parts:
            if part.get("parts"):
                extract_parts(part["parts"])

            # Nhận PDF theo kiểu khai báo HOẶC đuôi .pdf (giống is_pdf_part của main.py).
            # Vd. Bkav gửi PDF + XML đính kèm nhưng khai báo kiểu chung (application/octet-stream).
            if part.get("mimeType") == "application/pdf" or (part.get("filename") or "").lower().endswith(".pdf"):
                has_pdf = True

            if part.get("body", {}).get("data"):
                raw = base64.urlsafe_b64decode(
                    part["body"]["data"]
                ).decode("utf-8", errors="ignore")
                text += raw + "\n"

    if msg["payload"].get("parts"):
        extract_parts(msg["payload"]["parts"])
    else:
        data = msg["payload"]["body"].get("data")
        if data:
            text = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")

    # Lấy thêm tiêu đề + ngày từ chính lần gọi này (không tốn thêm request)
    subject, sender = "", ""
    for h in msg["payload"].get("headers", []):
        name = h.get("name", "").lower()
        if name == "subject":
            subject = _decode_mime_header(h.get("value", ""))
        elif name == "from":
            sender = _decode_mime_header(h.get("value", ""))

    mail_date = ""
    try:
        mail_date = datetime.fromtimestamp(int(msg["internalDate"]) / 1000).strftime("%Y-%m-%d %H:%M")
    except Exception:
        pass

    meta = {"subject": subject, "mail_date": mail_date, "sender": sender,
            "attachments": attachment_parts(msg["payload"])}
    return clean_text(text), has_pdf, meta


def get_label_id(service, label_name):
    results = service.users().labels().list(userId="me").execute(num_retries=GMAIL_RETRIES)
    labels = results.get("labels", [])
    for label in labels:
        if label["name"].lower() == label_name.lower():
            return label["id"]
    return None


def list_all_messages(service, label_id=None, query=None):
    """
    Liệt kê ID email (chỉ ID, rất nhẹ).
      - label_id: lọc theo label (chế độ cũ)
      - query   : Gmail search query, ví dụ "in:inbox after:1756684800"
    """
    messages = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "maxResults": 500, "pageToken": page_token}
        if label_id:
            kwargs["labelIds"] = [label_id]
        if query:
            kwargs["q"] = query
        results = service.users().messages().list(**kwargs).execute(num_retries=GMAIL_RETRIES)
        messages.extend(results.get("messages", []))

        page_token = results.get("nextPageToken")
        if not page_token:
            break

    return messages


# Fix #2: Hàm core dùng chung, thay thế 2 hàm worker bị trùng lặp
# include_content=True dùng cho run_scan(), False dùng cho main()
def _process_message(creds, msg_id, include_content=False, keywords=None):
    """
    Xử lý một email: lấy content, kiểm tra PDF, trích xuất supplier.
    Thành công: trả về dict kết quả.
    Lỗi: trả về {"id": ..., "error": "..."} để phía gọi ghi nhận và thử lại lần sau.
    """
    try:
        local_service = get_local_service(creds)
        content, has_pdf, meta = get_message_content_and_pdf(local_service, msg_id)
        supplier = extract_supplier(content) if not has_pdf else None

        # keyword_hit: người gửi / tiêu đề / nội dung có chứa một trong các từ khóa (vd. "petrolimex")
        keyword_hit = False
        if keywords:
            haystack = f"{meta.get('sender', '')} {meta.get('subject', '')} {content}".lower()
            keyword_hit = any(k.lower() in haystack for k in keywords if k)

        result = {
            "id": msg_id,
            "has_pdf": has_pdf,
            "supplier": supplier,
            "subject": meta.get("subject", ""),
            "mail_date": meta.get("mail_date", ""),
            "sender": meta.get("sender", ""),
            "keyword_hit": keyword_hit,
            # Danh sách file đính kèm (đã rút gọn) -> main.py tải PDF/XML không cần đọc lại email
            "attachments": meta.get("attachments", []),
        }

        # Chỉ giữ content khi cần thiết (tránh RAM leak).
        # main.py chỉ dùng content để tìm link hóa đơn -> chỉ cần với mail có supplier.
        # Quét cả Inbox có thể lên tới hàng nghìn mail nên bỏ content của mail không dùng tới.
        if include_content and supplier:
            result["content"] = content

        return result

    except Exception as e:
        print(f"Lỗi khi đọc email {msg_id}: {e}")
        return {"id": msg_id, "error": str(e)}


def _scan_ids(creds, ids, include_content, keywords=None):
    """Đọc song song danh sách ID. Trả về (kết quả thành công, danh sách lỗi)."""
    ok, errors = [], []
    total = len(ids)
    if not total:
        return ok, errors

    done = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=SCAN_WORKERS) as executor:
        futures = {
            executor.submit(_process_message, creds, mid, include_content, keywords): mid
            for mid in ids
        }
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result is None or "error" in result:
                errors.append(result or {"id": futures[future], "error": "UNKNOWN"})
            else:
                ok.append(result)
            done += 1
            if total >= 100 and done % 100 == 0:
                print(f"  ... đã đọc {done}/{total} email")
    return ok, errors


def main(label_name="xd", token_path="token.json", credentials_path="credentials.json"):
    creds = get_creds(token_path=token_path, credentials_path=credentials_path)
    service = build("gmail", "v1", credentials=creds)

    label_id = get_label_id(service, label_name)
    if not label_id:
        print(f"Không tìm thấy label '{label_name}'")
        return

    print("Đang lấy danh sách ID email...")
    messages = list_all_messages(service, label_id=label_id)
    if not messages:
        print("Không có email nào trong nhãn này.")
        return

    suppliers = []
    pdf_count = 0

    print(f"Bắt đầu quét song song {len(messages)} email...")

    ok, _errors = _scan_ids(creds, [m["id"] for m in messages], include_content=False)
    for result in ok:
        if result["has_pdf"]:
            pdf_count += 1
        elif result["supplier"]:
            suppliers.append(result["supplier"])

    counter = Counter(suppliers)

    print(f"Tổng số email đã quét: {len(messages)}")
    print(f"Số email có file PDF đính kèm: {pdf_count}")
    print("Các nhà cung cấp hóa đơn tìm thấy:")
    for supplier, count in counter.items():
        print(f"- {supplier}: {count} email")


def run_scan(label_name="xd", token_path="token.json", credentials_path="credentials.json",
             query=None, skip_ids=None, include_ids=None, keywords=None, extra_query=None):
    """
    Quét email và trả về nội dung + thông tin PDF/supplier.

    Nguồn email:
      - query=None      : quét theo label `label_name` (chế độ cũ)
      - query="in:inbox": quét theo Gmail search query (vd. Inbox)
      - extra_query     : điều kiện lọc thêm (vd. khoảng ngày) khi quét theo label

    Tránh quét lại:
      - skip_ids   : ID đã xử lý xong -> KHÔNG tải nội dung nữa (tiết kiệm request Gmail)
      - include_ids: ID cần xử lý lại (FAILED/PENDING...) -> luôn quét, kể cả khi
                     mail đã không còn nằm trong kết quả liệt kê (vd. đã archive khỏi Inbox)

    Nhận diện theo từ khóa:
      - keywords: danh sách từ khóa; mỗi mail có thêm cờ `keyword_hit` (True nếu người gửi,
                  tiêu đề hoặc nội dung chứa từ khóa) - dùng để nhận diện mail Petrolimex.
    """
    creds = get_creds(token_path=token_path, credentials_path=credentials_path)
    service = build("gmail", "v1", credentials=creds)

    if query is None:
        label_id = get_label_id(service, label_name)
        if not label_id:
            print(f"[WARN] Không tìm thấy label Gmail '{label_name}' -> không có email nào để quét.")
            return {"total": 0, "listed": 0, "skipped_known": 0, "pdf": 0,
                    "suppliers": Counter(), "data": [], "read_errors": []}
        messages = list_all_messages(service, label_id=label_id, query=extra_query)
    else:
        messages = list_all_messages(service, query=query)

    listed_ids = [m["id"] for m in messages]
    skip_ids = set(skip_ids or [])
    include_ids = set(include_ids or [])

    ids_to_scan = [mid for mid in listed_ids if mid not in skip_ids]
    already = set(ids_to_scan)
    ids_to_scan.extend(mid for mid in include_ids if mid not in already)

    ok, errors = _scan_ids(creds, ids_to_scan, include_content=True, keywords=keywords)

    suppliers = []
    pdf_count = 0
    for result in ok:
        if result["has_pdf"]:
            pdf_count += 1
        if result["supplier"]:
            suppliers.append(result["supplier"])

    return {
        "total": len(ids_to_scan),
        "listed": len(listed_ids),
        "skipped_known": len([mid for mid in listed_ids if mid in skip_ids]),
        "pdf": pdf_count,
        "suppliers": Counter(suppliers),
        "data": ok,
        "read_errors": errors,
    }


if __name__ == "__main__":
    main()
