# -*- coding: utf-8 -*-
"""
쿠팡이츠 월별 결제 내역 CSV 자동화

내가 지정한 연/월에 대해, Gmail로 온 쿠팡이츠 결제 안내 메일을 찾아서
CSV 파일로 정리한다. 쿠팡이츠 결제는 아래 3곳의 PG(결제 대행)사 중 한
곳을 통해 메일이 온다.

  1. NHN KCP        (pgadmcust@kcp.co.kr)
     - 메일 본문의 "구매상점명"이 '쿠팡이츠' 또는 '쿠팡이츠(레거시)'인 것
  2. 나이스페이먼츠  (nice_customer@nicepg.co.kr)
     - 메일 본문의 "주문번호"가 'ROCKET_PAY_DELIVERY_'로 시작하는 것
  3. 이지페이(KICC)  (easypay_noreturn@easypay.co.kr)
     - 메일에 담긴 "English Receipt" 페이지의 주문번호가
       'ROCKET_PAY_DELIVERY_'로 시작하는 것

이 프로그램은 내 Gmail 계정에서만(읽기 전용) 메일을 읽어오고, 이지페이의
경우 메일 속 영수증 페이지를 열어 주문번호만 추가로 확인한다. 확인한
내용은 모두 이 컴퓨터 안에서만 처리되며, 결과는 output 폴더의 CSV
파일로만 저장된다. 외부로 전송되거나 별도로 저장되지 않는다.

사용법:
    run.bat 을 더블클릭하거나, 터미널에서 `python coupang_eats_csv.py` 실행
"""

import base64
import csv
import html as html_lib
import os
import re
import sys
from datetime import datetime, timedelta

import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- 경로 설정 -----------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# 지메일 읽기 전용 권한만 요청 (메일 삭제/발송 등은 하지 않음)
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# CSV에 실제로 저장할 컬럼 순서
CSV_COLUMNS = ["구분", "결제일시", "PG사", "상점명", "상품명", "결제금액", "주문번호", "승인번호"]

# 쿠팡이츠 결제가 오는 3곳의 PG사 메일 발신자 주소
KCP_SENDER = "pgadmcust@kcp.co.kr"
NICEPG_SENDER = "nice_customer@nicepg.co.kr"
EASYPAY_SENDER = "easypay_noreturn@easypay.co.kr"

PG_NAME = {
    "kcp": "NHN KCP",
    "nicepg": "나이스페이먼츠",
    "easypay": "이지페이",
}

# 쿠팡이츠 주문번호는 항상 이 문자열로 시작한다 (나이스페이먼츠 / 이지페이 판별 기준)
DELIVERY_ORDER_PREFIX = "ROCKET_PAY_DELIVERY_"

# --- 이지페이 전용 제외 기준 ---
# 이지페이는 "English Receipt" 영수증 페이지에서 얻은 주문번호로 쿠팡이츠 여부를 가려낸다.
# ROCKET_PAY_DELIVERY_ 로 시작하지 않는 나머지 주문번호 중에서도,
#   - "ROCKET_PAY_" 로 시작하지만 "_DELIVERY_"가 아닌 것 (쿠팡 로켓배송 등)
#   - "311" 로 시작하는 것 (쿠팡 일반 상품 주문)
# 은 쿠팡이츠가 아니므로 제외한다. 그 외 주문번호(1023로 시작하는 쿠팡이츠 주문 등)는 포함한다.
EASYPAY_GENERAL_ROCKET_PREFIX = "ROCKET_PAY_"
EASYPAY_EXCLUDE_ORDER_PREFIX = "311"

# 이지페이 영수증 페이지 요청 시 사용할 타임아웃(초)
EASYPAY_RECEIPT_TIMEOUT = 15


# --- 1. 구글 로그인 / 인증 ------------------------------------------------
def get_gmail_service():
    """최초 1회만 브라우저 로그인이 뜨고, 이후에는 저장된 로그인 정보를 재사용한다."""
    if not os.path.exists(CREDENTIALS_PATH):
        print(f"[오류] '{os.path.basename(CREDENTIALS_PATH)}' 파일을 찾을 수 없습니다.")
        print("SETUP.md 안내를 참고해서 구글 클라우드 콘솔에서 인증 파일을 받아")
        print(f"이 폴더({BASE_DIR})에 넣어주세요.")
        sys.exit(1)

    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


# --- 2. 사용자에게 연/월 물어보기 -----------------------------------------
def ask_year_month():
    now = datetime.now()
    while True:
        raw_year = input(f"조회할 연도를 입력하세요 (예: {now.year}, 그냥 엔터=올해): ").strip()
        year_str = str(now.year) if raw_year == "" else raw_year

        raw_month = input("조회할 월을 입력하세요 (1~12): ").strip()

        if not (year_str.isdigit() and raw_month.isdigit()):
            print("→ 숫자로 입력해주세요. 다시 입력해주세요.\n")
            continue

        year = int(year_str)
        month = int(raw_month)

        if not (2000 <= year <= 2100) or not (1 <= month <= 12):
            print("→ 연도 또는 월 범위가 올바르지 않습니다. 다시 입력해주세요.\n")
            continue

        return year, month


# --- 3. Gmail 검색 --------------------------------------------------------
def month_range(year, month):
    start = datetime(year, month, 1)
    end = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
    return start, end


def build_query(sender, start, end, extra=""):
    # 시간대 경계에서 메일이 하루 밀리는 경우를 대비해 앞뒤로 하루씩 여유를 두고 검색한다.
    # 실제 대상 월인지는 메일 본문의 결제일시를 파싱해 다시 한번 정확히 걸러낸다.
    q_start = (start - timedelta(days=1)).strftime("%Y/%m/%d")
    q_end = (end + timedelta(days=1)).strftime("%Y/%m/%d")
    query = f"from:{sender} after:{q_start} before:{q_end}"
    if extra:
        query += f" {extra}"
    return query


def list_message_ids(service, query):
    ids = []
    page_token = None
    while True:
        resp = (
            service.users()
            .messages()
            .list(userId="me", q=query, pageToken=page_token, maxResults=100)
            .execute()
        )
        ids.extend(m["id"] for m in resp.get("messages", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return ids


def _decode_body_data(data):
    raw = base64.urlsafe_b64decode(data.encode("ASCII") + b"===")
    for encoding in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _find_part(payload, mime_type):
    if payload.get("mimeType") == mime_type and "data" in payload.get("body", {}):
        return _decode_body_data(payload["body"]["data"])
    for part in payload.get("parts") or []:
        found = _find_part(part, mime_type)
        if found is not None:
            return found
    return None


def get_message_body(service, msg_id):
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    payload = msg.get("payload", {})
    body = _find_part(payload, "text/html")
    if body is None:
        body = _find_part(payload, "text/plain")
    return body or ""


# --- 4. 메일 본문에서 항목 추출 -------------------------------------------
def extract_field(body, label):
    """<td>라벨</td><td>값</td> (또는 <th>라벨</th><td>값</td>) 형태의 표에서
    라벨 바로 다음 칸의 값을 뽑아낸다."""
    pattern = re.compile(rf">{re.escape(label)}</t[dh]>\s*<td[^>]*>(.*?)</td>", re.DOTALL)
    m = pattern.search(body)
    if not m:
        return ""
    text = re.sub(r"<[^>]+>", "", m.group(1))
    text = html_lib.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def parse_payment_datetime(text):
    """PG사마다 결제일시 표기 형식이 달라서 여러 형식을 순서대로 시도한다."""
    patterns = [
        r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(\d{1,2})시\s*(\d{1,2})분",  # KCP
        r"(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2}):(\d{1,2})",  # 나이스페이먼츠
        r"(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?",  # 이지페이
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if not m:
            continue
        parts = [int(g) if g else 0 for g in m.groups()]
        while len(parts) < 6:
            parts.append(0)
        y, mo, d, h, mi, s = parts[:6]
        try:
            return datetime(y, mo, d, h, mi, s)
        except ValueError:
            continue
    return None


def clean_amount(text):
    """'14,900 원', '32,200 원 (일시불)', '21,900원' 등을 '14,900원' 형태로 통일한다."""
    m = re.search(r"([\d][\d,]*)\s*원", text)
    if m:
        return f"{m.group(1)}원"
    return text.strip()


# --- 4-1. NHN KCP 메일 파싱 ------------------------------------------------
# KCP는 결제 메일과 취소 메일의 항목 이름(라벨)이 동일하고("결제일시", "결제금액" 등
# 그대로 쓰임), 취소 메일에만 "취소요청일시"가 추가로 붙는다. 이 필드의 유무로
# 결제/취소 메일을 구분한다. 부분취소인 경우 실제 취소 금액은 "부분취소금액"에 담긴다.
KCP_FIELDS = [
    "결제일시", "결제금액", "승인번호", "주문번호", "구매상점명", "주문상품명",
    "취소요청일시", "부분취소금액",
]
KCP_STORE_NAMES = {"쿠팡이츠", "쿠팡이츠(레거시)"}


def parse_kcp_message(body, msg_id, target_year, target_month, warn):
    values = {label: extract_field(body, label) for label in KCP_FIELDS}

    is_cancel = bool(values["취소요청일시"])
    kind = "취소" if is_cancel else "결제"
    dt_text = values["취소요청일시"] if is_cancel else values["결제일시"]

    dt = parse_payment_datetime(dt_text)
    if dt is None:
        warn(f"[KCP] 메일(id={msg_id})에서 {kind}일시를 해석하지 못해 건너뜁니다.")
        return None
    if not (dt.year == target_year and dt.month == target_month):
        return None

    if values["구매상점명"] not in KCP_STORE_NAMES:
        # 쿠팡이츠가 아닌 다른 KCP 결제 메일 (해당사항 없음, 조용히 제외)
        return None

    amount_text = (values["부분취소금액"] if is_cancel else "") or values["결제금액"]
    if not amount_text:
        warn(f"[KCP] 메일(id={msg_id})에서 {kind}금액을 찾지 못해 건너뜁니다.")
        return None

    return {
        "_dt": dt,
        "구분": kind,
        "결제일시": dt.strftime("%Y-%m-%d %H:%M"),
        "PG사": PG_NAME["kcp"],
        "상점명": values["구매상점명"],
        "상품명": values["주문상품명"],
        "결제금액": clean_amount(amount_text),
        "주문번호": values["주문번호"],
        "승인번호": values["승인번호"],
    }


# --- 4-2. 나이스페이먼츠 메일 파싱 -----------------------------------------
# 나이스페이먼츠도 취소 메일에 "취소요청일시"가 추가로 붙는다. 결제금액 라벨은
# 취소 메일에서는 "결제 금액"(띄어쓰기 있음)으로 바뀌어 못 찾으므로, 취소 메일에만
# 있는 "취소금액"(붙여쓰기) 라벨을 대신 사용한다.
NICEPG_FIELDS = [
    "결제일시", "결제금액", "승인번호", "주문번호", "상점명", "상품명",
    "취소요청일시", "취소금액",
]


def parse_nicepg_message(body, msg_id, target_year, target_month, warn):
    values = {label: extract_field(body, label) for label in NICEPG_FIELDS}

    is_cancel = bool(values["취소요청일시"])
    kind = "취소" if is_cancel else "결제"
    dt_text = values["취소요청일시"] if is_cancel else values["결제일시"]

    dt = parse_payment_datetime(dt_text)
    if dt is None:
        warn(f"[나이스페이먼츠] 메일(id={msg_id})에서 {kind}일시를 해석하지 못해 건너뜁니다.")
        return None
    if not (dt.year == target_year and dt.month == target_month):
        return None

    if not values["주문번호"].startswith(DELIVERY_ORDER_PREFIX):
        # 쿠팡 로켓배송 등 다른 주문 (쿠팡이츠 아님, 조용히 제외)
        return None

    amount_text = (values["취소금액"] if is_cancel else "") or values["결제금액"]
    if not amount_text:
        warn(f"[나이스페이먼츠] 메일(id={msg_id})에서 {kind}금액을 찾지 못해 건너뜁니다.")
        return None

    return {
        "_dt": dt,
        "구분": kind,
        "결제일시": dt.strftime("%Y-%m-%d %H:%M"),
        "PG사": PG_NAME["nicepg"],
        "상점명": values["상점명"],
        "상품명": values["상품명"],
        "결제금액": clean_amount(amount_text),
        "주문번호": values["주문번호"],
        "승인번호": values["승인번호"],
    }


# --- 4-3. 이지페이 메일 파싱 -----------------------------------------------
# 이지페이 메일은 결제금액 라벨이 "상품금액"으로 표기된다.
# 취소 메일은 "결제일시" 대신 "취소일시"가 쓰이고, 이번 취소분 금액은 "취소금액"에 담긴다.
EASYPAY_FIELDS = ["결제일시", "상품금액", "승인번호", "상호", "상품명", "취소일시", "취소금액"]


def _extract_easypay_control_no(body):
    """메일 본문의 'English Receipt' 링크에서 controlNo 값을 뽑아낸다."""
    m = re.search(r"controlNo=(\d+)[^\"'\s]*language_type=ENG", body)
    if m:
        return m.group(1)
    m = re.search(r"controlNo=(\d+)", body)
    return m.group(1) if m else None


def fetch_easypay_order_number(control_no):
    """이지페이 영수증 페이지(공개 페이지, 로그인 불필요)를 열어 주문번호만 읽어온다.
    메일 속 'English Receipt' 버튼을 누르는 것과 동일한 요청이며, 결과는 이 함수를
    호출한 곳에서만 사용하고 별도로 저장하지 않는다."""
    try:
        resp = requests.post(
            "https://office.easypay.co.kr/mcht/receipt/CardReceiptAction.do",
            data={
                "s_method": "",
                "controlNo": control_no,
                "controlWay": "",
                "tax_cd": "null",
                "language_type": "ENG",
            },
            timeout=EASYPAY_RECEIPT_TIMEOUT,
        )
        resp.encoding = resp.apparent_encoding or "utf-8"
        page = resp.text
    except requests.RequestException:
        return None

    m = re.search(r"Order Number</th>\s*<td[^>]*>(.*?)</td>", page, re.DOTALL)
    if not m:
        return None
    value = re.sub(r"<[^>]+>", "", m.group(1))
    return html_lib.unescape(value).strip()


def parse_easypay_message(body, msg_id, target_year, target_month, warn):
    values = {label: extract_field(body, label) for label in EASYPAY_FIELDS}

    is_cancel = bool(values["취소일시"])
    kind = "취소" if is_cancel else "결제"
    dt_text = values["취소일시"] if is_cancel else values["결제일시"]

    dt = parse_payment_datetime(dt_text)
    if dt is None:
        warn(f"[이지페이] 메일(id={msg_id})에서 {kind}일시를 해석하지 못해 건너뜁니다.")
        return None
    if not (dt.year == target_year and dt.month == target_month):
        return None

    control_no = _extract_easypay_control_no(body)
    if not control_no:
        warn(f"[이지페이] 메일(id={msg_id})에서 영수증 링크를 찾지 못해 건너뜁니다.")
        return None

    order_no = fetch_easypay_order_number(control_no)
    if not order_no:
        warn(f"[이지페이] 메일(id={msg_id})의 영수증 페이지에서 주문번호를 읽지 못해 건너뜁니다.")
        return None

    is_general_rocket_order = order_no.startswith(
        EASYPAY_GENERAL_ROCKET_PREFIX
    ) and not order_no.startswith(DELIVERY_ORDER_PREFIX)
    is_general_goods_order = order_no.startswith(EASYPAY_EXCLUDE_ORDER_PREFIX)
    if is_general_rocket_order or is_general_goods_order:
        # 쿠팡 로켓배송/일반 상품 주문 (쿠팡이츠 아님, 조용히 제외)
        return None

    amount_text = (values["취소금액"] if is_cancel else "") or values["상품금액"]
    if not amount_text:
        warn(f"[이지페이] 메일(id={msg_id})에서 {kind}금액을 찾지 못해 건너뜁니다.")
        return None

    return {
        "_dt": dt,
        "구분": kind,
        "결제일시": dt.strftime("%Y-%m-%d %H:%M"),
        "PG사": PG_NAME["easypay"],
        "상점명": values["상호"],
        "상품명": values["상품명"],
        "결제금액": clean_amount(amount_text),
        "주문번호": order_no,
        "승인번호": values["승인번호"],
    }


def _as_excel_text(value):
    """엑셀이 긴 숫자 문자열(주문번호 등)을 1.02311E+18 같은 지수 표기로 자동
    변환하지 못하도록, ="값" 형태의 텍스트 수식으로 감싼다. 메모장 등에서
    열면 ="..." 그대로 보이지만, 엑셀에서는 원래 숫자가 그대로 표시된다."""
    escaped = str(value).replace('"', '""')
    return f'="{escaped}"'


# --- 5. CSV 저장 -----------------------------------------------------------
def save_csv(rows, year, month):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"{year % 100:02d}{month:02d}00_쿠팡이츠결제내역_claudecli.csv"
    filepath = os.path.join(OUTPUT_DIR, filename)

    # UTF-8 BOM(utf-8-sig)으로 저장해야 엑셀에서 열었을 때 한글이 깨지지 않는다.
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for row in rows:
            out = {col: row[col] for col in CSV_COLUMNS}
            if out.get("주문번호"):
                out["주문번호"] = _as_excel_text(out["주문번호"])
            writer.writerow(out)

    return filepath


# --- 메인 ------------------------------------------------------------------
def main():
    print("=== 쿠팡이츠 월별 결제 내역 CSV 만들기 ===\n")

    year, month = ask_year_month()
    print(f"\n{year}년 {month}월 결제 내역을 조회합니다. 잠시만 기다려주세요...\n")

    service = get_gmail_service()
    start, end = month_range(year, month)

    warnings = []

    def warn(message):
        warnings.append(message)
        print(f"  [건너뜀] {message}")

    rows = []

    # 1) NHN KCP
    print("[1/3] NHN KCP 메일 확인 중...")
    kcp_ids = list_message_ids(service, build_query(KCP_SENDER, start, end, extra="subject:쿠팡이츠"))
    print(f"  대상 메일 {len(kcp_ids)}건")
    for msg_id in kcp_ids:
        body = get_message_body(service, msg_id)
        if not body:
            warn(f"[KCP] 메일(id={msg_id})의 내용을 읽을 수 없어 건너뜁니다.")
            continue
        row = parse_kcp_message(body, msg_id, year, month, warn)
        if row:
            rows.append(row)

    # 2) 나이스페이먼츠
    print("[2/3] 나이스페이먼츠 메일 확인 중...")
    nicepg_ids = list_message_ids(service, build_query(NICEPG_SENDER, start, end))
    print(f"  대상 메일 {len(nicepg_ids)}건")
    for msg_id in nicepg_ids:
        body = get_message_body(service, msg_id)
        if not body:
            warn(f"[나이스페이먼츠] 메일(id={msg_id})의 내용을 읽을 수 없어 건너뜁니다.")
            continue
        row = parse_nicepg_message(body, msg_id, year, month, warn)
        if row:
            rows.append(row)

    # 3) 이지페이
    print("[3/3] 이지페이 메일 확인 중... (건별로 영수증 페이지를 추가로 확인해서 다소 시간이 걸릴 수 있습니다)")
    easypay_ids = list_message_ids(service, build_query(EASYPAY_SENDER, start, end))
    print(f"  대상 메일 {len(easypay_ids)}건")
    for msg_id in easypay_ids:
        body = get_message_body(service, msg_id)
        if not body:
            warn(f"[이지페이] 메일(id={msg_id})의 내용을 읽을 수 없어 건너뜁니다.")
            continue
        row = parse_easypay_message(body, msg_id, year, month, warn)
        if row:
            rows.append(row)

    rows.sort(key=lambda r: r["_dt"])

    filepath = save_csv(rows, year, month)

    print()
    if rows:
        print(f"총 {len(rows)}건을 정리해서 저장했습니다.")
    else:
        print(f"{year}년 {month}월에는 쿠팡이츠 결제 내역이 없습니다. (항목만 있는 빈 CSV를 만들었습니다)")
    print(f"저장 위치: {filepath}")

    if warnings:
        print(f"\n※ {len(warnings)}건은 형식을 해석하지 못해 건너뛰었습니다. 위 [건너뜀] 내용을 확인해주세요.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n취소되었습니다.")
    except PermissionError as e:
        print(f"\n[오류] 파일에 저장하지 못했습니다: {e.filename}")
        print("→ 이 CSV 파일을 엑셀 등 다른 프로그램에서 열어놓은 상태라서 저장이 막혔을 가능성이 매우 높습니다.")
        print("   해당 파일을 닫은 뒤, 프로그램을 다시 실행해주세요.")
    except Exception as e:  # noqa: BLE001 - 비개발자용 프로그램이므로 원인을 그대로 보여준다
        print(f"\n[오류] 문제가 발생했습니다: {e}")
    finally:
        input("\n엔터 키를 누르면 창이 닫힙니다...")
