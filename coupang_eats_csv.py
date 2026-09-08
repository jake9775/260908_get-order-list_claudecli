# -*- coding: utf-8 -*-
"""
쿠팡이츠 월별 결제 내역 CSV 자동화

내가 지정한 연/월에 대해, Gmail로 온 쿠팡이츠 결제 안내 메일
("NHN KCP - 쿠팡이츠의 결제 내역입니다")을 찾아서 CSV 파일로 정리한다.

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

# 쿠팡이츠 결제 안내 메일에서 뽑아낼 항목 (메일 본문에 나오는 라벨 그대로)
FIELDS = ["결제일시", "결제금액", "카드종류", "승인번호", "주문번호", "주문상품명"]

# CSV에 실제로 저장할 컬럼 순서
CSV_COLUMNS = ["결제일시", "상품명", "결제금액", "카드종류", "주문번호", "승인번호"]


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


def build_query(start, end):
    # 시간대 경계에서 메일이 하루 밀리는 경우를 대비해 앞뒤로 하루씩 여유를 두고 검색한다.
    # 실제 대상 월인지는 메일 본문의 결제일시를 파싱해 다시 한번 정확히 걸러낸다.
    q_start = (start - timedelta(days=1)).strftime("%Y/%m/%d")
    q_end = (end + timedelta(days=1)).strftime("%Y/%m/%d")
    return f"from:pgadmcust@kcp.co.kr subject:쿠팡이츠 after:{q_start} before:{q_end}"


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
    """<td>라벨</td><td>값</td> 형태의 표에서 라벨 바로 다음 칸의 값을 뽑아낸다."""
    pattern = re.compile(rf">{re.escape(label)}</td>\s*<td[^>]*>(.*?)</td>", re.DOTALL)
    m = pattern.search(body)
    if not m:
        return ""
    text = re.sub(r"<[^>]+>", "", m.group(1))
    text = html_lib.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def parse_payment_datetime(text):
    m = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(\d{1,2})시\s*(\d{1,2})분", text)
    if not m:
        return None
    y, mo, d, h, mi = (int(x) for x in m.groups())
    try:
        return datetime(y, mo, d, h, mi)
    except ValueError:
        return None


def parse_message(service, msg_id, target_year, target_month, warn):
    body = get_message_body(service, msg_id)
    if not body:
        warn(f"메일(id={msg_id})의 내용을 읽을 수 없어 건너뜁니다.")
        return None

    values = {label: extract_field(body, label) for label in FIELDS}

    dt = parse_payment_datetime(values["결제일시"])
    if dt is None:
        warn(
            f"메일(id={msg_id})에서 결제일시를 해석하지 못해 건너뜁니다. "
            "(쿠팡/KCP 메일 형식이 바뀌었을 수 있습니다)"
        )
        return None

    if not (dt.year == target_year and dt.month == target_month):
        # 검색 시 하루씩 여유를 뒀기 때문에, 대상 월이 아닌 메일은 조용히 제외한다.
        return None

    if not values["결제금액"]:
        warn(f"메일(id={msg_id})에서 결제금액을 찾지 못해 건너뜁니다.")
        return None

    return {
        "_dt": dt,
        "결제일시": values["결제일시"],
        "상품명": values["주문상품명"],
        "결제금액": values["결제금액"],
        "카드종류": values["카드종류"],
        "주문번호": values["주문번호"],
        "승인번호": values["승인번호"],
    }


# --- 5. CSV 저장 -----------------------------------------------------------
def save_csv(rows, year, month):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"쿠팡이츠_{year}년{month:02d}월.csv"
    filepath = os.path.join(OUTPUT_DIR, filename)

    # UTF-8 BOM(utf-8-sig)으로 저장해야 엑셀에서 열었을 때 한글이 깨지지 않는다.
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row[col] for col in CSV_COLUMNS})

    return filepath


# --- 메인 ------------------------------------------------------------------
def main():
    print("=== 쿠팡이츠 월별 결제 내역 CSV 만들기 ===\n")

    year, month = ask_year_month()
    print(f"\n{year}년 {month}월 결제 내역을 조회합니다. 잠시만 기다려주세요...\n")

    service = get_gmail_service()

    start, end = month_range(year, month)
    query = build_query(start, end)
    ids = list_message_ids(service, query)
    print(f"관련 메일 {len(ids)}건을 찾았습니다. 내용을 확인하는 중...\n")

    warnings = []

    def warn(message):
        warnings.append(message)
        print(f"  [건너뜀] {message}")

    rows = []
    for msg_id in ids:
        row = parse_message(service, msg_id, year, month, warn)
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
    except Exception as e:  # noqa: BLE001 - 비개발자용 프로그램이므로 원인을 그대로 보여준다
        print(f"\n[오류] 문제가 발생했습니다: {e}")
    finally:
        input("\n엔터 키를 누르면 창이 닫힙니다...")
