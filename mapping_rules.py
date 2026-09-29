# -*- coding: utf-8 -*-
"""
상품명 매핑규칙 (rules/memo_mapping_구글플레이.csv, rules/memo_mapping_쿠팡이츠.csv).

결과 엑셀에서 **상품명이 빈칸인 행**에만, 규칙 파일의 상품명을 채운다.
(구글 결제 센터 메일처럼 본문에 상품명이 없는 결제에 이름을 붙이기 위한 용도)

파일 형식 (CSV, 첫 줄은 칸 이름):
    구글플레이: PG사,상점명,결제금액,구분,주문번호,결제방법,상품명
    쿠팡이츠:   PG사,상점명,결제금액,구분,주문번호,승인번호,상품명

  - 한 줄이 규칙 하나. 값이 적힌 조건 칸은 **전부** 만족해야 상품명이 채워진다.
  - 빈칸인 조건 칸은 무시한다. (조건이 하나도 없는 줄은 모든 행에 붙어버리므로 건너뜀)
  - 위에서부터 검사해서 **처음 맞는 규칙 하나만** 적용한다.
  - 상품명이 이미 있는 행은 건드리지 않는다. (공백만 들어 있어도 빈칸으로 본다)

조건 칸에 쓰는 법은 가계부 프로젝트(260902_autobudget_claudecli)와 같다:
  - 글자 조건      Google Cloud        그냥 적으면 '포함' (이 프로젝트는 모든 글자 칸이 '포함')
  - 둘 다 만족     YouTube AND 멤버십
  - 하나라도 만족   YouTube OR Google
  - 제외          NOT 취소
  - 묶기          (A OR B) AND NOT C   괄호 안에 AND나 OR가 있을 때만 '묶기'로 계산한다.
  - 큰따옴표       "Google Play"       글자가 완전히 똑같아야만 매칭 ('Google Payments'와는 매칭 안 됨)
  - 숫자 조건 (결제금액)   10000, >=10000, <5000, <>0, >=1000 AND <=5000, 7890 OR 9900

  * AND / OR / NOT 은 반드시 **대문자**로 쓴다.
  * 글자 조건은 **대소문자를 구분한다** (Google과 google은 다른 글자).

조건 해석 코드는 vendor_memo_rules.py 를 그대로 쓴다. 그 파일은 가계부 프로젝트의
scripts/memo_rules.py 를 2026-09-29에 한 글자도 바꾸지 않고 복사한 것이다.
(가계부 쪽 문법이 바뀌면 파일째 다시 복사한다)
"""
import csv
import io
from pathlib import Path

from vendor_memo_rules import _eval, _parse, _read_text


def ensure_template(path, header):
    """규칙 파일이 없으면 칸 이름 줄만 있는 빈 파일을 만든다. 실패해도 프로그램은 계속한다."""
    path = Path(path)
    if path.exists():
        return
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes((",".join(header) + "\n").encode("utf-8-sig"))
        print(f"  규칙 파일이 없어 새로 만들었습니다: {path}")
    except OSError as e:
        print(f"  [경고] 규칙 파일을 만들지 못했습니다: {e}")


def load(path, condition_columns, number_columns, result_column="상품명", label="상품명 규칙"):
    """[(줄번호, [(칸이름, 조건), ...], 결과값), ...] 를 돌려준다.
    문법이 틀린 줄은 경고하고 건너뛴다. 파일을 읽을 수 없으면 경고하고 규칙 없이 진행한다."""
    path = Path(path)
    if not path.exists():
        return []
    try:
        reader = csv.DictReader(io.StringIO(_read_text(path)))
        # 엑셀에서 헤더 칸에 실수로 공백이 섞여도 그 칸을 찾을 수 있도록 앞뒤 공백을 무시한다.
        if reader.fieldnames:
            reader.fieldnames = [(name or "").strip() for name in reader.fieldnames]
        known = set(condition_columns) | {result_column}
        for name in reader.fieldnames or []:
            if name and name not in known:
                print(f"  [경고] {label} 파일의 '{name}' 칸은 알 수 없는 칸이라 무시합니다. "
                      f"(쓸 수 있는 칸: {', '.join(condition_columns)}, {result_column})")
        items = list(reader)
    except (ValueError, OSError, csv.Error) as e:
        print(f"  [경고] {label} 파일을 읽지 못해 규칙 없이 진행합니다: {e} (UTF-8 CSV로 저장해주세요)")
        return []

    rules = []
    for line_no, item in enumerate(items, start=2):
        result = (item.get(result_column) or "").strip()
        conds = []
        col = ""
        try:
            for col in condition_columns:
                cell = (item.get(col) or "").strip()
                if cell:
                    conds.append((col, _parse(cell, col in number_columns)))
        except ValueError as e:
            print(f"  [경고] {label} {line_no}째 줄 '{col}' 칸을 건너뜁니다: {e}")
            continue
        if not conds and not result:
            continue  # 완전히 빈 줄
        if not result:
            print(f"  [경고] {label} {line_no}째 줄은 {result_column}이 비어 있어 건너뜁니다.")
            continue
        if not conds:
            print(f"  [경고] {label} {line_no}째 줄은 조건이 하나도 없어 건너뜁니다. "
                  f"(모든 행에 같은 {result_column}이 붙는 것을 막기 위함)")
            continue
        rules.append((line_no, conds, result))
    return rules


def _cell_value(row, col, number_columns):
    if col in number_columns:
        return int(row.get(col) or 0)
    return str(row.get(col) or "")


def apply(rows, rules, number_columns, exact_columns=(), result_column="상품명"):
    """결과 칸이 빈칸인 행에만, 위에서부터 처음 맞는 규칙의 값을 채운다. 채운 건수를 돌려준다."""
    filled = 0
    for row in rows:
        if str(row.get(result_column) or "").strip():
            continue
        for _line_no, conds, result in rules:
            if all(_eval(node, _cell_value(row, col, number_columns), col in exact_columns)
                   for col, node in conds):
                row[result_column] = result
                filled += 1
                break
    return filled
