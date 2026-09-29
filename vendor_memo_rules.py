# -*- coding: utf-8 -*-
"""
'조건 칸' 방식 메모 규칙 (생활비: rules/memo_mapping_생활비.csv).

파일 형식 (CSV, 첫 줄은 헤더):
    은행,거래구분,적요,출금,입금,메모

  - 한 줄이 규칙 하나. 값이 적힌 조건 칸은 **전부** 만족해야 메모가 적용된다.
  - **빈칸인 조건 칸은 무시**한다. (조건이 하나도 없는 줄은 모든 거래에 붙어버리므로 건너뜀)
  - 위에서부터 검사해서 **처음 맞는 규칙 하나만** 적용한다. 메모가 이미 있는 행은 건드리지 않는다.

조건 칸에 쓰는 법 (구글 검색 연산자처럼):
  - 글자 조건      쿠팡              그냥 적으면 '포함' (은행/거래구분/적요 모두)
                   스타벅스 강남      공백이 있어도 하나의 글자 조건 ('스타벅스 강남'이 통째로 들어있어야 함)
  - 둘 다 만족     스타벅스 AND 배달
  - 하나라도 만족   스타벅스 OR 이디야
  - 제외          NOT 배달           (예: 스타벅스 AND NOT 배달)
  - 묶기          (스타벅스 OR 이디야) AND NOT 배달     괄호 안에 AND나 OR가 있을 때만 '묶기'로 계산한다.
                                                    안에 AND/OR가 없는 괄호(예: 씨유(CU))는 그냥 글자 그대로 취급한다.
  - 큰따옴표       "바나나"           괄호처럼 묶기로 계산하고 싶지 않거나, '포함'이 아니라 정확히 그 글자와
                                    완전히 똑같아야만 매칭하고 싶을 때 큰따옴표로 감싼다.
                                    (예: "바나나"는 '바나나'와만 매칭되고, '바나나우유'에는 매칭되지 않는다)
  - 숫자 조건 (출금/입금)   7890, =7890, >=10000, <5000, <>0, >=1000 AND <=5000, 7890 OR 9900

  * AND / OR / NOT 은 반드시 **대문자**로 쓴다. (소문자 and/or/not 은 그냥 글자로 취급)
  * 우선순위는 NOT -> AND -> OR 순이다. 헷갈리면 괄호로 묶는다.
  * 글자 조건은 **대소문자를 구분한다** (GS25와 gs25는 다른 글자로 취급).
"""
import csv
import io
import re

CONDITION_COLUMNS = ("은행", "거래구분", "적요", "출금", "입금")
_NUMBER_COLUMNS = ("출금", "입금")
_EXACT_TEXT_COLUMNS = ()  # 그냥 적은 글자를 '정확히 같음'으로 보는 칸 (없음: 은행/거래구분/적요 모두 '포함')

_OPERATORS = ("AND", "OR", "NOT")
_NUM_RE = re.compile(r"^(>=|<=|<>|!=|>|<|=)?\s*(-?[\d,]+(?:\.\d+)?)$")
_QUOTE_CHARS = '"“”'


def _find_structural_parens(text):
    """'(' / ')' 중에서 '묶기'로 계산해야 하는 위치(글자 번호)의 집합을 찾는다.
    짝이 맞는 괄호 쌍 안에 (공백으로 구분된) AND 나 OR 가 하나라도 있어야만
    '묶기'로 취급하고, 그 외의 괄호(예: 씨유(CU))는 그냥 평범한 글자로 남겨둔다.
    (큰따옴표 안에 있는 내용은 괄호/AND/OR 검사에서 제외한다)"""
    n = len(text)
    in_quote = bytearray(n)
    i = 0
    while i < n:
        if text[i] in _QUOTE_CHARS:
            j = i + 1
            while j < n and text[j] not in _QUOTE_CHARS:
                j += 1
            end = min(j + 1, n)
            for k in range(i, end):
                in_quote[k] = 1
            i = end
        else:
            i += 1

    stack, pairs = [], []
    for idx, ch in enumerate(text):
        if in_quote[idx]:
            continue
        if ch == "(":
            stack.append(idx)
        elif ch == ")" and stack:
            pairs.append((stack.pop(), idx))

    structural = set()
    for open_idx, close_idx in pairs:
        visible = "".join(
            c if not m else " "
            for c, m in zip(text[open_idx + 1:close_idx], in_quote[open_idx + 1:close_idx])
        )
        words = visible.split()
        if "AND" in words or "OR" in words:
            structural.add(open_idx)
            structural.add(close_idx)
    return structural


def _tokenize(text):
    """조건 글자를 (종류, 값, ...) 조각들로 나눈다.
    종류: LP '(' / RP ')' / OP AND·OR·NOT / TERM (글자, 큰따옴표로 감쌌는지 여부).
    연산자·괄호가 아닌 단어들이 공백으로 이어져 있으면 하나의 글자 조건으로 합친다.
    AND/OR가 안에 없는 괄호는 '묶기'가 아니라 그냥 글자 취급한다(_find_structural_parens)."""
    structural = _find_structural_parens(text)
    tokens, pending = [], []

    def flush():
        if pending:
            tokens.append(("TERM", " ".join(pending), False))
            pending.clear()

    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
        elif ch in _QUOTE_CHARS:
            flush()
            j = i + 1
            while j < n and text[j] not in _QUOTE_CHARS:
                j += 1
            if j >= n:
                raise ValueError("큰따옴표가 닫히지 않았습니다")
            tokens.append(("TERM", text[i + 1:j], True))
            i = j + 1
        elif ch in "()" and i in structural:
            flush()
            tokens.append(("LP" if ch == "(" else "RP", ch))
            i += 1
        else:
            j = i
            while (j < n and not text[j].isspace() and text[j] not in _QUOTE_CHARS
                   and not (text[j] in "()" and j in structural)):
                j += 1
            word = text[i:j]
            i = j
            if word in _OPERATORS:
                flush()
                tokens.append(("OP", word))
            else:
                pending.append(word)
    flush()
    return tokens


def _literal(text, numeric, quoted=False):
    if text == "":
        raise ValueError("비어 있는 조건이 있습니다")
    if numeric:
        m = _NUM_RE.match(text)
        if not m:
            raise ValueError(f"숫자 조건이 아닙니다: {text}")
        op = m.group(1) or "="
        op = "<>" if op == "!=" else op
        return ("NUM", op, float(m.group(2).replace(",", "")))
    return ("TEXT", text, quoted)


class _Parser:
    """OR(가장 약함) < AND < NOT(가장 강함) 순서로 읽는 작은 파서."""

    def __init__(self, tokens, numeric):
        self.tokens, self.pos, self.numeric = tokens, 0, numeric

    def _peek(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def _is_op(self, name):
        tok = self._peek()
        return tok is not None and tok[0] == "OP" and tok[1] == name

    def parse(self):
        if not self.tokens:
            raise ValueError("조건이 비어 있습니다")
        node = self._or()
        tok = self._peek()
        if tok is not None:
            if tok[0] == "RP":
                raise ValueError("닫는 괄호 )가 더 많습니다 (괄호나 AND/OR가 들어간 글자는 큰따옴표로 감싸세요)")
            raise ValueError("조건과 조건 사이에 AND 또는 OR가 필요합니다 "
                             "(괄호가 들어간 글자는 큰따옴표로 감싸세요. 예: \"(주)에스알\")")
        return node

    def _or(self):
        nodes = [self._and()]
        while self._is_op("OR"):
            self.pos += 1
            nodes.append(self._and())
        return nodes[0] if len(nodes) == 1 else ("OR", nodes)

    def _and(self):
        nodes = [self._not()]
        while self._is_op("AND"):
            self.pos += 1
            nodes.append(self._not())
        return nodes[0] if len(nodes) == 1 else ("AND", nodes)

    def _not(self):
        if self._is_op("NOT"):
            self.pos += 1
            return ("NOT", self._not())
        return self._primary()

    def _primary(self):
        tok = self._peek()
        if tok is None:
            raise ValueError("AND / OR / NOT 뒤에 조건이 없습니다")
        kind = tok[0]
        if kind == "LP":
            self.pos += 1
            node = self._or()
            if self._peek() is None or self._peek()[0] != "RP":
                raise ValueError("여는 괄호 (에 짝이 되는 닫는 괄호 )가 없습니다")
            self.pos += 1
            return node
        if kind == "TERM":
            self.pos += 1
            return _literal(tok[1], self.numeric, tok[2])
        if kind == "RP":
            raise ValueError("괄호 ( ) 안이 비어 있거나 닫는 괄호 )가 너무 일찍 나왔습니다")
        raise ValueError(f"{tok[1]} 앞에 조건이 없습니다")


def _parse(text, numeric):
    """조건 글자를 검사용 구조로 바꾼다. 문법이 틀리면 ValueError."""
    return _Parser(_tokenize(text), numeric).parse()


def _compare(op, value, target):
    return {
        "=": value == target, "<>": value != target,
        ">": value > target, ">=": value >= target,
        "<": value < target, "<=": value <= target,
    }[op]


def _eval(node, value, exact_text):
    kind = node[0]
    if kind == "AND":
        return all(_eval(n, value, exact_text) for n in node[1])
    if kind == "OR":
        return any(_eval(n, value, exact_text) for n in node[1])
    if kind == "NOT":
        return not _eval(node[1], value, exact_text)
    if kind == "NUM":
        return _compare(node[1], value, node[2])
    # TEXT: 대소문자를 구분한다. 큰따옴표로 감쌌거나(quoted) exact_text 칸이면
    # '완전히 같음', 아니면 '포함'으로 본다.
    _, want, quoted = node
    have = str(value)
    return have == want if (exact_text or quoted) else want in have


def _read_text(path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp949"):  # 엑셀에서 'CSV(쉼표로 분리)'로 저장하면 cp949가 될 수 있다
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{path.name} 파일의 글자 인코딩을 읽을 수 없습니다 (UTF-8 CSV로 저장해주세요)")


def is_condition_format(path):
    """규칙 파일 헤더에 은행/거래구분/적요/출금/입금 칸이 있으면 '조건 칸' 방식이다."""
    if path is None or not path.exists():
        return False
    header = _read_text(path).lstrip("﻿").splitlines()[:1]
    if not header:
        return False
    names = {c.strip() for c in header[0].split(",")}
    return bool(names & set(CONDITION_COLUMNS))


def load(path):
    """[(줄번호, [(칸이름, 조건), ...], 메모), ...] 를 돌려준다. 문법이 틀린 줄은 경고하고 건너뛴다."""
    rules = []
    if path is None or not path.exists():
        return rules
    reader = csv.DictReader(io.StringIO(_read_text(path)))
    # 엑셀에서 헤더 칸에 실수로 공백이 섞여도(예: ' 입금 ') 그 칸을 통째로 못 찾아
    # 조건이 조용히 무시되는 일이 없도록, 칸 이름의 앞뒤 공백을 무시한다.
    if reader.fieldnames:
        reader.fieldnames = [(name or "").strip() for name in reader.fieldnames]
    for line_no, item in enumerate(reader, start=2):
        memo = (item.get("메모") or "").strip()
        conds = []
        col = ""
        try:
            for col in CONDITION_COLUMNS:
                cell = (item.get(col) or "").strip()
                if cell:
                    conds.append((col, _parse(cell, col in _NUMBER_COLUMNS)))
        except ValueError as e:
            print(f"  [경고] 메모 규칙 {line_no}째 줄 '{col}' 칸을 건너뜁니다: {e}")
            continue
        if not conds and not memo:
            continue  # 완전히 빈 줄
        if not memo:
            print(f"  [경고] 메모 규칙 {line_no}째 줄은 메모가 비어 있어 건너뜁니다.")
            continue
        if not conds:
            print(f"  [경고] 메모 규칙 {line_no}째 줄은 조건이 하나도 없어 건너뜁니다. "
                  f"(모든 거래에 같은 메모가 붙는 것을 막기 위함)")
            continue
        rules.append((line_no, conds, memo))
    return rules


def _cell_value(row, col):
    if col in _NUMBER_COLUMNS:
        return int(row.get(col) or 0)
    return str(row.get(col) or "")


def apply(rows, rules):
    """메모가 빈칸인 행에만, 위에서부터 처음 맞는 규칙의 메모를 채운다."""
    for row in rows:
        row.setdefault("메모", "")
        if row["메모"]:
            continue
        for _line_no, conds, memo in rules:
            if all(_eval(node, _cell_value(row, col), col in _EXACT_TEXT_COLUMNS) for col, node in conds):
                row["메모"] = memo
                break
