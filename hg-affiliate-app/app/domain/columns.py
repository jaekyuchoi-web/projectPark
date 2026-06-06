"""컬럼 해석 및 계정 분류 (도메인 공용 유틸).

입력 파일의 컬럼명이 제각각이므로 키워드로 핵심 컬럼을 찾아낸다.
또한 계정과목명을 38.x 버킷으로 분류하는 키워드 맵을 제공한다.
"""

from __future__ import annotations

import re

import pandas as pd

# 핵심 컬럼 후보 키워드
COLUMN_ALIASES: dict[str, list[str]] = {
    "partner": ["거래처", "거래처명", "상호", "업체", "거래상대", "계정상대", "성명"],
    "account": ["계정", "계정과목", "과목", "account"],
    "debit": ["차변", "차변금액", "debit", "출금"],
    "credit": ["대변", "대변금액", "credit", "입금"],
    "amount": ["금액", "잔액", "기말잔액", "당기잔액", "balance", "amount"],
    "date": ["일자", "전표일자", "거래일", "date"],
}


def resolve_column(df: pd.DataFrame, role: str) -> str | None:
    """역할(role)에 해당하는 실제 컬럼명을 찾는다.

    주의(원장형 데이터): '거래처코드/거래처명', '계정코드/계정명' 처럼 코드·이름
    컬럼이 함께 있는 경우, 매핑/키워드 매칭은 '이름' 기준이어야 하므로
    partner/account 역할은 '코드'·'번호' 컬럼을 피하고 '명' 컬럼을 우선한다.
    """
    aliases = COLUMN_ALIASES.get(role, [])
    cols = [str(c) for c in df.columns]
    avoid = ("코드", "번호") if role in ("partner", "account") else ()

    best: str | None = None
    for alias in aliases:
        for c in cols:
            if alias in c and not any(x in c for x in avoid):
                if role in ("partner", "account") and "명" in c:
                    return c  # 거래처명/계정명 등 '명' 컬럼 최우선
                if best is None:
                    best = c
    if best is not None:
        return best

    # 폴백: 회피 조건을 무시하고서라도(코드뿐일 때) 찾는다
    for alias in aliases:
        for c in cols:
            if alias in c:
                return c
    return None


# ── 거래처가 아닌 피벗/집계 산출물 라벨 (법인명에서 제외) ──────────
# 엑셀 피벗테이블/소계 행에서 흘러든 텍스트가 법인으로 오인되지 않도록 차단.
AGGREGATE_LABELS: set[str] = {
    "총합계", "합계", "소계", "총계", "누계", "계",
    "행레이블", "열레이블", "행 레이블", "열 레이블",
    "grandtotal", "total", "subtotal", "sum", "rowlabels", "columnlabels",
    "비어있음", "(비어있음)", "blank", "(blank)", "none", "nan",
    "전체", "기타", "미지정", "해당없음",
}


def is_aggregate_label(name: object) -> bool:
    """피벗/집계 산출물(총합계·행 레이블 등)이거나 법인명으로 볼 수 없는 값인지."""
    if name is None:
        return True
    s = str(name).strip()
    if not s or s.lower() == "nan":
        return True
    key = re.sub(r"[\s\u3000]", "", s).lower()
    if key in AGGREGATE_LABELS:
        return True
    # 괄호로만 묶인 라벨류: (비어 있음), (blank) 등
    inner = key.strip("()[]")
    if inner in AGGREGATE_LABELS:
        return True
    # 숫자/기호만으로 이루어진 값은 거래처명이 아님
    if re.fullmatch(r"[\d\W_]+", s):
        return True
    return False


def to_number(value) -> float:
    """문자열 금액을 float 으로. 괄호=음수, 콤마 제거. 실패 시 0."""
    if value is None:
        return 0.0
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace(",", "").replace("(", "").replace(")", "").replace("₩", "").replace("원", "")
    try:
        num = float(s)
    except ValueError:
        return 0.0
    return -num if neg else num


# ── 계정 분류 키워드 (38.1 거래내역) ──────────────────────────────
SALES_KEYWORDS = ["제품매출", "상품매출", "용역매출", "미디어", "렌탈료", "임대료수익"]
INTEREST_INCOME_KEYWORDS = ["이자수익"]
ACCRUED_INCOME_KEYWORDS = ["미수수익"]
PURCHASE_KEYWORDS = ["원재료"]
OTHER_EXPENSE_KEYWORDS = [
    "접대비", "기업업무추진비", "복리후생", "여비교통", "운반비",
    "지급임차료", "판매수수료", "지급수수료",
]
ASSET_ACQUIRE_KEYWORDS = ["건설중인자산", "상표권", "시설장치"]

# ── 38.2 채권채무 잔액 ─────────────────────────────────────────
BALANCE_BUCKETS: dict[str, list[str]] = {
    "매출채권": ["외상매출금", "받을어음"],
    "대여금": ["단기대여금", "장기대여금", "대여금"],
    "기타채권": ["미수금", "미수수익", "선급금"],
    "투자전환사채": ["전환사채(투자", "투자전환사채", "전환사채투자"],
    "기타채무": ["미지급금", "선수금", "예수금"],
    "발행전환사채": ["전환사채(발행", "발행전환사채", "전환사채발행"],
    "매입채무": ["외상매입금"],
}
# 매출채권에서 제외 (장기외상매출금)
BALANCE_EXCLUDE = {"매출채권": ["장기외상매출금"]}
ALLOWANCE_KEYWORDS = ["대손충당금"]


def match_bucket(account: str, keywords: list[str], exclude: list[str] | None = None) -> bool:
    a = re.sub(r"[\s\u3000]", "", str(account))
    if exclude:
        for ex in exclude:
            if re.sub(r"[\s\u3000]", "", ex) in a:
                return False
    for kw in keywords:
        if re.sub(r"[\s\u3000]", "", kw) in a:
            return True
    return False
