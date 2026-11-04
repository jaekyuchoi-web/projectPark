"""38.x 집계 (도메인 §4).

입력 프레임(정규화된 헤더)과 사업자명 매핑을 받아 정규(대표) 법인명별로
38.1~38.5 버킷을 산출한다. 산출 불가/불일치는 ErrorLog 에 기록한다.

용어 — "전기"(prev_balance)는 전년도 온기(연간 결산 = 4분기) 잔액이다.
38.3은 전기 충당금에 당기 원장 증감을 더하고, 상각비는 당기 원장에서만
집계한다. 미수수익 잔액 차이를 수익으로 역산하지 않는다.

주의: 어떤 산출 '수치'도 로그/화면으로 내보내지 않는다. 결과는 출력 엑셀에만 기록.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from . import columns as C
from .errors import ErrorLog


@dataclass
class Aggregate:
    canonical: str
    sales: float = 0.0          # 38.1 매출
    sales_other: float = 0.0    # 38.1 매출등 기타(이자수익, 충당금 변동 제외)
    purchase: float = 0.0       # 38.1 매입
    purchase_other: float = 0.0 # 38.1 매입등 기타
    ar: float = 0.0             # 38.2 매출채권
    loan: float = 0.0           # 38.2 대여금
    other_ar: float = 0.0       # 38.2 기타채권
    cb_invest: float = 0.0      # 38.2 투자전환사채
    other_ap: float = 0.0       # 38.2 기타채무
    cb_issued: float = 0.0      # 38.2 발행전환사채
    ap: float = 0.0             # 38.2 매입채무
    borrowing: float = 0.0
    allowance_opening: float = 0.0
    allowance_end: float = 0.0  # 38.3 당기말 대손충당금
    allowance_expense: float = 0.0  # 38.3 당기 대손상각비(환입이면 음수)
    fund_lending: float = 0.0   # 38.4 자금대여(당기 발생분)
    fund_collection: float = 0.0
    fund_borrowing: float = 0.0
    fund_repayment: float = 0.0
    fund_investment: float = 0.0
    fund_disposal: float = 0.0
    # 38.5 경영진 보상은 산출 불가 → 공란 (값 없음)


@dataclass
class AggregateResult:
    by_canonical: dict[str, Aggregate] = field(default_factory=dict)

    def get(self, canonical: str) -> Aggregate:
        if canonical not in self.by_canonical:
            self.by_canonical[canonical] = Aggregate(canonical=canonical)
        return self.by_canonical[canonical]


def _iter_rows(df: pd.DataFrame, mapping: dict[str, str], canonical: set[str]):
    """Yield related-party canonical names with tuple rows and column positions."""
    partner_col = C.resolve_column(df, "partner")
    if partner_col is None:
        return
    positions = C.array_column_positions(df)
    for row in C.array_rows(df):
        raw = str(C.array_row_value(row, positions, partner_col, "")).strip()
        if not raw or raw == "nan":
            continue
        canon = mapping.get(raw)
        if canon and canon in canonical:
            yield canon, row, positions


def aggregate_ledger(
    ledger: pd.DataFrame | None,
    prev_balance: pd.DataFrame | None,
    current_balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    result: AggregateResult,
    errors: ErrorLog,
) -> None:
    """38.1 거래내역 + 38.4 자금거래(당기 원장 기반)."""
    if ledger is None:
        errors.add("산출 불가", "당기 원장", "당기 거래 내역 소스가 없어 38.1/38.4 집계를 건너뜀",
                   "당기 계정별 원장 파일을 올려주세요.")
        return
    acc_col = C.resolve_column(ledger, "account")
    debit_col = C.resolve_column(ledger, "debit")
    credit_col = C.resolve_column(ledger, "credit")
    description_col = C.resolve_column(ledger, "description")
    if acc_col is None or (debit_col is None and credit_col is None):
        errors.add("형식 경고", "당기 원장", "계정/차변/대변 컬럼을 인식하지 못함",
                   "헤더에 계정과목·차변·대변 컬럼명이 포함되도록 정리해 주세요.")
        return

    opening = _balance_by_canon(prev_balance, mapping, canonical, C.ALLOWANCE_KEYWORDS)
    for canon in set(result.by_canonical) | set(opening):
        agg = result.get(canon)
        agg.allowance_opening = opening.get(canon, 0.0)
        agg.allowance_end = agg.allowance_opening
        agg.allowance_expense = 0.0

    for canon, row, positions in _iter_rows(ledger, mapping, canonical):
        acc = str(C.array_row_value(row, positions, acc_col, ""))
        debit = C.to_number(C.array_row_value(row, positions, debit_col))
        credit = C.to_number(C.array_row_value(row, positions, credit_col))
        debit, credit, shared_income = C.reclassify_shared_expense(
            acc, C.array_row_value(row, positions, description_col, ''), debit, credit,
        )
        agg = result.get(canon)

        # 매출 (대변-차변)
        if shared_income:
            agg.sales_other += credit - debit
        elif C.match_bucket(acc, C.SALES_KEYWORDS):
            agg.sales += credit - debit
        # 기타수익 = 이자수익·수입임대료·잡이익
        elif C.match_bucket(acc, C.OTHER_INCOME_KEYWORDS):
            agg.sales_other += credit - debit
        # 매입 = 원재료·상품·유무형자산 취득 (차변-대변)
        elif C.match_bucket(acc, C.PURCHASE_KEYWORDS):
            agg.purchase += debit - credit
        # 기타비용 = 지정된 비용 계정의 차변-대변
        elif C.is_purchase_other_eligible(acc, debit, credit):
            agg.purchase_other += debit - credit

        if C.match_bucket(acc, C.ALLOWANCE_KEYWORDS):
            agg.allowance_end += credit - debit
        if C.match_bucket(acc, C.ALLOWANCE_EXPENSE_KEYWORDS):
            agg.allowance_expense += debit - credit

        # 38.4: 당기 차변/대변을 각각 집계한다. 이월은 거래가 아니다.
        if not any(token in ' '.join(map(str, row)) for token in C.FUND_LENDING_CARRYFORWARD_TEXTS):
            if C.match_bucket(acc, C.LENDING_KEYWORDS):
                agg.fund_lending += debit
                agg.fund_collection += credit
            elif C.match_bucket(acc, C.BALANCE_BUCKETS['차입금']):
                agg.fund_borrowing += credit
                agg.fund_repayment += debit
            elif C.match_bucket(acc, C.BALANCE_BUCKETS['투자전환사채']):
                agg.fund_investment += debit
                agg.fund_disposal += credit

    actual = _balance_by_canon(current_balance, mapping, canonical, C.ALLOWANCE_KEYWORDS)
    for canon, agg in result.by_canonical.items():
        if current_balance is not None and abs(agg.allowance_end - actual.get(canon, 0.0)) > 1:
            errors.add('검증 불일치', f'{canon} / 대손충당금',
                       '전기잔액과 당기 원장으로 계산한 충당금이 검증용 잔액과 다름',
                       '38.3은 규칙에 따라 전기잔액+당기 증감을 유지했습니다. 조정분개를 확인하세요.')


def _balance_by_canon(
    balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    keywords: list[str],
    exclude: list[str] | None = None,
) -> dict[str, float]:
    """잔액명세서에서 특정 계정 키워드 합을 정규명별로."""
    out: dict[str, float] = {}
    if balance is None:
        return out
    acc_col = C.resolve_column(balance, "account")
    amt_col = C.resolve_column(balance, "amount")
    if acc_col is None or amt_col is None:
        return out
    for canon, row, positions in _iter_rows(balance, mapping, canonical):
        acc = str(C.array_row_value(row, positions, acc_col, ""))
        if C.match_bucket(acc, keywords, exclude):
            out[canon] = out.get(canon, 0.0) + C.to_number(
                C.array_row_value(row, positions, amt_col)
            )
    return out


def aggregate_balance(
    current_balance: pd.DataFrame | None,
    prev_balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    result: AggregateResult,
    errors: ErrorLog,
) -> None:
    """38.2 검증용 잔액 집계. 38.3은 aggregate_ledger에서 별도 계산한다."""
    if current_balance is None:
        errors.add("산출 불가", "당기말 잔액명세서", "검증용 잔액명세서가 없어 38.2/38.3 집계 불가",
                   "당기 채권채무 잔액 검증용 소스를 올려주세요.")
        return

    bucket_attr = {
        "매출채권": "ar", "대여금": "loan", "기타채권": "other_ar",
        "투자전환사채": "cb_invest", "기타채무": "other_ap",
        "발행전환사채": "cb_issued", "매입채무": "ap", "차입금": "borrowing",
    }
    for bucket, kws in C.BALANCE_BUCKETS.items():
        ex = C.BALANCE_EXCLUDE.get(bucket)
        vals = _balance_by_canon(current_balance, mapping, canonical, kws, ex)
        for canon, v in vals.items():
            setattr(result.get(canon), bucket_attr[bucket], v)

    # 38.3 시작점만 설정한다. 원장 없이는 상각비를 잔액 차이로 추정하지 않는다.
    prev_allow = _balance_by_canon(prev_balance, mapping, canonical, C.ALLOWANCE_KEYWORDS)
    for canon, opening in prev_allow.items():
        agg = result.get(canon)
        agg.allowance_opening = opening
        agg.allowance_end = opening

    # 투자/발행 전환사채 보유자 확인 필요 → 검토 항목
    if any(result.get(c).cb_invest or result.get(c).cb_issued for c in result.by_canonical):
        errors.add("자동 산출 불가", "전환사채", "사채 회차별 보유자(특관자) 확인이 필요함",
                   "사채 회차별 보유자 내역을 확인해 투자/발행 전환사채를 검증하세요.")
