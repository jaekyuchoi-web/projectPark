"""38.x 집계 (도메인 §4).

입력 프레임(정규화된 헤더)과 사업자명 매핑을 받아 정규(대표) 법인명별로
38.1~38.5 버킷을 산출한다. 산출 불가/불일치는 ErrorLog 에 기록한다.

용어 — "전기"(prev_balance) = 직전 결산년도(전년도 연간 결산 = 전년도 4분기)의 잔액.
결산은 연 1회뿐이므로 당기가 어느 분기든 전기는 항상 직전년도 4분기로 동일한 단일
기준선이다(분기별로 회전하지 않음). 따라서 아래 증감/역산(대손충당금·미수수익)과
재구성 검증은 모두 이 단일 전기말 기준선과 당기(YTD 누적) 사이의 차이로 계산한다.

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
    sales_other: float = 0.0    # 38.1 매출등 기타(이자수익)
    purchase: float = 0.0       # 38.1 매입
    purchase_other: float = 0.0 # 38.1 매입등 기타
    ar: float = 0.0             # 38.2 매출채권
    loan: float = 0.0           # 38.2 대여금
    other_ar: float = 0.0       # 38.2 기타채권
    cb_invest: float = 0.0      # 38.2 투자전환사채
    other_ap: float = 0.0       # 38.2 기타채무
    cb_issued: float = 0.0      # 38.2 발행전환사채
    ap: float = 0.0             # 38.2 매입채무
    allowance_end: float = 0.0  # 38.3 당기말 대손충당금
    allowance_expense: float = 0.0  # 38.3 당기 대손상각비
    fund_lending: float = 0.0   # 38.4 자금대여(당기 발생분)
    # 38.5 경영진 보상은 산출 불가 → 공란 (값 없음)


@dataclass
class AggregateResult:
    by_canonical: dict[str, Aggregate] = field(default_factory=dict)

    def get(self, canonical: str) -> Aggregate:
        if canonical not in self.by_canonical:
            self.by_canonical[canonical] = Aggregate(canonical=canonical)
        return self.by_canonical[canonical]


def _iter_rows(df: pd.DataFrame, mapping: dict[str, str], canonical: set[str]):
    """거래처를 정규명으로 매핑해, 특관자에 해당하는 행만 (canonical, row) 로 yield."""
    partner_col = C.resolve_column(df, "partner")
    if partner_col is None:
        return
    for _, row in df.iterrows():
        raw = str(row.get(partner_col, "")).strip()
        if not raw or raw == "nan":
            continue
        canon = mapping.get(raw)
        if canon and canon in canonical:
            yield canon, row


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
    if acc_col is None or (debit_col is None and credit_col is None):
        errors.add("형식 경고", "당기 원장", "계정/차변/대변 컬럼을 인식하지 못함",
                   "헤더에 계정과목·차변·대변 컬럼명이 포함되도록 정리해 주세요.")
        return

    accrued_net: dict[str, float] = {}  # 원장상 미수수익 순증 (대변-차변 아님: 자산이므로 차변-대변)

    for canon, row in _iter_rows(ledger, mapping, canonical):
        acc = str(row.get(acc_col, ""))
        debit = C.to_number(row.get(debit_col)) if debit_col else 0.0
        credit = C.to_number(row.get(credit_col)) if credit_col else 0.0
        agg = result.get(canon)

        # 매출 (대변-차변)
        if C.match_bucket(acc, C.SALES_KEYWORDS):
            agg.sales += credit - debit
        # 매출등 기타 = 이자수익
        elif C.match_bucket(acc, C.INTEREST_INCOME_KEYWORDS):
            agg.sales_other += credit - debit
        # 매입 = 원재료 (차변-대변)
        elif C.match_bucket(acc, C.PURCHASE_KEYWORDS):
            agg.purchase += debit - credit
        # 매입등 기타 = 기타비용 순액 + 자산취득
        elif C.match_bucket(acc, C.OTHER_EXPENSE_KEYWORDS):
            net = debit - credit
            # 지급임차료 전대(음수 라인) 제외 → 총액 처리
            if C.match_bucket(acc, ["지급임차료"]) and net < 0:
                pass
            else:
                agg.purchase_other += net
        elif C.match_bucket(acc, C.ASSET_ACQUIRE_KEYWORDS):
            agg.purchase_other += debit - credit

        # 미수수익 원장 순증 (잔액증감 역산용)
        if C.match_bucket(acc, C.ACCRUED_INCOME_KEYWORDS):
            accrued_net[canon] = accrued_net.get(canon, 0.0) + (debit - credit)

        # 38.4 자금대여 = 단기·장기대여금 차변 증가(이전 기간 블록/충당금 제외)
        if C.match_bucket(acc, C.LENDING_KEYWORDS, exclude=C.ALLOWANCE_KEYWORDS):
            text = " ".join(str(row.get(c, "")) for c in ledger.columns)
            if not any(token in text for token in C.FUND_LENDING_CARRYFORWARD_TEXTS) and debit > 0:
                agg.fund_lending += debit

    # 매출등 기타: 이자수익 + 미수수익 잔액증감 역산(당기말-전기이월-원장순증)
    prev_accrued = _balance_by_canon(prev_balance, mapping, canonical, C.ACCRUED_INCOME_KEYWORDS)
    cur_accrued = _balance_by_canon(current_balance, mapping, canonical, C.ACCRUED_INCOME_KEYWORDS)
    for canon, agg in result.by_canonical.items():
        delta = cur_accrued.get(canon, 0.0) - prev_accrued.get(canon, 0.0) - accrued_net.get(canon, 0.0)
        agg.sales_other += delta


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
    for canon, row in _iter_rows(balance, mapping, canonical):
        acc = str(row.get(acc_col, ""))
        if C.match_bucket(acc, keywords, exclude):
            out[canon] = out.get(canon, 0.0) + C.to_number(row.get(amt_col))
    return out


def aggregate_balance(
    current_balance: pd.DataFrame | None,
    prev_balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    result: AggregateResult,
    errors: ErrorLog,
) -> None:
    """38.2 채권채무 잔액 + 38.3 대손충당금 (당기말 검증용 잔액명세서 기준)."""
    if current_balance is None:
        errors.add("산출 불가", "당기말 잔액명세서", "검증용 잔액명세서가 없어 38.2/38.3 집계 불가",
                   "당기 채권채무 잔액 검증용 소스를 올려주세요.")
        return

    bucket_attr = {
        "매출채권": "ar", "대여금": "loan", "기타채권": "other_ar",
        "투자전환사채": "cb_invest", "기타채무": "other_ap",
        "발행전환사채": "cb_issued", "매입채무": "ap",
    }
    for bucket, kws in C.BALANCE_BUCKETS.items():
        ex = C.BALANCE_EXCLUDE.get(bucket)
        vals = _balance_by_canon(current_balance, mapping, canonical, kws, ex)
        for canon, v in vals.items():
            setattr(result.get(canon), bucket_attr[bucket], v)

    # 38.3 대손충당금
    cur_allow = _balance_by_canon(current_balance, mapping, canonical, C.ALLOWANCE_KEYWORDS)
    prev_allow = _balance_by_canon(prev_balance, mapping, canonical, C.ALLOWANCE_KEYWORDS)
    for canon, end in cur_allow.items():
        agg = result.get(canon)
        agg.allowance_end = end
        inc = end - prev_allow.get(canon, 0.0)
        agg.allowance_expense = inc if inc > 0 else 0.0

    # 투자/발행 전환사채 보유자 확인 필요 → 검토 항목
    if any(result.get(c).cb_invest or result.get(c).cb_issued for c in result.by_canonical):
        errors.add("자동 산출 불가", "전환사채", "사채 회차별 보유자(특관자) 확인이 필요함",
                   "사채 회차별 보유자 내역을 확인해 투자/발행 전환사채를 검증하세요.")
