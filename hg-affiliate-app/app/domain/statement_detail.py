"""Build the period-checked related-party transaction detail for sheet 39.1."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

import pandas as pd

from . import columns as C
from .period_extract import PERIOD_YEARMONTH_ATTR, Period, parse_year_month


class StatementDetailError(ValueError):
    """The filtered ledger cannot safely become statement detail."""


@dataclass(frozen=True)
class StatementDetailRow:
    sales_purchase: str | None
    funding: str | None
    receivable_payable: str | None
    income_expense: str | None
    bucket: str | None
    account_code: object
    account_name: object
    date: object
    description: object
    partner_code: object
    partner_name: object
    canonical_name: object
    debit: float
    credit: float
    balance: float

    def as_excel_row(self) -> list[object]:
        return [
            self.sales_purchase,
            self.funding,
            self.receivable_payable,
            self.income_expense,
            self.bucket,
            self.account_code,
            self.account_name,
            self.date,
            self.description,
            self.partner_code,
            self.partner_name,
            self.canonical_name,
            self.debit,
            self.credit,
            self.balance,
        ]


_RECEIVABLE_BUCKETS = {"매출채권", "대여금", "기타채권", "투자전환사채"}
_PAYABLE_BUCKETS = {"기타채무", "발행전환사채", "매입채무"}


def _exact_column(df: pd.DataFrame, *candidates: str) -> str | None:
    compact = {str(column).replace(" ", ""): str(column) for column in df.columns}
    for candidate in candidates:
        match = compact.get(candidate.replace(" ", ""))
        if match is not None:
            return match
    return None


def _excel_date(value: object) -> object:
    if isinstance(value, (dt.datetime, dt.date)):
        return value
    text = str(value).strip()
    separated = re.fullmatch(
        r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})(?:\s+00:00:00)?",
        text,
    )
    compact = re.fullmatch(r"(20\d{2})(\d{2})(\d{2})", text)
    match = separated or compact
    if match is None:
        return value
    try:
        return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return value


def _is_valid_year_month_pair(value: object) -> bool:
    if not isinstance(value, (tuple, list)) or len(value) != 2:
        return False
    year, month = value
    return (
        isinstance(year, int)
        and not isinstance(year, bool)
        and 2000 <= year <= 2099
        and isinstance(month, int)
        and not isinstance(month, bool)
        and 1 <= month <= 12
    )


def _classify(
    account: object,
) -> tuple[str | None, str | None, str | None, str | None, str | None]:
    name = str(account)
    sales_purchase = None
    funding = None
    receivable_payable = None
    income_expense = None
    bucket = None

    if C.match_bucket(name, C.SALES_KEYWORDS):
        sales_purchase, bucket = "매출", "매출"
    elif C.match_bucket(name, C.INTEREST_INCOME_KEYWORDS):
        income_expense, bucket = "이자수익", "이자수익"
    elif C.match_bucket(name, C.PURCHASE_KEYWORDS):
        sales_purchase, bucket = "매입", "매입"
    elif C.match_bucket(name, C.OTHER_EXPENSE_KEYWORDS) or C.match_bucket(
        name, C.ASSET_ACQUIRE_KEYWORDS
    ):
        income_expense, bucket = "비용", "기타비용"

    for balance_bucket, keywords in C.BALANCE_BUCKETS.items():
        if C.match_bucket(name, keywords, C.BALANCE_EXCLUDE.get(balance_bucket)):
            if balance_bucket in _RECEIVABLE_BUCKETS:
                receivable_payable = "채권"
            elif balance_bucket in _PAYABLE_BUCKETS:
                receivable_payable = "채무"
            if bucket is None:
                bucket = balance_bucket
            break

    if C.match_bucket(name, C.LENDING_KEYWORDS, C.ALLOWANCE_KEYWORDS):
        funding = "자금대여"

    return sales_purchase, funding, receivable_payable, income_expense, bucket


def build_statement_detail(
    ledger: pd.DataFrame,
    mapping: dict[str, str],
    canonical: set[str],
    period: Period,
) -> list[StatementDetailRow]:
    account_col = C.resolve_column(ledger, "account")
    partner_col = C.resolve_column(ledger, "partner")
    date_col = C.resolve_column(ledger, "date")
    debit_col = C.resolve_column(ledger, "debit")
    credit_col = C.resolve_column(ledger, "credit")
    balance_col = C.resolve_column(ledger, "amount")
    if None in (account_col, partner_col, date_col, debit_col, credit_col):
        raise StatementDetailError("39.1 상세 거래 필수 열을 식별하지 못했습니다.")

    account_code_col = _exact_column(ledger, "계정코드")
    description_col = _exact_column(ledger, "적요", "적요란")
    partner_code_col = _exact_column(ledger, "거래처코드")
    parsed_periods = ledger.attrs.get(PERIOD_YEARMONTH_ATTR)
    if parsed_periods is not None:
        if (
            not isinstance(parsed_periods, (list, tuple))
            or len(parsed_periods) != len(ledger)
            or not all(_is_valid_year_month_pair(value) for value in parsed_periods)
        ):
            raise StatementDetailError("상세 거래의 기간 메타데이터가 유효하지 않습니다.")
    rows: list[StatementDetailRow] = []

    for position, (_, source) in enumerate(ledger.iterrows()):
        partner_name = str(source.get(partner_col, "")).strip()
        canonical_name = mapping.get(partner_name)
        if canonical_name not in canonical:
            continue

        year_month = (
            parsed_periods[position]
            if parsed_periods is not None
            else parse_year_month(source.get(date_col))
        )
        if year_month is None or not period.contains(*year_month):
            raise StatementDetailError("선택한 누적 기간 밖의 상세 거래가 발견되었습니다.")

        classification = _classify(source.get(account_col))
        rows.append(
            StatementDetailRow(
                *classification,
                account_code=source.get(account_code_col, "") if account_code_col else "",
                account_name=source.get(account_col, ""),
                date=_excel_date(source.get(date_col, "")),
                description=source.get(description_col, "") if description_col else "",
                partner_code=source.get(partner_code_col, "") if partner_code_col else "",
                partner_name=partner_name,
                canonical_name=canonical_name,
                debit=C.to_number(source.get(debit_col)),
                credit=C.to_number(source.get(credit_col)),
                balance=C.to_number(source.get(balance_col)) if balance_col else 0.0,
            )
        )
    return rows
