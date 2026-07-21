"""Build the period-checked related-party transaction detail for sheet 39.1."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from ..excel_io import DEDUPLICATED_HEADER_BASES_ATTR
from . import columns as C
from .period_extract import PERIOD_YEARMONTH_ATTR, Period, parse_year_month


class StatementDetailError(ValueError):
    """The filtered ledger cannot safely become statement detail."""


def write_detail_cell(cell, value: object) -> None:
    """Write a detail value without treating user text as a formula."""
    if isinstance(value, str):
        if ILLEGAL_CHARACTERS_RE.search(value):
            raise StatementDetailError(
                "39.1 상세 거래를 안전하게 기록하지 못했습니다."
            )
        cell.value = value
        cell.data_type = "s"
        return
    cell.value = value


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


def _column_positions(
    df: pd.DataFrame, *columns: str | None
) -> dict[str, int]:
    """Resolve selected column names to positions for tuple-based row access."""
    positions: dict[str, list[int]] = {}
    for position, column in enumerate(df.columns):
        positions.setdefault(str(column), []).append(position)
    selected: dict[str, int] = {}
    for column in columns:
        if column is None:
            continue
        matches = positions.get(column, [])
        if len(matches) != 1:
            raise StatementDetailError("39.1 상세 거래 필수 열을 식별하지 못했습니다.")
        selected[column] = matches[0]
    return selected


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
    debit: float,
    credit: float,
    row_values: list[object],
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
    elif C.is_purchase_other_eligible(name, debit, credit):
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

    if C.is_fund_lending_eligible(name, debit, row_values):
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
    balance_col = C.resolve_ledger_balance_column(ledger)
    if None in (account_col, partner_col, date_col, debit_col, credit_col):
        raise StatementDetailError("39.1 상세 거래 필수 열을 식별하지 못했습니다.")

    account_code_col = _exact_column(ledger, "계정코드")
    description_col = _exact_column(ledger, "적요", "적요란")
    partner_code_col = _exact_column(ledger, "거래처코드")
    deduplicated_bases = ledger.attrs.get(DEDUPLICATED_HEADER_BASES_ATTR, frozenset())
    if not isinstance(deduplicated_bases, (set, frozenset)):
        raise StatementDetailError("39.1 상세 거래 필수 열을 식별하지 못했습니다.")
    selected_columns = (
        account_col,
        partner_col,
        date_col,
        debit_col,
        credit_col,
        balance_col,
        account_code_col,
        description_col,
        partner_code_col,
    )
    if any(
        column in deduplicated_bases
        for column in selected_columns
        if column is not None
    ):
        raise StatementDetailError("39.1 상세 거래 필수 열을 식별하지 못했습니다.")
    positions = _column_positions(
        ledger,
        account_col,
        partner_col,
        date_col,
        debit_col,
        credit_col,
        balance_col,
        account_code_col,
        description_col,
        partner_code_col,
    )
    account_position = positions[account_col]
    partner_position = positions[partner_col]
    date_position = positions[date_col]
    debit_position = positions[debit_col]
    credit_position = positions[credit_col]
    balance_position = positions.get(balance_col) if balance_col else None
    account_code_position = (
        positions.get(account_code_col) if account_code_col else None
    )
    description_position = (
        positions.get(description_col) if description_col else None
    )
    partner_code_position = (
        positions.get(partner_code_col) if partner_code_col else None
    )
    parsed_periods = ledger.attrs.get(PERIOD_YEARMONTH_ATTR)
    if parsed_periods is not None:
        if (
            not isinstance(parsed_periods, (list, tuple))
            or len(parsed_periods) != len(ledger)
            or not all(_is_valid_year_month_pair(value) for value in parsed_periods)
        ):
            raise StatementDetailError("상세 거래의 기간 메타데이터가 유효하지 않습니다.")
    rows: list[StatementDetailRow] = []

    for position, source in enumerate(C.array_rows(ledger)):
        partner_name = str(source[partner_position]).strip()
        canonical_name = mapping.get(partner_name)
        if canonical_name not in canonical:
            continue

        year_month = (
            parsed_periods[position]
            if parsed_periods is not None
            else parse_year_month(source[date_position])
        )
        if year_month is None or not period.contains(*year_month):
            raise StatementDetailError("선택한 누적 기간 밖의 상세 거래가 발견되었습니다.")

        debit = C.to_number(source[debit_position])
        credit = C.to_number(source[credit_position])
        classification = _classify(
            source[account_position], debit, credit, source
        )
        rows.append(
            StatementDetailRow(
                *classification,
                account_code=(
                    source[account_code_position]
                    if account_code_position is not None
                    else ""
                ),
                account_name=source[account_position],
                date=_excel_date(source[date_position]),
                description=(
                    source[description_position]
                    if description_position is not None
                    else ""
                ),
                partner_code=(
                    source[partner_code_position]
                    if partner_code_position is not None
                    else ""
                ),
                partner_name=partner_name,
                canonical_name=canonical_name,
                debit=debit,
                credit=credit,
                balance=(
                    C.to_number(source[balance_position])
                    if balance_position is not None
                    else 0.0
                ),
            )
        )
    return rows
