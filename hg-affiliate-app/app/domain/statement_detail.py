"""Build the period-checked related-party transaction detail for sheet 39.1."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, replace

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
    balance: float | None

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
_OPENING_LABEL = "[전기이월]"


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


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.casefold() == "nan" else text


def _account_key(value: object) -> str:
    return re.sub(r"[\s\u3000]", "", _clean_text(value)).casefold()


def detail_group_key(row: StatementDetailRow) -> tuple[str, str]:
    """Return the contiguous 39.1 balance group identity: (company, account)."""
    return (
        _clean_text(row.canonical_name),
        _account_key(row.account_name),
    )


def _balance_bucket(account: object) -> str | None:
    name = str(account)
    for bucket, keywords in C.BALANCE_BUCKETS.items():
        if C.match_bucket(name, keywords, C.BALANCE_EXCLUDE.get(bucket)):
            return bucket
    return None


def _short_opening_account_code(value: object) -> object:
    text = _clean_text(value)
    if not text:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if text.isdigit() and len(text) >= 7 and text.endswith("00"):
        return text[:-2]
    return text


def _ledger_opening_account_codes(
    ledger: pd.DataFrame,
    account_position: int,
    account_code_position: int | None,
) -> dict[str, object]:
    codes: dict[str, object] = {}
    if account_code_position is None:
        return codes
    for source in C.array_rows(ledger):
        account = source[account_position]
        code = _short_opening_account_code(source[account_code_position])
        if _clean_text(code):
            codes.setdefault(_account_key(account), code)
    return codes


def _opening_rows(
    prev_balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    account_codes: dict[str, object],
) -> list[StatementDetailRow]:
    if prev_balance is None or prev_balance.empty:
        return []

    account_col = C.resolve_column(prev_balance, "account")
    partner_col = C.resolve_column(prev_balance, "partner")
    amount_col = C.resolve_column(prev_balance, "amount")
    if None in (account_col, partner_col, amount_col):
        raise StatementDetailError("39.1 전기이월 필수 열을 식별하지 못했습니다.")
    partner_code_col = _exact_column(prev_balance, "거래처코드", "코드")
    selected_columns = (account_col, partner_col, amount_col, partner_code_col)
    deduplicated_bases = prev_balance.attrs.get(
        DEDUPLICATED_HEADER_BASES_ATTR, frozenset()
    )
    if not isinstance(deduplicated_bases, (set, frozenset)) or any(
        column in deduplicated_bases
        for column in selected_columns
        if column is not None
    ):
        raise StatementDetailError("39.1 전기이월 필수 열을 식별하지 못했습니다.")
    positions = _column_positions(prev_balance, *selected_columns)
    account_position = positions[account_col]
    partner_position = positions[partner_col]
    amount_position = positions[amount_col]
    partner_code_position = (
        positions.get(partner_code_col) if partner_code_col else None
    )

    by_group: dict[tuple[str, str], list[StatementDetailRow]] = {}
    for source in C.array_rows(prev_balance):
        partner_name = _clean_text(source[partner_position])
        canonical_name = mapping.get(partner_name)
        if canonical_name not in canonical:
            continue
        account_name = source[account_position]
        bucket = _balance_bucket(account_name)
        if bucket is None:
            continue
        amount = C.to_number(source[amount_position])
        if amount == 0.0:
            continue
        partner_code = (
            source[partner_code_position]
            if partner_code_position is not None
            else ""
        )
        is_credit = bucket in _PAYABLE_BUCKETS
        debit = 0.0 if is_credit else amount
        credit = amount if is_credit else 0.0
        classification = list(
            _classify(
                account_name,
                debit,
                credit,
                [account_name, partner_name, _OPENING_LABEL],
            )
        )
        if bucket == "대여금":
            classification[1] = "대여금"
        row = StatementDetailRow(
            *classification,
            account_code=account_codes.get(_account_key(account_name), ""),
            account_name=account_name,
            date=_OPENING_LABEL,
            description="",
            partner_code=partner_code,
            partner_name=partner_name,
            canonical_name=canonical_name,
            debit=debit,
            credit=credit,
            balance=None,
        )
        by_group.setdefault(detail_group_key(row), []).append(row)

    rows: list[StatementDetailRow] = []
    for group in by_group.values():
        aliases = {_clean_text(row.partner_name) for row in group}
        if len(aliases) > 1:
            first = group[0]
            rows.append(
                replace(
                    first,
                    debit=sum(row.debit for row in group),
                    credit=sum(row.credit for row in group),
                )
            )
        else:
            rows.extend(group)
    return rows


def uses_credit_balance(row: StatementDetailRow) -> bool:
    """Return whether a 39.1 group balance follows credit-minus-debit."""
    if row.receivable_payable == "채권" or row.bucket in _RECEIVABLE_BUCKETS:
        return False
    if row.receivable_payable == "채무" or row.bucket in _PAYABLE_BUCKETS:
        return True
    if row.sales_purchase == "매출" or row.income_expense == "이자수익":
        return True
    if row.bucket in {"기타수익", "자산매각"}:
        return True
    digits = re.sub(r"\D", "", _clean_text(row.account_code))
    return bool(digits) and digits[0] in {"2", "3", "4"}


def detail_balance_formula(
    row: StatementDetailRow,
    first_excel_row: int,
    last_excel_row: int,
) -> str:
    debit_sum = f"SUM(N{first_excel_row}:N{last_excel_row})"
    credit_sum = f"SUM(O{first_excel_row}:O{last_excel_row})"
    if uses_credit_balance(row):
        return f"={credit_sum}-{debit_sum}"
    return f"={debit_sum}-{credit_sum}"


def build_statement_detail(
    ledger: pd.DataFrame,
    mapping: dict[str, str],
    canonical: set[str],
    period: Period,
    prev_balance: pd.DataFrame | None = None,
) -> list[StatementDetailRow]:
    account_col = C.resolve_column(ledger, "account")
    partner_col = C.resolve_column(ledger, "partner")
    date_col = C.resolve_column(ledger, "date")
    debit_col = C.resolve_column(ledger, "debit")
    credit_col = C.resolve_column(ledger, "credit")
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
        account_code_col,
        description_col,
        partner_code_col,
    )
    account_position = positions[account_col]
    partner_position = positions[partner_col]
    date_position = positions[date_col]
    debit_position = positions[debit_col]
    credit_position = positions[credit_col]
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
    account_codes = _ledger_opening_account_codes(
        ledger, account_position, account_code_position
    )
    grouped: dict[tuple[str, str], list[StatementDetailRow]] = {}
    for opening in _opening_rows(
        prev_balance, mapping, canonical, account_codes
    ):
        grouped.setdefault(detail_group_key(opening), []).append(opening)

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
        row = StatementDetailRow(
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
            balance=None,
        )
        grouped.setdefault(detail_group_key(row), []).append(row)

    group_order = {key: index for index, key in enumerate(grouped)}
    company_order: dict[str, int] = {}
    for key in grouped:
        company_order.setdefault(key[0], group_order[key])

    rows: list[StatementDetailRow] = []
    for key in sorted(
        grouped, key=lambda item: (company_order[item[0]], group_order[item])
    ):
        group = grouped[key]
        last = group[-1]
        # The template's summary and 39.2 detail blocks intentionally reflect only
        # classified buckets.  Leaving P blank for an unclassified account preserves
        # that account-reflection boundary and keeps the 39.1 self-check auditable.
        if last.bucket is not None:
            total_debit = sum(row.debit for row in group)
            total_credit = sum(row.credit for row in group)
            balance = (
                total_credit - total_debit
                if uses_credit_balance(last)
                else total_debit - total_credit
            )
            group[-1] = replace(last, balance=balance)
        rows.extend(group)
    return rows
