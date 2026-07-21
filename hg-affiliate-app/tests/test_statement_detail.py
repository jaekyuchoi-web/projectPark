from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from app.domain.period_extract import Period, extract_current_period, parse_year_month
from app.domain.statement_detail import (
    StatementDetailError,
    build_statement_detail,
)


def _ledger(rows: list[tuple[str, str, str]]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "계정코드": "40100",
                "계정명": account,
                "날짜": date,
                "적요": description,
                "거래처코드": "V001",
                "거래처명": "특관자A",
                "차변": "0",
                "대변": "100",
                "잔액": "100",
            }
            for date, account, description in rows
        ]
    )


def test_q2_detail_is_january_through_june_not_april_through_june():
    ledger = _ledger(
        [
            ("2026-01-03", "상품매출", "January"),
            ("2026-03-31", "상품매출", "March"),
            ("2026-04-01", "상품매출", "April"),
            ("2026-06-30", "상품매출", "June"),
            ("2026-07-01", "상품매출", "July"),
            ("2025-12-31", "상품매출", "Prior year"),
        ]
    )
    extracted = extract_current_period(ledger, Period(2026, 2))
    assert extracted.ok is True

    rows = build_statement_detail(
        extracted.df,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert [parse_year_month(row.date) for row in rows] == [
        (2026, 1),
        (2026, 3),
        (2026, 4),
        (2026, 6),
    ]


def test_detail_uses_ai_rescued_period_metadata_without_reparsing():
    ledger = _ledger([("이천이십육년 유월", "상품매출", "AI rescued")])
    extracted = extract_current_period(
        ledger,
        Period(2026, 2),
        ai_parser=lambda values: {value: (2026, 6) for value in values},
    )
    assert extracted.ok is True

    rows = build_statement_detail(
        extracted.df,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert len(rows) == 1
    assert rows[0].date == "이천이십육년 유월"


def test_detail_excludes_non_related_parties():
    ledger = _ledger([("2026-05-01", "상품매출", "related")])
    general = ledger.copy()
    general.loc[0, "거래처명"] = "일반거래처"
    combined = pd.concat([ledger, general], ignore_index=True)

    rows = build_statement_detail(
        combined,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert len(rows) == 1
    assert rows[0].canonical_name == "특관자A"


def test_detail_maps_lending_row_to_existing_b_through_p_contract():
    ledger = _ledger([("2026-06-30", "단기대여금", "2026.1Q 정상 적요")])
    ledger.loc[0, "차변"] = "1,200"
    ledger.loc[0, "대변"] = "200"
    ledger.loc[0, "잔액"] = "1,000"

    row = build_statement_detail(
        ledger,
        mapping={"특관자A": "정규특관자A"},
        canonical={"정규특관자A"},
        period=Period(2026, 2),
    )[0]

    assert row.as_excel_row() == [
        None,
        "자금대여",
        "채권",
        None,
        "대여금",
        "40100",
        "단기대여금",
        dt.date(2026, 6, 30),
        "2026.1Q 정상 적요",
        "V001",
        "특관자A",
        "정규특관자A",
        1200.0,
        200.0,
        1000.0,
    ]


def test_detail_fails_closed_when_given_out_of_period_row():
    ledger = _ledger([("2026-07-01", "상품매출", "outside")])

    with pytest.raises(StatementDetailError, match="기간 밖"):
        build_statement_detail(
            ledger,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )


def test_detail_fails_closed_when_required_column_is_missing():
    ledger = _ledger([("2026-06-30", "상품매출", "missing")]).drop(columns=["차변"])

    with pytest.raises(StatementDetailError, match="필수 열"):
        build_statement_detail(
            ledger,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )


@pytest.mark.parametrize(
    "metadata",
    [
        (2026,),
        "not-a-pair",
        (2026, 13),
        ("2026", 6),
    ],
)
def test_detail_fails_closed_when_period_metadata_entry_is_malformed(metadata):
    ledger = _ledger([("2026-06-30", "상품매출", "do not leak this description")])
    ledger.attrs["period_year_months"] = [metadata]

    with pytest.raises(StatementDetailError, match="기간 메타데이터") as exc_info:
        build_statement_detail(
            ledger,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )

    assert "do not leak this description" not in str(exc_info.value)
