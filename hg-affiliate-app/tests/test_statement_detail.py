from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from app import excel_io
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


class _DeepcopyBomb:
    def __deepcopy__(self, memo):
        raise AssertionError("filtered ledger attrs must not be deep-copied")


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


def test_detail_avoids_iterrows_and_preserves_row_aligned_period_metadata(
    monkeypatch,
):
    ledger = _ledger(
        [
            ("2026-01-03", "상품매출", "first"),
            ("2026-06-30", "단기대여금", "second"),
            ("이천이십육년 유월", "상품매출", "AI rescued"),
        ]
    )
    ledger.loc[:, ["차변", "대변", "잔액"]] = [
        ["0", "100", "100"],
        ["1,200", "200", "1,000"],
        ["30", "0", "30"],
    ]
    metadata = [(2026, 1), (2026, 6), (2026, 6)]
    attrs = ledger.attrs
    ledger.attrs["period_year_months"] = metadata
    bomb = _DeepcopyBomb()
    ledger.attrs["copy_bomb"] = bomb

    def fail_if_iterrows_called(self):
        raise AssertionError("detail construction must not use DataFrame.iterrows()")

    monkeypatch.setattr(pd.DataFrame, "iterrows", fail_if_iterrows_called)

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert ledger.attrs is attrs
    assert ledger.attrs["period_year_months"] is metadata
    assert ledger.attrs["copy_bomb"] is bomb
    assert [
        (
            row.date,
            row.sales_purchase,
            row.funding,
            row.debit,
            row.credit,
            row.balance,
        )
        for row in rows
    ] == [
        (dt.date(2026, 1, 3), "매출", None, 0.0, 100.0, None),
        ("이천이십육년 유월", "매출", None, 30.0, 0.0, 70.0),
        (dt.date(2026, 6, 30), None, "자금대여", 1200.0, 200.0, 1000.0),
    ]


@pytest.mark.parametrize(
    "column,position",
    [
        ("차변", 7),
        ("거래처명", 6),
        ("적요", 4),
    ],
)
def test_detail_fails_closed_for_duplicate_selected_column(column, position):
    ledger = _ledger([("2026-06-30", "상품매출", "duplicate")])
    ledger.insert(position, column, "duplicate", allow_duplicates=True)

    with pytest.raises(StatementDetailError, match="필수 열") as exc_info:
        build_statement_detail(
            ledger,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )

    assert column not in str(exc_info.value)


@pytest.mark.parametrize("duplicate", ["차변", "거래처명", "적요"])
def test_detail_rejects_deduplicated_selected_source_headers_after_extraction(
    duplicate,
):
    headers = ["계정명", "날짜", "거래처명", "적요", "차변", "대변"]
    values = ["상품매출", "2026-06-30", "특관자A", "source duplicate", "0", "100"]
    position = headers.index(duplicate) + 1
    headers.insert(position, duplicate)
    values.insert(position, "duplicate")
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))

    assert normalized.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset(
        {duplicate}
    )
    extracted = extract_current_period(normalized, Period(2026, 2))
    assert extracted.ok is True
    assert extracted.df.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset(
        {duplicate}
    )

    with pytest.raises(StatementDetailError, match="필수 열") as exc_info:
        build_statement_detail(
            extracted.df,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )

    assert duplicate not in str(exc_info.value)


@pytest.mark.parametrize(
    ("first_header", "second_header", "first_value", "second_value"),
    [
        ("적요", "적 요", "PRIVATE_DESCRIPTION_ONE", "PRIVATE_DESCRIPTION_TWO"),
    ],
)
def test_detail_rejects_whitespace_equivalent_selected_headers_after_extraction(
    first_header,
    second_header,
    first_value,
    second_value,
):
    headers = [
        "계정명",
        "날짜",
        "거래처명",
        "차변",
        "대변",
        first_header,
        second_header,
    ]
    values = [
        "상품매출",
        "2026-06-30",
        "특관자A",
        "0",
        "100",
        first_value,
        second_value,
    ]
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))

    assert normalized.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset(
        {first_header, second_header}
    )
    extracted = extract_current_period(normalized, Period(2026, 2))
    assert extracted.ok is True
    assert extracted.df.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset(
        {first_header, second_header}
    )

    with pytest.raises(StatementDetailError, match="필수 열") as exc_info:
        build_statement_detail(
            extracted.df,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )

    error = str(exc_info.value)
    assert first_header not in error
    assert second_header not in error
    assert first_value not in error
    assert second_value not in error


def test_detail_allows_duplicate_source_running_balance_headers_because_they_are_ignored():
    headers = ["계정명", "날짜", "거래처명", "차변", "대변", "잔액", "잔 액"]
    values = ["상품매출", "2026-06-30", "특관자A", "0", "100", "999", "888"]
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))
    extracted = extract_current_period(normalized, Period(2026, 2))

    rows = build_statement_detail(
        extracted.df,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert len(rows) == 1
    assert rows[0].balance == 100.0


def test_detail_allows_unique_and_irrelevant_deduplicated_source_headers():
    headers = ["계정명", "날짜", "거래처명", "적요", "차변", "대변", "비고", "비고"]
    values = ["상품매출", "2026-06-30", "특관자A", "normal", "0", "100", "", ""]
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))
    extracted = extract_current_period(normalized, Period(2026, 2))

    assert normalized.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset({"비고"})
    assert extracted.ok is True
    assert len(
        build_statement_detail(
            extracted.df,
            mapping={"특관자A": "특관자A"},
            canonical={"특관자A"},
            period=Period(2026, 2),
        )
    ) == 1


def test_detail_allows_whitespace_equivalent_irrelevant_source_headers():
    headers = ["계정명", "날짜", "거래처명", "적요", "차변", "대변", "비고", "비 고"]
    values = ["상품매출", "2026-06-30", "특관자A", "normal", "0", "100", "one", "two"]
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))
    extracted = extract_current_period(normalized, Period(2026, 2))

    assert normalized.attrs[excel_io.DEDUPLICATED_HEADER_BASES_ATTR] == frozenset(
        {"비고", "비 고"}
    )
    assert extracted.ok is True
    rows = build_statement_detail(
        extracted.df,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert len(rows) == 1


def test_normalized_unique_source_headers_do_not_set_duplicate_contract():
    headers = ["계정명", "날짜", "거래처명", "적요", "차변", "대변"]
    values = ["상품매출", "2026-06-30", "특관자A", "normal", "0", "100"]
    normalized = excel_io.normalized_frame(pd.DataFrame([headers, values]))

    assert excel_io.DEDUPLICATED_HEADER_BASES_ATTR not in normalized.attrs


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


def test_detail_ignores_source_running_balance_and_calculates_group_balance():
    ledger = _ledger(
        [
            ("2026-01-03", "상품매출", "first sale"),
            ("2026-06-30", "상품매출", "second sale"),
        ]
    )
    ledger.loc[:, ["차변", "대변", "잔액"]] = [
        ["0", "100", "9,999"],
        ["20", "0", "8,888"],
    ]

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert [row.balance for row in rows] == [None, 80.0]


def test_detail_does_not_reflect_unclassified_account_in_balance_column():
    ledger = _ledger(
        [
            ("2026-01-03", "분류대상아님", "first"),
            ("2026-06-30", "분류대상아님", "second"),
        ]
    )
    ledger.loc[:, ["차변", "대변", "잔액"]] = [
        ["100", "0", "9,999"],
        ["0", "20", "8,888"],
    ]

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    assert all(row.bucket is None for row in rows)
    assert [row.balance for row in rows] == [None, None]


def test_detail_adds_asset_opening_to_debit_and_keeps_only_terminal_balance():
    ledger = _ledger(
        [
            ("2026-01-03", "외상매출금", "invoice"),
            ("2026-06-30", "외상매출금", "collection"),
        ]
    )
    ledger.loc[:, ["계정코드", "차변", "대변", "잔액"]] = [
        ["1080000", "100", "0", "9,999"],
        ["1080000", "0", "40", "8,888"],
    ]
    opening = pd.DataFrame(
        [
            {
                "계정과목": "외상매출금",
                "거래처": "특관자A",
                "거래처코드": "V001",
                "금액": "500",
            }
        ]
    )

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
        prev_balance=opening,
    )

    assert len(rows) == 3
    assert rows[0].date == "[전기이월]"
    assert rows[0].account_code == "10800"
    assert rows[0].debit == 500.0
    assert rows[0].credit == 0.0
    assert [row.balance for row in rows] == [None, None, 560.0]


def test_detail_adds_liability_opening_to_credit_and_uses_credit_nature_balance():
    ledger = _ledger([("2026-06-30", "미지급금", "payment")])
    ledger.loc[0, ["계정코드", "차변", "대변", "잔액"]] = [
        "2530000",
        "200",
        "0",
        "7,777",
    ]
    opening = pd.DataFrame(
        [
            {
                "계정과목": "미지급금",
                "거래처": "특관자A",
                "거래처코드": "V001",
                "금액": "500",
            }
        ]
    )

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
        prev_balance=opening,
    )

    assert rows[0].debit == 0.0
    assert rows[0].credit == 500.0
    assert [row.balance for row in rows] == [None, 300.0]


def test_detail_merges_opening_aliases_only_within_same_partner_code_and_account():
    ledger = _ledger([("2026-06-30", "외상매출금", "collection")])
    ledger.loc[0, ["계정코드", "거래처코드", "거래처명", "차변", "대변"]] = [
        "1080000",
        "V001",
        "특관자A",
        "0",
        "150",
    ]
    opening = pd.DataFrame(
        [
            {
                "계정과목": "외상매출금",
                "거래처": "특관자A",
                "거래처코드": "V001",
                "금액": "100",
            },
            {
                "계정과목": "외상매출금",
                "거래처": "특관자㈜",
                "거래처코드": "V001",
                "금액": "50",
            },
        ]
    )

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A", "특관자㈜": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
        prev_balance=opening,
    )

    assert len(rows) == 2
    assert rows[0].date == "[전기이월]"
    assert rows[0].debit == 150.0
    assert rows[1].balance == 0.0


def test_detail_row_categories_follow_aggregate_row_eligibility_rules():
    ledger = _ledger(
        [
            ("2026-06-01", "단기대여금", "전기 이월"),
            ("2026-06-02", "단기대여금", "대여금 상환"),
            ("2026-06-03", "대손충당금(단기대여금)", "충당금"),
            ("2026-06-04", "단기대여금", "당기 신규 대여"),
            ("2026-06-05", "지급임차료", "전대 차감"),
            ("2026-06-06", "보험료", "정상 비용"),
        ]
    )
    ledger.loc[:, ["차변", "대변"]] = [
        ["100", "0"],
        ["0", "100"],
        ["100", "0"],
        ["100", "0"],
        ["0", "100"],
        ["100", "0"],
    ]

    rows = build_statement_detail(
        ledger,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        period=Period(2026, 2),
    )

    by_description = {row.description: row for row in rows}
    assert by_description["전기 이월"].funding is None
    assert by_description["대여금 상환"].funding is None
    assert by_description["충당금"].funding is None
    assert by_description["당기 신규 대여"].funding == "자금대여"
    assert by_description["전대 차감"].income_expense is None
    assert by_description["전대 차감"].bucket is None
    assert by_description["정상 비용"].income_expense == "비용"
    assert by_description["정상 비용"].bucket == "기타비용"


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
