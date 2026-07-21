"""회계항목 분류 규칙 회귀 테스트.

첨부된 "특관자_회계항목_분류규칙.xlsx"의 38.1/38.4 규칙을 기준으로 한다.
실샘플 기반 테스트는 금액을 하드코딩하지 않고 샘플 원장에서 기대값을 계산한다.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from app import excel_io
from app.config import Settings
from app.domain import columns as C
from app.domain.aggregate import AggregateResult, aggregate_balance, aggregate_ledger
from app.domain.errors import ErrorLog
from app.domain.period_extract import Period, extract_current_period
from app.normalize import normalize_names
from app.pipeline import _balance_long_frame, _best_frame, _check_prior_period, _distinct_partners, run_pipeline

SAMPLE_Q1_DIR = Path(__file__).resolve().parents[2] / "_sample_input"
SAMPLE_25_3Q_DIR = Path(__file__).resolve().parents[2] / "_sample_input_25.3Q"


def _sample_q1(prefix: str) -> Path | None:
    if not SAMPLE_Q1_DIR.is_dir():
        return None
    for p in sorted(SAMPLE_Q1_DIR.iterdir()):
        if p.name.startswith(prefix) and p.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}:
            return p
    return None


def _sample_25_3q(prefix: str) -> Path | None:
    if not SAMPLE_25_3Q_DIR.is_dir():
        return None
    for p in sorted(SAMPLE_25_3Q_DIR.iterdir()):
        if p.name.startswith(prefix) and p.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}:
            return p
    return None


def _sample_25_3q_slots() -> dict[str, Path | None]:
    return {
        "prev_balance": _sample_25_3q("1)"),
        "current_ledger": _sample_25_3q("2)"),
        "current_balance": _sample_25_3q("3)"),
        "name_pivot": _sample_25_3q("4)"),
        "prev_statement": _sample_25_3q("5)"),
    }


def test_381_keywords_follow_reference_workbook():
    assert C.match_bucket("보험료", C.OTHER_EXPENSE_KEYWORDS)
    assert C.match_bucket("건설중인자산_무형", C.ASSET_ACQUIRE_KEYWORDS)
    assert not C.match_bucket("임대료수익", C.SALES_KEYWORDS)
    assert not C.match_bucket("기업업무추진비", C.OTHER_EXPENSE_KEYWORDS)


def test_normalized_frame_prefers_accounting_header_labels_over_dense_data_rows():
    pd = pytest.importorskip("pandas")
    raw = pd.DataFrame(
        [
            ["미수수익명세서", None, None, None, None, None, None, None, None, None, None, None],
            ["(2026년 3월 31일 현재)", None, None, None, None, None, None, None, None, None, None, None],
            ["회사명", None, None, None, None, None, None, "(단위 : 원)", None, None, None, None],
            ["코드", "거래처", "내용", "이율", "금액", "대손충당금", "잔액", "비고", "사업장", None, None, None],
            ["89004", "앨리브랜즈주식회사", "단기대여금", "0.046", "5695178", "0", "5695178", None, "본사", "본사", "123", "456"],
        ]
    )

    normalized = excel_io.normalized_frame(raw)

    assert C.resolve_column(normalized, "partner") == "거래처"
    assert C.resolve_column(normalized, "amount") == "금액"


@pytest.mark.skipif(_sample_25_3q("2)") is None, reason="_sample_input_25.3Q 샘플 원장 없음")
def test_sample_25_3q_insurance_rows_feed_purchase_other():
    ledger_path = _sample_25_3q("2)")
    name_pivot_path = _sample_25_3q("4)")
    assert ledger_path is not None

    ledger = _best_frame(ledger_path)
    acc_col = C.resolve_column(ledger, "account")
    partner_col = C.resolve_column(ledger, "partner")
    debit_col = C.resolve_column(ledger, "debit")
    credit_col = C.resolve_column(ledger, "credit")
    assert acc_col and partner_col and debit_col and credit_col

    insurance = ledger[ledger[acc_col].astype(str).str.contains("보험료", na=False)].copy()
    assert not insurance.empty

    distinct_partners = sorted(
        {
            str(v).strip()
            for v in ledger[partner_col].tolist()
            if str(v).strip() and str(v).strip().lower() != "nan"
        }
    )
    norm = normalize_names(
        distinct_partners,
        name_pivot_path,
        Settings(openai_api_key=None, openai_model="gpt-4.1-mini"),
    )
    canonical = norm.canonical or set(norm.mapping.values())
    insurance = insurance[
        insurance[partner_col].map(lambda v: norm.mapping.get(str(v).strip()) in canonical)
    ].copy()
    assert not insurance.empty

    expected = sum(
        C.to_number(row[debit_col]) - C.to_number(row[credit_col])
        for _, row in insurance.iterrows()
    )
    assert abs(expected) > 0

    result = AggregateResult()
    aggregate_ledger(
        insurance,
        prev_balance=None,
        current_balance=None,
        mapping=norm.mapping,
        canonical=canonical,
        result=result,
        errors=ErrorLog(),
    )

    actual = sum(agg.purchase_other for agg in result.by_canonical.values())
    assert actual == pytest.approx(expected)


@pytest.mark.skipif(_sample_25_3q("2)") is None, reason="_sample_input_25.3Q 샘플 원장 없음")
def test_sample_25_3q_pipeline_produces_outputs(tmp_path):
    slot_paths = _sample_25_3q_slots()
    assert all(slot_paths.values())

    outcome = run_pipeline(
        slot_paths,
        Settings(openai_api_key=None, openai_model="gpt-4.1-mini"),
        tmp_path,
        {k: v.name for k, v in slot_paths.items() if v is not None},
        period=Period(2025, 3),
    )

    assert outcome.ok is True
    assert {"당기_특관자_명세서", "지배구조", "오류목록"} <= set(outcome.outputs)
    assert all(path.exists() for path in outcome.outputs.values())


@pytest.mark.skipif(_sample_25_3q("1)") is None, reason="_sample_input_25.3Q 전기 샘플 없음")
def test_sample_25_3q_prior_year_end_is_valid_prev_under_final_rule():
    slot_paths = _sample_25_3q_slots()
    errors = ErrorLog()

    _check_prior_period(
        slot_paths,
        {k: v.name for k, v in slot_paths.items() if v is not None},
        errors,
        selected_period=Period(2025, 3),
    )

    assert errors.count == 0


def test_fund_lending_ignores_prior_half_carryforward_block():
    ledger = pytest.importorskip("pandas").DataFrame(
        [
            {
                "계정명": "단기대여금",
                "거래처명": "특관자A",
                "차변": "1000",
                "대변": "0",
                "적요": "전반기 자금대여 블록",
            },
            {
                "계정명": "단기대여금",
                "거래처명": "특관자A",
                "차변": "2000",
                "대변": "0",
                "적요": "당기 신규 대여",
            },
        ]
    )
    result = AggregateResult()
    aggregate_ledger(
        ledger,
        prev_balance=None,
        current_balance=None,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").fund_lending == 2000


def test_aggregate_ledger_avoids_iterrows_for_period_metadata(monkeypatch):
    pd = pytest.importorskip("pandas")
    ledger = pd.DataFrame(
        [
            {
                "계정명": "상품매출",
                "거래처명": "특관자A",
                "차변": "0",
                "대변": "100",
            }
        ]
    )
    metadata = [(2026, 6)]
    attrs = ledger.attrs
    ledger.attrs["period_year_months"] = metadata

    def fail_if_iterrows_called(self):
        raise AssertionError("ledger aggregation must not use DataFrame.iterrows()")

    monkeypatch.setattr(pd.DataFrame, "iterrows", fail_if_iterrows_called)

    result = AggregateResult()
    aggregate_ledger(
        ledger,
        prev_balance=None,
        current_balance=None,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert ledger.attrs is attrs
    assert ledger.attrs["period_year_months"] is metadata
    assert result.get("특관자A").sales == 100.0


def test_fund_lending_ignores_allowance_accounts():
    ledger = pytest.importorskip("pandas").DataFrame(
        [
            {
                "계정명": "대손충당금(단기대여금)",
                "거래처명": "특관자A",
                "차변": "1000",
                "대변": "0",
            },
            {
                "계정명": "장기대여금",
                "거래처명": "특관자A",
                "차변": "2000",
                "대변": "0",
            },
        ]
    )
    result = AggregateResult()
    aggregate_ledger(
        ledger,
        prev_balance=None,
        current_balance=None,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").fund_lending == 2000


def test_aggregate_row_eligibility_excludes_nonqualifying_detail_categories():
    pd = pytest.importorskip("pandas")
    ledger = pd.DataFrame(
        [
            {"계정명": "단기대여금", "거래처명": "특관자A", "차변": 100, "대변": 0, "적요": "전기 이월"},
            {"계정명": "단기대여금", "거래처명": "특관자A", "차변": 0, "대변": 100, "적요": "상환"},
            {"계정명": "대손충당금(단기대여금)", "거래처명": "특관자A", "차변": 100, "대변": 0, "적요": "충당금"},
            {"계정명": "단기대여금", "거래처명": "특관자A", "차변": 100, "대변": 0, "적요": "신규 대여"},
            {"계정명": "지급임차료", "거래처명": "특관자A", "차변": 0, "대변": 100, "적요": "전대 차감"},
            {"계정명": "보험료", "거래처명": "특관자A", "차변": 100, "대변": 0, "적요": "정상 비용"},
        ]
    )
    result = AggregateResult()

    aggregate_ledger(
        ledger,
        prev_balance=None,
        current_balance=None,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").fund_lending == 100
    assert result.get("특관자A").purchase_other == 100


def test_interest_income_is_not_offset_when_accrued_balance_source_is_absent():
    pd = pytest.importorskip("pandas")
    ledger = pd.DataFrame(
        [
            {"계정명": "미수수익", "거래처명": "특관자A", "차변": "100", "대변": "0"},
            {"계정명": "이자수익", "거래처명": "특관자A", "차변": "0", "대변": "100"},
        ]
    )
    result = AggregateResult()

    aggregate_ledger(
        ledger,
        prev_balance=None,
        current_balance=None,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").sales_other == 100


def test_accrued_income_adjustment_ignores_allowance_account_entries():
    pd = pytest.importorskip("pandas")
    ledger = pd.DataFrame(
        [
            {"계정명": "미수수익", "거래처명": "특관자A", "차변": "100", "대변": "0"},
            {"계정명": "대손충당금(미수수익)", "거래처명": "특관자A", "차변": "0", "대변": "100"},
            {"계정명": "이자수익", "거래처명": "특관자A", "차변": "0", "대변": "100"},
        ]
    )
    prev_balance = pd.DataFrame([{"거래처": "특관자A", "계정과목": "미수수익", "금액": 0}])
    current_balance = pd.DataFrame([{"거래처": "특관자A", "계정과목": "미수수익", "금액": 100}])
    result = AggregateResult()

    aggregate_ledger(
        ledger,
        prev_balance=prev_balance,
        current_balance=current_balance,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").sales_other == 100


@pytest.mark.skipif(_sample_q1("2)") is None, reason="_sample_input 샘플 원장 없음")
def test_sample_q1_sales_other_keeps_interest_income_and_excludes_lux_allowance():
    settings = Settings(openai_api_key=None, openai_model="gpt-4.1-mini")
    prev_balance = _balance_long_frame(_sample_q1("1)"))
    current_balance = _balance_long_frame(_sample_q1("3)"))
    ledger = _best_frame(_sample_q1("2)"))
    extracted = extract_current_period(ledger, Period(2026, 1))
    assert extracted.ok is True

    norm = normalize_names(
        _distinct_partners([extracted.df, current_balance, prev_balance]),
        _sample_q1("4)"),
        settings,
    )
    canonical = norm.canonical or set(norm.mapping.values())
    result = AggregateResult()
    aggregate_balance(current_balance, prev_balance, norm.mapping, canonical, result, ErrorLog())
    aggregate_ledger(extracted.df, prev_balance, current_balance, norm.mapping, canonical, result, ErrorLog())

    assert result.get("(주)프레시코").sales_other == 147_338_630
    assert result.get("(주)룩스앤메이코스메틱").sales_other == 0


def test_allowance_decrease_is_not_reclassified_to_sales_other():
    pd = pytest.importorskip("pandas")
    prev_balance = pd.DataFrame(
        [
            {"거래처": "특관자A", "계정과목": "미수수익", "금액": 7_264_104_108},
            {"거래처": "특관자A", "계정과목": "대손충당금", "금액": 7_264_104_108},
        ]
    )
    current_balance = pd.DataFrame(
        [
            {"거래처": "특관자A", "계정과목": "미수수익", "금액": 6_500_000_000},
            {"거래처": "특관자A", "계정과목": "대손충당금", "금액": 6_500_000_000},
        ]
    )
    ledger = pd.DataFrame(columns=["계정명", "거래처명", "차변", "대변"])
    result = AggregateResult()
    errors = ErrorLog()
    aggregate_balance(
        current_balance,
        prev_balance,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=errors,
    )

    aggregate_ledger(
        ledger,
        prev_balance=prev_balance,
        current_balance=current_balance,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=errors,
    )

    assert result.get("특관자A").sales_other == 0


def test_allowance_reversal_stays_in_allowance_expense_as_negative():
    pd = pytest.importorskip("pandas")
    prev_balance = pd.DataFrame(
        [{"거래처": "특관자A", "계정과목": "대손충당금", "금액": 7_264_104_108}]
    )
    current_balance = pd.DataFrame(
        [{"거래처": "특관자A", "계정과목": "대손충당금", "금액": 6_500_000_000}]
    )
    result = AggregateResult()

    aggregate_balance(
        current_balance,
        prev_balance,
        mapping={"특관자A": "특관자A"},
        canonical={"특관자A"},
        result=result,
        errors=ErrorLog(),
    )

    assert result.get("특관자A").allowance_end == 6_500_000_000
    assert result.get("특관자A").allowance_expense == -764_104_108
