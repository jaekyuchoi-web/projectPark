"""회계항목 분류 규칙 회귀 테스트.

첨부된 "특관자_회계항목_분류규칙.xlsx"의 38.1/38.4 규칙을 기준으로 한다.
실샘플 기반 테스트는 금액을 하드코딩하지 않고 샘플 원장에서 기대값을 계산한다.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from app.config import Settings
from app.domain import columns as C
from app.domain.aggregate import AggregateResult, aggregate_ledger
from app.domain.errors import ErrorLog
from app.domain.period_extract import Period
from app.normalize import normalize_names
from app.pipeline import _best_frame, run_pipeline

SAMPLE_25_3Q_DIR = Path(__file__).resolve().parents[2] / "_sample_input_25.3Q"


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
