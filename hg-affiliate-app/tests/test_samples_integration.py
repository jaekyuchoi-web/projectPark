"""실제 _sample_input 파일을 사용한 통합 TDD.

파일명 접두('1)'~'5)')로 입력 소스 종류를 식별한다. 샘플은 민감데이터라
저장소에 없을 수 있으므로, 없으면 전체 모듈을 skip 한다.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from app.config import Settings
from app.domain import columns as C
from app.domain.errors import ErrorLog
from app.domain.period_extract import Period, extract_current_period
from app.pipeline import _best_frame, _check_prior_period, run_pipeline

SAMPLE_DIR = Path(__file__).resolve().parents[2] / "_sample_input"


def _sample(prefix: str) -> Path | None:
    if not SAMPLE_DIR.is_dir():
        return None
    for p in sorted(SAMPLE_DIR.glob("*.xlsx")):
        if p.name.startswith(prefix):
            return p
    return None


PREV_BAL = _sample("1)")     # 전기이월소스 (25.4Q 잔액명세서)
LEDGER = _sample("2)")       # 당기거래내역소스 (26.1Q 계정별원장)
CUR_BAL = _sample("3)")      # 채권채무잔액검증용 (26.1Q 잔액명세서)
NAME_PIVOT = _sample("4)")   # 특관자상호정리
PREV_STMT = _sample("5)")    # 전기 특관자명세서

pytestmark = pytest.mark.skipif(LEDGER is None, reason="_sample_input 샘플 파일 없음")


def test_real_ledger_date_column_resolves():
    # 실제 원장의 날짜 컬럼명은 "날짜" — date 별칭에 포함되어 해석되어야 한다
    df = _best_frame(LEDGER)
    assert C.resolve_column(df, "date") == "날짜"


def test_real_ledger_extract_q1():
    # 26.1Q 원장 → 당기(2026 1Q) 추출: 전 거래행 채택, 기간외 0, 빈(소계)행 제외
    df = _best_frame(LEDGER)
    r = extract_current_period(df, Period(2026, 1))
    assert r.ok is True
    assert r.dropped_out_of_period == 0
    assert r.kept > 0
    # 회계 항등식: 채택 + 기간외 + 빈행 = 전체 행수
    assert r.kept + r.dropped_out_of_period + r.dropped_blank == len(df)


def test_real_ledger_extract_q4_superset_of_q1():
    # 4Q(1~12월)는 1Q를 포함 → 1Q와 동일하게 전 거래행 채택(이 원장은 1~3월뿐)
    df = _best_frame(LEDGER)
    q1 = extract_current_period(df, Period(2026, 1))
    q4 = extract_current_period(df, Period(2026, 4))
    assert q4.ok is True
    assert q4.kept == q1.kept
    assert q4.dropped_out_of_period == 0


def test_real_ledger_wrong_year_extracts_nothing():
    # 2025년을 당기로 선택하면 이 원장(2026)에서 채택 0 (차단 아님, 정상 추출)
    df = _best_frame(LEDGER)
    r = extract_current_period(df, Period(2025, 4))
    assert r.ok is True
    assert r.kept == 0


# ── 전기 가드: 실제 파일명 기준(25.4Q vs 26.1Q) ────────────────────
def _guard_count(prev_path, prev_name, selected_year):
    e = ErrorLog()
    _check_prior_period(
        {"prev_balance": prev_path, "current_balance": CUR_BAL},
        {"prev_balance": prev_name, "current_balance": CUR_BAL.name if CUR_BAL else None},
        e, selected_year=selected_year,
    )
    return e.count


def test_real_prev_25q4_no_false_warning():
    # 정상: 전기=25.4Q, 당기 2026 → 경고 0 (stale 표제 날짜에 오탐하지 않음)
    assert _guard_count(PREV_BAL, PREV_BAL.name, 2026) == 0


def test_real_wrong_prev_same_quarter_warns():
    # 잘못: 전기 슬롯에 26.1Q(동일 분기)를 올림 → 분기 표식 경고
    assert _guard_count(CUR_BAL, CUR_BAL.name, 2026) >= 1


# ── 전체 파이프라인 해피패스(실샘플) ───────────────────────────────
@pytest.mark.skipif(CUR_BAL is None, reason="당기말 잔액 샘플 없음")
def test_real_pipeline_q1_produces_outputs(tmp_path):
    slot_paths = {
        "prev_balance": PREV_BAL, "current_ledger": LEDGER, "current_balance": CUR_BAL,
        "name_pivot": NAME_PIVOT, "prev_statement": PREV_STMT,
    }
    slot_filenames = {k: (v.name if v else None) for k, v in slot_paths.items()}
    settings = Settings(openai_api_key=None, openai_model="gpt-4.1-mini")  # 키 없음 → 휴리스틱
    outcome = run_pipeline(slot_paths, settings, tmp_path, slot_filenames, period=Period(2026, 1))
    assert outcome.ok is True
    assert "날짜" not in outcome.message  # 날짜 추출에서 차단되지 않음
    assert "당기_특관자_명세서" in outcome.outputs
    for p in outcome.outputs.values():
        assert Path(p).exists()
