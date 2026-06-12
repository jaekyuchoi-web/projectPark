"""run_pipeline 의 당기 추출 연동 테스트 (외부 서비스 불필요)."""

from __future__ import annotations

from openpyxl import Workbook

from app.config import Settings
from app.domain.period_extract import Period
from app.pipeline import run_pipeline

_SETTINGS = Settings(openai_api_key=None, openai_model="m")


def _make_ledger(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(["계정명", "날짜", "거래처명", "차변", "대변"])
    for r in rows:
        ws.append(r)
    wb.save(path)


def _slots(ledger=None):
    return {
        "prev_balance": None,
        "current_ledger": ledger,
        "current_balance": None,
        "name_pivot": None,
        "prev_statement": None,
    }


def test_run_pipeline_blocks_on_unparseable_ledger_date(tmp_path):
    # 날짜 해석 불가 셀 + AI 키 없음 → 엄격 차단(다운로드 차단)
    led = tmp_path / "ledger.xlsx"
    _make_ledger(led, [
        ["현금", "2026-01-10", "거래처A", "1000", "0"],
        ["현금", "날짜미상", "거래처B", "1000", "0"],
    ])
    outcome = run_pipeline(_slots(led), _SETTINGS, tmp_path, period=Period(2026, 1))
    assert outcome.ok is False
    assert "날짜" in outcome.message
    assert "오류목록" in outcome.outputs


def test_run_pipeline_no_block_when_dates_clean(tmp_path):
    # 모든 날짜 파싱 가능 → 추출 단계에서 차단되지 않음(이후 단계로 진행)
    led = tmp_path / "ledger.xlsx"
    _make_ledger(led, [
        ["현금", "2026-01-10", "거래처A", "1000", "0"],
        ["현금", "2026-02-20", "거래처B", "1000", "0"],
    ])
    outcome = run_pipeline(_slots(led), _SETTINGS, tmp_path, period=Period(2026, 1))
    # 날짜 차단 메시지가 아니어야 한다(다른 사유로 ok 여부는 무관)
    assert "날짜" not in outcome.message
