"""run_pipeline 의 당기 추출 연동 테스트 (외부 서비스 불필요)."""

from __future__ import annotations

from openpyxl import Workbook, load_workbook

from app.config import Settings
from app.domain.period_extract import Period, parse_year_month
from app.domain.statement_detail import StatementDetailError
from app.output_check import CheckResult
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


def _make_period_template(path):
    wb = Workbook()
    wb.active.title = "특관자"
    detail = wb.create_sheet("39.1")
    pivot = wb.create_sheet("39.2")
    detail["B15"] = "매입매출"
    detail["F15"] = "구분계정과목"
    detail["I15"] = "날짜"
    detail["J15"] = "적요"
    detail["J16"] = "OLD_Q1_SENTINEL"
    detail["P12"] = "=SUM(P16:P510)"
    pivot["D3"] = "특관자A"
    wb.save(path)


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


def test_run_pipeline_q2_uses_january_through_june_for_summary_and_detail(
    tmp_path, monkeypatch
):
    from app import pipeline
    from app.domain import statement

    template = tmp_path / "template.xlsx"
    _make_period_template(template)
    monkeypatch.setattr(statement, "TEMPLATE_PATH", template)
    monkeypatch.setattr(
        pipeline,
        "verify_output",
        lambda _: CheckResult(ok=True, reason="test", recalculated=False),
    )
    seen: dict[str, object] = {}
    original_detail = pipeline.build_statement_detail
    original_aggregate = pipeline.aggregate_ledger

    def capture_detail(ledger, **kwargs):
        seen["detail"] = ledger
        return original_detail(ledger, **kwargs)

    def capture_aggregate(ledger, *args, **kwargs):
        seen["aggregate"] = ledger
        return original_aggregate(ledger, *args, **kwargs)

    monkeypatch.setattr(pipeline, "build_statement_detail", capture_detail)
    monkeypatch.setattr(pipeline, "aggregate_ledger", capture_aggregate)
    ledger = tmp_path / "ledger.xlsx"
    _make_ledger(
        ledger,
        [
            ["상품매출", "2026-01-03", "특관자A", "0", "100"],
            ["상품매출", "2026-06-30", "특관자A", "0", "200"],
            ["상품매출", "2026-07-01", "특관자A", "0", "400"],
        ],
    )

    outcome = run_pipeline(
        _slots(ledger),
        _SETTINGS,
        tmp_path,
        period=Period(2026, 2),
    )

    assert outcome.ok is True
    assert seen["aggregate"] is seen["detail"]
    statement_path = outcome.outputs["당기_특관자_명세서"]
    assert statement_path.name == "당기_특관자_명세서_2026_2Q.xlsx"
    workbook = load_workbook(statement_path, data_only=False)
    detail = workbook["39.1"]
    assert [
        parse_year_month(detail["I16"].value),
        parse_year_month(detail["I17"].value),
    ] == [
        (2026, 1),
        (2026, 6),
    ]
    assert detail["I18"].value is None
    assert detail["A1"].value == "2026년 2분기 누적 (1~6월)"
    assert workbook["39.2"]["E3"].value == 300.0
    assert "OLD_Q1_SENTINEL" not in {
        detail.cell(row=row, column=10).value
        for row in range(16, detail.max_row + 1)
    }
    workbook.close()


def test_run_pipeline_blocks_detail_failure_without_leaking_accounting_values(
    tmp_path, monkeypatch
):
    from app import pipeline

    secret = "sensitive transaction description 12345"
    ledger = tmp_path / "ledger.xlsx"
    _make_ledger(ledger, [["상품매출", "2026-06-30", "특관자A", "0", "100"]])

    def fail_detail(*args, **kwargs):
        raise StatementDetailError(secret)

    monkeypatch.setattr(pipeline, "build_statement_detail", fail_detail)

    outcome = run_pipeline(_slots(ledger), _SETTINGS, tmp_path, period=Period(2026, 2))

    assert outcome.ok is False
    assert outcome.message == "당기 상세 거래 검증에 실패하여 다운로드를 차단했습니다."
    assert set(outcome.outputs) == {"오류목록"}
    error_book = load_workbook(outcome.outputs["오류목록"], data_only=True)
    error_values = {
        str(cell.value)
        for row in error_book.active.iter_rows()
        for cell in row
        if cell.value is not None
    }
    error_book.close()
    assert "당기 상세 검증 실패" in error_values
    assert secret not in error_values
