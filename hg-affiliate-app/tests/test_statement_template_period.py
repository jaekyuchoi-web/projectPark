from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile

import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from app.domain.aggregate import AggregateResult
from app.domain.period_extract import Period
from app.domain.statement_detail import StatementDetailRow
from app.domain.statement_template import fill_statement_template


def _write_template(path: Path) -> None:
    wb = Workbook()
    wb.active.title = "특관자"
    detail = wb.create_sheet("39.1")
    pivot = wb.create_sheet("39.2")
    detail["B15"] = "매입매출"
    detail["F15"] = "구분계정과목"
    detail["I15"] = "날짜"
    detail["J15"] = "적요"
    detail["B16"] = "OLD_Q1_SENTINEL"
    detail["J16"] = "OLD_Q1_SENTINEL"
    detail["B16"].fill = PatternFill("solid", fgColor="FFFF00")
    detail["C4"] = "=SUMIFS($N$16:$N$510,$F$16:$F$510,$B4)"
    detail["P12"] = "=SUM(P16:P510)"
    pivot["D3"] = "특관자A"
    wb.save(path)


def _row(date: str, description: str) -> StatementDetailRow:
    return StatementDetailRow(
        sales_purchase="매출",
        funding=None,
        receivable_payable=None,
        income_expense=None,
        bucket="매출",
        account_code="40100",
        account_name="상품매출",
        date=date,
        description=description,
        partner_code="V001",
        partner_name="특관자A",
        canonical_name="특관자A",
        debit=0.0,
        credit=100.0,
        balance=100.0,
    )


def test_template_replaces_stale_q1_detail_with_q2_ytd_rows(tmp_path):
    template = tmp_path / "template.xlsx"
    _write_template(template)

    wb, unmatched = fill_statement_template(
        template,
        AggregateResult(),
        Period(2026, 2),
        [_row("2026-01-03", "2026.1Q 정상 적요"), _row("2026-06-30", "June")],
    )

    detail = wb["39.1"]
    assert unmatched == []
    assert detail["A1"].value == "2026년 2분기 누적 (1~6월)"
    assert detail["I16"].value == "2026-01-03"
    assert detail["I17"].value == "2026-06-30"
    assert detail["J16"].value == "2026.1Q 정상 적요"
    assert "OLD_Q1_SENTINEL" not in {
        detail.cell(row=row, column=column).value
        for row in range(16, detail.max_row + 1)
        for column in range(2, 17)
    }
    assert detail["C4"].value == "=SUMIFS($N$16:$N$17,$F$16:$F$17,$B4)"
    assert detail["P12"].value == "=SUM(P16:P17)"
    assert detail["B17"].fill.fgColor.rgb == detail["B16"].fill.fgColor.rgb


def test_existing_broken_template_is_not_silently_hidden_by_fallback(tmp_path, monkeypatch):
    from app.domain import statement

    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not an xlsx")
    monkeypatch.setattr(statement, "TEMPLATE_PATH", broken)

    with pytest.raises(BadZipFile):
        statement.build_statement(AggregateResult(), Period(2026, 2), [])
