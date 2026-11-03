from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from app.domain.aggregate import AggregateResult
from app.domain.period_extract import Period
from app.domain.statement import build_statement


def test_prior_styles_are_copied_across_workbooks_safely(tmp_path):
    source = tmp_path / 'prior.xlsx'
    prior = Workbook()
    prior.active.title = '특관자'
    for i in range(1, 50):
        prior.active.cell(i, 1, '전기 참고').font = Font(size=i + 8)
    prior.active['D119'] = '#REF!'
    prior.save(source)
    result = AggregateResult()
    result.get('A').fund_collection = 50
    output, _ = build_statement(result, Period(2026, 3), [], template_path=source)
    path = tmp_path / 'output.xlsx'
    output.save(path)
    loaded = load_workbook(path)
    assert loaded['전기_참고']['A51'].font.sz == 57
    assert loaded['전기_참고']['D121'].data_type != 'e'
