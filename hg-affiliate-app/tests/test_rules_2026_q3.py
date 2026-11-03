import pandas as pd
import pytest

from app.domain.aggregate import AggregateResult, aggregate_balance, aggregate_ledger
from app.domain.errors import ErrorLog
from app.domain.statement_detail import _classify


def aggregate(rows, opening=None, closing=None):
    ledger = pd.DataFrame(rows, columns=['계정명', '차변', '대변'])
    ledger['거래처명'] = 'A'
    result = AggregateResult()
    errors = ErrorLog()
    aggregate_balance(closing, opening, {'A': 'A'}, {'A'}, result, errors)
    aggregate_ledger(ledger, opening, closing, {'A': 'A'}, {'A'}, result, errors)
    return result.get('A')


def test_new_purchase_and_income_rules():
    a = aggregate([('상품', 100, 0), ('건설중인자산_무형', 200, 0),
                   ('통신비', 30, 0), ('수입임대료', 0, 40), ('잡이익', 0, 50)])
    assert a.purchase == 300
    assert a.purchase_other == 30
    assert a.sales_other == 90
    assert _classify('건설중인자산_무형', 200, 0, [])[4] == '매입'


def test_allowance_rollforward_and_expense_are_independent():
    opening = pd.DataFrame([['A', '대손충당금(단기대여금)', 100]],
                           columns=['거래처', '계정과목', '금액'])
    closing = pd.DataFrame([['A', '대손충당금(단기대여금)', 999]],
                           columns=['거래처', '계정과목', '금액'])
    a = aggregate([('대손충당금(단기대여금)', 20, 5),
                   ('기타의대손상각비', 0, 8)], opening, closing)
    assert a.allowance_end == 85
    assert a.allowance_expense == -8


def test_funding_reports_both_sides_without_netting():
    a = aggregate([('단기대여금', 100, 20), ('장기대여금', 50, 30),
                   ('단기차입금', 40, 200), ('FVPL측정금융자산', 300, 60)])
    assert a.fund_lending == 150
    assert a.fund_collection == 50
    assert a.fund_borrowing == 200
    assert a.fund_repayment == 40
    assert a.fund_investment == 300
    assert a.fund_disposal == 60


def test_new_balance_buckets_do_not_absorb_allowance_or_duplicates():
    closing = pd.DataFrame([['A', '임차보증금(유동)', 10], ['A', '미수금(기타)', 20],
                            ['A', '미지급비용', 30], ['A', '단기차입금', 40],
                            ['A', 'FVPL금융자산', 50], ['A', '대손충당금(미수금)', 5]],
                           columns=['거래처', '계정과목', '금액'])
    a = aggregate([], closing=closing)
    assert a.other_ar == 30
    assert a.other_ap == 30
    assert a.borrowing == 40
    assert a.cb_invest == 50


def test_revised_report_includes_new_parties_and_funding(tmp_path):
    from openpyxl import Workbook
    from app.domain.statement import build_statement
    from app.domain.period_extract import Period
    source = tmp_path / 'prior.xlsx'
    wb = Workbook()
    wb.active.title = '특관자'
    wb.active['B9'] = '종속기업'
    wb.active['C9'] = 'A'
    wb.save(source)
    result = AggregateResult()
    result.get('A').allowance_end = 85
    result.get('A').allowance_expense = -8
    result.get('새로운법인').fund_borrowing = 200
    wb, unmatched = build_statement(result, Period(2026, 3), [], template_path=source)
    assert not unmatched
    assert '전기_참고' in wb.sheetnames
    main = wb['특관자']
    strings = [c.value for row in main for c in row if isinstance(c.value, str)]
    assert any('38.3' in x for x in strings)
    assert any('38.4' in x for x in strings)
    assert '새로운법인' in strings
    assert any('SUMIF' in x for x in strings)


def test_pivot_recovery_source_cannot_overwrite_aliases(tmp_path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from app.normalize import seed_from_name_pivot
    wb = Workbook()
    wb.active.append(['대표법인'])
    wb.active.append(['별칭'])
    wb.active['A2'].alignment = Alignment(indent=1)
    raw = wb.create_sheet('_복구_피벗원본')
    raw.append(['거래처명'])
    raw.append(['별칭'])
    path = tmp_path / 'pivot.xlsx'
    wb.save(path)
    seed = seed_from_name_pivot(path)
    assert seed['별칭'] == '대표법인'
    assert '거래처명' not in seed


def test_name_table_reads_explicit_alias_column_without_importing_titles(tmp_path):
    from openpyxl import Workbook
    from app.normalize import seed_from_name_pivot
    wb = Workbook()
    ws = wb.active
    ws.append(['HLB글로벌(주) 특수관계자 지배구조'])
    ws.append([])
    ws.append(['정규(대표) 법인명', '구분', 'HLB글로벌과의 관계', '분개장 이표기(별칭)', '비고'])
    ws.append(['대표법인', '기타특수관계자', '특수관계자', '옛이름, 지점명', '검토'])
    ws.append(['에이치엘비오션테크', '', '', '바다중공업,에이치엘비오션테크'])
    path = tmp_path / 'names.xlsx'
    wb.save(path)
    seed = seed_from_name_pivot(path)
    assert seed == {'대표법인': '대표법인', '옛이름': '대표법인', '지점명': '대표법인',
                    '에이치엘비오션테크': '에이치엘비오션테크', '바다중공업': '에이치엘비오션테크'}
