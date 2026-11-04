import pandas as pd
import pytest
from app.domain.aggregate import AggregateResult, aggregate_ledger
from app.domain.errors import ErrorLog
from app.domain.period_extract import Period, extract_current_period
from app.domain.statement_detail import build_statement_detail, detail_group_key, detail_balance_formula


def calculate(rows):
    ledger = pd.DataFrame(rows, columns=['계정명','적요','차변','대변','날짜'])
    ledger['거래처명'] = '특관자A'
    ledger = extract_current_period(ledger, Period(2026,3)).df
    result = AggregateResult()
    aggregate_ledger(ledger,None,None,{'특관자A':'특관자A'},{'특관자A'},result,ErrorLog())
    detail = build_statement_detail(ledger,{'특관자A':'특관자A'},{'특관자A'},Period(2026,3))
    return result.get('특관자A'), detail, ledger


def test_negative_shared_fee_moves_to_positive_credit_and_income_separate_from_cost():
    agg, detail, ledger = calculate([
        ['지급수수료','일반 비용',80,0,'2026-01-03'],
        ['지급수수료','공동경비 정산',-100,0,'2026-03-01'],
        ['지급수수료','9월 공동경비',-20,0,'2026-09-30'],
        ['지급수수료','공동경비',-999,0,'2026-10-01'],
    ])
    assert agg.sales_other == 120
    assert agg.purchase_other == 80
    expense = [r for r in detail if r.bucket=='기타비용']
    income = [r for r in detail if r.bucket=='기타수익']
    assert len(expense)==1 and expense[0].balance==80
    assert [(r.debit,r.credit,r.balance) for r in income]==[(0,100,None),(0,20,120)]
    assert detail_group_key(expense[0]) != detail_group_key(income[0])
    assert detail_balance_formula(income[-1],17,18)=='=SUM(O17:O18)-SUM(N17:N18)'
    assert ledger['차변'].tolist()==[80,-100,-20]  # 원본 프레임은 수정하지 않는다.


@pytest.mark.parametrize('account,description,debit,credit,cost,income',[
    ('지급수수료','공동경비',30,0,30,0),
    ('지급수수료','공동경비',0,10,-10,0),
    ('지급수수료','일반 환급',-30,0,-30,0),
    ('보험료','공동경비',-30,0,-30,0),
    ('지급수수료','공동 구매 수수료',-30,0,-30,0),
    ('보험료','공동 경비',-30,0,-30,0),
    ('지급수수료','그룹 공동 경비',30,0,30,0),
    ('지급수수료','공동경비',-30,7,0,37),
])
def test_only_specified_account_description_and_negative_debit_qualify(account,description,debit,credit,cost,income):
    agg,detail,_=calculate([[account,description,debit,credit,'2026-09-01']])
    assert agg.purchase_other==cost
    assert agg.sales_other==income
    if income:
        assert (detail[0].debit,detail[0].credit)==(0,income)
    else:
        assert (detail[0].debit,detail[0].credit)==(debit,credit)


@pytest.mark.parametrize('description', [
    '공동경비',
    '그룹공동경비',
    '공동 경비',
    '공동경비(25년 기준 50% 적용)',
    '그룹 공동\t경비\n(25년 기준)',
    '공동\u00a0경비',
])
def test_statement_392_and_391_agree_for_mixed_fee_account(tmp_path, description):
    from openpyxl import Workbook
    from app.domain.statement import build_statement
    agg, detail, _ = calculate([
        ['지급수수료','일반 비용',80,0,'2026-01-03'],
        ['지급수수료',description,-120,0,'2026-09-01'],
    ])
    source=tmp_path/'prior.xlsx'
    prior=Workbook(); prior.active.title='특관자'; prior.save(source)
    result=AggregateResult(by_canonical={'특관자A':agg})
    wb,_=build_statement(result,Period(2026,3),detail,template_path=source)
    assert wb['39.2']['C4'].value==120  # 매출등 기타
    assert wb['39.2']['E4'].value==80   # 매입등 기타
    assert wb['39.1']['P3'].value=='=SUM(N3:N3)-SUM(O3:O3)'
    assert wb['39.1']['P4'].value=='=SUM(O4:O4)-SUM(N4:N4)'
    assert wb['39.1']['N4'].value==0
    assert wb['39.1']['O4'].value==120
    assert wb['39.1']['F4'].value=='기타수익'
    assert wb['39.1']['J4'].value==description
    assert wb['39.1']['B2'].value == '매입매출'
    assert wb['39.1'].auto_filter.ref == 'B2:P4'
    for name in ['39.1', '39.2']:
        assert wb[name].freeze_panes is None
