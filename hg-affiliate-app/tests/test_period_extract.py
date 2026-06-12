"""period_extract 단위 테스트 (TDD).

당기 기간 정의: 해당년도 1월 ~ 분기말월 (1Q→3, 2Q→6, 3Q→9, 4Q→12).
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from app.domain.period_extract import (
    Period,
    extract_current_period,
    parse_period,
    parse_year_month,
)


# ── parse_period: 요청값 검증 ────────────────────────────────────────
def test_parse_period_valid():
    assert parse_period(2026, 3) == Period(2026, 3)
    assert parse_period("2026", "1") == Period(2026, 1)


@pytest.mark.parametrize("y,q", [
    (1999, 1), (2100, 1), (2026, 0), (2026, 5),
    ("x", 1), (2026, "q"), (None, 1), (2026, None),
])
def test_parse_period_invalid(y, q):
    with pytest.raises(ValueError):
        parse_period(y, q)


def _ledger(dates: list) -> pd.DataFrame:
    """'일자' 컬럼(컬럼 별칭에 이미 존재)을 가진 원장형 프레임."""
    return pd.DataFrame({
        "계정명": ["현금"] * len(dates),
        "일자": dates,
        "거래처명": ["거래처"] * len(dates),
        "차변": ["1000"] * len(dates),
        "대변": ["0"] * len(dates),
    })


# ── parse_year_month: 결정론 파서 ────────────────────────────────────
@pytest.mark.parametrize("value,expected", [
    ("2026-03-15", (2026, 3)),
    ("2026.3.15", (2026, 3)),
    ("2026/03/15", (2026, 3)),
    ("20260315", (2026, 3)),
    ("2026년 3월 15일", (2026, 3)),
    ("2026-03", (2026, 3)),
    ("2026-12-31", (2026, 12)),
    (dt.datetime(2026, 6, 30), (2026, 6)),
    (dt.date(2026, 9, 1), (2026, 9)),
])
def test_parse_year_month_ok(value, expected):
    assert parse_year_month(value) == expected


@pytest.mark.parametrize("value", ["", "   ", "nan", None, "합계", "거래처명", "abc"])
def test_parse_year_month_unparseable(value):
    assert parse_year_month(value) is None


def test_parse_year_month_excel_serial():
    # 엑셀 serial 46112 = 2026-03-15 (1900 날짜체계, 1899-12-30 기준)
    assert parse_year_month(46112) == (2026, 3)


# ── Period 모델 ─────────────────────────────────────────────────────
@pytest.mark.parametrize("quarter,end_month", [(1, 3), (2, 6), (3, 9), (4, 12)])
def test_period_end_month(quarter, end_month):
    assert Period(2026, quarter).end_month == end_month


@pytest.mark.parametrize("quarter,month,inside", [
    (1, 3, True), (1, 4, False),       # 1Q = 1~3월
    (2, 6, True), (2, 7, False),       # 2Q = 1~6월
    (3, 9, True), (3, 10, False),      # 3Q = 1~9월
    (4, 12, True),                     # 4Q = 1~12월
    (2, 1, True),                      # 누적: 1월 포함
])
def test_period_contains_same_year(quarter, month, inside):
    assert Period(2026, quarter).contains(2026, month) is inside


def test_period_contains_other_year_excluded():
    # 선택년도와 다른 연도는 제외 (전년도 12월, 익년 1월 모두 당기 아님)
    assert Period(2026, 1).contains(2025, 12) is False
    assert Period(2026, 4).contains(2027, 1) is False


# ── extract_current_period: 추출 + 엄격 차단 ─────────────────────────
def test_extract_keeps_all_in_period():
    df = _ledger(["2026-01-12", "2026-02-26", "2026-03-31"])
    r = extract_current_period(df, Period(2026, 1))
    assert r.ok is True
    assert r.kept == 3
    assert len(r.df) == 3
    assert r.dropped_out_of_period == 0


def test_extract_filters_out_of_period_rows():
    # 2Q(1~6월) 선택: 1·5월 채택, 7월·2025-12·2027 제외
    df = _ledger(["2026-01-10", "2026-05-20", "2026-07-01", "2025-12-31", "2027-01-01"])
    r = extract_current_period(df, Period(2026, 2))
    assert r.ok is True
    assert r.kept == 2
    assert r.dropped_out_of_period == 3
    assert set(r.df["일자"]) == {"2026-01-10", "2026-05-20"}


def test_extract_blank_date_rows_dropped_not_failure():
    df = _ledger(["2026-02-01", "", None, "2026-03-01"])
    r = extract_current_period(df, Period(2026, 1))
    assert r.ok is True
    assert r.kept == 2
    assert r.dropped_blank == 2


def test_extract_no_date_column_blocks():
    df = pd.DataFrame({"계정명": ["현금"], "금액": ["1000"]})
    r = extract_current_period(df, Period(2026, 1))
    assert r.ok is False
    assert "날짜" in r.reason


def test_extract_unparseable_without_ai_blocks_strict():
    df = _ledger(["2026-02-01", "날짜미상", "2026-03-01"])
    r = extract_current_period(df, Period(2026, 1))  # ai_parser 없음
    assert r.ok is False
    assert r.reason  # 사유 존재


def test_extract_unparseable_rescued_by_ai():
    df = _ledger(["2026-02-01", "이천이십육년 삼월", "2026-03-01"])
    ai = lambda vals: {v: (2026, 3) for v in vals}  # AI가 구제
    r = extract_current_period(df, Period(2026, 1), ai_parser=ai)
    assert r.ok is True
    assert r.kept == 3
    assert r.parsed_ai == 1


def test_extract_ai_still_fails_blocks():
    df = _ledger(["2026-02-01", "도저히모름", "2026-03-01"])
    ai = lambda vals: {v: None for v in vals}  # AI도 실패
    r = extract_current_period(df, Period(2026, 1), ai_parser=ai)
    assert r.ok is False
    assert r.reason
