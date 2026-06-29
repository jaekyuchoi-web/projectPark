"""전기 기준 가드(period_check) 검증.

최종 도메인 정의:
- 전기 = 전년도 온기/4분기
- 같은 해 직전분기 자료는 참고할 수 있으나 전기가 아니다
pytest 가 없을 수 있어 plain assert + __main__ 러너로 동작한다.

실행: .venv/bin/python tests/test_period_check.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.domain.period_check import analyze_period, evaluate_prior_period  # noqa: E402


def _n(prev: str, cur: str) -> int:
    return len(evaluate_prior_period(prev, cur))


def test_analyze_basic():
    info = analyze_period("2026년 1분기말 채권채무 잔액명세서")
    assert info.quarters == {1}
    assert info.years == [2026]
    assert info.has_year_end is False

    ye = analyze_period("2025년 연간 결산 잔액명세서")
    assert ye.has_year_end is True

    q4 = analyze_period("2025 4분기 잔액명세서")
    assert q4.quarters == {4}
    assert q4.has_year_end is True  # 4분기뿐이면 연말 신호


def test_prior_year_end_no_warning_for_all_quarters():
    # 모든 당기 분기에서 전기는 전년도 온기/4분기다.
    assert _n("2025년 결산 잔액명세서", "2026년 1분기말 잔액명세서") == 0
    assert _n("2025 4분기 잔액명세서", "2026 1분기 잔액명세서") == 0
    assert _n("2025 4분기 잔액명세서", "2026 2분기 잔액명세서") == 0
    assert _n("2025 연간결산 잔액명세서", "2026 3분기 잔액명세서") == 0
    assert _n("2025 4분기 연간 잔액명세서", "2026 4분기 잔액명세서") == 0


def test_same_year_previous_quarter_warns_after_q1():
    # 2026년 2Q 주석 작성 시 2026년 1Q는 참고 자료일 뿐 전기가 아니다.
    assert _n("2026년 1분기말 잔액명세서", "2026년 2분기말 잔액명세서") >= 1
    assert _n("2026 2Q 잔액명세서", "2026 3Q 잔액명세서") >= 1


def test_same_year_q4_wrong_for_q1():
    # 1분기는 같은 해 4분기가 아니라 직전년도 온기/4분기를 전기로 쓴다.
    n = _n("2026 4분기 연간 잔액명세서", "2026 1분기 잔액명세서")
    assert n >= 1


def test_prior_year_file_with_stray_print_date_no_false_positive():
    # 직전년도(2025) 결산본에 출력일(2026)이 섞여도, 2025가 남아 오탐하지 않는다
    n = _n("2025년 결산 잔액명세서 출력일 2026-01-10", "2026년 1분기말 잔액명세서")
    assert n == 0


# ── 선택 당기 기간 연동 (selected_year/selected_quarter) ───────────
def _ny(prev: str, cur: str, year: int, quarter: int) -> int:
    return len(evaluate_prior_period(prev, cur, selected_year=year, selected_quarter=quarter))


def test_selected_year_correct_prev_no_warning():
    # 당기 2026 1~4Q 선택 → 기대 전기 = 2025 온기/4Q.
    for quarter in (1, 2, 3, 4):
        assert _ny("2025년 4분기 연간 결산 잔액명세서", "무관", 2026, quarter) == 0


def test_selected_period_after_q1_rejects_same_year_previous_quarter():
    # 당기 2025 3Q 선택 → 전기는 2024 온기/4Q, 2025 2Q가 아니다.
    warnings = evaluate_prior_period(
        "2025년 2분기 잔액명세서",
        "무관",
        selected_year=2025,
        selected_quarter=3,
    )
    assert warnings
    assert "2024년 4분기/온기" in warnings[0].content


def test_selected_year_wrong_prev_year_warns():
    # 당기 2026 1Q인데 전기가 2024 결산 → 기대 2025 불일치 경고
    assert _ny("2024년 4분기 연간 결산 잔액명세서", "무관", 2026, 1) >= 1


def test_selected_year_prev_same_year_quarter_warns_twice():
    # 당기 2026 1Q인데 전기 슬롯에 2026 1Q를 올림 → 기대 전기 불일치
    assert _ny("2026년 1분기말 잔액명세서", "무관", 2026, 1) >= 1


def test_selected_period_prior_year_end_warns_after_q1():
    # 당기 2025 3Q라도 전기 슬롯의 2024 4Q 온기는 정상이다.
    assert _ny("2024년 4분기 연간 결산 잔액명세서", "무관", 2025, 3) == 0


def test_selected_year_takes_precedence_over_current_text_heuristic():
    # selected period가 주어지면 current_text 연도 비교(휴리스틱)에 의존하지 않는다
    # 전기=2025결산, current_text에 연도 없음 → 그래도 무경고
    assert _ny("2025년 결산", "분기말 잔액명세서", 2026, 1) == 0


def _run():
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    failed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except AssertionError as e:  # noqa: PERF203
            failed += 1
            print(f"FAIL {fn.__name__}: {e}")
    print(f"\n{len(fns) - failed}/{len(fns)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(_run())
