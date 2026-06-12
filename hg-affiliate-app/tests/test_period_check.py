"""전기 기준 가드(period_check) 검증.

"전기" = 직전 결산년도(전년도 4분기). 당기가 2026 1~4분기면 전기는 모두 2025 4분기.
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


def test_correct_prior_year_end_no_warning():
    # 직전 결산(전년도 4분기/연간)을 전기로 쓴 정상 케이스 → 경고 없음
    assert _n("2025년 결산 잔액명세서", "2026년 1분기말 잔액명세서") == 0
    assert _n("2025 4분기 잔액명세서", "2026 2분기 잔액명세서") == 0
    assert _n("2025 연간결산 잔액명세서", "2026 3분기 잔액명세서") == 0
    assert _n("2025 4분기 연간 잔액명세서", "2026 4분기 잔액명세서") == 0


def test_user_definition_all_quarters_same_prev():
    # 핵심 정의: 당기 2026 1~4분기 모두 전기=2025 4분기 → 어느 분기든 경고 없음
    prev = "2025년 4분기 연간 결산 잔액명세서"
    for q in (1, 2, 3, 4):
        assert _n(prev, f"2026년 {q}분기말 잔액명세서") == 0, f"{q}분기에서 오탐"


def test_wrong_previous_quarter_used_warns():
    # 2분기 보고에 전분기(2026 1분기)를 전기로 올림 → (A)분기 + (B)동일연도 = 2건
    assert _n("2026년 1분기말 잔액명세서", "2026년 2분기말 잔액명세서") == 2
    # 1Q/3Q 표기도 동일하게 잡힌다
    assert _n("2026 1Q 잔액명세서", "2026 3Q 잔액명세서") == 2


def test_wrong_same_year_year_end_only_year_warning():
    # 전기를 같은 해 연말(4분기)로 잘못 지정(연도만 동일) → (A)는 면제, (B)만 1건
    n = _n("2026 4분기 연간 잔액명세서", "2026 1분기 잔액명세서")
    assert n == 1


def test_prior_year_file_with_stray_print_date_no_false_positive():
    # 직전년도(2025) 결산본에 출력일(2026)이 섞여도, 2025가 남아 오탐하지 않는다
    n = _n("2025년 결산 잔액명세서 출력일 2026-01-10", "2026년 1분기말 잔액명세서")
    assert n == 0


# ── 선택 당기 년도 연동 (selected_year) ─────────────────────────────
def _ny(prev: str, cur: str, year: int) -> int:
    return len(evaluate_prior_period(prev, cur, selected_year=year))


def test_selected_year_correct_prev_no_warning():
    # 당기 2026 선택 → 기대 전기 = 2025 4Q. "25.4Q" 전기면 무경고
    assert _ny("2025년 4분기 연간 결산 잔액명세서", "무관", 2026) == 0


def test_selected_year_wrong_prev_year_warns():
    # 당기 2026인데 전기가 2024 결산 → 기대 2025 불일치 경고 1건
    assert _ny("2024년 4분기 연간 결산 잔액명세서", "무관", 2026) == 1


def test_selected_year_prev_same_year_quarter_warns_twice():
    # 당기 2026인데 전기 슬롯에 2026 1분기를 올림 → 분기(A) + 연도(Y) 2건
    assert _ny("2026년 1분기말 잔액명세서", "무관", 2026) == 2


def test_selected_year_takes_precedence_over_current_text_heuristic():
    # selected_year 주어지면 current_text 연도 비교(휴리스틱)에 의존하지 않는다
    # 전기=2025결산, current_text에 연도 없음 → 그래도 무경고
    assert _ny("2025년 결산", "분기말 잔액명세서", 2026) == 0


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
