"""전기 기준 점검 (정의 가드).

최종 도메인 정의상 "전기"는 전년도 온기(연간 결산 = 4분기)다.
같은 해 직전분기 자료는 참고할 수 있으나 전기가 아니다.

업로드된 '전기 이월 소스'가 이 기준과 다르게 들어오면 대손상각비·미수수익
역산·자금대여가 잘못 계산될 수 있다. 이를 막기 위해, 파일명/시트명에서
연도·분기 표식을 읽어 '전기 소스가 최종규칙의 기대 기간이 아닐 가능성'을
경고로 띄운다(차단 아님).

이 모듈은 순수 로직만 담는다(파일 IO 없음) — 텍스트(blob)는 호출측에서 구성한다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# "1분기" / "1 분기" / "1Q" / "Q1" (대소문자 무시). 분기 숫자만 추출.
_QUARTER_RE = re.compile(r"([1-4])\s*분기|([1-4])\s*[Qq](?![A-Za-z0-9])|[Qq]\s*([1-4])")
# 연도(2000~2099). 표제/파일명 영역에서만 수집해 오탐을 줄인다.
# 한글이 붙어도("2026년") 잡히도록 \b 대신 숫자 비인접 룩어라운드를 쓴다.
_YEAR_RE = re.compile(r"(?<!\d)(20\d{2})(?!\d)")
# 파일명 표식의 2자리 연도 + 분기(예: 25.3Q, 24_4Q, 26년1분기).
_SHORT_YEAR_QUARTER_RE = re.compile(
    r"(?<!\d)(\d{2})\s*(?:[._\-/]|년\s*)\s*([1-4])\s*(?:[Qq]|분기)"
)
# 연말(직전 결산) 신호 — 이 표식이 있으면 '분기처럼 보여도' 결산본으로 본다.
_YEAR_END_TOKENS = ("연간", "결산", "온기", "12월", "12 월", "기말결산")


@dataclass
class PeriodInfo:
    quarters: set[int] = field(default_factory=set)
    years: list[int] = field(default_factory=list)
    has_year_end: bool = False


def analyze_period(text: str) -> PeriodInfo:
    """표제/파일명 텍스트에서 분기·연도·연말표식을 추출한다."""
    if not text:
        return PeriodInfo()
    quarters: set[int] = set()
    for m in _QUARTER_RE.finditer(text):
        for g in m.groups():
            if g:
                quarters.add(int(g))
    years = {int(y) for y in _YEAR_RE.findall(text)}
    for yy, q in _SHORT_YEAR_QUARTER_RE.findall(text):
        years.add(2000 + int(yy))
        quarters.add(int(q))
    years = sorted(years)
    has_year_end = any(t in text for t in _YEAR_END_TOKENS)
    # 분기 표식이 4분기뿐이면 그것도 연말(직전 결산) 신호로 본다.
    if quarters == {4}:
        has_year_end = True
    return PeriodInfo(quarters=quarters, years=years, has_year_end=has_year_end)


@dataclass
class PriorPeriodWarning:
    content: str
    action: str


_ACTION = (
    "전기는 전년도 온기(연간 결산 = 4분기)여야 합니다. "
    "예) 당기 2026년 1~4분기는 모두 2025년 온기/4분기를 전기로 사용하세요. "
    "같은 해 직전분기 자료는 참고 자료일 수 있으나 전기 슬롯에는 넣지 마세요."
)


def _expected_prior(selected_year: int, selected_quarter: int) -> tuple[int, int, bool]:
    """(기대 연도, 기대 분기, 온기/연말 허용 여부)."""
    return selected_year - 1, 4, True


def _expected_label(year: int, quarter: int, allow_year_end: bool) -> str:
    label = f"{year}년 {quarter}분기"
    return f"{label}/온기" if allow_year_end else label


def _infer_single_quarter(text: str) -> int | None:
    quarters = analyze_period(text).quarters
    return next(iter(quarters)) if len(quarters) == 1 else None


def _selected_period_warning(
    prev: PeriodInfo,
    selected_year: int,
    selected_quarter: int,
) -> PriorPeriodWarning | None:
    expected_year, expected_quarter, allow_year_end = _expected_prior(selected_year, selected_quarter)
    expected = _expected_label(expected_year, expected_quarter, allow_year_end)
    problems: list[str] = []

    if prev.years and expected_year not in prev.years:
        problems.append(f"표식 연도: {prev.years}")

    if prev.quarters:
        if allow_year_end:
            if expected_quarter not in prev.quarters and not prev.has_year_end:
                problems.append(f"표식 분기: {sorted(prev.quarters)}")
        elif expected_quarter not in prev.quarters:
            problems.append(f"표식 분기: {sorted(prev.quarters)}")
    elif allow_year_end and not prev.has_year_end:
        problems.append("분기/온기 표식 없음")

    if not problems:
        return None
    return PriorPeriodWarning(
        content=f"전기 이월 소스에서 기대 전기({expected})를 확인하지 못했습니다({', '.join(problems)}).",
        action=_ACTION,
    )


def evaluate_prior_period(
    prev_text: str,
    current_text: str,
    selected_year: int | None = None,
    selected_quarter: int | None = None,
) -> list[PriorPeriodWarning]:
    """전기 소스가 최종규칙의 기대 기간이 아닐 정황을 경고 목록으로 반환한다.

    selected_year가 주어지면 최종규칙 기준으로 정밀 검증한다. selected_quarter는
    호환성을 위해 받지만, 전기는 모든 분기에서 전년도 온기/4분기로 동일하다.
    경고는 정보성(검토 권고)일 뿐 처리를 차단하지 않는다.
    """
    warnings: list[PriorPeriodWarning] = []
    prev = analyze_period(prev_text)

    if selected_year is not None:
        warning = _selected_period_warning(prev, selected_year, selected_quarter or 1)
        return [warning] if warning is not None else []

    # 선택 기간이 없으면 current_text의 단일 연도·분기 표식으로 휴리스틱 판정한다.
    cur = analyze_period(current_text)
    if len(cur.years) == 1 and len(cur.quarters) == 1:
        warning = _selected_period_warning(prev, cur.years[0], next(iter(cur.quarters)))
        return [warning] if warning is not None else []

    return warnings
