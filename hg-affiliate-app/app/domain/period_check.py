"""전기 기준 점검 (정의 가드).

"전기" = 직전 결산년도(전년도 연간 결산 = 전년도 4분기). 결산은 연 1회뿐이므로
당기가 어느 분기든 전기는 항상 직전년도 4분기로 동일하다(분기별로 회전하지 않음).

업로드된 '전기 이월 소스'가 실수로 직전년도 결산이 아니라 동일/직전 '분기'
(예: 2026 1분기)로 들어오면 대손상각비·미수수익 역산·자금대여가 분기 증분으로
과소계상된다. 이를 막기 위해, 파일명/시트명/표제행 텍스트에서 연도·분기 표식을
읽어 '전기 소스가 직전 결산년도가 아닐 가능성'을 경고로 띄운다(차단 아님).

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
    years = sorted({int(y) for y in _YEAR_RE.findall(text)})
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
    "전기는 '직전 결산년도'(전년도 연간 결산 = 전년도 4분기)여야 합니다. "
    "예) 당기 2026년 1~4분기는 모두 2025년 4분기(연간 결산)를 전기로 사용하세요. "
    "전기 슬롯에 직전년도 결산(4분기) 파일이 맞는지 확인하세요."
)


def evaluate_prior_period(
    prev_text: str,
    current_text: str,
    selected_year: int | None = None,
) -> list[PriorPeriodWarning]:
    """전기 소스가 '직전 결산년도'가 아닐 정황을 경고 목록으로 반환한다.

    selected_year(당기 선택 년도)가 주어지면 기대 전기연도 = selected_year-1 로
    정밀 검증한다. 없으면 current_text 와의 연도 비교(휴리스틱)로 판정한다.
    경고는 정보성(검토 권고)일 뿐 처리를 차단하지 않는다.
    """
    warnings: list[PriorPeriodWarning] = []
    prev = analyze_period(prev_text)

    # (A) 전기 소스가 동일/직전 '분기'(1~3분기)로 보이고 연말(결산) 표식이 없음.
    mid_quarters = {q for q in prev.quarters if q in (1, 2, 3)}
    if mid_quarters and not prev.has_year_end:
        qtxt = "·".join(f"{q}분기" for q in sorted(mid_quarters))
        warnings.append(PriorPeriodWarning(
            content=(
                f"전기 이월 소스가 직전 결산(전년도 연간·4분기)이 아니라 '{qtxt}'(동일/직전 분기)로 "
                "보입니다. 전기는 분기와 무관하게 직전년도 4분기(연간 결산)로 동일해야 합니다."
            ),
            action=_ACTION,
        ))

    if selected_year is not None:
        # (Y-정밀) 선택 당기 년도 기준 기대 전기연도 = selected_year-1.
        expected = selected_year - 1
        if prev.years and expected not in prev.years:
            warnings.append(PriorPeriodWarning(
                content=(
                    f"전기 이월 소스에서 기대 전기연도({expected})를 확인하지 못했습니다"
                    f"(표식 연도: {prev.years}). 당기={selected_year}의 전기는 "
                    f"{expected}년 4분기(연간 결산)여야 합니다."
                ),
                action=_ACTION,
            ))
    else:
        # (B-휴리스틱) 전기 소스에 '당기보다 이른 연도'가 전혀 없음.
        #   직전년도 파일이면 그 연도(예: 2025)가 표제에 남아 min < 당기연도가 된다.
        cur = analyze_period(current_text)
        if prev.years and cur.years and min(prev.years) >= max(cur.years):
            warnings.append(PriorPeriodWarning(
                content=(
                    f"전기 이월 소스에서 당기({max(cur.years)})보다 이른 연도가 확인되지 않습니다"
                    f"(전기 표식 연도: {min(prev.years)}). 전기는 직전년도여야 합니다."
                ),
                action=_ACTION,
            ))

    return warnings
