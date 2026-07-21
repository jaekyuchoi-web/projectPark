"""당기 기간 정의 및 당기 데이터 추출.

당기 기간(누적): 해당년도 1월 ~ 분기말월.
  1Q=1~3월, 2Q=1~6월, 3Q=1~9월, 4Q=1~12월.

날짜 추출은 거래 발생일이 행마다 있는 '당기 원장'에만 적용한다(잔액명세서는
시점 스냅샷이라 제외). 결정론 파서로 먼저 처리하고, 비어있지 않은데 파싱 실패한
셀은 AI(ChatGPT) 보조 파서로 재시도한다. 그래도 남으면 엄격 차단한다.

값/수치(행 내용)는 로그·화면으로 내보내지 않는다.
"""

from __future__ import annotations

import datetime as dt
import json
import re
import time
from collections.abc import Callable
from dataclasses import dataclass

import pandas as pd

from ..config import Settings
from . import columns as C

# 엑셀 1900 날짜체계 기준일(윤년 버그 호환 위해 1899-12-30)
_EXCEL_EPOCH = dt.datetime(1899, 12, 30)
PERIOD_YEARMONTH_ATTR = "period_year_months"

# "2026-03-15" / "2026.3.15" / "2026/3" / "2026년 3월" 등에서 년·월 추출
_YM_SEP_RE = re.compile(r"(20\d{2})\s*[-./년]\s*(\d{1,2})")
# "20260315" / "202603" 같은 붙임 표기
_YM_COMPACT_RE = re.compile(r"^(20\d{2})(\d{2})(\d{2})?$")


def parse_year_month(value) -> tuple[int, int] | None:
    """값에서 (year, month) 를 결정론적으로 추출한다. 실패 시 None."""
    if value is None:
        return None
    # datetime / date 객체
    if isinstance(value, (dt.datetime, dt.date)):
        return value.year, value.month
    # 엑셀 serial (숫자)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            d = _EXCEL_EPOCH + dt.timedelta(days=float(value))
        except (OverflowError, ValueError):
            return None
        if 1900 <= d.year <= 2099:
            return d.year, d.month
        return None

    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None

    m = _YM_COMPACT_RE.match(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        return (y, mo) if 1 <= mo <= 12 else None

    m = _YM_SEP_RE.search(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        return (y, mo) if 1 <= mo <= 12 else None

    return None


@dataclass(frozen=True)
class Period:
    """Calendar-year-to-date reporting period.

    NON-NEGOTIABLE: quarters start in January. 2Q is January through June,
    never April through June only.
    """

    year: int
    quarter: int

    @property
    def start_month(self) -> int:
        return 1

    @property
    def end_month(self) -> int:
        return self.quarter * 3

    @property
    def label(self) -> str:
        return (
            f"{self.year}년 {self.quarter}분기 누적 "
            f"({self.start_month}~{self.end_month}월)"
        )

    def contains(self, year: int, month: int) -> bool:
        return year == self.year and self.start_month <= month <= self.end_month


def parse_period(year, quarter) -> Period:
    """요청값(년도/분기)을 검증해 Period 로 만든다. 잘못되면 ValueError."""
    try:
        y, q = int(year), int(quarter)
    except (TypeError, ValueError):
        raise ValueError("당기 년도/분기를 숫자로 입력하세요.")
    if not (2000 <= y <= 2099):
        raise ValueError("당기 년도는 2000~2099 범위여야 합니다.")
    if q not in (1, 2, 3, 4):
        raise ValueError("당기 분기는 1~4 중 하나여야 합니다.")
    return Period(y, q)


@dataclass
class ExtractResult:
    """당기 추출 결과. ok=False 이면 다운로드 차단."""

    df: pd.DataFrame
    ok: bool
    reason: str = ""
    kept: int = 0
    dropped_out_of_period: int = 0
    dropped_blank: int = 0
    parsed_ai: int = 0  # AI 보조 파서가 구제한 행 수


# AI 보조 파서: 미파싱 문자열 목록 → {문자열: (year, month) | None}
AiDateParser = Callable[[list[str]], dict[str, "tuple[int, int] | None"]]


def extract_current_period(
    df: pd.DataFrame,
    period: Period,
    *,
    ai_parser: AiDateParser | None = None,
) -> ExtractResult:
    """원장 프레임에서 당기(period)에 해당하는 행만 추출한다.

    - 날짜 컬럼 미식별 → 차단.
    - 빈/합계행(날짜 빈값) → 비거래로 제외(실패 아님).
    - 값이 있는데 결정론 파싱 실패한 셀 → AI 보조 파서로 재시도.
    - 그래도 미파싱 셀이 남으면 → 엄격 차단.
    """
    empty = df.iloc[0:0]
    date_col = C.resolve_column(df, "date")
    if date_col is None:
        return ExtractResult(
            df=empty, ok=False,
            reason="당기 원장에서 날짜(전표일자) 컬럼을 식별하지 못해 기간 추출이 불가합니다.",
        )

    parsed: dict[int, tuple[int, int]] = {}
    failures: dict[int, str] = {}
    blank = 0

    for i, v in enumerate(df[date_col].tolist()):
        s = "" if v is None else str(v).strip()
        if not s or s.lower() == "nan":
            blank += 1
            continue
        ym = parse_year_month(v)
        if ym is None:
            failures[i] = s
        else:
            parsed[i] = ym

    parsed_ai = 0
    if failures and ai_parser is not None:
        ai_map = ai_parser(sorted(set(failures.values()))) or {}
        for i, s in list(failures.items()):
            ym = ai_map.get(s)
            if ym is not None:
                parsed[i] = ym
                del failures[i]
                parsed_ai += 1

    if failures:
        return ExtractResult(
            df=empty, ok=False,
            reason=(
                f"당기 원장에서 날짜를 해석하지 못한 셀이 {len(failures)}건 남아 "
                "처리를 중단했습니다(엄격 차단)."
            ),
            dropped_blank=blank, parsed_ai=parsed_ai,
        )

    keep = sorted(i for i, ym in parsed.items() if period.contains(*ym))
    filtered = df.iloc[keep].reset_index(drop=True)
    filtered.attrs[PERIOD_YEARMONTH_ATTR] = [parsed[index] for index in keep]
    return ExtractResult(
        df=filtered, ok=True,
        kept=len(keep),
        dropped_out_of_period=len(parsed) - len(keep),
        dropped_blank=blank,
        parsed_ai=parsed_ai,
    )


# ── AI 보조 날짜 파서 (normalize.py 패턴) ────────────────────────────
def _parse_date_json(text: str, batch: list[str]) -> dict[str, tuple[int, int] | None]:
    """모델 JSON('YYYY-MM' 또는 null)을 {표기: (year, month) | None} 로 변환."""
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if not m:
            return {}
        try:
            data = json.loads(m.group(0))
        except json.JSONDecodeError:
            return {}
    if not isinstance(data, dict):
        return {}
    out: dict[str, tuple[int, int] | None] = {}
    for k, v in data.items():
        out[str(k)] = parse_year_month(v) if isinstance(v, str) else None
    return out


def _build_date_prompt(batch: list[str]) -> str:
    return (
        "다음은 회계 원장의 '전표일자' 표기 목록이다. 약식·한글·구분자 혼용 등으로 "
        "다양하게 적혀 있을 수 있다. 각 표기에서 거래가 발생한 연·월을 추출하라. "
        "날짜로 해석 가능하면 'YYYY-MM' 문자열로, 도저히 날짜로 볼 수 없으면 null 로 매핑하라.\n\n"
        f"표기 목록: {json.dumps(batch, ensure_ascii=False)}\n\n"
        '반드시 JSON 객체만 출력: {"표기": "YYYY-MM" 또는 null, ...}'
    )


def _date_chat_with_retry(client, model: str, prompt: str, retries: int = 3) -> str | None:
    for attempt in range(retries):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": "너는 한국 회계 날짜 파서다. JSON만 출력한다."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0,
                response_format={"type": "json_object"},
            )
            return resp.choices[0].message.content
        except Exception:  # noqa: BLE001 - 레이트리밋/일시오류 재시도
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    return None


def _openai_parse_dates(
    values: list[str],
    settings: Settings,
    client=None,
) -> dict[str, tuple[int, int] | None]:
    """ChatGPT 로 미파싱 날짜 표기를 (year, month) 로 해석. 배치 + 재시도."""
    if client is None:
        from openai import OpenAI  # 지연 임포트
        client = OpenAI(api_key=settings.openai_api_key)
    result: dict[str, tuple[int, int] | None] = {}
    batch_size = 60
    for i in range(0, len(values), batch_size):
        batch = values[i : i + batch_size]
        text = _date_chat_with_retry(client, settings.openai_model, _build_date_prompt(batch))
        if text:
            result.update(_parse_date_json(text, batch))
    return result


def build_ai_date_parser(settings: Settings) -> AiDateParser | None:
    """OPENAI 키가 있으면 AI 보조 파서를, 없으면 None 을 반환한다."""
    if not settings.has_openai_key:
        return None

    def parser(values: list[str]) -> dict[str, tuple[int, int] | None]:
        return _openai_parse_dates(values, settings)

    return parser
