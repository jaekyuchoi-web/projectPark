"""내용 기반 자동 분류 (절대 준수 e).

파일명이 제각각일 수 있으므로, 파일명 힌트 + 시트명 + 헤더 키워드를 종합해
5개 슬롯 중 어디에 해당하는지 점수화하여 추정한다.
사용자는 화면에서 슬롯 매핑을 수동 교정할 수 있다.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import SLOT_KEYS
import pandas as pd

from .excel_io import read_workbook, sheet_names

# 슬롯별 키워드 (파일명·시트명·헤더에 등장하면 가점)
SLOT_KEYWORDS: dict[str, list[str]] = {
    "prev_balance": ["전기", "이월", "전기말", "잔액명세", "기초", "전년"],
    "current_ledger": ["원장", "거래", "분개", "총계정", "당기거래", "ledger", "전표"],
    "current_balance": ["당기말", "검증", "잔액명세", "당기잔액", "기말", "채권채무"],
    "name_pivot": ["상호", "정리", "특관자상호", "정규", "명칭", "거래처정리"],
    "prev_statement": ["특관자", "명세서", "전기명세", "지배구조", "39.1", "39.2", "특수관계"],
}


@dataclass
class ClassifyResult:
    slot: str | None       # 추정 슬롯키 (없으면 None)
    scores: dict[str, int] # 슬롯별 점수 (디버그/표시용 아님, 내부)


def _text_blob(
    path: Path,
    original_name: str,
    sheets: dict[str, pd.DataFrame] | None = None,
) -> str:
    """파일명 + 시트명 + 상단 헤더 텍스트만 모은다(값 데이터는 포함하지 않음)."""
    parts: list[str] = [Path(original_name).stem]
    loaded = sheets
    if loaded is None:
        try:
            parts.extend(sheet_names(path))
        except Exception:  # noqa: BLE001
            pass
        try:
            loaded = read_workbook(path)
        except Exception:  # noqa: BLE001
            loaded = None
    else:
        parts.extend(list(loaded.keys()))

    if loaded:
        for df in loaded.values():
            if df is None or df.empty:
                continue
            head_rows = df.head(12)
            for _, row in head_rows.iterrows():
                for v in row.tolist():
                    s = str(v).strip()
                    if s and s != "nan" and not _looks_numeric(s):
                        parts.append(s)
    return " ".join(parts).lower()


def _looks_numeric(s: str) -> bool:
    t = s.replace(",", "").replace("-", "").replace(".", "").replace("(", "").replace(")", "")
    return t.isdigit()


def classify(
    path: Path,
    original_name: str,
    sheets: dict[str, pd.DataFrame] | None = None,
) -> ClassifyResult:
    blob = _text_blob(path, original_name, sheets=sheets)
    scores: dict[str, int] = {k: 0 for k in SLOT_KEYS}
    for slot, kws in SLOT_KEYWORDS.items():
        for kw in kws:
            if kw.lower() in blob:
                scores[slot] += 1

    # 동률/0점 처리: 특정 강한 신호 보정
    # '명세서' 가 있는데 '전기/이월' 신호가 약하면 전기 명세서로
    best_slot = max(scores, key=lambda k: scores[k])
    if scores[best_slot] == 0:
        return ClassifyResult(slot=None, scores=scores)
    return ClassifyResult(slot=best_slot, scores=scores)


def auto_assign(classified: dict[str, ClassifyResult]) -> dict[str, str]:
    """여러 파일의 분류 결과를 슬롯에 1:1 배정한다.

    각 슬롯에 가장 점수가 높은 파일을 배정하되, 충돌 시 점수 우선.
    반환: {file_id: slot_key}
    """
    # (file_id, slot, score) 후보를 점수 내림차순으로 정렬해 그리디 배정
    candidates: list[tuple[int, str, str]] = []
    for file_id, res in classified.items():
        for slot, score in res.scores.items():
            if score > 0:
                candidates.append((score, file_id, slot))
    candidates.sort(reverse=True)

    assigned_slot: set[str] = set()
    assigned_file: set[str] = set()
    result: dict[str, str] = {}
    for score, file_id, slot in candidates:
        if file_id in assigned_file or slot in assigned_slot:
            continue
        result[file_id] = slot
        assigned_file.add(file_id)
        assigned_slot.add(slot)
    return result
