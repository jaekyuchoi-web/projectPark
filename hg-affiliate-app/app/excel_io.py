"""다중 포맷 엑셀/CSV 읽기.

확장자에 따라 적절한 엔진으로 워크북을 연다.
- .xlsx/.xlsm : openpyxl
- .xls        : xlrd
- .xlsb       : pyxlsb
- .csv        : pandas (인코딩 자동 추정)

읽기는 항상 "모든 시트 -> DataFrame" 형태로 정규화하여 반환한다.
대용량 원장 대비 dtype 추론은 보수적으로 처리한다.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


class ExcelReadError(Exception):
    """워크북 열기/읽기 실패."""


def _read_csv(path: Path) -> dict[str, pd.DataFrame]:
    last_err: Exception | None = None
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
            return {"Sheet1": df}
        except Exception as exc:  # noqa: BLE001 - 인코딩 후보 순회
            last_err = exc
    raise ExcelReadError(f"CSV 인코딩 인식 실패: {last_err}")


def _engine_for(ext: str) -> str:
    return {
        ".xlsx": "openpyxl",
        ".xlsm": "openpyxl",
        ".xls": "xlrd",
        ".xlsb": "pyxlsb",
    }[ext]


def read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    """모든 시트를 {시트명: DataFrame} 으로 읽는다. 값은 문자열로 보존."""
    ext = path.suffix.lower()
    if ext == ".csv":
        return _read_csv(path)
    if ext not in {".xlsx", ".xlsm", ".xls", ".xlsb"}:
        raise ExcelReadError(f"지원하지 않는 확장자: {ext}")

    engine = _engine_for(ext)
    try:
        sheets = pd.read_excel(
            path,
            sheet_name=None,
            engine=engine,
            dtype=str,
            header=None,  # 헤더 위치가 제각각이라 원시 매트릭스로 읽음
        )
    except Exception as exc:  # noqa: BLE001
        raise ExcelReadError(f"워크북 열기 실패({engine}): {exc}") from exc

    if not sheets:
        raise ExcelReadError("시트가 존재하지 않습니다.")
    return sheets


def sheet_names(path: Path) -> list[str]:
    """시트명 목록만 빠르게 확인 (검증용)."""
    ext = path.suffix.lower()
    if ext == ".csv":
        return ["Sheet1"]
    engine = _engine_for(ext)
    try:
        xls = pd.ExcelFile(path, engine=engine)
        return list(xls.sheet_names)
    except Exception as exc:  # noqa: BLE001
        raise ExcelReadError(f"시트 목록 확인 실패: {exc}") from exc


def detect_header_row(df: pd.DataFrame, max_scan: int = 15) -> int:
    """상단에서 헤더로 보이는 행 인덱스를 추정한다.

    먼저 회계 파일의 대표 헤더(거래처/계정/금액 등)를 찾고, 없으면
    비어있지 않은 셀이 가장 많은 행을 헤더로 간주한다.
    """
    header_tokens = {
        "partner": {"거래처", "거래처명", "상호", "업체", "거래상대", "계정상대", "성명"},
        "account": {"계정", "계정과목", "계정명", "과목", "account"},
        "amount": {"금액", "잔액", "기말잔액", "당기잔액", "balance", "amount"},
        "debit_credit": {"차변", "차변금액", "debit", "대변", "대변금액", "credit"},
    }
    limit = min(max_scan, len(df))
    for r in range(limit):
        values = {str(v).strip() for v in df.iloc[r].tolist() if str(v).strip() and str(v) != "nan"}
        has_partner = bool(values & header_tokens["partner"])
        has_account = bool(values & header_tokens["account"])
        has_amount = bool(values & header_tokens["amount"])
        has_debit_credit = bool(values & header_tokens["debit_credit"])
        if (has_partner and has_amount) or (has_account and (has_amount or has_debit_credit)):
            return r

    best_row = 0
    best_count = -1
    for r in range(limit):
        count = int(df.iloc[r].map(lambda v: bool(str(v).strip()) and str(v) != "nan").sum())
        if count > best_count:
            best_count = count
            best_row = r
    return best_row


def normalized_frame(df: pd.DataFrame) -> pd.DataFrame:
    """추정 헤더 행을 컬럼명으로 승격한 DataFrame 을 반환한다."""
    if df.empty:
        return df
    hr = detect_header_row(df)
    header = [str(v).strip() for v in df.iloc[hr].tolist()]
    body = df.iloc[hr + 1 :].copy()
    body.columns = _dedupe(header)
    body = body.reset_index(drop=True)
    return body


def _dedupe(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for name in names:
        base = name if name and name != "nan" else "col"
        if base in seen:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
        else:
            seen[base] = 0
            out.append(base)
    return out
