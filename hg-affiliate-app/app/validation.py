"""입력 파일 선행 검증 (절대 준수 a).

업로드 즉시 다음을 순서대로 검사한다.
  1) 0바이트/확장자 확인
  2) 매직바이트(파일 시그니처) 확인
  3) 워크북 열기 성공 여부
  4) 시트/헤더(비어있지 않은 내용) 존재 여부
  5) 손상·암호화 파일 차단

실패 시 사용자에게 사유와 함께 알리고 처리를 진행하지 않는다.
파일 '내용'은 결과에 담지 않는다. (b/f)
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .config import ALLOWED_EXTENSIONS
from .excel_io import ExcelReadError, read_workbook

# 매직바이트 시그니처
_ZIP_SIG = b"PK\x03\x04"
_OLE_SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# 내용 존재 여부는 상단 일부 행만 샘플 (대용량 시트 전체 스캔 방지)
_CONTENT_SAMPLE_ROWS = 30


@dataclass
class ValidationResult:
    ok: bool
    filename: str
    ext: str
    reasons: list[str]
    sheet_count: int = 0


def _read_head(path: Path, n: int = 8) -> bytes:
    with path.open("rb") as f:
        return f.read(n)


def _sheet_has_content(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    sample = df.head(_CONTENT_SAMPLE_ROWS)
    for v in sample.to_numpy().ravel():
        s = str(v).strip()
        if s and s != "nan":
            return True
    return False


def validate_file(path: Path, original_name: str) -> ValidationResult:
    """검증만 수행. 분류용 시트가 필요하면 validate_file_with_sheets 를 사용."""
    result, _ = validate_file_with_sheets(path, original_name)
    return result


def validate_file_with_sheets(
    path: Path,
    original_name: str,
) -> tuple[ValidationResult, dict[str, pd.DataFrame] | None]:
    """검증 + 워크북 1회 읽기. 성공 시 시트 dict 를 함께 반환(분류 재사용)."""
    ext = Path(original_name).suffix.lower()
    reasons: list[str] = []

    try:
        size = path.stat().st_size
    except OSError:
        return ValidationResult(False, original_name, ext, ["파일을 읽을 수 없습니다."]), None
    if size == 0:
        return ValidationResult(False, original_name, ext, ["0바이트(빈) 파일입니다."]), None
    if ext not in ALLOWED_EXTENSIONS:
        return (
            ValidationResult(
                False,
                original_name,
                ext,
                [f"지원하지 않는 확장자입니다: {ext or '(없음)'} (허용: xlsx, xlsm, xls, xlsb, csv)"],
            ),
            None,
        )

    head = _read_head(path)
    if ext in {".xlsx", ".xlsm"}:
        if head.startswith(_OLE_SIG):
            return (
                ValidationResult(
                    False,
                    original_name,
                    ext,
                    ["암호화되었거나 형식이 일치하지 않는 파일로 의심됩니다(OLE 컨테이너)."],
                ),
                None,
            )
        if not head.startswith(_ZIP_SIG):
            return (
                ValidationResult(
                    False,
                    original_name,
                    ext,
                    ["엑셀(zip) 시그니처가 아닙니다. 파일이 손상되었거나 확장자가 잘못되었습니다."],
                ),
                None,
            )
    elif ext == ".xlsb":
        if not head.startswith(_ZIP_SIG):
            return (
                ValidationResult(
                    False,
                    original_name,
                    ext,
                    ["xlsb 시그니처가 아닙니다. 파일이 손상되었을 수 있습니다."],
                ),
                None,
            )
    elif ext == ".xls":
        if not head.startswith(_OLE_SIG):
            return (
                ValidationResult(
                    False,
                    original_name,
                    ext,
                    ["xls(OLE2) 시그니처가 아닙니다. 파일이 손상되었을 수 있습니다."],
                ),
                None,
            )

    try:
        sheets = read_workbook(path)
    except ExcelReadError as exc:
        msg = str(exc)
        if "encrypt" in msg.lower() or "password" in msg.lower():
            return (
                ValidationResult(
                    False,
                    original_name,
                    ext,
                    ["암호화된 파일입니다. 보호를 해제 후 다시 올려주세요."],
                ),
                None,
            )
        return (
            ValidationResult(False, original_name, ext, [f"워크북 열기에 실패했습니다: {msg}"]),
            None,
        )

    if not sheets:
        return ValidationResult(False, original_name, ext, ["시트가 존재하지 않습니다."]), None

    has_content = any(_sheet_has_content(df) for df in sheets.values())
    if not has_content:
        return (
            ValidationResult(
                False,
                original_name,
                ext,
                ["내용(헤더/데이터)이 비어 있습니다."],
                len(sheets),
            ),
            None,
        )

    return ValidationResult(True, original_name, ext, reasons, len(sheets)), sheets
