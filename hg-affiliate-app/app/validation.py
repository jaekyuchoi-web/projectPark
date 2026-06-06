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

from .config import ALLOWED_EXTENSIONS
from .excel_io import ExcelReadError, read_workbook, sheet_names

# 매직바이트 시그니처
_ZIP_SIG = b"PK\x03\x04"          # xlsx/xlsm/xlsb(=zip 기반)
_OLE_SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # 구형 xls(=OLE2)
# 암호화된 OOXML 도 OLE 컨테이너(CDF) 로 저장됨 → xlsx 인데 OLE 시그니처면 암호화 의심


@dataclass
class ValidationResult:
    ok: bool
    filename: str
    ext: str
    reasons: list[str]            # 실패/경고 사유 (수치·내용 미포함)
    sheet_count: int = 0


def _read_head(path: Path, n: int = 8) -> bytes:
    with path.open("rb") as f:
        return f.read(n)


def validate_file(path: Path, original_name: str) -> ValidationResult:
    ext = Path(original_name).suffix.lower()
    reasons: list[str] = []

    # 1) 0바이트 / 확장자
    try:
        size = path.stat().st_size
    except OSError:
        return ValidationResult(False, original_name, ext, ["파일을 읽을 수 없습니다."])
    if size == 0:
        return ValidationResult(False, original_name, ext, ["0바이트(빈) 파일입니다."])
    if ext not in ALLOWED_EXTENSIONS:
        return ValidationResult(
            False, original_name, ext,
            [f"지원하지 않는 확장자입니다: {ext or '(없음)'} (허용: xlsx, xlsm, xls, xlsb, csv)"],
        )

    # 2) 매직바이트
    head = _read_head(path)
    if ext in {".xlsx", ".xlsm"}:
        if head.startswith(_OLE_SIG):
            return ValidationResult(False, original_name, ext,
                                    ["암호화되었거나 형식이 일치하지 않는 파일로 의심됩니다(OLE 컨테이너)."])
        if not head.startswith(_ZIP_SIG):
            return ValidationResult(False, original_name, ext,
                                    ["엑셀(zip) 시그니처가 아닙니다. 파일이 손상되었거나 확장자가 잘못되었습니다."])
    elif ext == ".xlsb":
        if not head.startswith(_ZIP_SIG):
            return ValidationResult(False, original_name, ext,
                                    ["xlsb 시그니처가 아닙니다. 파일이 손상되었을 수 있습니다."])
    elif ext == ".xls":
        if not head.startswith(_OLE_SIG):
            return ValidationResult(False, original_name, ext,
                                    ["xls(OLE2) 시그니처가 아닙니다. 파일이 손상되었을 수 있습니다."])
    # csv 는 시그니처 검사 생략

    # 3) 워크북 열기 + 4) 시트/헤더 존재
    try:
        names = sheet_names(path)
        if not names:
            return ValidationResult(False, original_name, ext, ["시트가 존재하지 않습니다."])
        sheets = read_workbook(path)
    except ExcelReadError as exc:
        msg = str(exc)
        if "encrypt" in msg.lower() or "password" in msg.lower():
            return ValidationResult(False, original_name, ext, ["암호화된 파일입니다. 보호를 해제 후 다시 올려주세요."])
        return ValidationResult(False, original_name, ext, [f"워크북 열기에 실패했습니다: {msg}"])

    # 비어있지 않은 내용(헤더) 1개 이상 확인
    has_content = False
    for df in sheets.values():
        if df is not None and not df.empty:
            non_empty = df.map(lambda v: bool(str(v).strip()) and str(v) != "nan").to_numpy().any()
            if non_empty:
                has_content = True
                break
    if not has_content:
        return ValidationResult(False, original_name, ext, ["내용(헤더/데이터)이 비어 있습니다."], len(names))

    return ValidationResult(True, original_name, ext, reasons, len(names))
