"""처리 파이프라인 오케스트레이션 (§3).

검증/분류는 라우트에서 선행하고, 여기서는 매핑된 슬롯 파일로
정규화 → 38.x 집계 → 자체검증 → 출력 생성 → 출력 안전성 검증을 수행한다.
산출 수치는 반환·로그하지 않으며, 출력 파일 경로와 '건수' 요약만 반환한다.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from . import excel_io
from .config import Settings
from .domain import columns as C
from .domain.aggregate import AggregateResult, aggregate_balance, aggregate_ledger
from .domain.error_report import build_error_report
from .domain.errors import ErrorLog
from .domain.governance import build_governance
from .domain.statement import build_statement
from .normalize import normalize_names
from .output_check import verify_output


@dataclass
class PipelineOutcome:
    ok: bool
    message: str
    outputs: dict[str, Path]   # 이름 -> 경로
    error_count: int           # 검토/오류 건수(수치 아님)


def _best_frame(path: Path | None) -> pd.DataFrame | None:
    """워크북에서 거래처 컬럼을 가진(없으면 가장 큰) 시트를 정규화해 반환."""
    if path is None:
        return None
    try:
        sheets = excel_io.read_workbook(path)
    except excel_io.ExcelReadError:
        return None
    best: pd.DataFrame | None = None
    best_score = -1
    for df in sheets.values():
        if df is None or df.empty:
            continue
        norm = excel_io.normalized_frame(df)
        score = len(norm)
        if C.resolve_column(norm, "partner") is not None:
            score += 1_000_000  # 거래처 컬럼 보유 시트 우선
        if score > best_score:
            best_score = score
            best = norm
    return best


def _balance_long_frame(path: Path | None) -> pd.DataFrame | None:
    """잔액명세서(시트=계정 구조)를 거래처/계정과목/금액 long 프레임으로 변환.

    실제 잔액명세서는 한 워크북에 '외상매출금/받을어음/단기대여금/미수금/선급금/
    외상매입금/미지급금/예수금...' 처럼 **시트마다 1개 계정**이 들어있고, 각 시트에는
    거래처와 금액만 있고 계정 컬럼이 없다. 이를 다음 규칙으로 평면화한다.
      - 시트에 (계정 컬럼 + 금액 컬럼)이 이미 있으면(원장형) 그대로 누적
      - 없으면 시트명을 계정과목으로 사용
      - 시트 안에 '대손충당금' 컬럼이 있으면 별도의 '대손충당금' 계정행으로 추가(38.3)
    이렇게 하면 기존 38.2/38.3 집계가 계정명 기준으로 정상 매칭된다.
    """
    if path is None:
        return None
    try:
        sheets = excel_io.read_workbook(path)
    except excel_io.ExcelReadError:
        return None

    rows: list[pd.DataFrame] = []
    for sname, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        df = excel_io.normalized_frame(raw)
        pcol = C.resolve_column(df, "partner")
        amtcol = C.resolve_column(df, "amount")
        if pcol is None or amtcol is None:
            continue

        acol = C.resolve_column(df, "account")
        if acol is not None and acol != amtcol:
            # 이미 계정/금액 컬럼이 있는 평면형 시트
            sub = pd.DataFrame({
                "거래처": df[pcol].astype(str),
                "계정과목": df[acol].astype(str),
                "금액": df[amtcol],
            })
            rows.append(sub)
        else:
            # 시트명을 계정과목으로 사용
            rows.append(pd.DataFrame({
                "거래처": df[pcol].astype(str),
                "계정과목": str(sname).strip(),
                "금액": df[amtcol],
            }))

        # 시트 내 대손충당금 컬럼 → 별도 계정행 (38.3 대손충당금)
        allow_col = next((str(c) for c in df.columns if "대손충당금" in str(c)), None)
        if allow_col is not None and allow_col != amtcol:
            rows.append(pd.DataFrame({
                "거래처": df[pcol].astype(str),
                "계정과목": "대손충당금",
                "금액": df[allow_col],
            }))

    if not rows:
        return None
    return pd.concat(rows, ignore_index=True)


def _distinct_partners(frames: list[pd.DataFrame | None]) -> list[str]:
    names: set[str] = set()
    for df in frames:
        if df is None:
            continue
        col = C.resolve_column(df, "partner")
        if col is None:
            continue
        for v in df[col].tolist():
            s = str(v).strip()
            if s and s != "nan" and not C.is_aggregate_label(s):
                names.add(s)
    return sorted(names)


def _verify_reconstruction(
    prev_balance: pd.DataFrame | None,
    ledger: pd.DataFrame | None,
    current_balance: pd.DataFrame | None,
    mapping: dict[str, str],
    canonical: set[str],
    errors: ErrorLog,
) -> None:
    """(전기이월+당기거래) 재구성 잔액 ↔ 검증용(당기말) 대조.

    불일치는 검증용에 맞추고(이미 38.2는 검증용 기준) 오류목록에 기록한다.
    대표적으로 매출채권/매입채무에 대해 점검.
    """
    if prev_balance is None or ledger is None or current_balance is None:
        return
    checks = {
        "매출채권": (["외상매출금", "받을어음"], "asset"),
        "매입채무": (["외상매입금"], "liability"),
    }
    acc_col = C.resolve_column(ledger, "account")
    debit_col = C.resolve_column(ledger, "debit")
    credit_col = C.resolve_column(ledger, "credit")
    if acc_col is None:
        return

    from .domain.aggregate import _balance_by_canon  # 재사용

    for bucket, (kws, kind) in checks.items():
        prev_b = _balance_by_canon(prev_balance, mapping, canonical, kws)
        cur_b = _balance_by_canon(current_balance, mapping, canonical, kws)
        # 원장 순변동
        ledger_net: dict[str, float] = {}
        partner_col = C.resolve_column(ledger, "partner")
        if partner_col:
            for _, row in ledger.iterrows():
                raw = str(row.get(partner_col, "")).strip()
                canon = mapping.get(raw)
                if not canon or canon not in canonical:
                    continue
                if C.match_bucket(str(row.get(acc_col, "")), kws):
                    d = C.to_number(row.get(debit_col)) if debit_col else 0.0
                    c = C.to_number(row.get(credit_col)) if credit_col else 0.0
                    delta = (d - c) if kind == "asset" else (c - d)
                    ledger_net[canon] = ledger_net.get(canon, 0.0) + delta
        for canon in canonical:
            recon = prev_b.get(canon, 0.0) + ledger_net.get(canon, 0.0)
            actual = cur_b.get(canon, 0.0)
            if abs(recon - actual) > 1.0:
                errors.add(
                    "검증 불일치", f"{canon} / {bucket}",
                    "재구성(전기이월+당기거래) 잔액이 검증용 당기말 잔액과 일치하지 않음",
                    "검증용 잔액명세서를 기준으로 채택했습니다. 원장/이월 분개를 확인하세요.",
                )


def run_pipeline(
    slot_paths: dict[str, Path | None],
    settings: Settings,
    output_dir: Path,
) -> PipelineOutcome:
    errors = ErrorLog()

    # 1) 슬롯별 프레임 로드
    #   - 잔액명세서(전기/당기)는 '시트=계정' 구조 → long 프레임으로 평면화
    #   - 당기 원장은 평면 long(계정명/거래처명/차변/대변) → 가장 큰 시트 사용
    prev_balance = _balance_long_frame(slot_paths.get("prev_balance"))
    ledger = _best_frame(slot_paths.get("current_ledger"))
    current_balance = _balance_long_frame(slot_paths.get("current_balance"))

    if current_balance is None:
        errors.add("형식 경고", "당기말 잔액명세서", "검증용 잔액명세서를 읽지 못함", "파일/시트/헤더를 확인하세요.")

    # 2) 사업자명 정규화
    distinct = _distinct_partners([ledger, current_balance, prev_balance])
    norm = normalize_names(distinct, slot_paths.get("name_pivot"), settings)
    canonical = norm.canonical or set(norm.mapping.values())
    if not canonical:
        errors.add("산출 불가", "특관자 집합", "정규(대표) 법인명 집합을 구성하지 못함",
                   "특관자 상호 정리 파일을 보완하거나 OPENAI_API_KEY를 설정하세요.")
    errors.extend_unmatched(norm.unmatched)

    # 3~5) 집계
    result = AggregateResult()
    aggregate_balance(current_balance, prev_balance, norm.mapping, canonical, result, errors)
    aggregate_ledger(ledger, prev_balance, current_balance, norm.mapping, canonical, result, errors)

    # 6) 자체 검증
    _verify_reconstruction(prev_balance, ledger, current_balance, norm.mapping, canonical, errors)

    # 38.5 경영진 보상 안내(공란)
    errors.add("자동 산출 불가", "주요 경영진 보상(38.5)",
               "인사·급여(임원 식별/배분) 자료가 없어 숫자 공란 처리", "임원 급여 자료로 별도 보완하세요.")

    # 7) 출력 생성
    statement_wb, stmt_unmatched = build_statement(result)
    for canon in stmt_unmatched:
        errors.add(
            "템플릿 미반영", canon,
            "데이터가 있으나 명세서 양식의 특수관계자 행에 없어 자동 반영되지 못함",
            "특관자 양식에 해당 법인 행을 수기로 추가하거나, 상호정리 표기를 양식과 일치시키세요.",
        )
    governance_wb = build_governance(norm.mapping, canonical)

    statement_path = output_dir / "당기_특관자_명세서.xlsx"
    governance_path = output_dir / "지배구조.xlsx"
    statement_wb.save(statement_path)
    governance_wb.save(governance_path)

    # 8) 출력 안전성 검증 (명세서·지배구조)
    for label, p in (("당기 특관자 명세서", statement_path), ("지배구조", governance_path)):
        check = verify_output(p)
        if not check.ok:
            errors.add("출력 검증 실패", label, check.reason, "관리자에게 문의하세요(다운로드 차단).")
            error_wb = build_error_report(errors)
            error_path = output_dir / "오류목록.xlsx"
            error_wb.save(error_path)
            return PipelineOutcome(
                ok=False,
                message=f"{label} 출력 검증에 실패하여 다운로드를 차단했습니다.",
                outputs={"오류목록": error_path},
                error_count=errors.count,
            )

    # 오류목록은 항상 생성
    error_wb = build_error_report(errors)
    error_path = output_dir / "오류목록.xlsx"
    error_wb.save(error_path)

    return PipelineOutcome(
        ok=True,
        message="명세서 생성을 완료했습니다.",
        outputs={
            "당기_특관자_명세서": statement_path,
            "지배구조": governance_path,
            "오류목록": error_path,
        },
        error_count=errors.count,
    )
