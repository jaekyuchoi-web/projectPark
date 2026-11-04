"""당기 특관자 명세서 생성 (출력물 1).

시트 구성: 특관자 / 39.1 / 39.2 / _집계(데이터)
- _집계: 정규명별 38.x 산출값(데이터 소스)
- 특관자: SUMIF 로 _집계를 자동 합산(사업보고서 게재용). 수식만 사용해 재현성 확보.
- 39.2: 정규명 x 항목 피벗.

수식은 동일 워크북 내 _집계 시트만 참조하므로 LibreOffice 재계산 시 오류가 없다.
값/수치는 화면·로그로 내보내지 않으며 엑셀 파일에만 기록한다.
"""

from __future__ import annotations

from collections.abc import Sequence
from copy import copy
from datetime import date, datetime
from math import ceil
from pathlib import Path
from unicodedata import east_asian_width

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .aggregate import AggregateResult
from .period_extract import Period
from .statement_detail import (
    StatementDetailRow,
    detail_balance_formula,
    detail_group_key,
    write_detail_cell,
)
from .statement_template import fill_statement_template, write_period_label

# 고정 템플릿(정답 양식). 존재하면 이 양식으로 산출, 없으면 단순표 폴백.
TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "assets" / "statement_template.xlsx"

# (표시 헤더, Aggregate 속성) — 38.5 경영진보상은 공란이라 값 컬럼 없음
COLUMNS: list[tuple[str, str]] = [
    ("매출", "sales"),
    ("매출등 기타", "sales_other"),
    ("매입", "purchase"),
    ("매입등 기타", "purchase_other"),
    ("매출채권", "ar"),
    ("대여금", "loan"),
    ("기타채권", "other_ar"),
    ("투자전환사채", "cb_invest"),
    ("기타채무", "other_ap"),
    ("발행전환사채", "cb_issued"),
    ("매입채무", "ap"),
    ("대손충당금", "allowance_end"),
    ("대손상각비", "allowance_expense"),
    ("자금대여", "fund_lending"),
    ("차입금", "borrowing"),
    ("전기말 대손충당금", "allowance_opening"),
    ("자금대여상환", "fund_collection"),
    ("차입", "fund_borrowing"),
    ("차입상환", "fund_repayment"),
    ("투자전환사채취득", "fund_investment"),
    ("투자전환사채처분", "fund_disposal"),
]

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center")
_NUMFMT = "#,##0"


def _style_header(cell) -> None:
    cell.fill = _HEADER_FILL
    cell.font = _HEADER_FONT
    cell.alignment = _CENTER
    cell.border = _BORDER


def build_statement(
    result: AggregateResult,
    period: Period,
    detail_rows: Sequence[StatementDetailRow],
    template_path: Path | None = None,
) -> tuple[Workbook, list[str]]:
    """당기 특관자 명세서 워크북을 만든다.

    Returns:
        (workbook, unmatched) — 고정 템플릿 사용 시 템플릿 거래처에 매칭되지 못한
        정규명 목록(수기 추가 검토). 폴백 경로에선 빈 목록.
    """
    if template_path is not None:
        return _build_revised(result, period, detail_rows, template_path), []
    if TEMPLATE_PATH.exists():
        return fill_statement_template(TEMPLATE_PATH, result, period, detail_rows)
    return _build_legacy(result, period, detail_rows), []


def _build_revised(result, period, detail_rows, template_path):
    """2026-3Q 규칙: 모든 버킷/거래처를 표시하고 전기 양식은 참고로 보존."""
    wb = _build_legacy(result, period, detail_rows, detail_header_row=2)
    del wb['특관자']
    del wb['_집계']
    ws = wb.create_sheet('특관자', 0)
    ws.append(['HLB글로벌(주) 특수관계자 명세서'])
    ws.append([period.label])
    ws['J2'] = '단위: 원'
    ws.merge_cells('A1:J1')
    ws.merge_cells('A2:H2')
    ws['A1'].font = Font(size=16, bold=True, color='1F4E79')
    prior = load_workbook(template_path, data_only=True)
    groups = {}
    if '특관자' in prior.sheetnames:
        original = prior['특관자']
        for row in original.iter_rows(min_col=2, max_col=3):
            group, name = (c.value for c in row)
            if isinstance(group, str) and isinstance(name, str) and ('기업' in group or '특수관계' in group):
                from ..normalize import _norm_key
                groups[_norm_key(name)] = group
        ref = wb.create_sheet('전기_참고')
        ref.append(['제공된 전기 명세서 참고값 — 당기 집계에 포함하지 않음'])
        from copy import copy
        for row in original.iter_rows():
            for cell in row:
                if cell.value is not None:
                    value = '원본 수식 오류(확인 필요)' if cell.data_type == 'e' else cell.value
                    target = ref.cell(cell.row + 2, cell.column, value)
                    target.font = copy(cell.font)
                    target.fill = copy(cell.fill)
                    target.border = copy(cell.border)
                    target.alignment = copy(cell.alignment)
                    target.number_format = cell.number_format
        for key, dim in original.column_dimensions.items():
            ref.column_dimensions[key].width = dim.width
        for column in range(4, 27):
            ref.column_dimensions[get_column_letter(column)].width = 20
        ref.column_dimensions['C'].width = 35
        ref.column_dimensions['R'].width = 35
        ref.sheet_view.showGridLines = False
        for row in ref:
            for cell in row:
                cell.alignment = Alignment(vertical='center', wrap_text=False)
        for merged in original.merged_cells.ranges:
            ref.merge_cells(start_row=merged.min_row+2, end_row=merged.max_row+2,
                            start_column=merged.min_col, end_column=merged.max_col)
    prior.close()
    sections = [
        ('38.1 특수관계자와의 주요 거래내역',
         [('매출', 'sales'), ('기타수익', 'sales_other'), ('매입', 'purchase'), ('기타비용', 'purchase_other')]),
        ('38.2 특수관계자와의 채권·채무 잔액',
         [('매출채권', 'ar'), ('대여금', 'loan'), ('기타채권', 'other_ar'), ('투자전환사채', 'cb_invest'),
          ('매입채무', 'ap'), ('기타채무', 'other_ap'), ('차입금', 'borrowing'), ('발행전환사채', 'cb_issued')]),
        ('38.3 특수관계자 채권 관련 대손충당금',
         [('전기말 충당금', 'allowance_opening'), ('당기말 충당금', 'allowance_end'), ('당기 대손상각비', 'allowance_expense')]),
        ('38.4 특수관계자와의 자금거래 내역',
         [('자금대여', 'fund_lending'), ('자금대여상환', 'fund_collection'), ('차입', 'fund_borrowing'),
          ('차입상환', 'fund_repayment'), ('투자전환사채취득', 'fund_investment'), ('투자전환사채처분', 'fund_disposal')]),
    ]
    source_columns = {attr: get_column_letter(i) for i, (_, attr) in enumerate(COLUMNS, 2)}
    from ..normalize import _norm_key
    for title, columns in sections:
        ws.append([])
        ws.append([title])
        ws.cell(ws.max_row, 1).font = Font(size=12, bold=True, color='1F4E79')
        if title.startswith('38.3'):
            ws.append(['당기말 충당금 = 전기잔액 + 당기 대변 − 차변 / 대손상각비 = 당기 차변 − 대변'])
        if title.startswith('38.4'):
            ws.append(['당기 거래만 집계(이월 제외). 취득·대여와 회수·상환을 각각 표시.'])
        ws.append(['구분', '특수관계자명'] + [label for label, _ in columns])
        for cell in ws[ws.max_row][:len(columns) + 2]:
            _style_header(cell)
        first = ws.max_row + 1
        names = [name for name, agg in sorted(result.by_canonical.items())
                 if any(getattr(agg, attr) for _, attr in columns)]
        for name in names:
            r = ws.max_row + 1
            ws.cell(r, 1, groups.get(_norm_key(name), '특수관계자'))
            ws.cell(r, 2, name)
            for c, (_, attr) in enumerate(columns, 3):
                letter = source_columns[attr]
                ws.cell(r, c, f'=SUMIF(\'39.2\'!$A$4:$A${max(4, len(result.by_canonical)+3)},B{r},\'39.2\'!${letter}$4:${letter}${max(4, len(result.by_canonical)+3)})')
                ws.cell(r, c).number_format = '#,##0;[Red](#,##0);"-"'
        if names:
            last = ws.max_row
            r = last + 1
            ws.cell(r, 2, '합계')
            for c in range(3, len(columns) + 3):
                letter = get_column_letter(c)
                ws.cell(r, c, f'=SUM({letter}{first}:{letter}{last})').number_format = '#,##0;[Red](#,##0);"-"'
        else:
            ws.append(['해당 거래 없음'])
    ws.append([])
    ws.append(['38.5 주요 경영진 보상'])
    ws.append(['급여·퇴직급여 자료 미제공 — 별도 확인 필요'])
    ws.column_dimensions['A'].width = 19
    ws.column_dimensions['B'].width = 35
    for c in range(3, 11):
        ws.column_dimensions[get_column_letter(c)].width = 22
    ws.freeze_panes = 'C3'
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:J{ws.max_row}'
    detail = wb['39.1']
    detail.auto_filter.ref = f'B2:P{max(3, detail.max_row)}'
    for row in detail.iter_rows(min_row=3, min_col=14, max_col=16):
        for cell in row:
            cell.number_format = '#,##0;[Red](#,##0);"-"'
    _fit_table_cells(detail, header_row=2, first_column=2)
    _fit_table_cells(wb['39.2'], header_row=3, first_column=1)
    wb['39.2'].sheet_view.showGridLines = False
    wb.active = 0
    return wb


def _fit_table_cells(ws, header_row: int, first_column: int) -> None:
    """한글 표시 폭과 줄바꿈을 반영해 표의 열 너비·행 높이를 맞춘다."""
    def displayed(cell):
        if cell.value is None:
            return ''
        if cell.data_type == 'f':
            return '0' * 18  # 잔액 수식의 숫자 표시 공간
        if isinstance(cell.value, (datetime, date)):
            return cell.value.strftime('%Y-%m-%d')
        if isinstance(cell.value, (int, float)):
            return f'{cell.value:,.0f}'
        return str(cell.value)

    def width(text):
        return sum(2 if east_asian_width(c) in ('W', 'F') else 1 for c in text)

    ws.freeze_panes = None
    for column in range(first_column, ws.max_column + 1):
        cells = list(ws.iter_cols(min_col=column, max_col=column, min_row=header_row))[0]
        widest = max((width(line) for cell in cells for line in displayed(cell).split('\n')), default=0)
        limit = 80 if ws.title == '39.1' and column == 10 else 48
        ws.column_dimensions[get_column_letter(column)].width = min(limit, max(12, widest + 4))
    for row in ws.iter_rows(min_row=header_row, min_col=first_column):
        height = 20
        for cell in row:
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = 'center'
            cell.alignment = alignment
            available = ws.column_dimensions[cell.column_letter].width - 4
            lines = sum(max(1, ceil(width(line) / available)) for line in displayed(cell).split('\n'))
            height = max(height, lines * ((cell.font.sz or 11) + 4) + 5)
        ws.row_dimensions[row[0].row].height = height


def _build_legacy(
    result: AggregateResult,
    period: Period,
    detail_rows: Sequence[StatementDetailRow],
    detail_header_row: int = 15,
) -> Workbook:
    wb = Workbook()

    data_ws = wb.active
    data_ws.title = "_집계"
    _write_data_sheet(data_ws, result)

    main_ws = wb.create_sheet("특관자")
    _write_main_sheet(main_ws, result, data_ws.title)

    pivot_ws = wb.create_sheet("39.2")
    _write_pivot_sheet(pivot_ws, result)

    note_ws = wb.create_sheet("39.1")
    write_period_label(note_ws, period)
    headers = [
        "매입매출", "자금거래", "채권채무", "수익.비용", "구분계정과목",
        "계정코드", "계정과목", "날짜", "적요", "거래처코드", "거래처명",
        "회사명(본지점합산)", "차변", "대변", "잔액",
    ]
    for column, header in enumerate(headers, start=2):
        _style_header(note_ws.cell(row=detail_header_row, column=column, value=header))
    group_first_row = detail_header_row + 1
    previous_group_key: tuple[str, str, str] | None = None
    for row_number, detail in enumerate(detail_rows, start=detail_header_row + 1):
        current_group_key = detail_group_key(detail)
        if current_group_key != previous_group_key:
            group_first_row = row_number
        previous_group_key = current_group_key
        for column, value in enumerate(detail.as_excel_row(), start=2):
            cell = note_ws.cell(row=row_number, column=column)
            if column == 16 and detail.balance is not None:
                cell.value = detail_balance_formula(
                    detail, group_first_row, row_number
                )
            else:
                write_detail_cell(cell, value)
    # _집계 시트는 데이터 소스이므로 숨김
    data_ws.sheet_state = "hidden"
    wb.active = wb.sheetnames.index("특관자")
    return wb


def _write_data_sheet(ws, result: AggregateResult) -> None:
    headers = ["정규명"] + [h for h, _ in COLUMNS]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        _style_header(c)
    names = sorted(result.by_canonical.keys())
    for r, name in enumerate(names, start=2):
        agg = result.by_canonical[name]
        ws.cell(row=r, column=1, value=name).border = _BORDER
        for col, (_, attr) in enumerate(COLUMNS, start=2):
            cell = ws.cell(row=r, column=col, value=round(getattr(agg, attr), 0))
            cell.number_format = _NUMFMT
            cell.border = _BORDER
    ws.cell(row=1, column=1).comment = None


def _write_main_sheet(ws, result: AggregateResult, data_title: str) -> None:
    ws["A1"] = "HLB글로벌(주) 당기 특수관계자 명세서"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS) + 1)

    header_row = 3
    headers = ["정규(대표) 법인명"] + [h for h, _ in COLUMNS]
    for col, h in enumerate(headers, start=1):
        _style_header(ws.cell(row=header_row, column=col, value=h))

    names = sorted(result.by_canonical.keys())
    n = len(names)
    first_data = header_row + 1
    for i, name in enumerate(names):
        r = first_data + i
        ws.cell(row=r, column=1, value=name).border = _BORDER
        # _집계 시트의 같은 정규명 행을 SUMIF 로 자동 합산
        for col in range(2, len(COLUMNS) + 2):
            col_letter = get_column_letter(col)
            crit_range = f"'{data_title}'!$A:$A"
            sum_range = f"'{data_title}'!${col_letter}:${col_letter}"
            formula = f"=SUMIF({crit_range},$A{r},{sum_range})"
            cell = ws.cell(row=r, column=col, value=formula)
            cell.number_format = _NUMFMT
            cell.border = _BORDER

    # 합계 행
    if n:
        total_r = first_data + n
        ws.cell(row=total_r, column=1, value="합계").font = Font(bold=True)
        for col in range(2, len(COLUMNS) + 2):
            col_letter = get_column_letter(col)
            rng = f"{col_letter}{first_data}:{col_letter}{first_data + n - 1}"
            cell = ws.cell(row=total_r, column=col, value=f"=SUM({rng})")
            cell.number_format = _NUMFMT
            cell.font = Font(bold=True)
            cell.border = _BORDER

    ws.column_dimensions["A"].width = 28
    for col in range(2, len(COLUMNS) + 2):
        ws.column_dimensions[get_column_letter(col)].width = 14


def _write_pivot_sheet(ws, result: AggregateResult) -> None:
    ws["A1"] = "39.2 특수관계자별 거래/잔액 피벗"
    ws["A1"].font = Font(bold=True)
    header_row = 3
    headers = ["정규명"] + [h for h, _ in COLUMNS]
    for col, h in enumerate(headers, start=1):
        _style_header(ws.cell(row=header_row, column=col, value=h))
    for i, name in enumerate(sorted(result.by_canonical.keys())):
        r = header_row + 1 + i
        agg = result.by_canonical[name]
        ws.cell(row=r, column=1, value=name).border = _BORDER
        for col, (_, attr) in enumerate(COLUMNS, start=2):
            cell = ws.cell(row=r, column=col, value=round(getattr(agg, attr), 0))
            cell.number_format = _NUMFMT
            cell.border = _BORDER
    ws.column_dimensions["A"].width = 28
