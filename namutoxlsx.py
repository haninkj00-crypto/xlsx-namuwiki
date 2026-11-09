#!/usr/bin/env python3
"""Convert a NamuMark table block to an Excel .xlsx workbook."""
from __future__ import annotations

import argparse
import re
import sys
import textwrap
import warnings

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print(textwrap.dedent("""
        [오류] openpyxl 모듈이 없습니다.
          python -m pip install openpyxl
          Debian/Ubuntu: sudo apt install python3-openpyxl
    """).strip(), file=sys.stderr)
    raise SystemExit(1)

from namu_common import ParsedCell, TableData, parse_table, strip_markup_details

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

EXCEL_MAX_COLS = 16384
EXCEL_MAX_ROWS = 1048576


def _to_px(value, base_px: float | None = None) -> float | None:
    if value is None:
        return None
    s = str(value).strip().lower()
    m = re.fullmatch(r"([0-9.]+)(px|pt|em|%)?", s)
    if not m:
        return None
    number = float(m.group(1))
    unit = m.group(2) or "px"
    if unit == "%":
        return base_px * number / 100.0 if base_px is not None else None
    if unit == "pt":
        return number * 96.0 / 72.0
    if unit == "em":
        return number * 16.0
    return number


def _excel_width_from_px(px: float) -> float:
    # Approximate Excel's character-width units using Calibri 11 metrics.
    return max(0.1, (px - 5.0) / 7.0)


def _set_font(cell, *, bold=None, italic=None, underline=None, strike=None, color=None):
    old = cell.font
    cell.font = Font(
        name=old.name, sz=old.sz, b=old.b if bold is None else bold,
        i=old.i if italic is None else italic,
        u=old.u if underline is None else ("single" if underline else None),
        strike=old.strike if strike is None else strike,
        color=("FF" + color if isinstance(color, str) and re.fullmatch(r"[0-9A-Fa-f]{6}", color) else (color or old.color)), vertAlign=old.vertAlign,
        charset=old.charset, family=old.family, scheme=old.scheme,
        outline=old.outline, shadow=old.shadow, condense=old.condense,
        extend=old.extend,
    )


def _set_fill(cell, color):
    if color:
        cell.fill = PatternFill(fill_type="solid", fgColor="FF" + color)


def _apply_border(ws, rows: int, cols: int, color: str):
    side = Side(style="thin", color="FF" + color)
    border = Border(left=side, right=side, top=side, bottom=side)
    for row in ws.iter_rows(min_row=1, max_row=rows, min_col=1, max_col=cols):
        for cell in row:
            cell.border = border


def _build_grid(table: TableData, max_columns: int = EXCEL_MAX_COLS):
    """Place cells on a real occupancy grid; no arbitrary +50 padding."""
    if not table.rows:
        return [], 0, 0
    occupancy: set[tuple[int, int]] = set()
    anchors: dict[tuple[int, int], ParsedCell] = {}
    num_rows = len(table.rows)
    max_cols = 0

    for r, cells in enumerate(table.rows):
        col = 0
        for cell in cells:
            # A new cell starts at the first free position. Its whole colspan
            # must also be free of rowspans originating above this row.
            while True:
                while (r, col) in occupancy:
                    col += 1
                if all((r, col + dc) not in occupancy for dc in range(cell.colspan)):
                    break
                col += 1
            end_col = col + cell.colspan
            if end_col > max_columns:
                raise ValueError(
                    f"표의 열 수가 제한({max_columns})을 초과합니다. 셀 {r + 1}행의 병합 속성을 확인하십시오."
                )
            end_row = r + cell.rowspan
            if end_row > EXCEL_MAX_ROWS:
                raise ValueError("rowspan을 포함한 행 수가 Excel 최대 행 수를 초과합니다.")

            # Detect overlaps rather than silently overwriting another cell.
            for rr in range(r, end_row):
                for cc in range(col, end_col):
                    if (rr, cc) in occupancy:
                        raise ValueError(f"병합 셀이 겹칩니다: 행 {rr + 1}, 열 {cc + 1}")
            anchors[(r, col)] = cell
            for rr in range(r, end_row):
                for cc in range(col, end_col):
                    if (rr, cc) != (r, col):
                        occupancy.add((rr, cc))
            col = end_col
            max_cols = max(max_cols, col)
            num_rows = max(num_rows, end_row)

    if max_cols > max_columns:
        raise ValueError(f"표 너비가 최대 열 제한({max_columns})을 초과합니다.")
    grid = [[None for _ in range(max_cols)] for _ in range(num_rows)]
    for (r, c), cell in anchors.items():
        grid[r][c] = cell
    return grid, num_rows, max_cols


def _write_grid(
    table: TableData,
    grid,
    num_rows: int,
    num_cols: int,
    output_path: str,
    sheet_name: str,
    include_caption: bool = True,
):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    caption_text = strip_markup_details(table.caption).text if table.caption and include_caption else ""
    row_offset = 1 if caption_text else 0
    if caption_text:
        c = ws.cell(1, 1, caption_text)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if num_cols > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, num_cols))

    table_style = table.table_attrs
    table_bg = table_style.get("tablebgcolor")
    table_color = table_style.get("tablecolor")
    border_color = table_style.get("tablebordercolor")
    column_styles: dict[str, dict[int, str]] = {"colbgcolor": {}, "colcolor": {}}

    # A numeric table width is used as a soft total-width hint, not a strict
    # browser pixel layout. Explicit cell widths take precedence.
    table_px = _to_px(table_style.get("tablewidth"))
    if table_px and num_cols:
        default_px = table_px / num_cols
        for col in range(1, num_cols + 1):
            ws.column_dimensions[get_column_letter(col)].width = _excel_width_from_px(default_px)

    for r in range(num_rows):
        row_data = [x for x in grid[r] if x is not None]
        row_style: dict[str, str] = {}
        header_row = any(cell_data.options.get("thead") for cell_data in row_data)
        for cell_data in row_data:
            if cell_data.options.get("rowbgcolor"):
                row_style["rowbgcolor"] = cell_data.options["rowbgcolor"]
            if cell_data.options.get("rowcolor"):
                row_style["rowcolor"] = cell_data.options["rowcolor"]
        if row_style.get("rowbgcolor"):
            for col in range(1, num_cols + 1):
                _set_fill(ws.cell(r + 1 + row_offset, col), row_style["rowbgcolor"])
        if row_style.get("rowcolor"):
            for col in range(1, num_cols + 1):
                _set_font(ws.cell(r + 1 + row_offset, col), color=row_style["rowcolor"])

        for col, cell_data in enumerate(grid[r]):
            if cell_data is None:
                continue
            er, ec = r + 1 + row_offset, col + 1
            details = strip_markup_details(cell_data.content)
            cell = ws.cell(er, ec, details.text)
            cell.alignment = Alignment(wrap_text=True, vertical="top")

            # Table/row/column/cell precedence mirrors NamuMark's broad model.
            opts = cell_data.options
            if opts.get("colbgcolor"):
                column_styles["colbgcolor"][col] = opts["colbgcolor"]
            if opts.get("colcolor"):
                column_styles["colcolor"][col] = opts["colcolor"]
            bg = opts.get("bgcolor") or column_styles["colbgcolor"].get(col) or row_style.get("rowbgcolor") or table_bg
            fg = details.color or opts.get("color") or column_styles["colcolor"].get(col) or row_style.get("rowcolor") or table_color
            _set_fill(cell, bg)
            _set_font(
                cell,
                bold=details.bold or header_row or bool(opts.get("thead")),
                italic=details.italic,
                underline=details.underline,
                strike=details.strike,
                color=fg,
            )
            if details.hyperlink:
                cell.hyperlink = details.hyperlink

            horizontal = opts.get("align")
            vertical = opts.get("valign")
            cell.alignment = Alignment(
                horizontal=horizontal if horizontal in {"left", "center", "right"} else "general",
                vertical=vertical if vertical in {"top", "center", "bottom"} else "top",
                wrap_text=True,
            )

            width_px = _to_px(opts.get("width"), table_px)
            if width_px and cell_data.colspan == 1:
                ws.column_dimensions[get_column_letter(ec)].width = _excel_width_from_px(width_px)
            elif width_px and cell_data.colspan > 1:
                each = _excel_width_from_px(width_px / cell_data.colspan)
                for cc in range(ec, ec + cell_data.colspan):
                    ws.column_dimensions[get_column_letter(cc)].width = each

            height_px = _to_px(opts.get("height"))
            if height_px:
                ws.row_dimensions[er].height = height_px * 0.75

            if cell_data.colspan > 1 or cell_data.rowspan > 1:
                ws.merge_cells(
                    start_row=er, start_column=ec,
                    end_row=er + cell_data.rowspan - 1,
                    end_column=ec + cell_data.colspan - 1,
                )

    # Table defaults apply also to unfilled cells. Apply them after writing.
    for row in ws.iter_rows(min_row=1 + row_offset, max_row=num_rows + row_offset,
                            min_col=1, max_col=max(1, num_cols)):
        for cell in row:
            if table_bg and cell.fill.fill_type is None:
                _set_fill(cell, table_bg)
            if table_color and cell.font.color is None:
                _set_font(cell, color=table_color)
    if border_color and num_rows and num_cols:
        _apply_border(ws, num_rows + row_offset, num_cols, border_color)

    # Freeze panes for long tables and set a readable default view.
    if num_rows >= 2 and any(c.options.get("thead") for c in table.rows[0]):
        ws.freeze_panes = f"A{2 + row_offset}"
    ws.sheet_view.showGridLines = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    wb.save(output_path)
    wb.close()


def namu_to_xlsx(text: str, output_path: str, sheet_name: str = "Sheet1",
                 strict: bool = False, max_columns: int = EXCEL_MAX_COLS,
                 include_caption: bool = True):
    table = parse_table(text, strict=strict)
    if not table.rows:
        raise ValueError("표 행(||로 시작하는 행)을 찾지 못했습니다.")
    grid, rows, cols = _build_grid(table, max_columns=max_columns)
    if not cols:
        raise ValueError("표의 열 수를 계산하지 못했습니다.")
    _write_grid(table, grid, rows, cols, output_path, sheet_name, include_caption)
    return rows, cols, table.warnings


def build_parser():
    p = argparse.ArgumentParser(
        prog="namutoxlsx",
        description="나무위키 표 문법을 Excel(.xlsx)로 변환합니다.",
        epilog="예시:\n  python namutoxlsx.py input.txt output.xlsx\n  cat raw.txt | python namutoxlsx.py - out.xlsx\n  python namutoxlsx.py input.txt out.xlsx --sheet 시간표",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("input", help="입력 텍스트 파일 경로 (- 이면 stdin)")
    p.add_argument("output", help="출력 .xlsx 파일 경로")
    p.add_argument("--sheet", default="Sheet1", help="시트 이름 (기본: Sheet1)")
    p.add_argument("--strict", action="store_true", help="잘못된 표 구문이 있으면 복구하지 않고 중단")
    p.add_argument("--max-columns", type=int, default=EXCEL_MAX_COLS,
                   help=f"최대 표 열 수 (기본: {EXCEL_MAX_COLS})")
    p.add_argument("--no-caption", action="store_true", help="|캡션| 형식의 캡션 행을 생략")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.max_columns < 1 or args.max_columns > EXCEL_MAX_COLS:
            raise ValueError(f"--max-columns는 1~{EXCEL_MAX_COLS} 범위여야 합니다.")
        if args.input == "-":
            if sys.stdin.isatty():
                print("나무위키 표 텍스트를 입력한 뒤 Ctrl+D(Linux/macOS) 또는 Ctrl+Z 후 Enter(Windows)를 누르세요.", file=sys.stderr)
            source = sys.stdin.read()
        else:
            with open(args.input, "r", encoding="utf-8-sig") as f:
                source = f.read()
        rows, cols, notes = namu_to_xlsx(
            source, args.output, sheet_name=args.sheet, strict=args.strict,
            max_columns=args.max_columns, include_caption=not args.no_caption,
        )
        for note in notes:
            print("[경고] " + note, file=sys.stderr)
        print(f"저장 완료: {args.output} ({rows}행 × {cols}열)", file=sys.stderr)
        return 0
    except Exception as exc:
        print(f"[오류] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
