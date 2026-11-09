#!/usr/bin/env python3
"""Excel (.xlsx) -> NamuMark table converter."""
from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import warnings

try:
    import openpyxl
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.datetime import to_excel as datetime_to_excel, from_excel as datetime_from_excel
except ImportError:
    print("[오류] openpyxl 모듈이 없습니다. `python -m pip install openpyxl`로 설치하세요.", file=sys.stderr)
    raise SystemExit(1)

from namu_common import apply_excel_font_style, color_to_hex, escape_namuwiki_syntax, format_excel_number

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def _as_text(value) -> str:
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d" if value.time() == dt.time(0, 0) else "%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")
    return "" if value is None else str(value)


def _expand_merge_range(bounds, merged_ranges, policy: str):
    min_col, min_row, max_col, max_row = bounds
    changed = True
    while changed:
        changed = False
        for merged in merged_ranges:
            a, b, c, d = range_boundaries(str(merged))
            overlaps = not (c < min_col or a > max_col or d < min_row or b > max_row)
            contained = a >= min_col and b >= min_row and c <= max_col and d <= max_row
            if overlaps and not contained:
                if policy == "error":
                    raise ValueError(f"선택 범위가 병합 셀 {merged}의 일부만 포함합니다.")
                if policy == "expand":
                    new = (min(min_col, a), min(min_row, b), max(max_col, c), max(max_row, d))
                    if new != (min_col, min_row, max_col, max_row):
                        min_col, min_row, max_col, max_row = new
                        changed = True
    return min_col, min_row, max_col, max_row


def _temporal_to_text(value, cell, *, date_mode, epoch):
    """Convert an Excel date/time only when the caller explicitly opts in.

    `iso` is a deterministic normalized representation, not Excel's rendered
    display. `serial` emits the underlying Excel day serial. Neither attempts
    to reproduce locale-dependent Excel formatting.
    """
    if date_mode == "stop":
        raise ValueError(
            f"{cell.coordinate}: 날짜/시간 형식 셀({cell.number_format!r})을 발견해 변환을 중단했습니다. "
            "날짜 표시 방식은 Excel·운영체제·지역 설정에 따라 달라질 수 있습니다. "
            "수동 변환하려면 --date-format iso 또는 --date-format serial을 지정하십시오."
        )

    if date_mode == "iso":
        if isinstance(value, (dt.datetime, dt.date, dt.time)):
            return _as_text(value)
        if isinstance(value, dt.timedelta):
            return str(value)
        # A date-formatted cell may still be exposed as a numeric serial.
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return _as_text(datetime_from_excel(value, epoch))
        return str(value)

    if date_mode == "serial":
        if isinstance(value, dt.timedelta):
            return str(value.total_seconds() / 86400)
        if isinstance(value, (dt.datetime, dt.date, dt.time)):
            return str(datetime_to_excel(value, epoch))
        return str(value)

    raise ValueError(f"알 수 없는 date_mode: {date_mode!r}")


def _number_to_text(value, cell, *, number_mode, strict_format):
    """Convert a numeric value only with an explicit caller-selected policy."""
    if number_mode == "stop":
        raise ValueError(
            f"{cell.coordinate}: 숫자형 셀을 발견해 변환을 중단했습니다 "
            f"(값={value!r}, 숫자 서식={cell.number_format!r}). "
            "숫자·통화·백분율 등의 표시 결과는 Excel 및 지역 설정에 따라 달라질 수 있습니다. "
            "원시값을 사용하려면 --number-format raw, 기존 근사 포맷터를 사용하려면 "
            "--number-format approx를 명시하십시오."
        )
    if number_mode == "raw":
        return str(value)
    if number_mode in {"approx", "display"}:
        if number_mode == "display":
            print(
                f"[경고] {cell.coordinate}: --number-format display는 Excel 실제 화면 재현이 아니라 "
                "내장 근사 포맷터의 호환 별칭입니다.", file=sys.stderr,
            )
        try:
            return format_excel_number(value, cell.number_format)
        except ValueError as exc:
            message = (
                f"{cell.coordinate}: 근사 숫자 포맷터가 이 서식을 지원하지 않습니다: "
                f"{cell.number_format!r} ({exc})"
            )
            if strict_format:
                raise ValueError(message) from exc
            print(f"[경고] {message}; 이 셀만 원시 숫자로 출력합니다.", file=sys.stderr)
            return str(value)
    raise ValueError(f"알 수 없는 number_mode: {number_mode!r}")


def _value_to_text(
    cell, formula_cell, *, fx_mode, formula_fallback, number_mode,
    date_mode, strict_format, epoch,
):
    value = cell.value
    is_formula = (
        formula_cell is not None and isinstance(formula_cell.value, str)
        and formula_cell.value.startswith("=")
    )
    if fx_mode == "output" and is_formula and value is None:
        msg = (
            f"{cell.coordinate}: 수식 캐시값이 없습니다. openpyxl은 수식을 계산하지 않으므로 "
            "Excel/LibreOffice에서 재계산 후 저장해야 결과값을 얻을 수 있습니다."
        )
        if formula_fallback == "error":
            raise ValueError(msg)
        print("[경고] " + msg, file=sys.stderr)
        return formula_cell.value if formula_fallback == "formula" else ""

    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date, dt.time, dt.timedelta)) or getattr(cell, "is_date", False):
        return _temporal_to_text(value, cell, date_mode=date_mode, epoch=epoch)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return _number_to_text(value, cell, number_mode=number_mode, strict_format=strict_format)
    result = str(value)
    if fx_mode == "input" and getattr(cell, "quotePrefix", False):
        result = "'" + result
    return result


def _cell_attrs(cell, colspan: int, rowspan: int) -> str:
    attrs = []
    alignment = getattr(cell, "alignment", None)
    horizontal = getattr(alignment, "horizontal", None)
    vertical = getattr(alignment, "vertical", None)
    if horizontal == "center":
        attrs.append("<:>")
    elif horizontal == "right":
        attrs.append("<)>")

    fill = getattr(cell, "fill", None)
    if fill and fill.fill_type == "solid":
        bg = color_to_hex(getattr(fill, "fgColor", None)) or color_to_hex(getattr(fill, "start_color", None))
        if bg:
            attrs.append(f"<bgcolor=#{bg},#{bg}>")

    if colspan > 1:
        attrs.append(f"<-{colspan}>")
    if rowspan > 1:
        if vertical == "top":
            attrs.append(f"<^|{rowspan}>")
        elif vertical == "bottom":
            attrs.append(f"<v|{rowspan}>")
        else:
            attrs.append(f"<|{rowspan}>")
    elif vertical == "top":
        attrs.append("<^|1>")
    elif vertical == "bottom":
        attrs.append("<v|1>")
    return "".join(attrs)


def _escape_cell(cell, raw: str, apply_options) -> str:
    escaped = escape_namuwiki_syntax(raw, apply_options)
    text = apply_excel_font_style(cell, escaped)
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "[br]")


def excel_to_namuwiki(
    file_path: str,
    cell_range: str | None = None,
    sheet_name: str | None = None,
    apply_options=None,
    fx_mode: str = "output",
    merge_range: str = "expand",
    formula_fallback: str = "formula",
    number_mode: str = "stop",
    date_mode: str = "stop",
    strict_format: bool = False,
    hyperlinks: str = "display",
    wrapper: bool = True,
) -> str:
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
    if fx_mode not in {"input", "output"}:
        raise ValueError("--fx는 input 또는 output이어야 합니다.")
    if number_mode not in {"stop", "raw", "approx", "display"}:
        raise ValueError("number_mode는 stop, raw, approx 또는 호환 별칭 display여야 합니다.")
    if date_mode not in {"stop", "iso", "serial"}:
        raise ValueError("date_mode는 stop, iso, serial 중 하나여야 합니다.")

    wb_formula = openpyxl.load_workbook(file_path, data_only=False)
    wb_value = wb_formula if fx_mode == "input" else openpyxl.load_workbook(file_path, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in wb_formula.sheetnames:
                raise KeyError(f"시트를 찾을 수 없습니다: {sheet_name}")
            wf, ws = wb_formula[sheet_name], wb_value[sheet_name]
        else:
            wf, ws = wb_formula.worksheets[0], wb_value.worksheets[0]

        if cell_range:
            original = range_boundaries(cell_range)
            if any(v is None for v in original):
                raise ValueError(f"잘못된 셀 범위: {cell_range}")
            bounds = _expand_merge_range(original, ws.merged_cells.ranges, merge_range)
            if bounds != original and merge_range == "expand":
                print(
                    f"[정보] 병합 셀을 보존하기 위해 범위를 {get_column_letter(bounds[0])}{bounds[1]}:"
                    f"{get_column_letter(bounds[2])}{bounds[3]}로 확장했습니다.", file=sys.stderr,
                )
            min_col, min_row, max_col, max_row = bounds
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        if min_col < 1 or min_row < 1 or max_col < min_col or max_row < min_row:
            raise ValueError("셀 범위가 올바르지 않습니다.")
        if max_col > 16384 or max_row > 1048576:
            raise ValueError("Excel 워크시트의 최대 행/열 범위를 벗어났습니다.")

        merge_map = {}
        for merged in ws.merged_cells.ranges:
            a, b, c, d = range_boundaries(str(merged))
            intersects = not (c < min_col or a > max_col or d < min_row or b > max_row)
            if not intersects:
                continue
            inside = a >= min_col and b >= min_row and c <= max_col and d <= max_row
            visible_top, visible_left = max(b, min_row), max(a, min_col)
            visible_bottom, visible_right = min(d, max_row), min(c, max_col)
            if not inside and merge_range == "clip":
                print(f"[경고] 병합 셀 {merged}를 선택 범위 안쪽으로 잘라냅니다.", file=sys.stderr)
            for r in range(visible_top, visible_bottom + 1):
                for col in range(visible_left, visible_right + 1):
                    merge_map[(r, col)] = {
                        "anchor": (visible_top, visible_left),
                        "source_anchor": (b, a),
                        "is_anchor": (r == visible_top and col == visible_left),
                        "rowspan": visible_bottom - visible_top + 1,
                        "colspan": visible_right - visible_left + 1,
                    }

        lines = ['{{{#!wiki style="word-break: keep-all"'] if wrapper else []
        for r in range(min_row, max_row + 1):
            cells = []
            for c in range(min_col, max_col + 1):
                merge = merge_map.get((r, c))
                if merge and not merge["is_anchor"]:
                    continue
                source_row, source_col = (merge["source_anchor"] if merge else (r, c))
                cell = ws.cell(source_row, source_col)
                formula_cell = wf.cell(source_row, source_col)
                raw = _value_to_text(
                    cell, formula_cell, fx_mode=fx_mode, formula_fallback=formula_fallback,
                    number_mode=number_mode, date_mode=date_mode,
                    strict_format=strict_format, epoch=wb_formula.epoch,
                )
                if hyperlinks == "external" and getattr(cell, "hyperlink", None):
                    target = getattr(cell.hyperlink, "target", None)
                    if target and target.startswith(("http://", "https://", "ftp://")):
                        label = apply_excel_font_style(cell, escape_namuwiki_syntax(raw, apply_options))
                        content = f"[{target} {label}]"
                    else:
                        content = _escape_cell(cell, raw, apply_options)
                else:
                    content = _escape_cell(cell, raw, apply_options)
                attrs = _cell_attrs(
                    cell,
                    merge["colspan"] if merge else 1,
                    merge["rowspan"] if merge else 1,
                )
                cells.append(f"||{attrs} {content} ")
            # Retain rows even when every position is occupied by a prior rowspan.
            lines.append("".join(cells) + "||" if cells else "||")

        if wrapper:
            lines.append("}}}")
        return "\n".join(lines)
    finally:
        if wb_value is not wb_formula:
            wb_value.close()
        wb_formula.close()


def build_parser():
    p = argparse.ArgumentParser(
        prog="xlxstonamu",
        description="Excel(.xlsx) 표를 나무위키 표 문법으로 변환합니다.",
        epilog=("예시:\n  xlxstonamu.py table.xlsx\n  xlxstonamu.py table.xlsx -r A1:D10 -s Sheet2\n"
                "  xlxstonamu.py table.xlsx -a intlink extlink --fx input\n"
                "  xlxstonamu.py table.xlsx -o table.txt --number-format raw\n"
                "  xlxstonamu.py table.xlsx -o table.txt --number-format approx --date-format iso"),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("file", help="입력 Excel 파일 (.xlsx)")
    p.add_argument("-r", "--range", dest="cell_range", help="셀 범위 (예: A1:D5)")
    p.add_argument("-s", "--sheet", help="시트명 (기본: 첫 번째 시트)")
    p.add_argument("-o", "--output", help="출력 텍스트 파일 (기본: stdout)")
    p.add_argument("-a", "--apply", nargs="+", default=[],
                   choices=["intlink", "extlink", "footnote", "macro", "table", "all"],
                   help="이스케이프하지 않고 원형으로 유지할 문법")
    p.add_argument("--fx", choices=["input", "output"], default="output",
                   help="수식 문자열(input) 또는 저장된 결과값(output, 기본값)")
    p.add_argument("--formula-fallback", choices=["formula", "blank", "error"], default="formula",
                   help="수식 캐시가 없을 때 처리 (기본: 수식 문자열)")
    p.add_argument("--merge-range", choices=["expand", "error", "clip"], default="expand",
                   help="범위가 병합 셀을 일부만 포함할 때 (기본: 범위 확장)")
    p.add_argument(
        "--number-format", dest="number_mode",
        choices=["stop", "raw", "approx", "display"], default="stop",
        help=("숫자 셀 처리: stop=기본값(즉시 중단), raw=저장값, "
              "approx=내장 근사 포맷터(수동 선택). display는 approx의 호환 별칭"),
    )
    p.add_argument(
        "--date-format", dest="date_mode", choices=["stop", "iso", "serial"], default="stop",
        help="날짜/시간 셀 처리: stop=기본값(즉시 중단), iso=표준 문자열, serial=Excel 일련번호",
    )
    p.add_argument("--strict-format", action="store_true", help="근사 숫자 포맷터가 서식을 지원하지 않으면 중단")
    p.add_argument("--hyperlinks", choices=["display", "external"], default="display",
                   help="Excel 하이퍼링크를 표시 텍스트로만 출력하거나 외부 링크로 보존")
    p.add_argument("--no-wrapper", action="store_true", help="바깥 #!wiki 블록을 생성하지 않음")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    allowed = [x for x in args.apply if x != "all"]
    if "all" in args.apply:
        allowed = ["intlink", "extlink", "footnote", "macro", "table"]
    try:
        output = excel_to_namuwiki(
            args.file, cell_range=args.cell_range, sheet_name=args.sheet,
            apply_options=allowed, fx_mode=args.fx, merge_range=args.merge_range,
            formula_fallback=args.formula_fallback, number_mode=args.number_mode,
            date_mode=args.date_mode, strict_format=args.strict_format, hyperlinks=args.hyperlinks,
            wrapper=not args.no_wrapper,
        )
        if args.output:
            with open(args.output, "w", encoding="utf-8", newline="\n") as f:
                f.write(output + "\n")
        else:
            print(output)
        return 0
    except Exception as exc:
        print(f"[오류] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
