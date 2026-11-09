#!/usr/bin/env python3
"""Shared helpers for table-focused NamuMark <-> XLSX conversion.

This is not a complete implementation of the the seed parser. It handles the
common table grammar conservatively and preserves unsupported constructs as
plain text where possible.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import html
import re
from typing import Optional


class NamuSyntaxError(ValueError):
    pass


FORMAT_TOKENS = ("'''", "''", "__", "~~", "--", "^^", ",,")
MACRO_NAMES = {
    "age", "anchor", "dday", "include", "youtube", "nicovideo",
    "kakaotv", "pagecount", "navertv", "vimeo", "ruby", "br",
    "clearfix", "date", "datetime", "footnote", "tableofcontents",
    "각주", "목차",
}
CSS_BASIC_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000",
    "green": "008000", "blue": "0000FF", "yellow": "FFFF00",
    "gray": "808080", "grey": "808080", "silver": "C0C0C0",
    "maroon": "800000", "purple": "800080", "fuchsia": "FF00FF",
    "lime": "00FF00", "olive": "808000", "navy": "000080",
    "teal": "008080", "aqua": "00FFFF", "orange": "FFA500",
    "transparent": None,
}


def normalize_hex_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip().lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{3}", value):
        value = "".join(ch * 2 for ch in value)
    if re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return value.upper()
    return CSS_BASIC_COLORS.get(value.lower())


def color_to_hex(color) -> Optional[str]:
    """Resolve openpyxl RGB/indexed colors; theme/tint colors are not guessed."""
    if color is None:
        return None
    typ = getattr(color, "type", None)
    if typ == "rgb":
        raw = str(getattr(color, "rgb", ""))
        if re.fullmatch(r"[0-9A-Fa-f]{8}", raw):
            return raw[-6:].upper()
        if re.fullmatch(r"[0-9A-Fa-f]{6}", raw):
            return raw.upper()
    if typ == "indexed":
        try:
            from openpyxl.styles.colors import COLOR_INDEX
            return str(COLOR_INDEX[int(color.indexed)])[-6:].upper()
        except Exception:
            return None
    return None


def _escape_token(token: str) -> str:
    return "".join("\\" + ch for ch in token)


def _balanced_end(text: str, start: int, opening: str, closing: str) -> Optional[int]:
    depth = 0
    i = start
    while i < len(text):
        if text[i] == "\\":
            i += 2
        elif text.startswith(opening, i):
            depth += 1
            i += len(opening)
        elif text.startswith(closing, i):
            depth -= 1
            i += len(closing)
            if depth == 0:
                return i
        else:
            i += 1
    return None


def _macro_end(text: str, start: int) -> Optional[int]:
    m = re.match(r"\[([A-Za-z가-힣]+)", text[start:])
    if not m or m.group(1).lower() not in MACRO_NAMES:
        return None
    i = start + len(m.group(0))
    depth = 0
    while i < len(text):
        if text[i] == "\\":
            i += 2
            continue
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth = max(0, depth - 1)
        elif text[i] == "]" and depth == 0:
            return i + 1
        i += 1
    return None


def escape_namuwiki_syntax(value, apply_options=None) -> str:
    """Escape control markup in literal spreadsheet data.

    `apply_options` may contain intlink/extlink/footnote/macro/table or all.
    Formatting tokens are always neutralized; Excel font styles are wrapped
    after this function so the generated wrappers remain active.
    """
    if value is None:
        return ""
    src = str(value)
    allowed = set(apply_options or ())
    if "all" in allowed:
        allowed.update({"intlink", "extlink", "footnote", "macro", "table"})

    out: list[str] = []
    i = 0
    line_start = True
    while i < len(src):
        if src[i] == "\\":
            out.append("\\\\")
            i += 1
            line_start = False
            continue
        token = next((t for t in FORMAT_TOKENS if src.startswith(t, i)), None)
        if token:
            out.append(_escape_token(token))
            i += len(token)
            line_start = False
            continue
        if src.startswith(("{{{", "}}}"), i):
            out.append(_escape_token(src[i:i + 3]))
            i += 3
            line_start = False
            continue
        if src.startswith("[[", i):
            end = _balanced_end(src, i, "[[", "]]" )
            if "intlink" in allowed and end is not None:
                out.append(src[i:end]); i = end
            else:
                out.append("\\[\\["); i += 2
            line_start = False
            continue
        if src.startswith(("[http://", "[https://", "[ftp://"), i):
            end = src.find("]", i + 1)
            if "extlink" in allowed and end >= 0:
                out.append(src[i:end + 1]); i = end + 1
            else:
                out.append("\\["); i += 1
            line_start = False
            continue
        if src.startswith("[*", i):
            end = src.find("]", i + 2)
            if "footnote" in allowed and end >= 0:
                out.append(src[i:end + 1]); i = end + 1
            else:
                out.append("\\["); i += 1
            line_start = False
            continue
        if src[i] == "[":
            end = _macro_end(src, i)
            if "macro" in allowed and end is not None:
                out.append(src[i:end]); i = end
            else:
                out.append("\\["); i += 1
            line_start = False
            continue
        if src.startswith("||", i):
            out.append("||" if "table" in allowed else r"\|\|")
            i += 2
            line_start = False
            continue
        if line_start:
            if src.startswith(("#redirect", "#넘겨주기"), i):
                token = "#redirect" if src.startswith("#redirect", i) else "#넘겨주기"
                out.append("\\" + token[0] + token[1:]); i += len(token); line_start = False
                continue
            block = next((t for t in ("##", "----", "==") if src.startswith(t, i)), None)
            if block:
                out.append(_escape_token(block)); i += len(block); line_start = False
                continue
            list_match = re.match(r"(?:\d+|[a-zA-Z]|i|I)\.\s", src[i:])
            if list_match:
                marker = list_match.group(0).rstrip()
                out.append("\\" + marker); i += len(marker); line_start = False
                continue
            if src[i] in ">*":
                out.append("\\" + src[i]); i += 1; line_start = False
                continue
        ch = src[i]
        out.append(ch)
        i += 1
        line_start = ch in "\r\n" or (line_start and ch in " \t")
    return "".join(out)


def apply_excel_font_style(cell, text: str, include_color: bool = True) -> str:
    """Apply Excel font styles as NamuMark wrappers around already escaped text."""
    if not text or getattr(cell, "font", None) is None:
        return text
    f = cell.font
    if getattr(f, "strike", False):
        text = f"~~{text}~~"
    if getattr(f, "underline", None):
        text = f"__{text}__"
    if getattr(f, "italic", False):
        text = f"''{text}''"
    if getattr(f, "bold", False):
        text = f"'''{text}'''"
    if include_color:
        color = color_to_hex(getattr(f, "color", None))
        if color:
            text = f"{{{{{{#{color},#{color} {text}}}}}}}"
    return text


def _split_format_sections(fmt: str) -> list[str]:
    out, buf = [], []
    quote = False
    bracket = 0
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"' and (i == 0 or fmt[i - 1] != "\\"):
            quote = not quote
        elif not quote and ch == "[":
            bracket += 1
        elif not quote and ch == "]" and bracket:
            bracket -= 1
        if ch == ";" and not quote and bracket == 0:
            out.append("".join(buf)); buf = []
        else:
            buf.append(ch)
        i += 1
    out.append("".join(buf))
    return out


def format_excel_number(value, number_format: str) -> str:
    """Format common Excel numeric formats; unsupported patterns raise ValueError."""
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return str(value)
    fmt = number_format or "General"
    if fmt in ("General", "@", ""):
        return str(value)
    sections = _split_format_sections(fmt)
    if any(re.search(r"\[(?:<=|>=|<>|=|<|>)\s*[-+]?\d+(?:\.\d+)?\]", part) for part in sections):
        raise ValueError(f"조건부 숫자 서식은 지원하지 않습니다: {number_format!r}")
    section = sections[0]
    number = value
    if value < 0 and len(sections) > 1:
        section, number = sections[1], abs(value)
    elif value == 0 and len(sections) > 2:
        section, number = sections[2], 0
    section = re.sub(r"\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s*\d+)\]", "", section, flags=re.I)
    if section == "":
        return ""
    if re.search(r"\[[^\]]+\]", section):
        raise ValueError(f"지원하지 않는 숫자 서식 지시자: {number_format!r}")
    section = re.sub(r'"([^"]*)"', lambda m: m.group(1), section)
    section = section.replace("\\", "").replace("_", "").replace("*", "")
    percent = section.count("%")
    number *= 100 ** percent
    section = section.replace("%", "")
    m = re.search(r"[0#?][0#?,.]*[0#?]?(?:E[+-]0+)?", section, re.I)
    if not m:
        raise ValueError(f"지원하지 않는 숫자 서식: {number_format!r}")
    pattern = m.group(0)
    prefix, suffix = section[:m.start()], section[m.end():]
    scale_match = re.search(r",+$", pattern)
    if scale_match:
        scale_count = len(scale_match.group(0))
        number /= 1000 ** scale_count
        pattern = pattern[:-scale_count]
    scientific = "E" in pattern.upper()
    grouping = "," in pattern.split(".", 1)[0]
    if scientific:
        mantissa, exponent = re.split("[Ee]", pattern, maxsplit=1)
        decimals = len(mantissa.split(".", 1)[1]) if "." in mantissa else 0
        exp_digits = len(exponent.lstrip("+-"))
        rendered = f"{number:.{decimals}E}"
        mant, exp = rendered.split("E")
        rendered = f"{mant}E{'+' if int(exp) >= 0 else '-'}{abs(int(exp)):0{exp_digits}d}"
    else:
        decimals = len(pattern.split(".", 1)[1].split("E", 1)[0]) if "." in pattern else 0
        min_dec = sum(x == "0" for x in (pattern.split(".", 1)[1] if "." in pattern else ""))
        int_pattern = pattern.split(".", 1)[0].split("E", 1)[0].replace(",", "")
        min_int = int_pattern.count("0")
        rendered = f"{number:,.{decimals}f}" if grouping else f"{number:.{decimals}f}"
        if decimals > min_dec and "." in pattern and "#" in pattern.split(".", 1)[1]:
            rendered = rendered.rstrip("0").rstrip(".")
        int_part, dot, dec_part = rendered.partition(".")
        sign = "-" if int_part.startswith("-") else ""
        bare_int = int_part.lstrip("-")
        if len(bare_int.replace(",", "")) < min_int:
            bare_int = bare_int.zfill(min_int)
        rendered = sign + bare_int + (dot + dec_part if dot else "")
    return prefix + rendered + suffix + ("%" * percent)


@dataclass
class ParsedCell:
    attrs: str
    content: str
    implicit_colspan: int = 1
    colspan: int = 1
    rowspan: int = 1
    options: dict = field(default_factory=dict)


@dataclass
class TableData:
    rows: list[list[ParsedCell]]
    table_attrs: dict = field(default_factory=dict)
    caption: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass
class TextData:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: Optional[str] = None
    hyperlink: Optional[str] = None


def _find_pipe_runs(line: str) -> list[tuple[int, int]]:
    runs = []
    i = 0
    braces = links = brackets = 0
    while i < len(line):
        if line[i] == "\\":
            i += 2; continue
        # Table attributes such as <|2>, <^|1>, and CSS-like values may
        # contain a literal pipe; it is not a cell separator.
        if line[i] == "<" and i + 1 < len(line) and (line[i + 1].isalpha() or line[i + 1] in "-|^v:)("):
            tag_end = line.find(">", i + 1)
            if tag_end >= 0:
                i = tag_end + 1
                continue
        if line.startswith("{{{", i):
            braces += 1; i += 3; continue
        if line.startswith("}}}", i) and braces:
            braces -= 1; i += 3; continue
        if braces:
            i += 1; continue
        if line.startswith("[[", i):
            links += 1; i += 2; continue
        if line.startswith("]]", i) and links:
            links -= 1; i += 2; continue
        if links:
            i += 1; continue
        if line[i] == "[": brackets += 1
        elif line[i] == "]" and brackets: brackets -= 1
        if line[i] == "|" and not brackets:
            j = i
            while j < len(line) and line[j] == "|": j += 1
            runs.append((i, j - i)); i = j
        else:
            i += 1
    return runs


def split_cell_attributes(content: str) -> tuple[str, str]:
    content = content.lstrip()
    i, attrs = 0, []
    while i < len(content) and content[i] == "<":
        end = content.find(">", i + 1)
        if end < 0: break
        attrs.append(content[i:end + 1]); i = end + 1
    return "".join(attrs), content[i:].strip()


def parse_cell_options(attr_str: str) -> dict:
    result = {}
    for tag in re.findall(r"<([^<>]*)>", attr_str):
        tag = tag.strip()
        if tag in {"thead", "sortable", "nopad", "keepall", "rowkeepall", "colkeepall"}:
            result[tag] = True
        elif "=" in tag:
            key, val = tag.split("=", 1)
            key = key.strip().lower()
            if key.startswith("table "):
                key = "table" + key[6:].replace(" ", "")
            result[key] = val.strip()
        elif tag.startswith("color#"):
            result["color"] = "#" + tag[6:]
        elif tag in {"(", ":", ")"}:
            result["align"] = {"(": "left", ":": "center", ")": "right"}[tag]
        elif tag.startswith(("<-", "<|", "^|", "v|")):
            result[tag] = True
    for marker, align in (("<(>", "left"), ("<:>", "center"), ("<)>", "right")):
        if marker in attr_str: result["align"] = align
    if re.search(r"<\^\|\d+>", attr_str): result["valign"] = "top"
    elif re.search(r"<v\|\d+>", attr_str): result["valign"] = "bottom"
    elif re.search(r"<\|\d+>", attr_str): result["valign"] = "center"
    for key in ("bgcolor", "colbgcolor", "rowbgcolor", "tablebgcolor", "color", "colcolor", "rowcolor", "tablecolor", "tablebordercolor"):
        if key in result:
            result[key] = normalize_hex_color(str(result[key]).split(",", 1)[0])
    return result


def parse_row(line: str, strict: bool = False) -> list[ParsedCell]:
    line = line.strip()
    if not line.startswith("||"): return []
    if line == "||": return []
    runs = _find_pipe_runs(line)
    if not runs or runs[0][0] != 0: return []
    if any(count % 2 for _, count in runs):
        raise NamuSyntaxError(f"행 구분자는 짝수 개 파이프로 구성되어야 합니다: {line!r}")
    cells = []
    for idx in range(len(runs) - 1):
        pos, count = runs[idx]
        next_pos = runs[idx + 1][0]
        attrs, body = split_cell_attributes(line[pos + count:next_pos])
        cells.append(ParsedCell(attrs, body, implicit_colspan=max(1, count // 2)))
    last_pos, last_count = runs[-1]
    after = line[last_pos + last_count:]
    if after.strip():
        if strict: raise NamuSyntaxError(f"행 마지막에 닫는 ||가 없습니다: {line!r}")
        attrs, body = split_cell_attributes(after)
        cells.append(ParsedCell(attrs, body, implicit_colspan=max(1, last_count // 2)))
    for cell in cells:
        mcol = re.search(r"<-(\d+)", cell.attrs)
        mrow = re.search(r"<(?:\^|v)?\|(\d+)", cell.attrs)
        cell.colspan = int(mcol.group(1)) if mcol else cell.implicit_colspan
        cell.rowspan = int(mrow.group(1)) if mrow else 1
        if cell.colspan < 1 or cell.rowspan < 1:
            raise NamuSyntaxError("colspan/rowspan은 1 이상이어야 합니다.")
        cell.options = parse_cell_options(cell.attrs)
    return cells


def _update_brace_stack(line: str, stack: list[str]) -> None:
    i = 0
    while i < len(line):
        if line[i] == "\\": i += 2; continue
        if line.startswith("{{{", i):
            tail = line[i + 3:]
            m = re.match(r"#!([A-Za-z]+)", tail)
            mode = m.group(1).lower() if m else "literal"
            # A wiki code-class wrapper is a literal sample, not a table host.
            if mode == "wiki" and re.search(r'\bclass\s*=\s*["\']code["\']', tail):
                mode = "literal"
            stack.append(mode); i += 3; continue
        if line.startswith("}}}", i) and stack:
            stack.pop(); i += 3; continue
        i += 1


def extract_table_lines(text: str) -> list[str]:
    """Extract the first top-level table; allow outer wiki/folding/if wrappers.

    Tables nested inside a cell are not flattened into their parent's grid.
    """
    stack: list[str] = []
    rows: list[str] = []
    base_depth = None
    fenced = False
    for raw in text.splitlines():
        s = raw.strip()
        if s.startswith("##"):
            continue
        if s.startswith(("```", "~~~")):
            fenced = not fenced; continue
        if fenced: continue
        blocked = any(mode not in {"wiki", "folding", "if"} for mode in stack)
        depth = len(stack)
        is_row = s.startswith("||") and not blocked
        if base_depth is None and is_row:
            base_depth = depth
        if base_depth is not None:
            if is_row and depth == base_depth:
                rows.append(s)
            elif not s:
                pass
            elif s.startswith("}}}") and depth <= base_depth:
                if rows: break
            elif s.startswith("{{{#!") and depth <= base_depth:
                pass
            elif depth > base_depth:
                pass  # nested wiki block within a parent cell
            elif rows:
                break
        _update_brace_stack(raw, stack)
    return rows


def parse_table(text: str, strict: bool = False) -> TableData:
    lines = extract_table_lines(text)
    rows, notes = [], []
    for n, line in enumerate(lines, 1):
        try:
            parsed = parse_row(line, strict=strict)
            if any("{{{#!if" in cell.content for cell in parsed):
                notes.append(f"행 {n}: #!if 조건식은 평가하지 않고 본문 텍스트를 추출합니다.")
            # Keep a bare `||` as an empty row so rowspan geometry is retained.
            rows.append(parsed)
        except NamuSyntaxError as exc:
            if strict: raise
            notes.append(f"행 {n}: {exc}")
    table_attrs = {}
    if rows:
        opts = rows[0][0].options
        for key in ("tablewidth", "tablealign", "tablebgcolor", "tablecolor", "tablebordercolor"):
            if key in opts: table_attrs[key] = opts[key]
    caption_match = re.search(r"(?m)^\s*\|([^|\n]+)\|\s*(?=\|\|)", text)
    caption = caption_match.group(1).strip() if caption_match else ""
    return TableData(rows, table_attrs, caption, notes)


def _split_link(inner: str) -> tuple[str, str]:
    i = nested = 0
    while i < len(inner):
        if inner[i] == "\\": i += 2; continue
        if inner.startswith("[[", i): nested += 1; i += 2; continue
        if inner.startswith("]]", i) and nested: nested -= 1; i += 2; continue
        if inner[i] == "|" and nested == 0: return inner[:i], inner[i + 1:]
        i += 1
    return inner, inner


def _css_style(header: str) -> dict:
    m = re.search(r'\bstyle\s*=\s*"([^"]*)"', header)
    if not m: return {}
    out = {}
    for item in m.group(1).split(";"):
        if ":" not in item: continue
        k, v = (x.strip().lower() for x in item.split(":", 1))
        if k == "color": out["color"] = normalize_hex_color(v)
        elif k == "background-color": out["bgcolor"] = normalize_hex_color(v)
        elif k == "font-weight" and v in {"bold", "700", "800", "900"}: out["bold"] = True
        elif k == "font-style" and v == "italic": out["italic"] = True
        elif k == "text-decoration":
            out["underline"] = "underline" in v
            out["strike"] = "line-through" in v
        elif k == "text-align": out["align"] = v
    return out


def _merge_style(dest: TextData, src: dict) -> None:
    for key in ("bold", "italic", "underline", "strike"):
        if src.get(key): setattr(dest, key, True)
    if src.get("color"): dest.color = src["color"]


def strip_markup_details(source: str) -> TextData:
    """Flatten NamuMark to visible cell text and infer cell-wide font styling."""
    result = TextData("")

    def walk(s: str, inherited=None) -> str:
        style = dict(inherited or {})
        out = []
        i = 0
        while i < len(s):
            if s[i] == "\\":
                if i + 1 < len(s): out.append(s[i + 1]); i += 2
                else: out.append("\\"); i += 1
                continue
            if s.startswith("{{{", i):
                end = _balanced_end(s, i, "{{{", "}}}")
                if end:
                    inner = s[i + 3:end - 3]
                    header, sep, body = inner.partition("\n")
                    local = dict(style)
                    if header.startswith("#!"):
                        feature = header[2:].split(None, 1)[0].lower() if header[2:].strip() else ""
                        payload = body if sep else ""
                        if feature == "wiki":
                            local.update(_css_style(header))
                            out.append(walk(payload, local))
                        elif feature == "folding":
                            label = re.search(r"\[([^\]]*)\]", header)
                            if label:
                                out.append(label.group(1))
                                if payload: out.append("\n")
                            out.append(walk(payload, local))
                        elif feature == "style":
                            pass  # CSS definition block has no visible text.
                        elif feature == "html":
                            html_source = payload if sep else re.sub(r"^#!html\s*", "", header, flags=re.I)
                            html_source = re.sub(r"<br\s*/?>", "\n", html_source, flags=re.I)
                            out.append(html.unescape(re.sub(r"<[^>]*>", "", html_source)))
                        elif feature in {"syntax", "latex"}:
                            # Code/TeX is displayed as source, not reparsed.
                            out.append(payload)
                        elif feature == "if":
                            # The condition is intentionally not evaluated.
                            out.append(walk(payload, local))
                        else:
                            out.append(walk(payload, local))
                    else:
                        cm = re.match(r"#([#\w,.-]+)\s+(.*)", inner, re.S)
                        sm = re.match(r"[+-]\d\s+(.*)", inner, re.S)
                        if cm:
                            local["color"] = normalize_hex_color(cm.group(1).split(",", 1)[0])
                            out.append(walk(cm.group(2), local))
                        elif sm:
                            out.append(walk(sm.group(1), local))
                        else:
                            # Plain {{{...}}} is literal/monospace text.
                            out.append(inner)
                    _merge_style(result, local)
                    i = end; continue
            if s.startswith("[[", i):
                end = _balanced_end(s, i, "[[", "]]" )
                if end:
                    target, label = _split_link(s[i + 2:end - 2])
                    target = target.strip()
                    if not target.lower().startswith(("파일:", "file:", "분류:", "category:")):
                        if target.startswith(("http://", "https://", "ftp://")):
                            result.hyperlink = target
                        out.append(walk(label, style))
                    i = end; continue
            if s.startswith(("[http://", "[https://", "[ftp://"), i):
                end = s.find("]", i + 1)
                if end >= 0:
                    parts = s[i + 1:end].split(None, 1)
                    result.hyperlink = parts[0]
                    out.append(walk(parts[1] if len(parts) > 1 else parts[0], style))
                    i = end + 1; continue
            if s.startswith("[*", i):
                end = s.find("]", i + 2)
                if end >= 0: i = end + 1; continue
            if s.startswith("[br]", i):
                out.append("\n"); i += 4; continue
            if s.startswith("[", i):
                m = re.match(r"\[(include|youtube|nicovideo|kakaotv|navertv|vimeo|anchor|ruby)\(", s[i:], re.I)
                if m:
                    macro_name = m.group(1).lower()
                    args_start = i + len(m.group(0))
                    j = args_start; depth = 1; first_comma = None
                    while j < len(s):
                        if s[j] == "\\": j += 2; continue
                        if s[j] == "(": depth += 1
                        elif s[j] == ")":
                            depth -= 1
                            if depth == 0:
                                if macro_name == "ruby":
                                    first_arg = s[args_start:(first_comma if first_comma is not None else j)].strip()
                                    out.append(walk(first_arg, style))
                                j += 2 if j + 1 < len(s) and s[j + 1] == "]" else 1
                                break
                        elif s[j] == "," and depth == 1 and first_comma is None:
                            first_comma = j
                        j += 1
                    i = j; continue
            pair = next((p for p in FORMAT_TOKENS if s.startswith(p, i)), None)
            if pair:
                end = s.find(pair, i + len(pair))
                if end >= 0:
                    key = {"'''": "bold", "''": "italic", "__": "underline", "~~": "strike", "--": "strike"}.get(pair)
                    if key:
                        style[key] = True
                        _merge_style(result, style)
                    out.append(walk(s[i + len(pair):end], style))
                    i = end + len(pair); continue
            out.append(s[i]); i += 1
        return "".join(out)

    result.text = walk(source).strip()
    return result


def strip_markup(source: str) -> str:
    return strip_markup_details(source).text
