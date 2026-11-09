# NamuMark ↔ Excel table converter

`xlxstonamu.py` converts an `.xlsx` worksheet/range into NamuMark table syntax.
`namutoxlsx.py` reads the first table block in a NamuMark text file and writes
an `.xlsx` workbook. Both scripts use `namu_common.py`; keep the three Python
files in the same directory.

## Requirements

- Python 3.9 or newer is recommended.
- `openpyxl` 3.1 or newer.

```cmd
python -m pip install openpyxl
```

Debian/Ubuntu system package alternative:

```sh
sudo apt install python3-openpyxl
```

## Excel → NamuMark

```sh
python3 xlxstonamu.py workbook.xlsx
python3 xlxstonamu.py workbook.xlsx -s Sheet2 -r A1:F30
python3 xlxstonamu.py workbook.xlsx -s Sheet2 -r A1:F30 -o table.namu
# 숫자/날짜가 있는 파일은 처리 방식을 명시해야 합니다.
python3 xlxstonamu.py workbook.xlsx -o table.namu --number-format raw --date-format iso
```

Options:

- `-s/--sheet`: worksheet name; defaults to the first worksheet.
- `-r/--range`: A1-style range; defaults to the worksheet's used dimensions.
- `-o/--output`: write UTF-8 text to a file; defaults to stdout.
- `--fx output|input`: output the saved formula result or the formula expression.
  `output` is the default. `openpyxl` does not calculate formulas.
- `--formula-fallback formula|blank|error`: behavior when a formula has no
  cached result. The default emits the formula expression with a warning.
- `--merge-range expand|error|clip`: behavior when the selected range cuts
  through a merged cell. Default `expand` preserves the whole merge.
- `--number-format stop|raw|approx`: default `stop` aborts the entire conversion
  when any numeric cell is encountered. This includes General numbers, currency,
  accounting, and percentages because their displayed form can depend on Excel,
  application, OS, and regional settings. `raw` explicitly emits the stored
  numeric value. `approx` explicitly invokes the retained Python formatter for
  common formats; it is an approximation, not Excel's rendering engine.
  `display` remains as a compatibility alias for `approx` and emits a warning.
- `--date-format stop|iso|serial`: default `stop` aborts when a date/time-formatted
  cell is encountered. `iso` explicitly emits a normalized ISO-like string;
  `serial` emits the underlying Excel day serial. Neither reproduces Excel's
  locale-dependent date/time display.
- `--strict-format` makes the optional approximate numeric formatter stop on
  unsupported patterns. Without it, the formatter warns and uses the raw number
  for that cell only.
- `--hyperlinks display|external`: show a hyperlink's display text only (default)
  or generate an external NamuMark link.
- `-a/--apply`: leave selected user-entered NamuMark syntax unescaped. Choices:
  `intlink`, `extlink`, `footnote`, `macro`, `table`, or `all`.
- `--no-wrapper`: omit the outer `#!wiki style="word-break: keep-all"` block.

By default, literal cell content is escaped so text such as `[[page]]`, `||`,
`[date]`, `'''`, and `{{{` does not unexpectedly become markup. Excel font
bold/italic/underline/strike and RGB/indexed font colors are applied afterward
as deliberate NamuMark formatting. Solid RGB/indexed fills, horizontal and
vertical alignment, and merged ranges are also exported.

## NamuMark → Excel

```sh
python3 namutoxlsx.py table.namu output.xlsx
python3 namutoxlsx.py table.namu output.xlsx --sheet 시간표
cat table.namu | python3 namutoxlsx.py - output.xlsx
```

```cmd
type table.namu | python3 namutoxlsx.py - output.xlsx
```


Options:

- `--sheet`: output worksheet name (default `Sheet1`).
- `--strict`: stop instead of recovering from malformed table rows.
- `--max-columns`: guard against accidental oversized spans (default Excel's
  maximum of 16,384 columns).
- `--no-caption`: ignore a `|caption|` table caption.

The parser recognizes ordinary `||` rows, repeated-pipe implicit colspans,
`<-N>` colspans, `<|N>` rowspans, common table/row/column/cell colors, text
colors, `<thead>`, alignment, cell width/height, table border color, and common
inline formatting. External links become Excel hyperlinks. Images and footnotes
are not embedded; file links are omitted from the cell text. `#!folding` labels
and bodies are flattened into text.

## Conversion model and limitations

This is a **table-oriented converter**, not a complete the seed renderer or a
lossless XLSX serialization format.

1. The reverse parser extracts the first top-level table. A table may be inside
   an outer `#!wiki`, `#!folding`, or `#!if` wrapper. Literal/code blocks are
   skipped. Nested tables inside a table cell are not flattened into the
   parent grid; extract and convert them separately.
2. Mixed inline formatting within one cell is flattened to text and represented
   as cell-wide Excel styling. A run-by-run rich-text model is not produced.
3. `#!if` expressions are not evaluated. Their body text is extracted and a
   warning is emitted. `#!syntax`/`#!latex` bodies are treated as text, not run
   or rendered. `#!html` is reduced to visible text with a lightweight tag
   removal pass.
4. Include/video macros, images, and footnotes are not executed or embedded.
   Dynamic macros such as `[date]` remain literal because there is no live wiki
   renderer available.
5. Excel formulas are not recalculated. In `--fx output` mode, only a value
   previously cached by Excel/LibreOffice can be read. Save/recalculate the
   workbook in a spreadsheet application to refresh cached values.
6. Locale- or application-dependent Excel values are guarded by default:
   conversion stops at the first numeric, currency, percentage, date, or time
   cell. The caller must explicitly choose `--number-format raw|approx` and/or
   `--date-format iso|serial`. The retained numeric formatter handles only a
   subset of common patterns and is not a reproduction of Excel's renderer.
   Conditional, elapsed-time, locale-specific, and exotic formats remain
   unsupported by that optional approximation.
7. RGB and indexed colors are supported. Theme/tint colors are not resolved;
   the Excel → NamuMark direction writes the same explicit color for light and
   dark themes. Conditional formatting, rich text, individual border topology,
   charts, images, print settings, and hidden rows/columns are not converted.
8. Column widths and row heights in NamuMark tables are approximated in Excel
   units. Browser CSS pixels and Excel character-width units are not identical.
9. Round-tripping is semantic, not byte-for-byte: the result aims to preserve
   visible cell text, basic styling, and merge geometry, not workbook internals.

## Tests

Run the built-in regression suite from this directory:

```sh
python -m unittest discover -s tests -v
```

The tests cover parsing of pipe runs and attributes, escaped delimiters,
wrapper/code-block table extraction, font/color/link conversion, number formats,
merged-cell geometry, formula cache behavior, and a representative round trip.
