# fastxlsxio

**A lightweight, high-performance Python library for fast XLSX I/O, built with Rust.**
*The project aims to provide xlsx writing with the same functionality as Python's `xlsxwriter`, only faster, while keeping the reading implementation as it came from the `fastxlsx` project.*

- **Writing** — a thin Python layer over [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) 0.99.
  Constant-memory mode, per-column type and format caching, GIL released while rows are flushed.
- **Reading** — credited to `fastxlsx`, no work of its own here (see [Reading](#reading)). Built on
  [`calamine`](https://github.com/tafia/calamine); reads rectangular ranges straight into typed values / numpy
  arrays, and can read many workbooks in parallel.
- **Typed errors** — every error raised while writing is an `XIOError` subclass that also inherits the matching builtin
  (`ValueError`, `OverflowError`, ...), and carries the cell address when there is one.

> **Status:** fastxlsxio exposes only a small slice of what `rust_xlsxwriter` can do (values, `num_format`, bold,
> merged text cells, column widths). See [Not implemented yet](#not-implemented-yet) for the full list.

## Installation

Python 3.10 – 3.14. CI builds wheels for Linux (glibc and musl: x86_64, x86, aarch64), Windows (x64, x86) and
macOS (x86_64, arm64).

### PyPI Install

```bash
pip install fastxlsxio
```

### Source Build (Requires Rust Toolchain)

```bash
git clone https://github.com/JamesCrumble/fastxlsxio.git
cd fastxlsxio
pip install .
```

## Quick start

### Write

```python
from datetime import date

from fastxlsxio import ColTypeHint, XIOWWorkbook

wb = XIOWWorkbook("report.xlsx")

# Dates are plain numbers in Excel; without a number format they show up as 45293.
wb.add_format({"num_format": "dd.mm.yyyy"}, bind_to_datatype=ColTypeHint.Date)
money = wb.add_format({"num_format": "#,##0.00"})
header = wb.add_format({"bold": True})

ws = wb.add_worksheet("Report")
ws.write_row(0, 0, ["id", "name", "price", "created"], formats=[header] * 4)

rows = [
    {"id": 1, "name": "apple", "price": 1.5, "created": date(2024, 1, 2)},
    {"id": 2, "name": "pear", "price": 2.25, "created": date(2024, 1, 3)},
]
# formats[i] is applied to the i-th column of every row; dict rows are written by their values
ws.write_rows(1, 0, rows, formats=[None, None, money, None])

ws.set_column_width(1, 20)
ws.merge_range(4, 0, 4, 3, "Total")

wb.save()
```

### Read

```python
from fastxlsxio import DShape, DType, RangeInfo, XIORWorkbook

wb = XIORWorkbook("report.xlsx")
print(wb.sheetnames)  # ['Report']
ws = wb.get("Report")  # by name or by index

ws.read_value(RangeInfo((0, 0), DShape.Row(4), dtype=DType.Str))  # ['id', 'name', 'price', 'created']
ws.read_value(RangeInfo((1, 0), DShape.Column(2), dtype=DType.Int))  # array([1, 2])
ws.cell_value("B2", dtype=DType.Str)  # 'apple'
```

## Writing

All `row` / `col` arguments are **0-based**. Everything below lives in `fastxlsxio.XIOW*`.

| Class | Purpose |
| --- | --- |
| `XIOWWorkbook(filepath=None, options=None)` | Entry point. `add_worksheet`, `add_format`, `get_by_idx`, `get_by_name`, `sheetnames`, `save(path=None)`. |
| `XIOWWorksheet` | Returned by `add_worksheet`. `write_cell`, `write_row`, `write_rows`, `write_column`, `merge_range`, `set_column_width`, `set_column_format`, `set_row_format`. |
| `XIOFormat` | A cell format from `XIOWWorkbook.add_format({...})` (or `XIOFormat.from_properties({...})`). |
| `XIOWOptions` | Write-mode switches, see below. |
| `ColTypeHint` | Cell kind (`Float`, `Int`, `String`, `Bool`, `Blank`, `Date`, `Time`, `DateTime`, `Sequence`); used to bind a default format to a type. |

### Methods

| Method | What it does |
| --- | --- |
| `write_cell(row, col, value, format=None)` | One value. |
| `write_row(row, col, value, formats=None)` | Values left to right. `value` is a list, tuple, dict (its *values*, keys are ignored) or any iterable. |
| `write_rows(row, col, value, formats=None)` | Many rows at once, in chunks of 1024, with the GIL released while a chunk is written. `value` must be a **sequence** of rows (list / tuple), not a generator — batch one with `itertools.batched`. Each row may be a list, tuple, dict or iterable. |
| `write_column(row, col, value)` | Values top to bottom. Only with `constant_memory=False`; no `format` argument yet. |
| `merge_range(r1, c1, r2, c2, value, format=None)` | Merge a range (ends included). The value is always written as text. |
| `set_column_width(col, width)` | Width in characters. |
| `set_column_format(col, format)` / `set_row_format(row, format)` | Default format of a column / row; cells with their own format keep it. |
| `XIOWWorkbook.add_format(props, bind_to_datatype=None)` | Create a format. With `bind_to_datatype=ColTypeHint.X` it also becomes the default for every cell of that type written afterwards (re-binding replaces it). |

Supported format properties: `num_format` (`str`) and `bold` (`True`). Anything else raises `XIOUnsupportedError`.

### Options

```python
from fastxlsxio import XIOWOptions, XIOWWorkbook

options = XIOWOptions(constant_memory=True, cache_typehints_write_optimization=True)
wb = XIOWWorkbook("out.xlsx", options=options)
ws = wb.add_worksheet("fast")                                   # uses the workbook options
ws2 = wb.add_worksheet("random-access", XIOWOptions(constant_memory=False))  # per-sheet override
```

| Option | Default | Effect |
| --- | --- | --- |
| `constant_memory` | `True` | Rows are flushed to a temporary file as you write, so memory stays flat on huge sheets. Write strictly **top to bottom**, and `write_column` is unavailable. |
| `cache_col_formats` | `True` | Applies a format to a column once instead of re-setting it for every cell. |
| `cache_row_formats` | `False` | The same, per row. Cannot be combined with `cache_col_formats` (`XIOArgumentError`). |
| `cache_typehints_write_optimization` | `True` | Remembers the value type last seen in each column so the next cells skip type detection. A column that turns out mixed falls back to the general path automatically. |

Options are readable and settable as attributes (`ws.options`, `wb.options`).

### Python → Excel value mapping

| Python value | Written as |
| --- | --- |
| `float` | number |
| `int` | number if `abs(value) <= 2**53`; larger ints (Excel can't hold them exactly) are written as **text** digits |
| `str` | text |
| `bool` | boolean |
| `None` | blank — a cell is only emitted if a format applies |
| `datetime`, `date`, `time` (and subclasses such as `pandas.Timestamp`) | Excel date-time number. Needs a `num_format` to look like a date, see the quick start. |
| `decimal.Decimal` | number (converted through `float`) |
| `list` / `tuple` / other non-`str` sequence | one text cell, items joined with `", "` — see [gotchas](#behavior-notes-and-gotchas) |
| anything else (`dict`, enums, `UUID`, numpy scalars, ...) | `str(value)` as text |

### Threads and the GIL

`write_rows` converts a chunk of Python objects with the GIL held, then releases it while the chunk is written, and
`save()` releases it for the whole save. A worksheet being written from one thread raises `XIOStateError` if another
thread touches it, and so does `add_worksheet` while any sheet is mid-write.

**Write a worksheet completely before adding the next one.** Adding a worksheet can move the earlier ones in memory,
and their handles then stop working (`XIOStateError`) — including handles returned by `get_by_name` / `get_by_idx`.

## Reading

> **Origin of this part.** The reading implementation is credited to the original
> [fastxlsx](https://github.com/shuangluoxss/fastxlsx) project. fastxlsxio has no work of its own in reading so far,
> and none is planned for now. Writing is different: it was rewritten from scratch here and has nothing in common
> with fastxlsx.

```python
from fastxlsxio import DShape, DType, RangeInfo, XIORWorkbook, read_many
```

| Piece | Notes |
| --- | --- |
| `XIORWorkbook(path)` | `sheetnames`, `worksheets`, `n_sheets`, `path`; `get(idx_or_name)`, `get_by_idx`, `get_by_name`. |
| `XIORWorksheet` | `title`, `n_rows`, `n_cols`; `read_value(range_info)`, `read_values([...] or {...})`, `cell_value(addr, *, dtype, strict)`. |
| `RangeInfo(pos, data_shape=DShape.Scalar(), *, dtype=DType.Any, strict=True)` | `pos` is `(row, col)`, 0-based; negative values count back from the end of the sheet. |
| `DShape` | `Scalar()`, `Row(n_cols)`, `Column(n_rows)`, `Matrix(n_rows, n_cols)`. |
| `DType` | `Int`, `Float`, `Str`, `Bool`, `Date`, `DateTime`, `Any`. With `strict=True` a cell that doesn't match `dtype` raises; with `strict=False` a default value is used. |
| `XIORWorkbook.read_worksheets({...})` | Several sheets (by name or index), each with a list or dict of ranges. |
| `read_many({path: {sheet: [...]}})` | Several workbooks **in parallel** (rayon). Mirrors the input structure. |
| `addr_to_idx("C3")` / `idx_to_addr(2, 2)` | `(2, 2)` ↔ `"C3"`. |

`Int` and `Float` ranges come back as numpy arrays; `Str`, `Date`, `DateTime` and `Any` ranges as plain (nested) lists.
Numbers are stored as floats in XLSX, so `DType.Any` returns `1.0` for a cell holding 1 — ask for `DType.Int` when
you want ints.

```python
wb = XIORWorkbook("report.xlsx")
wb.read_worksheets({
    "Report": {
        "header": RangeInfo((0, 0), DShape.Row(4), dtype=DType.Str),
        "price": RangeInfo((1, 2), DShape.Column(2), dtype=DType.Float),
    },
})
# {'Report': {'header': ['id', 'name', 'price', 'created'], 'price': array([1.5 , 2.25])}}

read_many({"a.xlsx": {0: [RangeInfo((0, 0), dtype=DType.Str)]}})
```

## Errors

Every error fastxlsxio raises while writing is a `fastxlsxio.XIOError`. Subclasses also inherit the builtin they replaced, so existing
`except ValueError:` handlers keep working.

| Exception | Also a | Raised when |
| --- | --- | --- |
| `XIOConversionError` | `ValueError` | A value can't become a cell. `.address` (e.g. `"B3"`), `.row`, `.col`, `.value`, `.reason`. |
| `XIOOverflowError` | `XIOConversionError`, `OverflowError` | A number or row/column index doesn't fit. |
| `XIOArgumentError` | `ValueError` | Bad argument, e.g. both format caches on, no path for `save()`. |
| `XIONotFoundError` | `XIOArgumentError` | No worksheet at that index / with that name. |
| `XIOIterableError` | `XIOArgumentError` | A row / column you passed isn't iterable. |
| `XIOUnsupportedError` | `NotImplementedError` | Unsupported format property, or `write_column` in constant-memory mode. |
| `XIOStateError` | `RuntimeError` | Worksheet is busy in another thread, or its handle is stale. |
| `XIOXlsxError` | `ValueError` | The backend refused: Excel limits, bad sheet name, bad merge range, date outside 1900–9999. |
| `XIOFileError` | `FileExistsError` | The file can't be saved, or the workbook is invalid at save time (e.g. duplicate sheet names). |

Bad `row` / `col` arguments (negative, too large, wrong type) raise a plain `OverflowError` / `TypeError`.
Reading still raises builtin exceptions (opening a missing or corrupt file raises `FileExistsError`).

## Behavior notes and gotchas

Things that work as designed today but may surprise you:

- **Constant-memory mode silently drops writes to an earlier row.** Once row `n` is written, writes to rows `< n` are
  ignored by the backend — no exception, only a message on stderr. It is the default mode, so write in row order.
- **Dates need a `num_format`.** Without one (for instance via `bind_to_datatype`), a date cell shows Excel's serial number.
- **Dates are truncated:** `tzinfo` is ignored, sub-second precision is dropped, and the year must be within 1900–9999
  (`XIOXlsxError` otherwise).
- **numpy scalars, enums and `bytes` are written as text.** `np.float64(1.5)`, `np.int64(3)`, an `IntEnum` member
  fail the exact-type checks and go through `str()`. Convert with `.item()` / `.value` first. `bytes` is a sequence,
  so `b"xy"` becomes `120, 121`.
- **Sequences are flattened to one text cell.** Items are joined with `", "`, `None` becomes an empty item, floats are
  rounded to 3 decimals, and `bool` items are rendered as `Да` / `Нет` (hard-coded). A sequence made only of
  `datetime`s is joined with `"; "` as `dd.mm.yyyy, hh:mm:ss`.
- **`merge_range` always writes text** (`str(value)`), and with no format it uses Excel's default one.
- **`bold` only accepts `True`.** `{"bold": False}` raises `XIOUnsupportedError`; simply leave the key out.
- **`write_rows` takes sequences, not iterators.** A generator or iterator raises `TypeError`.
- **Ints beyond `±2**53` become text**, so they survive intact but aren't numeric in Excel.

## Not implemented yet

fastxlsxio wraps [`rust_xlsxwriter` 0.99.0](https://crates.io/crates/rust_xlsxwriter/0.99.0) (crate features enabled:
`chrono`, `zlib`, `ryu`, `constant_memory`). The tables below list what that crate offers and fastxlsxio does **not**
expose to Python yet. The status column reads Yes, Partial or No.

### Overview

| Area | Status |
| --- | --- |
| Write numbers, text, booleans, blanks, dates / times | Yes |
| `num_format` and `bold` | Yes |
| Merged ranges | Partial (text only) |
| Column width, column / row default format | Partial (width in characters only) |
| Constant-memory mode | Yes |
| Formulas, hyperlinks, rich strings, checkboxes | No |
| Full cell formatting (fonts, colors, borders, alignment, ...) | No |
| Row height, hidden rows / columns, autofit, outline groups | No |
| Freeze panes, zoom, tab color, sheet visibility, sheet view modes | No |
| Print / page setup, headers and footers | No |
| Tables, autofilters | No |
| Conditional formatting, data validation | No |
| Charts, sparklines, chartsheets | No |
| Images, shapes / textboxes, buttons, cell notes | No |
| Worksheet / range protection | No |
| Document properties, defined names, themes | No |
| Save to an in-memory buffer / writer | No (`save()` takes a path only) |
| VBA macros (`.xlsm`) | No |
| Modify / append to an existing file | No (not possible in `rust_xlsxwriter` either) |

### Cell data

| `rust_xlsxwriter` API | Notes |
| --- | --- |
| `write_formula`, `write_formula_with_format`, `set_formula_result`, `set_formula_result_default` | Formulas with cached results |
| `write_array_formula*`, `write_dynamic_formula*`, `write_dynamic_array_formula*` | Array and Excel 365 dynamic-array formulas |
| `write_url`, `write_url_with_text`, `write_url_with_format`, `write_url_with_options`, `Url` | Hyperlinks |
| `write_rich_string`, `write_rich_string_with_format` | Multi-format strings |
| `insert_checkbox`, `insert_checkbox_with_format` | Checkbox cells |
| `embed_image`, `embed_image_with_format` | Images embedded in a cell |
| `clear_cell`, `clear_cell_format` | Remove a written cell / its format |
| `set_nan_value`, `set_infinity_value`, `set_neg_infinity_value` | Custom text for NaN / ±inf (Excel can't store them, so `rust_xlsxwriter` writes default strings such as `"NAN"`) |
| `merge_range` with non-text values | Only text is merged today |
| `write_column` with a format | Raises `XIOArgumentError` |

### Cell formatting (`Format`)

Only `set_num_format` and `set_bold` are wired up. Missing:

| Group | `rust_xlsxwriter` API |
| --- | --- |
| Font | `set_italic`, `set_underline`, `set_font_strikethrough`, `set_font_name`, `set_font_size`, `set_font_color`, `set_font_family`, `set_font_charset`, `set_font_scheme`, `set_font_script` |
| Alignment | `set_align` (horizontal / vertical), `set_indent`, `set_rotation`, `set_text_wrap`, `set_shrink`, `set_reading_direction` |
| Fill | `set_background_color`, `set_foreground_color`, `set_pattern` |
| Borders | `set_border`, `set_border_top/bottom/left/right` (+ `_color`), `set_border_color`, `set_border_diagonal`, `set_border_diagonal_color`, `set_border_diagonal_type` |
| Protection | `set_locked`, `set_unlocked`, `set_hidden` |
| Other | `set_num_format_index` (built-in format indices), `set_quote_prefix`, `set_hyperlink`, `set_checkbox`, `Format::merge`, the `unset_*` family |
| Applying formats | `set_cell_format`, `set_range_format`, `set_range_format_with_border`, `set_column_range_format`, `set_default_format` (worksheet and workbook) |

### Rows, columns and layout

| Group | `rust_xlsxwriter` API |
| --- | --- |
| Column width | `set_column_width_pixels`, `set_column_range_width`, `set_column_range_width_pixels`, `set_column_autofit_width` |
| Autofit | `autofit`, `autofit_to_max_width`, `set_autofit_max_width`, `set_autofit_max_row` |
| Row height | `set_row_height`, `set_row_height_pixels`, `set_default_row_height`, `set_default_row_height_pixels` |
| Hide / show | `set_column_hidden`, `set_column_range_hidden`, `set_row_hidden`, `set_row_unhidden`, `hide_unused_rows` |
| Outline groups | `group_rows`, `group_rows_collapsed`, `group_columns`, `group_columns_collapsed`, `group_symbols_above`, `group_symbols_to_left` |

### Sheet view and behavior

| `rust_xlsxwriter` API | Notes |
| --- | --- |
| `set_freeze_panes`, `set_freeze_panes_top_cell` | Frozen rows / columns |
| `set_selection`, `set_top_left_cell`, `set_active`, `set_selected`, `set_first_tab` | Selection and which sheet / tab is shown |
| `set_zoom`, `set_screen_gridlines`, `set_right_to_left` | View options |
| `set_view_normal`, `set_view_page_layout`, `set_view_page_break_preview` | View modes |
| `set_tab_color`, `set_hidden`, `set_very_hidden` | Tab color and sheet visibility |
| `ignore_error`, `ignore_error_range` | Suppress Excel's green-triangle warnings |

### Print and page setup

`set_landscape`, `set_portrait`, `set_paper_size`, `set_margins`, `set_print_area`, `set_print_scale`,
`set_print_fit_to_pages`, `set_repeat_rows`, `set_repeat_columns`, `set_print_gridlines`, `set_print_headings`,
`set_print_center_horizontally`, `set_print_center_vertically`, `set_print_black_and_white`, `set_print_draft`,
`set_print_first_page_number`, `set_page_order`, `set_page_breaks`, `set_vertical_page_breaks`,
`set_header`, `set_footer`, `set_header_image`, `set_footer_image`, `set_header_footer_scale_with_doc`,
`set_header_footer_align_with_page`.

### Tables, filters, validation, conditional formatting

| Feature | `rust_xlsxwriter` API |
| --- | --- |
| Worksheet tables | `add_table`, `Table`, `TableColumn`, `TableStyle`, `TableFunction` |
| Autofilters | `autofilter`, `filter_column`, `filter_automatic_off`, `FilterCondition`, `FilterCriteria`, `FilterData` |
| Data validation | `add_data_validation`, `DataValidation`, `DataValidationRule`, `DataValidationErrorStyle` |
| Conditional formatting | `add_conditional_format` with `ConditionalFormatCell`, `Blank`, `Error`, `Duplicate`, `Formula`, `Average`, `Top`, `Text`, `Date`, `2ColorScale`, `3ColorScale`, `DataBar`, `IconSet`, `CustomIcon` |

### Graphics and annotations

| Feature | `rust_xlsxwriter` API |
| --- | --- |
| Charts | `Chart` (+ series, axes, legend, trendlines, error bars, data labels, data table, ...), `insert_chart`, `insert_chart_with_offset` |
| Chartsheets | `Workbook::add_chartsheet`, `Worksheet::new_chartsheet` |
| Sparklines | `add_sparkline`, `add_sparkline_group`, `Sparkline` |
| Images | `insert_image`, `insert_image_with_offset`, `insert_image_fit_to_cell`, `insert_image_fit_to_cell_centered`, `insert_background_image`, `Image` |
| Shapes / textboxes | `insert_shape`, `insert_shape_with_offset`, `Shape` |
| Buttons (for macros) | `insert_button`, `insert_button_with_offset`, `Button` |
| Cell notes | `insert_note`, `show_all_notes`, `set_default_note_author`, `Note` |

### Protection

`protect`, `protect_with_password`, `protect_with_options` (`ProtectionOptions`), `unprotect_range`,
`unprotect_range_with_options`; format-level `set_locked` / `set_hidden` (see above).

### Workbook level

| `rust_xlsxwriter` API | Notes |
| --- | --- |
| `save_to_buffer`, `save_to_writer` | Save to bytes / a writer instead of a path (`BytesIO`, web responses) |
| `set_properties`, `DocProperties`, `CustomProperty` | Author, title, custom properties, ... |
| `define_name` | Named ranges and constants |
| `use_custom_theme`, `use_excel_2023_theme` | Workbook themes |
| `add_vba_project`, `add_vba_project_with_signature`, `set_vba_name` | Macros / `.xlsm` |
| `add_worksheet_with_low_memory` | Second memory-saving mode: also keeps one row in memory, but strings go to the shared string table instead of inline — smaller files and standard Excel output, a bit more memory with many unique strings |
| `set_tempdir` | Where constant-memory sheets keep their temporary files (defaults to the OS temp dir, which may be a RAM disk) |
| `use_zip_large_file` | ZIP64 for very large outputs |
| `read_only_recommended` | "Open as read-only" prompt |

### Not applicable from Python

- `serialize`, `serialize_headers*`, `deserialize_headers*`, `XlsxSerialize` derive (the `serde` crate feature):
  Rust struct serialization. The closest Python equivalent today is `write_rows` with dict rows.

### Optional crate features that are not enabled

`zmij` (a faster float formatter that `rust_xlsxwriter` recommends over the deprecated `ryu`, ~10% faster on numeric
sheets), `enhanced_autofit` (only useful once autofit is exposed), `rust_decimal`, `jiff`, `polars`, `serde`, `wasm`.

### Reading side

Inherited from `fastxlsx` (see [Reading](#reading)); none of the points below are planned for now.

- XLSX only. `calamine` can also open XLS / XLSB / ODS, but only `Xlsx` is wired up.
- Values only: no formulas, formats, merged-cell info, charts or images.
- The whole worksheet range is loaded into memory when a sheet is opened; there is no streaming reader.
- No modify / append / update of existing files.

## Development

```bash
pip install maturin
maturin develop --release        # build the extension into the current virtualenv
pip install ".[dev]"             # pytest, ruff, xlsxwriter
pytest
```

The benchmark script [`sandbox.py`](sandbox.py) writes 300 000 rows × 17 columns to three sheets with both `fastxlsxio`
and `xlsxwriter` and prints the timings (needs `xlsxwriter` and Python 3.12+ for `itertools.batched`).
`pip install ".[bench]"` pulls in the other libraries used for comparisons.

Type stubs live in [`python/fastxlsxio/`](python/fastxlsxio) (`*.pyi`).

## License

Apache-2.0, see [LICENSE](LICENSE).
