from collections.abc import Iterable
from typing import Any, ClassVar

# Bad `row`/`col` arguments (negative, too big, wrong type) raise a plain OverflowError / TypeError.

class ColTypeHint:
    """Cell value kind. `Unknown` and `Dynamic` are internal and can't be bound to a format."""

    Float: ClassVar[ColTypeHint]
    Int: ClassVar[ColTypeHint]
    String: ClassVar[ColTypeHint]
    Bool: ClassVar[ColTypeHint]
    Blank: ClassVar[ColTypeHint]
    Date: ClassVar[ColTypeHint]
    Time: ClassVar[ColTypeHint]
    DateTime: ClassVar[ColTypeHint]
    Sequence: ClassVar[ColTypeHint]
    Unknown: ClassVar[ColTypeHint]
    Dynamic: ClassVar[ColTypeHint]
    def __eq__(self, other: object) -> bool: ...
    def __int__(self) -> int: ...
    def __hash__(self) -> int: ...

class XIOWOptions:
    """
    - `constant_memory` - flush rows to disk as you go: write top to bottom, no `write_column`.
    - `cache_typehints_write_optimization` - assume a column keeps one type (None is fine).
    - `cache_col_formats` / `cache_row_formats` - don't re-set a format a column / row already has.

    Raises:
        XIOArgumentError: both `cache_col_formats` and `cache_row_formats` are on.
    """

    def __init__(
        self,
        constant_memory: bool = True,
        cache_col_formats: bool = True,
        cache_row_formats: bool = False,
        cache_typehints_write_optimization: bool = True,
    ) -> None: ...
    @property
    def constant_memory(self) -> bool: ...
    @constant_memory.setter
    def constant_memory(self, value: bool) -> None: ...
    @property
    def cache_col_formats(self) -> bool: ...
    @cache_col_formats.setter
    def cache_col_formats(self, value: bool) -> None: ...
    @property
    def cache_row_formats(self) -> bool: ...
    @cache_row_formats.setter
    def cache_row_formats(self, value: bool) -> None: ...
    @property
    def cache_typehints_write_optimization(self) -> bool: ...
    @cache_typehints_write_optimization.setter
    def cache_typehints_write_optimization(self, value: bool) -> None: ...

class XIOFormat:
    """Cell format. Get one from `XIOWWorkbook.add_format`."""

    @staticmethod
    def from_properties(properties: dict[str, Any]) -> XIOFormat:
        """Only `num_format` (str) and `bold` (bool) work for now.

        Raises:
            XIOUnsupportedError: any other key.
            TypeError: a value has the wrong type (e.g. `num_format=5`).
        """

class XIOWWorksheet:
    """Get it from `XIOWWorkbook.add_worksheet`. Write a sheet fully before adding the next one,
    older handles stop working (`XIOStateError`).

    Ints beyond +-2**53 are written as text (Excel can't hold them exactly); unknown objects as `str(value)`.
    """

    @property
    def name(self) -> str: ...
    @property
    def constant_memory(self) -> bool: ...
    @property
    def options(self) -> XIOWOptions: ...
    def write_cell(self, row: int, col: int, value: Any, format: XIOFormat | None = None) -> None:
        """Write one value. `row` and `col` are 0-based.

        Raises:
            XIOConversionError: the value can't be converted.
            XIOXlsxError: the cell is outside Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def merge_range(
        self,
        first_row: int,
        first_col: int,
        last_row: int,
        last_col: int,
        value: Any,
        format: XIOFormat | None = None,
    ) -> None:
        """Merge a range (0-based, ends included) and put `value` in it as text.

        Raises:
            XIOXlsxError: single-cell range, overlaps another merge, or out of Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def write_row(self, row: int, col: int, value: Iterable[Any], formats: list[XIOFormat | None] | None = None) -> None:
        """Write values left to right from (`row`, `col`). `formats[i]` is for the i-th value.

        Raises:
            XIOIterableError: `value` isn't iterable.
            XIOOverflowError: the row runs past the last column index that fits (u16).
            XIOConversionError: a value can't be converted.
            XIOXlsxError: a cell is outside Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def write_rows(
        self,
        row: int,
        col: int,
        value: Iterable[Iterable[Any]],
        formats: list[XIOFormat | None] | None = None,
    ) -> None:
        """Write many rows from (`row`, `col`). `formats[i]` applies to the i-th column of every row.
        Rows are written in chunks of 1024: on error the earlier chunks are already in.

        Raises:
            XIOIterableError: a row isn't iterable.
            XIOOverflowError: a row or column index doesn't fit (u32 / u16).
            XIOConversionError: a value can't be converted. `.address` points at it.
            XIOXlsxError: a cell is outside Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def write_column(self, row: int, col: int, value: Iterable[Any], format: XIOFormat | None = None) -> None:
        """Write values top to bottom from (`row`, `col`). Not available in `constant_memory` mode.

        Raises:
            XIOUnsupportedError: the worksheet is in `constant_memory` mode.
            XIOArgumentError: `format` was passed (not supported here yet).
            XIOIterableError: `value` isn't iterable.
            XIOOverflowError: the column runs past the last row index that fits (u32).
            XIOConversionError: a value can't be converted.
            XIOXlsxError: a cell is outside Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def set_column_width(self, col: int, width: float) -> None:
        """Set a column's width (in characters).

        Raises:
            XIOXlsxError: column is out of Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def set_column_format(self, col: int, format: XIOFormat) -> None:
        """Default format for a column; cells with their own format keep it.

        Raises:
            XIOXlsxError: column is out of Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

    def set_row_format(self, row: int, format: XIOFormat) -> None:
        """Same as `set_column_format`, but for a row.

        Raises:
            XIOXlsxError: row is out of Excel's limits.
            XIOStateError: the sheet is busy or stale.
        """

class XIOWWorkbook:
    """Add sheets, write into them, `save()`."""

    def __init__(self, filepath: str | None = None, options: XIOWOptions | None = None) -> None:
        """`filepath` is the default for `save()`, `options` the default for new worksheets."""

    @property
    def options(self) -> XIOWOptions: ...
    @property
    def sheetnames(self) -> list[str]:
        """Names of all worksheets, in order."""

    def add_worksheet(self, name: str, options: XIOWOptions | None = None) -> XIOWWorksheet:
        """Add a sheet. `options` overrides the workbook's options for this sheet only.

        Raises:
            XIOXlsxError: bad name (empty, over 31 chars, contains `[ ] : * ? / \\`).
            XIOStateError: another worksheet is being written right now.
        """

    def add_format(
        self,
        properties: dict[str, Any],
        bind_to_datatype: ColTypeHint | None = None,
    ) -> XIOFormat:
        """Properties: `num_format` (str), `bold` (True).

        With `bind_to_datatype` it becomes the default for every cell of that type written from
        now on (re-binding replaces it).

        Raises:
            XIOArgumentError: `bind_to_datatype` is `ColTypeHint.Unknown` or `Dynamic`.
            XIOUnsupportedError: unknown property.
        """

    def get_by_idx(self, idx: int) -> XIOWWorksheet:
        """Sheet by 0-based index.

        Raises:
            XIONotFoundError: no sheet at that index.
        """

    def get_by_name(self, name: str) -> XIOWWorksheet:
        """Sheet by name.

        Raises:
            XIONotFoundError: no sheet with that name.
        """

    def save(self, path: str | None = None) -> None:
        """`path` wins over the constructor's `filepath`.

        Raises:
            XIOArgumentError: no path given anywhere.
            XIOFileError: can't write the file, or the workbook is invalid (e.g. two sheets
                with the same name). `.reason` has the details.
        """
