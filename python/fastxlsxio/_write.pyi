from collections.abc import Iterable
from typing import Any, ClassVar

class ColTypeHint:
    """Classifier of the Python value written into a cell, used internally
    for column type-hint caching and as the key for type-bound default
    formats registered via `XIOWWorkbook.add_format(..., bind_to_datatype=...)`.

    `Unknown` is an internal placeholder for not-yet-seen columns and cannot
    be used as a `bind_to_datatype` value.
    """

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
    def __eq__(self, other: object) -> bool: ...
    def __int__(self) -> int: ...
    def __hash__(self) -> int: ...

class XIOWOptions:
    """
    `constant_memory` - flush buffer on disk. Chaotic writes are not allowed. More info in rust_xlsxwriter doc.\n
    `cache_typehints_write_optimization` - Expected to have static data types on columns. Only nullable allowed

    `cache_col_formats` - cache cells format upon columns\n
    `cache_row_formats` - cache cells format upon rows\n
    P.S. Both values `cache_col_formats` and `cache_row_formats` should not be used at the same time due prioritized col branching\n
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
    @staticmethod
    def from_properties(properties: dict[str, Any]) -> XIOFormat: ...

class XIOWWorksheet:
    """xlsxwriter worksheet-like class"""

    def __init__(self, title: str) -> None:
        """Initialize a new worksheet with the specified title.

        Parameters
        ----------
        title : str
            The title of the worksheet.
        """

    @property
    def name(self) -> str: ...
    @property
    def constant_memory(self) -> bool: ...
    def write_cell(self, row: int, col: int, value: Any, format: XIOFormat | None = None) -> None:
        """Write a value to a specific cell in the worksheet.

        Parameters
        ----------
        row : int
            0-based row index.
        col : int
            0-based column index.
        value : Any
            The value to write to the cell.
        format : XIOFormat | None = None
            value format.
        """

    def merge_range(first_row: int, first_col: int, last_row: int, last_col: int, value: Any, format: XIOFormat | None = None) -> None:
        """Merge a range of cells.

        Parameters
        ----------
        first_row : int
            The first row of the range. (All zero indexed.)
        first_col : int
            The first column of the range.
        last_row : int
            The last row of the range.
        last_col : int
            The last column of the range.
        value : Any
            The value to write to the cell.
        format : XIOFormat | None = None
            value format.
        """

    def write_row(self, row: int, col: int, value: Iterable[Any], formats: list[XIOFormat] | None = None) -> None:
        """Write a row of values starting from a specific cell.

        Parameters
        ----------
        row : int
            0-based row index.
        col : int
            0-based starting column index.
        value : Iterable[Any]
            The 1D iterable of cell values to write.
        formats : list[XIOFormat] | None = None
            values formats.
        """

    def write_rows(self, row: int, col: int, value: Iterable[Iterable[Any]], formats: list[XIOFormat] | None = None) -> None:
        """Write multiple rows of values starting from a specific cell.

        Parameters
        ----------
        row : int
            0-based starting row index.
        col : int
            0-based starting column index.
        value : Iterable[Iterable[Any]]
            The 2D iterable of row values to write.
        formats : list[XIOFormat] | None = None
            values formats.
        """

    def write_column(self, row: int, col: int, value: Iterable[Any], format: XIOFormat | None = None) -> None:
        """Write a column of values starting from a specific cell.

        Parameters
        ----------
        row : int
            0-based starting row index.
        col : int
            0-based column index.
        value : Iterable[Any]
            The 1D iterable of column values to write.
        format : XIOFormat | None = None
            values format.
        """

    def write_columns(self, row: int, col: int, value: Iterable[Iterable[Any]], formats: list[XIOFormat] | None = None) -> None:
        """Write multiple columns of values starting from a specific cell.

        Parameters
        ----------
        row : int
            0-based starting row index.
        col : int
            0-based starting column index.
        value : Iterable[Iterable[Any]]
            The columns of values to write.
        formats : list[XIOFormat] | None = None
            values formats.
        """

    def set_column_width(self, col: int, width: float) -> bool:
        """
        Set the width of a single column

        Args:
            col:   First column (zero-indexed).
            width: Column width.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """

    def set_column_format(self, col: int, format: XIOFormat) -> None:
        """
        Set the format for a column of cells.

        The `set_column_format()` method is used to change the default format of a
        column. Any unformatted data written to that column will then adopt that
        format. Formatted data written to the column will maintain its own cell
        format. See the example below.

        A future version of this library may support automatic merging of
        explicit cell formatting with the column formatting but that isn't
        currently supported.

        # Parameters

        - `col`: The zero indexed column number.
        - `format`: The [`XIOFormat`] property for the cell.
        """
    def set_row_format(self, row: int, format: XIOFormat) -> None:
        """
        Set the format for a row of cells.

        The `set_row_format()` method is used to change the default format of a
        row. Any unformatted data written to that row will then adopt that
        format. Formatted data written to the row will maintain its own cell
        format. See the example below.

        A future version of this library may support automatic merging of
        explicit cell formatting with the row formatting but that isn't
        currently supported.

        # Parameters

        - `row`: The zero indexed row number.
        - `format`: The [`XIOFormat`] property for the cell.
        """

class XIOWWorkbook:
    """xlsxwriter workbook-like class"""

    def __init__(self, filepath: str | None, options: XIOWOptions | None = None) -> None: ...
    def add_worksheet(self, name: str, options: XIOWOptions | None = None) -> XIOWWorksheet:
        """Create a new worksheet with the specified name.

        Parameters
        ----------
        name : str
            The name of the new worksheet.
        options : XIOWOptions, uses default XIOWOptions
            options which control some logic

        Returns
        -------
        XIOWWorksheet
            The newly created worksheet.
        """
    def add_format(
        self,
        properties: dict[str, Any],
        bind_to_datatype: ColTypeHint | None = None,
    ) -> XIOFormat:
        """Adds format to workbook

        Parameters
        ----------
        properties : dict[str, Any]
            "num_format" : str
                format for numbers
            "bold" : bool
                set or unset bold to text. By default is False so True is only logical value here
        bind_to_datatype : ColTypeHint | None = None
            When provided, this format becomes the default for every cell of
            that type written afterwards (in this and any worksheet created
            after this call) that isn't given an explicit `format`. Calling
            this again with the same `bind_to_datatype` replaces the
            previously bound format. Raises ValueError for `ColTypeHint.Unknown`.

        Returns
        -------
        XIOFormat
            The format class to use in XIOWWorksheet.write_... methods
        """
    def get_by_idx(self, idx: int) -> XIOWWorksheet:
        """Get a worksheet by its 0-based index."""

    def get_by_name(self, name: str) -> XIOWWorksheet:
        """Get a worksheet by its name."""

    def save(self, path: str | None = None) -> None:
        """Save the workbook to the specified file path."""

    @property
    def sheetnames(self) -> list[str]:
        """Get the names of all worksheets in the workbook."""
