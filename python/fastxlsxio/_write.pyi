from collections.abc import Iterable
from typing import Any

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
        format : list[XIOFormat] | None = None
            values format.
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
        format : list[XIOFormat] | None = None
            values format.
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

    def write_matrix(self, row: int, col: int, value: Iterable[Iterable[Any]]) -> None:
        """Write a matrix of values starting from a specific cell.

        Parameters
        ----------
        row : int
            0-based starting row index.
        col : int
            0-based starting column index.
        value : Iterable[Iterable[Any]]
            The matrix of values to write.
        """

class XIOWWorkbook:
    """xlsxwriter workbook-like class"""

    def __init__(self) -> None: ...
    def add_worksheet(self, name: str, constant_memory: bool = False) -> XIOWWorksheet:
        """Create a new worksheet with the specified name.

        Parameters
        ----------
        name : str
            The name of the new worksheet.
        constant_memory : bool, default False
            Enabling constant memory optimization feature.

        Returns
        -------
        XIOWWorksheet
            The newly created worksheet.
        """
    def add_format(self, properties: dict[str, Any]) -> XIOFormat:
        """Adds format to workbook

        Parameters
        ----------
        properties : dict[str, Any]
            "num_format" : str
                format for numbers
            "bold" : bool
                set or unset bold to text. By default is False so True is only logical value here

        Returns
        -------
        XIOFormat
            The format class to use in XIOWWorksheet.write_... methods
        """
    def get_by_idx(self, idx: int) -> XIOWWorksheet:
        """Get a worksheet by its 0-based index."""

    def get_by_name(self, name: str) -> XIOWWorksheet:
        """Get a worksheet by its name."""

    def save(self, path: str) -> None:
        """Save the workbook to the specified file path."""

    @property
    def sheetnames(self) -> list[str]:
        """Get the names of all worksheets in the workbook."""
