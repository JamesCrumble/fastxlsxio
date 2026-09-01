from enum import IntEnum
from typing import NamedTuple

class DType(IntEnum):
    """Enumeration for data types."""

    Int: int
    Float: int
    Str: int
    Bool: int
    Date: int
    DateTime: int
    Any: int

class DShape:
    """Class to describe the shape of data."""

    class Scalar(NamedTuple): ...

    class Row(NamedTuple):
        n_cols: int

    class Column(NamedTuple):
        n_rows: int

    class Matrix(NamedTuple):
        n_rows: int
        n_cols: int
