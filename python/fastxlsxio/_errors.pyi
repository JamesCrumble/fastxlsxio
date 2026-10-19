class XIOError(Exception):
    """Base of everything fastxlsxio raises. Subclasses also inherit the builtin they replaced."""

class XIOConversionError(XIOError, ValueError):
    """A value can't become a cell. `address` is e.g. "B3", None if unknown."""

    value: str
    py_type: str
    target: str
    reason: str
    row: int | None
    col: int | None
    @property
    def address(self) -> str | None: ...
    def __init__(
        self,
        value: str,
        py_type: str,
        target: str,
        reason: str = "",
        row: int | None = None,
        col: int | None = None,
    ) -> None: ...

class XIOOverflowError(XIOConversionError, OverflowError):
    """A number or row/column index doesn't fit the target type."""

class XIOArgumentError(XIOError, ValueError): ...

class XIONotFoundError(XIOArgumentError):
    kind: str
    key: str
    def __init__(self, kind: str, key: str) -> None: ...

class XIOIterableError(XIOArgumentError):
    """A row / column you passed isn't iterable."""

    what: str
    got: str
    reason: str
    def __init__(self, what: str, got: str, reason: str) -> None: ...

class XIOUnsupportedError(XIOError, NotImplementedError): ...

class XIOStateError(XIOError, RuntimeError):
    """Worksheet is busy in another thread or its handle is stale."""

class XIOXlsxError(XIOError, ValueError):
    """The xlsx backend refused: Excel limits, bad sheet name, bad merge range..."""

    action: str
    reason: str
    row: int | None
    col: int | None
    @property
    def address(self) -> str | None: ...
    def __init__(self, action: str, reason: str, row: int | None = None, col: int | None = None) -> None: ...

class XIOFileError(XIOError, FileExistsError):
    """Can't save the file, or the workbook is invalid at save time (e.g. duplicate sheet names)."""

    action: str
    path: str
    reason: str
    def __init__(self, action: str, path: str, reason: str) -> None: ...
