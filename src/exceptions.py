__all__ = (
    "XIOError",
    "XIOConversionError",
    "XIOOverflowError",
    "XIOArgumentError",
    "XIONotFoundError",
    "XIOIterableError",
    "XIOUnsupportedError",
    "XIOStateError",
    "XIOXlsxError",
    "XIOFileError",
)


def _column_name(col: int) -> str:
    name = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        name = chr(ord("A") + rem) + name
    return name


def _at(row: "int | None", col: "int | None") -> str:
    if row is None or col is None:
        return ""
    return f" at {_column_name(col)}{row + 1} (row={row}, col={col})"


class XIOError(Exception):
    def __init__(self, *args):
        super().__init__(*args)

    def __str__(self) -> str:
        return str(self.args[0]) if self.args else ""


class XIOConversionError(XIOError, ValueError):
    def __init__(self, value, py_type, target, reason="", row=None, col=None):
        self.value = value
        self.py_type = py_type
        self.target = target
        self.reason = reason
        self.row = row
        self.col = col
        XIOError.__init__(self, value, py_type, target, reason, row, col)

    @property
    def address(self) -> "str | None":
        if self.row is None or self.col is None:
            return None
        return f"{_column_name(self.col)}{self.row + 1}"

    def __str__(self) -> str:
        msg = f"cannot convert {self.py_type} value {self.value} to {self.target}{_at(self.row, self.col)}"
        return f"{msg}: {self.reason}" if self.reason else msg


class XIOOverflowError(XIOConversionError, OverflowError):
    def __str__(self) -> str:
        return f"{self.py_type} value {self.value} does not fit into {self.target}{_at(self.row, self.col)}"


class XIOArgumentError(XIOError, ValueError):
    pass


class XIONotFoundError(XIOArgumentError):
    def __init__(self, kind, key):
        self.kind = kind
        self.key = key
        XIOError.__init__(self, kind, key)

    def __str__(self) -> str:
        return f"{self.kind} {self.key} not found"


class XIOIterableError(XIOArgumentError):
    def __init__(self, what, got, reason):
        self.what = what
        self.got = got
        self.reason = reason
        XIOError.__init__(self, what, got, reason)

    def __str__(self) -> str:
        return f"cannot write {self.what} from object of type {self.got}: {self.reason}"


class XIOUnsupportedError(XIOError, NotImplementedError):
    pass


class XIOStateError(XIOError, RuntimeError):
    pass


class XIOXlsxError(XIOError, ValueError):
    def __init__(self, action, reason, row=None, col=None):
        self.action = action
        self.reason = reason
        self.row = row
        self.col = col
        XIOError.__init__(self, action, reason, row, col)

    @property
    def address(self) -> "str | None":
        if self.row is None or self.col is None:
            return None
        return f"{_column_name(self.col)}{self.row + 1}"

    def __str__(self) -> str:
        return f"failed to {self.action}{_at(self.row, self.col)}: {self.reason}"


class XIOFileError(XIOError, FileExistsError):
    def __new__(cls, *args):
        # OSError.__new__ would interpret the args as (errno, strerror, filename)
        return super().__new__(cls)

    def __init__(self, action, path, reason):
        self.action = action
        self.path = path
        self.reason = reason
        XIOError.__init__(self, action, path, reason)
        self.args = (action, path, reason)

    def __str__(self) -> str:
        return f"failed to {self.action} {self.path}: {self.reason}"
