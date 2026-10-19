//! `XIOError` carries error parameters (value, types, cell); the Python classes in `exceptions.py` build the message from them.

use std::ffi::CString;
use std::fmt;

use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use rust_xlsxwriter::{ColNum, RowNum, XlsxError};

const EXCEPTIONS_SRC: &str = include_str!("exceptions.py");

const EXCEPTION_NAMES: &[&str] = &[
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
];

fn exceptions_module<'py>(py: Python<'py>) -> PyResult<Bound<'py, PyModule>> {
    static MODULE: PyOnceLock<Py<PyModule>> = PyOnceLock::new();
    let module = MODULE.get_or_try_init(py, || {
        let code = CString::new(EXCEPTIONS_SRC).expect("exceptions.py contains a NUL byte");
        PyModule::from_code(py, code.as_c_str(), c"exceptions.py", c"fastxlsxio").map(Bound::unbind)
    })?;
    Ok(module.bind(py).clone())
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    let exceptions = exceptions_module(m.py())?;
    for name in EXCEPTION_NAMES {
        m.add(*name, exceptions.getattr(*name)?)?;
    }
    Ok(())
}

fn instantiate<'py, A>(py: Python<'py>, name: &str, args: A) -> PyErr
where
    A: pyo3::call::PyCallArgs<'py>,
{
    let created = exceptions_module(py)
        .and_then(|m| m.getattr(name))
        .and_then(|cls| cls.call1(args));
    match created {
        Ok(instance) => PyErr::from_value(instance),
        Err(e) => e,
    }
}

const PREVIEW_MAX_CHARS: usize = 100;

#[derive(Clone, Copy, Debug)]
pub struct CellPos {
    pub row: RowNum,
    pub col: ColNum,
}

impl CellPos {
    fn parts(cell: &Option<CellPos>) -> (Option<RowNum>, Option<ColNum>) {
        (cell.map(|c| c.row), cell.map(|c| c.col))
    }
}

/// Truncated `repr()`; never fails
pub fn value_preview(value: &Bound<'_, PyAny>) -> String {
    let repr = match value.repr() {
        Ok(r) => r.to_string_lossy().into_owned(),
        Err(_) => return format!("<unrepresentable {}>", type_name(value)),
    };
    match repr.char_indices().nth(PREVIEW_MAX_CHARS) {
        Some((end, _)) => format!("{}…", &repr[..end]),
        None => repr,
    }
}

pub fn type_name(value: &Bound<'_, PyAny>) -> String {
    value
        .get_type()
        .qualname()
        .map(|n| n.to_string())
        .unwrap_or_else(|_| "<unknown>".to_string())
}

pub enum XIOError {
    Overflow { value: String, py_type: String, target: &'static str, cell: Option<CellPos> },
    Conversion { value: String, py_type: String, target: &'static str, reason: String, cell: Option<CellPos> },
    InvalidArgument(String),
    NotFound { kind: &'static str, key: String },
    InvalidIterable { what: &'static str, got: String, reason: String },
    Unsupported(String),
    State(String),
    Xlsx { action: &'static str, cell: Option<CellPos>, source: XlsxError },
    File { action: &'static str, path: String, reason: String },
    /// Python exception passed through as is
    Python(PyErr),
}

pub type XIOResult<T> = Result<T, XIOError>;

impl XIOError {
    pub fn overflow(value: &Bound<'_, PyAny>, target: &'static str) -> Self {
        Self::Overflow { value: value_preview(value), py_type: type_name(value), target, cell: None }
    }

    pub fn index_overflow(what: &str, value: impl fmt::Display, target: &'static str) -> Self {
        Self::Overflow { value: value.to_string(), py_type: format!("{what} index"), target, cell: None }
    }

    pub fn conversion(value: &Bound<'_, PyAny>, target: &'static str, reason: impl fmt::Display) -> Self {
        Self::Conversion {
            value: value_preview(value),
            py_type: type_name(value),
            target,
            reason: reason.to_string(),
            cell: None,
        }
    }

    pub fn invalid_argument(msg: impl Into<String>) -> Self {
        Self::InvalidArgument(msg.into())
    }

    pub fn unsupported(msg: impl Into<String>) -> Self {
        Self::Unsupported(msg.into())
    }

    pub fn state(msg: impl Into<String>) -> Self {
        Self::State(msg.into())
    }

    pub fn not_found(kind: &'static str, key: impl fmt::Display) -> Self {
        Self::NotFound { kind, key: key.to_string() }
    }

    pub fn not_iterable(what: &'static str, value: &Bound<'_, PyAny>, source: &PyErr) -> Self {
        Self::InvalidIterable { what, got: type_name(value), reason: source.to_string() }
    }

    pub fn xlsx(action: &'static str, source: XlsxError) -> Self {
        Self::Xlsx { action, cell: None, source }
    }

    pub fn file(action: &'static str, path: impl fmt::Debug, reason: impl fmt::Display) -> Self {
        Self::File { action, path: format!("{path:?}"), reason: reason.to_string() }
    }

    pub fn at(mut self, row: RowNum, col: ColNum) -> Self {
        match &mut self {
            Self::Overflow { cell, .. } | Self::Conversion { cell, .. } | Self::Xlsx { cell, .. } => {
                *cell = Some(CellPos { row, col });
            }
            _ => {}
        }
        self
    }
}

impl fmt::Debug for XIOError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Overflow { value, py_type, target, cell } => {
                write!(f, "Overflow({py_type} {value} -> {target}, {cell:?})")
            }
            Self::Conversion { value, py_type, target, reason, cell } => {
                write!(f, "Conversion({py_type} {value} -> {target}, {reason}, {cell:?})")
            }
            Self::InvalidArgument(m) | Self::Unsupported(m) | Self::State(m) => {
                write!(f, "{m}")
            }
            Self::NotFound { kind, key } => write!(f, "NotFound({kind} {key})"),
            Self::InvalidIterable { what, got, reason } => write!(f, "InvalidIterable({what}, {got}, {reason})"),
            Self::Xlsx { action, cell, source } => write!(f, "Xlsx({action}, {cell:?}, {source})"),
            Self::File { action, path, reason } => write!(f, "File({action} {path}: {reason})"),
            Self::Python(e) => write!(f, "Python({e})"),
        }
    }
}

impl From<PyErr> for XIOError {
    fn from(e: PyErr) -> Self {
        Self::Python(e)
    }
}

impl From<XIOError> for PyErr {
    fn from(e: XIOError) -> PyErr {
        Python::attach(|py| match e {
            XIOError::Python(err) => err,
            XIOError::Overflow { value, py_type, target, cell } => {
                let (row, col) = CellPos::parts(&cell);
                instantiate(py, "XIOOverflowError", (value, py_type, target, "", row, col))
            }
            XIOError::Conversion { value, py_type, target, reason, cell } => {
                let (row, col) = CellPos::parts(&cell);
                instantiate(py, "XIOConversionError", (value, py_type, target, reason, row, col))
            }
            XIOError::InvalidArgument(msg) => instantiate(py, "XIOArgumentError", (msg,)),
            XIOError::NotFound { kind, key } => instantiate(py, "XIONotFoundError", (kind, key)),
            XIOError::InvalidIterable { what, got, reason } => {
                instantiate(py, "XIOIterableError", (what, got, reason))
            }
            XIOError::Unsupported(msg) => instantiate(py, "XIOUnsupportedError", (msg,)),
            XIOError::State(msg) => instantiate(py, "XIOStateError", (msg,)),
            XIOError::Xlsx { action, cell, source } => {
                let (row, col) = CellPos::parts(&cell);
                instantiate(py, "XIOXlsxError", (action, source.to_string(), row, col))
            }
            XIOError::File { action, path, reason } => instantiate(py, "XIOFileError", (action, path, reason)),
        })
    }
}
