use std::fmt::{Write};

use crate::error::{XIOError, XIOResult};

use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateAccess, PyDateTime, PyFloat, PySequence, PyTime, PyTimeAccess, PyString};
use rust_xlsxwriter::{ExcelDateTime};

pub type XlsxInt = i64;

// Excel numbers are f64: ints above this lose digits
pub const XLSX_MAX_EXACT_INT: XlsxInt = 1 << 53;
pub type XlsxFloat = f64;

fn round3(val: XlsxFloat) -> XlsxFloat {
    (val * 1000.0).round() / 1000.0
}

pub fn pydecimal_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> XIOResult<XlsxFloat> {
    value
        .call_method0("__float__")
        .and_then(|f| f.extract::<XlsxFloat>())
        .map_err(|e| XIOError::conversion(value, "f64", e))
}

pub fn pytime_xlsx_format<'py>(t: &Bound<'py, PyTime>) -> XIOResult<ExcelDateTime> {
    ExcelDateTime::from_hms(
        t.get_hour() as u16,
        t.get_minute() as u8,
        t.get_second() as u8,
    )
    .map_err(|e| XIOError::xlsx("convert datetime", e))
}

pub fn pydate_xlsx_format<'py>(d: &Bound<'py, PyDate>) -> XIOResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(
        d.get_year() as u16,
        d.get_month() as u8,
        d.get_day() as u8,
    )
    .map_err(|e| XIOError::xlsx("convert datetime", e))
}

pub fn pydatetime_xlsx_format<'py>(dt: &Bound<'py, PyDateTime>) -> XIOResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(
        dt.get_year() as u16,
        dt.get_month() as u8,
        dt.get_day() as u8,
    )
    .map_err(|e| XIOError::xlsx("convert datetime", e))?
    .and_hms(
        dt.get_hour() as u16,
        dt.get_minute() as u8,
        dt.get_second() as u16,
    )
    .map_err(|e| XIOError::xlsx("convert datetime", e))
}

pub fn pybool_xlsx_format<'py>(b: &Bound<'py, PyBool>) -> XIOResult<&'py str> {
    Ok(if b.is_true() { "Да" } else { "Нет" })
}

/// "dd.mm.yyyy, hh:mm:ss" items joined by "; ", or None if some item isn't a datetime
fn pydatetimes_xlsx_format<'py>(seq: &Bound<'py, PySequence>) -> XIOResult<Option<String>> {
    let mut buffer = String::with_capacity(64);
    for (i, v) in seq.try_iter()?.enumerate() {
        let value = v?;
        let Ok(dt) = value.cast::<PyDateTime>() else {
            return Ok(None);
        };
        if i > 0 {
            buffer.push_str("; ");
        }
        let _ = write!(
            buffer,
            "{:02}.{:02}.{:04}, {:02}:{:02}:{:02}",
            dt.get_day(),
            dt.get_month(),
            dt.get_year(),
            dt.get_hour(),
            dt.get_minute(),
            dt.get_second(),
        );
    }
    Ok(Some(buffer))
}

/// Datetimes are joined by "; ", everything else by ", "
pub fn pyone_dimensional_iter_xlsx_format<'py>(seq: &Bound<'py, PySequence>) -> XIOResult<String> {
    if let Some(datetimes) = pydatetimes_xlsx_format(seq)? {
        return Ok(datetimes);
    }

    let mut buffer = String::with_capacity(64);
    for (i, v) in seq.try_iter()?.enumerate() {
        let value = v?;
        if i > 0 {
            buffer.push_str(", ");
        }

        if value.is_none() {
            continue;
        }

        let ptr = value.as_ptr();
        unsafe {
            if ffi::PyFloat_CheckExact(ptr) != 0 {
                let f: &Bound<'py, PyFloat> = value.cast_unchecked();
                let _ = write!(buffer, "{}", round3(f.value()));
            } else if ffi::PyLong_CheckExact(ptr) != 0 && ffi::PyBool_Check(ptr) == 0 {
                let s = value.str()?;
                buffer.push_str(&s.to_string_lossy());
            } else if ffi::PyUnicode_CheckExact(ptr) != 0 {
                let s: &Bound<'py, PyString> = value.cast_unchecked();
                let _ = write!(buffer, "{}", s);
            } else if ffi::PyBool_Check(ptr) != 0 {
                buffer.push_str(pybool_xlsx_format(&value.cast_unchecked())?);
            } else {
                let s = value.str()?;
                buffer.push_str(&s.to_string_lossy());
            }
        }

    }
    Ok(buffer)
}
