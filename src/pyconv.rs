use std::fmt::{Write};

use pyo3::ffi;
use pyo3::exceptions::{PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateAccess, PyDateTime, PyFloat, PySequence, PyTime, PyTimeAccess, PyString};
use rust_xlsxwriter::{ExcelDateTime};

pub type XlsxInt = i32;
pub type XlsxFloat = f64;

fn round3(val: XlsxFloat) -> XlsxFloat {
    (val * 1000.0).round() / 1000.0
}

pub fn pydecimal_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<XlsxFloat, PyErr> {
    let val: XlsxFloat = value.call_method0("__float__")?.extract()?;
    Ok(val)
}

pub fn pytime_xlsx_format<'py>(t: &Bound<'py, PyTime>) -> Result<ExcelDateTime, PyErr> {
    ExcelDateTime::from_hms(
        t.get_hour() as u16,
        t.get_minute() as u8,
        t.get_second() as u8,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))
}

pub fn pydate_xlsx_format<'py>(d: &Bound<'py, PyDate>) -> Result<ExcelDateTime, PyErr> {
    ExcelDateTime::from_ymd(
        d.get_year() as u16,
        d.get_month() as u8,
        d.get_day() as u8,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))
}

pub fn pydatetime_xlsx_format<'py>(dt: &Bound<'py, PyDateTime>) -> Result<ExcelDateTime, PyErr> {
    ExcelDateTime::from_ymd(
        dt.get_year() as u16,
        dt.get_month() as u8,
        dt.get_day() as u8,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))?
    .and_hms(
        dt.get_hour() as u16,
        dt.get_minute() as u8,
        dt.get_second() as u16,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))
}

pub fn pybool_xlsx_format<'py>(b: &Bound<'py, PyBool>) -> Result<&'py str, PyErr> {
    Ok(if b.is_true() { "Да" } else { "Нет" })
}

pub fn pyone_dimensional_iter_xlsx_format<'py>(seq: &Bound<'py, PySequence>) -> Result<String, PyErr> {
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
                let i: XlsxInt = value.extract()?;
                let _ = write!(buffer, "{}", i);
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
