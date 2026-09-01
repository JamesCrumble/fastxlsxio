use std::fmt::{Write};
use pyo3::exceptions::{PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateAccess, PyDateTime, PyFloat, PyInt, PyTime, PyTimeAccess};
use rust_xlsxwriter::{ExcelDateTime};

fn round3(val: f64) -> f64 {
    (val * 1000.0).round() / 1000.0
}

pub fn pydecimal_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<f64, PyErr> {
    let val: f64 = value.call_method0("__float__")?.extract()?;
    Ok(val)
}

pub fn pytime_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<ExcelDateTime, PyErr> {
    let t = value.downcast::<PyTime>()?;
    ExcelDateTime::from_hms(
        t.get_hour() as u16,
        t.get_minute() as u8,
        t.get_second() as u8,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))
}

pub fn pydate_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<ExcelDateTime, PyErr> {
    let d = value.downcast::<PyDate>()?;
    ExcelDateTime::from_ymd(
        d.get_year() as u16,
        d.get_month() as u8,
        d.get_day() as u8,
    )
    .map_err(|e| PyValueError::new_err(e.to_string()))
}

pub fn pydatetime_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<ExcelDateTime, PyErr> {
    let dt = value.downcast::<PyDateTime>()?;
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

pub fn pybool_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<&'py str, PyErr> {
    Ok(if value.extract::<bool>()? { "Да" } else { "Нет" })
}

pub fn pyone_dimensional_iter_xlsx_format<'py>(value: &Bound<'py, PyAny>) -> Result<String, PyErr> {
    let mut buffer = String::with_capacity(64);
    let seq = value.try_iter().unwrap();
    for (i, element) in seq.enumerate() {
        let elem = element?;
        if i > 0 {
            buffer.push_str(", ");
        }

        if elem.is_none() {
            continue;
        } else if elem.is_instance_of::<PyBool>() {
            buffer.push_str(pybool_xlsx_format(&elem)?);
        } else if elem.is_instance_of::<PyFloat>() {
            let f: f64 = elem.extract()?;
            let _ = write!(buffer, "{}", round3(f));
        } else if elem.is_instance_of::<PyInt>() {
            let i: i64 = elem.extract()?;
            let _ = write!(buffer, "{}", i);
        } else if let Ok(s) = elem.extract::<&str>() {
            buffer.push_str(s);
        } else {
            let s = elem.str()?;
            buffer.push_str(&s.to_string_lossy());
        }
    }
    Ok(buffer)
}