use crate::write::pyconv::*;

use std::borrow::Cow;
use std::sync::LazyLock;

use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::exceptions::{PyFileExistsError, PyRuntimeError, PyValueError};
use pyo3::types::{PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyTuple, PyList, PySequence, PyString, PyTime};
use rust_xlsxwriter::{Workbook, Worksheet, Format, ExcelDateTime, RowNum, ColNum};

pub static DEFAULT_FORMAT: LazyLock<Format> = LazyLock::new(Format::default);

#[pyclass(from_py_object, get_all, set_all)]
#[derive(Clone, Debug, PartialEq)]
pub struct XIOWOptions {
    pub constant_memory: bool,
    pub cache_col_formats: bool,
    pub cache_row_formats: bool,
    pub cache_typehints_write_optimization: bool,
}

impl XIOWOptions {
    fn default() -> Self {
        Self {
            constant_memory: true,
            cache_col_formats: true,
            cache_row_formats: false,
            cache_typehints_write_optimization: true,
        }
    }
}

#[pymethods]
impl XIOWOptions {

    #[new]
    #[pyo3(signature = (
        constant_memory = true,
        cache_col_formats = true,
        cache_row_formats = false,
        cache_typehints_write_optimization = true,
    ))]
    pub fn new(
        constant_memory: Option<bool>,
        cache_col_formats: Option<bool>,
        cache_row_formats: Option<bool>,
        cache_typehints_write_optimization: Option<bool>,
    ) -> PyResult<Self> {
        let mut options = XIOWOptions::default();
        if let Some(v) = constant_memory {
            options.constant_memory = v;
        }
        if let Some(v) = cache_col_formats {
            options.cache_col_formats = v;
        }
        if let Some(v) = cache_row_formats {
            options.cache_row_formats = v;
        }
        if let Some(v) = cache_typehints_write_optimization {
            options.cache_typehints_write_optimization = v;
        }

        if options.cache_row_formats && options.cache_col_formats {
            return Err(PyValueError::new_err("Recommended to not use both of caches. Col cache is prioritized in branching even if both enabled"))
        }

        Ok(options)
    }

    fn __repr__(&self) -> String {
        format!(
            "XIOWOptions(constant_memory={}, cache_col_formats={}, cache_row_formats={}, cache_typehints_write_optimization={})",
            self.constant_memory, self.cache_col_formats, self.cache_row_formats, self.cache_typehints_write_optimization
        )
    }
}

#[pyclass(from_py_object, eq, eq_int)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColTypeHint {
    Float,
    Int,
    String,
    Bool,
    Blank,

    // unsupported fast convertions
    Date,
    Time,
    DateTime,
    Sequence,

    Unknown,
    // column proved heterogeneous
    Dynamic,
}

impl ColTypeHint {
    pub const COUNT: usize = 11;

    fn as_index(self) -> usize {
        match self {
            ColTypeHint::Float => 0,
            ColTypeHint::Int => 1,
            ColTypeHint::String => 2,
            ColTypeHint::Bool => 3,
            ColTypeHint::Blank => 4,
            ColTypeHint::Date => 5,
            ColTypeHint::Time => 6,
            ColTypeHint::DateTime => 7,
            ColTypeHint::Sequence => 8,
            ColTypeHint::Unknown => 9,
            ColTypeHint::Dynamic => 10,
        }
    }

    pub fn from_excel_cell(cell: &ExcelCell) -> Self {
        match cell {
            ExcelCell::Float(_) => ColTypeHint::Float,
            ExcelCell::Int(_) => ColTypeHint::Int,
            ExcelCell::String(_) => ColTypeHint::String,
            ExcelCell::Bool(_) => ColTypeHint::Bool,
            ExcelCell::Date(_) => ColTypeHint::Date,
            ExcelCell::Time(_) => ColTypeHint::Time,
            ExcelCell::DateTime(_) => ColTypeHint::DateTime,
            ExcelCell::Sequence(_) => ColTypeHint::Sequence,
            ExcelCell::Blank => ColTypeHint::Blank,
        }
    }
}

pub enum ExcelCell<'a> {
    String(Cow<'a, str>),
    Float(XlsxFloat),
    Int(XlsxInt),
    Bool(bool),
    DateTime(ExcelDateTime),
    Date(ExcelDateTime),
    Time(ExcelDateTime),
    Sequence(String),
    Blank,
}

impl<'a> ExcelCell<'a> {

    pub fn from_py(value: &'a Bound<'a, PyAny>) -> PyResult<Self> {
        let ptr = value.as_ptr();

        unsafe {
            if ffi::PyFloat_CheckExact(ptr) != 0 {
                let f: &Bound<'a, PyFloat> = value.cast_unchecked();
                return Ok(ExcelCell::Float(f.value()));
            }
            if ffi::PyLong_CheckExact(ptr) != 0 && ffi::PyBool_Check(ptr) == 0 {
                let val: XlsxInt = value.extract()?;
                return Ok(ExcelCell::Int(val));
            }
            if ffi::PyUnicode_CheckExact(ptr) != 0 {
                let s: &Bound<'a, PyString> = value.cast_unchecked();
                return Ok(ExcelCell::String(Cow::Borrowed(s.to_str()?)));
            }
            if value.is_none() {
                return Ok(ExcelCell::Blank);
            }
            if ffi::PyBool_Check(ptr) != 0 {
                let b: &Bound<'a, PyBool> = value.cast_unchecked();
                return Ok(ExcelCell::Bool(b.is_true()));
            }
        }

        if let Ok(dt) = value.cast::<PyDateTime>() {
            return Ok(ExcelCell::DateTime(pydatetime_xlsx_format(dt)?));
        }
        if let Ok(d) = value.cast::<PyDate>() {
            return Ok(ExcelCell::Date(pydate_xlsx_format(d)?));
        }
        if let Ok(t) = value.cast::<PyTime>() {
            return Ok(ExcelCell::Time(pytime_xlsx_format(t)?));
        }
        if let Ok(t) = value.cast::<PySequence>() {
            return Ok(ExcelCell::Sequence(pyone_dimensional_iter_xlsx_format(t)?));
        }
        if value.get_type().name()? == "Decimal" {
            return Ok(ExcelCell::Float(pydecimal_xlsx_format(value)?));
        }

        Ok(ExcelCell::String(Cow::Owned(value.to_string())))
    }

    pub fn from_py_hinted(value: &'a Bound<'a, PyAny>, hint: Option<&ColTypeHint>) -> PyResult<(Self, ColTypeHint)> {
        if let Some(h) = hint {
            if *h == ColTypeHint::Dynamic {
                let cell = Self::from_py(value)?;
                return Ok((cell, ColTypeHint::Dynamic));
            }

            if value.is_none() {
                return Ok((ExcelCell::Blank, ColTypeHint::Blank));
            }

            let has_fast_path = matches!(
                h,
                ColTypeHint::Float
                    | ColTypeHint::Int
                    | ColTypeHint::String
                    | ColTypeHint::Bool
                    | ColTypeHint::Date
                    | ColTypeHint::Time
                    | ColTypeHint::DateTime
            );

            if has_fast_path {
                let ptr = value.as_ptr();
                unsafe {
                    match h {
                        ColTypeHint::Float if ffi::PyFloat_CheckExact(ptr) != 0 => {
                            let f: &Bound<'a, PyFloat> = value.cast_unchecked();
                            return Ok((ExcelCell::Float(f.value()), ColTypeHint::Float));
                        }
                        ColTypeHint::Int if ffi::PyLong_CheckExact(ptr) != 0 && ffi::PyBool_Check(ptr) == 0 => {
                            let i: XlsxInt = value.extract()?;
                            return Ok((ExcelCell::Int(i), ColTypeHint::Int));
                        }
                        ColTypeHint::String if ffi::PyUnicode_CheckExact(ptr) != 0 => {
                            let s: &Bound<'a, PyString> = value.cast_unchecked();
                            return Ok((ExcelCell::String(Cow::Borrowed(s.to_str()?)), ColTypeHint::String));
                        }
                        ColTypeHint::Bool if ffi::PyBool_Check(ptr) != 0 => {
                            let b: &Bound<'a, PyBool> = value.cast_unchecked();
                            return Ok((ExcelCell::Bool(b.is_true()), ColTypeHint::Bool));
                        }
                        ColTypeHint::Date if ffi::PyDate_CheckExact(ptr) != 0 => {
                            let d: &Bound<'a, PyDate> = value.cast_unchecked();
                            return Ok((ExcelCell::Date(pydate_xlsx_format(d)?), ColTypeHint::Date));
                        }
                        ColTypeHint::Time if ffi::PyTime_CheckExact(ptr) != 0 => {
                            let t: &Bound<'a, PyTime> = value.cast_unchecked();
                            return Ok((ExcelCell::Time(pytime_xlsx_format(t)?), ColTypeHint::Time));
                        }
                        ColTypeHint::DateTime if ffi::PyDateTime_CheckExact(ptr) != 0 => {
                            let dt: &Bound<'a, PyDateTime> = value.cast_unchecked();
                            return Ok((ExcelCell::DateTime(pydatetime_xlsx_format(dt)?), ColTypeHint::DateTime));
                        }
                        _ => {
                            let cell = Self::from_py(value)?;
                            return Ok((cell, ColTypeHint::Dynamic));
                        }
                    }
                }
            }
        }

        let cell = Self::from_py(value)?;
        let new_hint = ColTypeHint::from_excel_cell(&cell);
        Ok((cell, new_hint))
    }

    pub fn write(
        &self,
        worksheet: &mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let res: Result<_, _> = match (self, format) {
            (ExcelCell::String(s), None) => worksheet.write_string(row, col, s.as_ref()),
            (ExcelCell::Blank, None) => Ok(worksheet),
            (ExcelCell::Int(n), None) => worksheet.write_number(row, col, *n),
            (ExcelCell::Float(n), None) => worksheet.write_number(row, col, *n),
            (ExcelCell::Bool(b), None) => {worksheet.write_boolean(row, col, *b)},
            (ExcelCell::Date(dt), None) => worksheet.write_datetime(row, col, dt),
            (ExcelCell::Time(dt), None) => worksheet.write_datetime(row, col, dt),
            (ExcelCell::DateTime(dt), None) => worksheet.write_datetime(row, col, dt),
            (ExcelCell::Sequence(s), None) => worksheet.write_string(row, col, s),

            (ExcelCell::String(s), Some(fmt)) => worksheet.write_string_with_format(row, col, s.as_ref(), fmt),
            (ExcelCell::Blank, Some(fmt)) => worksheet.write_blank(row, col, fmt),
            (ExcelCell::Int(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Float(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Bool(b), Some(fmt)) => worksheet.write_boolean_with_format(row, col, *b, fmt),
            (ExcelCell::Date(dt), Some(fmt)) => worksheet.write_datetime_with_format(row, col, dt, fmt),
            (ExcelCell::Time(dt), Some(fmt)) => worksheet.write_datetime_with_format(row, col, dt, fmt),
            (ExcelCell::DateTime(dt), Some(fmt)) => worksheet.write_datetime_with_format(row, col, dt, fmt),
            (ExcelCell::Sequence(s), Some(fmt)) => worksheet.write_string_with_format(row, col, s, fmt),
        };

        res.map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(())
    }
}

#[pyclass]
pub struct XIOFormat {
    rs_format: Format
}

impl XIOFormat {
    fn new(rs_format: Format) -> Self {
        Self {
            rs_format: rs_format,
        }
    }
}

macro_rules! unpack_format {
    ($format:expr, $rs_format:ident) => {
        let _guard = $format.map(|f| f.borrow());
        let $rs_format = _guard.as_deref().map(|f| &f.rs_format);
    };
}

macro_rules! unpack_formats {
    ($formats:expr, $rs_formats:ident) => {
        let _guards = match $formats {
            Some(list) => {
                let mut vec = Vec::with_capacity(list.len());
                for item in list.iter() {
                    if item.is_none() {
                        vec.push(None);
                    } else {
                        vec.push(Some(item.cast::<XIOFormat>()?.borrow()));
                    }
                }
                vec
            }
            None => Vec::new(),
        };

        let $rs_formats: Vec<Option<&Format>> = _guards
            .iter()
            .map(|g| g.as_deref().map(|f| &f.rs_format))
            .collect();
    };
}

macro_rules! extract_format_by_offset {
    ($rs_formats:expr, $offset:expr, $rs_format:ident) => {
        let $rs_format = $rs_formats.get($offset).copied().flatten();
    };
}

#[pymethods]
impl XIOFormat {

    #[staticmethod]
    #[pyo3(signature = (properties))]
    pub fn from_properties(properties: &Bound<'_, PyDict>) -> PyResult<Self> {
        let mut rs_format = Format::new();

        for (key, value) in properties.iter() {
            let k: String = key.extract()?;
            if k == "num_format" {
                let num_format: String = value.extract()?;
                rs_format = rs_format.set_num_format(num_format);
            } else if k == "bold" && value.extract()? {
                rs_format = rs_format.set_bold();
            } else {
                return Err(PyRuntimeError::new_err(
                    format!("Set format by {} is not implemented", k),
                ));
            }
        }

        Ok(Self::new(rs_format))
    }
}

// Owned by XIOWWorkbook, referenced by XIOWWorksheet through a raw pointer
// so cloning a worksheet handle never copies grown caches.
#[derive(Default)]
struct WorksheetCaches {
    col_hints: Vec<ColTypeHint>,
    col_formats_setted: Vec<Option<Format>>,
    row_formats_setted: Vec<Option<Format>>,
}

fn resolve_format_cache<'a>(
    cache: &mut Vec<Option<Format>>,
    idx: usize,
    format: Option<&'a Format>,
) -> (bool, Option<&'a Format>) {
    if idx >= cache.len() {
        cache.resize(idx + 1, None);
    }

    match (&cache[idx], format) {
        (None, Some(fmt)) => {
            cache[idx] = Some(fmt.clone());
            (true, Some(fmt))
        }
        (Some(cached), Some(fmt)) if cached == fmt => (false, None),
        _ => (false, format),
    }
}

#[pyclass(from_py_object, weakref)]
#[derive(Clone)]
pub struct XIOWWorksheet {
    worksheet: *mut Worksheet,
    caches: *mut WorksheetCaches,
    options: XIOWOptions,
    datatype_format_bindings: *const [Option<Format>; ColTypeHint::COUNT],
}
unsafe impl Send for XIOWWorksheet {}
unsafe impl Sync for XIOWWorksheet {}

impl XIOWWorksheet {

    #[inline(always)]
    fn worksheet_ref(&self) -> &Worksheet {
        assert!(!self.worksheet.is_null(), "ERROR: Something went wrong. Cannot use worksheet which is deallocated!!!");
        unsafe { &*self.worksheet }
    }

    #[inline(always)]
    fn worksheet_refmut(&self) -> &mut Worksheet {
        assert!(!self.worksheet.is_null(), "ERROR: Something went wrong. Cannot use worksheet which is deallocated!!!");
        unsafe { &mut *self.worksheet }
    }

    #[inline(always)]
    fn caches_refmut(&self) -> &mut WorksheetCaches {
        assert!(!self.caches.is_null(), "ERROR: Something went wrong. Cannot use worksheet which is deallocated!!!");
        unsafe { &mut *self.caches }
    }

    fn new(
        worksheet: &mut Worksheet,
        name: String,
        options: XIOWOptions,
        caches: *mut WorksheetCaches,
        datatype_format_bindings: *const [Option<Format>; ColTypeHint::COUNT],
    ) -> Self {
        let _ = worksheet.set_name(name).map_err(|e| PyValueError::new_err(e.to_string()));
        Self {
            worksheet: worksheet as *mut Worksheet,
            caches,
            options,
            datatype_format_bindings,
        }
    }

    fn _write_cell_rs(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let colidx = col as usize;
        let rowidx = row as usize;
        let caches = self.caches_refmut();

        let cell: ExcelCell = if self.options.cache_typehints_write_optimization {
            let cached_hint = caches.col_hints.get(colidx);
            let (cell, hint) = ExcelCell::from_py_hinted(value, cached_hint)?;
            if colidx >= caches.col_hints.len() {
                caches.col_hints.resize(colidx + 1, ColTypeHint::Unknown);
            }
            caches.col_hints[colidx] = hint;
            cell
        } else {
            ExcelCell::from_py(value)?
        };

        let format = match format {
            Some(f) => Some(f),
            None => {
                let hint = ColTypeHint::from_excel_cell(&cell);
                unsafe { (*self.datatype_format_bindings)[hint.as_index()].as_ref() }
            }
        };

        let (should_be_set, cell_format) = if self.options.cache_col_formats {
            resolve_format_cache(&mut caches.col_formats_setted, colidx, format)
        } else if self.options.cache_row_formats {
            resolve_format_cache(&mut caches.row_formats_setted, rowidx, format)
        } else {
            (false, format)
        };

        let worksheet = self.worksheet_refmut();

        if should_be_set {
            let fmt = cell_format.expect("format must be Some when should_be_set is true");
            if self.options.cache_col_formats {
                worksheet.set_column_format(col, fmt).map_err(|e| PyValueError::new_err(e.to_string()))?;
            } else {
                worksheet.set_row_format(row, fmt).map_err(|e| PyValueError::new_err(e.to_string()))?;
            }
        }
        cell.write(worksheet, row, col, cell_format)
    }
}

#[pymethods]
impl XIOWWorksheet {

    #[getter]
    pub fn constant_memory(&self) -> bool {
        self.options.constant_memory
    }

    #[getter]
    pub fn name(&self) -> String {
        self.worksheet_ref().name()
    }

    #[getter]
    pub fn options(&self) -> XIOWOptions {
        self.options.clone()
    }

    #[pyo3(signature = (row, col, value, format = None))]
    pub fn write_cell<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PyAny>,
        format: Option<&Bound<'py, XIOFormat>>,
    ) -> PyResult<()> {
        unpack_format!(format, rs_format);
        self._write_cell_rs(row, col, value, rs_format)
    }

    #[pyo3(signature = (row, col, value, formats = None))]
    pub fn write_row<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PyAny>,
        formats: Option<&Bound<'py, PyList>>,
    ) -> PyResult<()> {
        unpack_formats!(formats, rs_formats);

        if let Ok(list) = value.cast::<PyList>() {
            for (offset, item) in list.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, &item, rs_format)?;
            }
            return Ok(());
        }

        if let Ok(tuple) = value.cast::<PyTuple>() {
            for (offset, item) in tuple.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, &item, rs_format)?;
            }
            return Ok(());
        }

        let iter_result = if let Ok(dict) = value.cast::<PyDict>() {
            dict.values().try_iter()
        } else {
            value.try_iter()
        }.map_err(|e| {
            PyValueError::new_err(format!("Cannot write row from invalid object: {}", e))
        })?;

        for (offset, item) in iter_result.enumerate() {
            extract_format_by_offset!(rs_formats, offset, rs_format);
            // was .unwrap() - now errors instead of panicking
            self._write_cell_rs(row, col + offset as ColNum, &item?, rs_format)?;
        }

        Ok(())
    }

    #[pyo3(signature = (row, col, value, formats = None))]
    pub fn write_rows<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PySequence>,
        formats: Option<&Bound<'py, PyList>>,
    ) -> PyResult<()> {
        let rows_iter = value.try_iter()?;
        for (r_offset, row_obj_res) in rows_iter.enumerate() {
            let row_obj = row_obj_res?;
            let current_row = row + r_offset as RowNum;
            self.write_row(current_row, col, &row_obj, formats)?;
        }

        Ok(())
    }

    #[pyo3(signature = (row, col, value, format = None))]
    pub fn write_column<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PyAny>,
        format: Option<&Bound<'py, XIOFormat>>,
    ) -> PyResult<()> {
        if self.options.constant_memory {
            return Err(PyRuntimeError::new_err(
                "Cannot write columnar data in constant_memory worksheet mode",
            ));
        }
        if format.is_some() {
            return Err(PyValueError::new_err(
                "Formatted write not supported for write_column",
            ));
        }

        let iter = value.try_iter().map_err(|e| {
            PyValueError::new_err(format!("Cannot write column from invalid object: {}", e))
        })?;

        for (offset, item) in iter.enumerate() {
            self._write_cell_rs(row + offset as RowNum, col, &item?, None)?;
        }

        Ok(())
    }

    #[pyo3(signature = (first_row, first_col, last_row, last_col, value, format = None))]
    pub fn merge_range<'py>(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        value: &Bound<'py, PyAny>,
        format: Option<&Bound<'py, XIOFormat>>,
    ) -> PyResult<()> {
        unpack_format!(format, rs_format);
        let worksheet = self.worksheet_refmut();

        let py_str = value.str()?;
        let rust_str: &str = py_str.to_str()?;

        worksheet.merge_range(first_row, first_col, last_row, last_col, rust_str, rs_format.unwrap_or(&DEFAULT_FORMAT))
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }

    #[pyo3(signature = (col, width))]
    fn set_column_width(
        &self,
        col: ColNum,
        width: XlsxFloat
    ) -> PyResult<()> {
        self.worksheet_refmut().set_column_width(col, width).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(())
    }

    #[pyo3(signature = (col, format))]
    fn set_column_format<'py>(&mut self, col: ColNum, format: &Bound<'py, XIOFormat>) -> PyResult<()> {
        unpack_format!(Some(format), rs_format);
        let fmt = rs_format.unwrap();

        if self.options.cache_col_formats {
            let caches = self.caches_refmut();
            let idx = col as usize;
            if idx >= caches.col_formats_setted.len() {
                caches.col_formats_setted.resize(idx + 1, None);
            }
            // explicit override - unconditional, unlike resolve_format_cache's
            // opportunistic "first format wins" behavior for per-cell writes
            caches.col_formats_setted[idx] = Some(fmt.clone());
        }
        self.worksheet_refmut().set_column_format(col, fmt).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(())
    }

    #[pyo3(signature = (row, format))]
    fn set_row_format<'py>(&mut self, row: RowNum, format: &Bound<'py, XIOFormat>) -> PyResult<()> {
        unpack_format!(Some(format), rs_format);
        let fmt = rs_format.unwrap();

        if self.options.cache_row_formats {
            let caches = self.caches_refmut();
            let idx = row as usize;
            if idx >= caches.row_formats_setted.len() {
                caches.row_formats_setted.resize(idx + 1, None);
            }
            caches.row_formats_setted[idx] = Some(fmt.clone());
        }
        self.worksheet_refmut().set_row_format(row, fmt).map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(())
    }

    fn __repr__(&self) -> String {
        format!("<XIOWWorksheet \"{}\">", self.name())
    }
}

#[pyclass]
pub struct XIOWWorkbook {
    filepath: Option<String>,
    workbook: Workbook,
    worksheets: Vec<XIOWWorksheet>,
    worksheet_caches: Vec<Box<WorksheetCaches>>,
    pub options: XIOWOptions,
    datatype_format_bindings: Box<[Option<Format>; ColTypeHint::COUNT]>,
}

impl XIOWWorkbook {
    fn get_sheetnames_string(&mut self) -> String {
        let sheetnames = self.worksheets.iter()
            .map(|ws| format!("\"{}\"", ws.name()))
            .collect::<Vec<String>>();
        format!("[{}]", sheetnames.join(", "))
    }
}

#[pymethods]
impl XIOWWorkbook {

    #[new]
    #[pyo3(signature = (filepath = None, options = None))]
    fn new(filepath: Option<String>, options: Option<XIOWOptions>) -> Self {
        Self {
            filepath,
            workbook: Workbook::new(),
            worksheets: Vec::new(),
            worksheet_caches: Vec::new(),
            options: options.unwrap_or(XIOWOptions::default()),
            datatype_format_bindings: Box::new(std::array::from_fn(|_| None)),
        }
    }

    #[getter]
    fn sheetnames(&mut self) -> Vec<String> {
        self.worksheets
            .iter()
            .map(|ws| ws.name())
            .collect()
    }

    #[getter]
    pub fn options(&self) -> XIOWOptions {
        self.options.clone()
    }

    #[pyo3(signature = (properties, bind_to_datatype = None))]
    fn add_format<'py>(&mut self, properties: &Bound<'py, PyDict>, bind_to_datatype: Option<ColTypeHint>) -> PyResult<XIOFormat> {
        let format = XIOFormat::from_properties(properties)?;

        if let Some(hint) = bind_to_datatype {
            if matches!(hint, ColTypeHint::Unknown | ColTypeHint::Dynamic) {
                return Err(PyValueError::new_err(
                    "Cannot bind a format to ColTypeHint.Unknown or ColTypeHint.Dynamic — they're internal placeholders, not real cell types",
                ));
            }
            self.datatype_format_bindings[hint.as_index()] = Some(format.rs_format.clone());
        }

        Ok(format)
    }

    #[pyo3(signature = (name, options = None))]
    fn add_worksheet(&mut self, name: String, options: Option<XIOWOptions>) -> PyResult<XIOWWorksheet> {
        let ws_options = options.unwrap_or(self.options.clone());
        let worksheet = if ws_options.constant_memory {
            self.workbook.add_worksheet_with_constant_memory()
        } else {
            self.workbook.add_worksheet()
        };

        let mut caches = Box::new(WorksheetCaches::default());
        let caches_ptr: *mut WorksheetCaches = caches.as_mut();
        self.worksheet_caches.push(caches);

        let sheet = XIOWWorksheet::new(
            worksheet,
            name,
            ws_options,
            caches_ptr,
            self.datatype_format_bindings.as_ref() as *const _,
        );
        self.worksheets.push(sheet.clone());

        Ok(sheet)
    }

    #[pyo3(signature = (filepath = None))]
    fn save(&mut self, py: Python<'_>, filepath: Option<String>) -> PyResult<()> {
        if filepath.is_none() && self.filepath.is_none() {
            return Err(PyValueError::new_err("Expected provided path or inited filepath. Got nothing."));
        }

        let path;
        match filepath {
            Some(p) => {path = p}
            None => {path = self.filepath.clone().unwrap()}
        }

        py.detach(|| {self.workbook.save(path)}).map_err(|e| PyFileExistsError::new_err(e.to_string()))
    }

    fn get_by_idx(&mut self, idx: usize) -> PyResult<XIOWWorksheet> {
        let sheet = self
            .worksheets
            .get(idx)
            .ok_or_else(|| PyValueError::new_err(format!("Worksheet at index {} not found", idx)))?;

        Ok(sheet.clone())
    }

    fn get_by_name(&mut self, name: String) -> PyResult<XIOWWorksheet> {
        let sheet = self
            .worksheets
            .iter()
            .find(|ws| ws.name() == name)
            .ok_or_else(|| PyValueError::new_err(format!("Worksheet '{}' not found", name)))?;

        Ok(sheet.clone())
    }

    fn __repr__(&mut self) -> String {
        format!(
            "<XIOWWorkbook(sheetnames={})>",
            self.get_sheetnames_string()
        )
    }
}