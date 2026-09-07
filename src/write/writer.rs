use crate::write::pyconv::*;

use std::borrow::Cow;
use std::sync::LazyLock;

use pyo3::exceptions::PyNotImplementedError;
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
        cache_typehints_write_optimization = true,
    ))]
    pub fn new(
        constant_memory: Option<bool>,
        cache_col_formats: Option<bool>,
        cache_typehints_write_optimization: Option<bool>,
    ) -> PyResult<Self> {
        let mut options = XIOWOptions::default();
        if let Some(v) = constant_memory {
            options.constant_memory = v;
        }
        if let Some(v) = cache_col_formats {
            options.cache_col_formats = v;
        }
        if let Some(v) = cache_typehints_write_optimization {
            options.cache_typehints_write_optimization = v;
        }

        if options.cache_row_formats {
            return Err(PyNotImplementedError::new_err("cache_row_formats are not implemented and cannot be enabled."))
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColTypeHint {
    Float,
    Int,
    String,
    Bool,
    Blank,

    // unsupported fast convertions
    DateTime,
    Sequence,

    Unknown,
}

impl ColTypeHint {
    pub fn from_excel_cell(cell: &ExcelCell) -> Self {
        match cell {
            ExcelCell::Float(_) => ColTypeHint::Float,
            ExcelCell::Int(_) => ColTypeHint::Int,
            ExcelCell::String(_) => ColTypeHint::String,
            ExcelCell::Bool(_) => ColTypeHint::Bool,
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
    Sequence(String),
    Blank,
}

impl<'a> ExcelCell<'a> {

    pub fn from_py(elem: &'a Bound<'a, PyAny>) -> PyResult<Self> {
        let ptr = elem.as_ptr();

        unsafe {
            // as fast as possible for regular data types without downcast/cast
            if ffi::PyFloat_CheckExact(ptr) != 0 {
                let f: &Bound<'a, PyFloat> = elem.cast_unchecked();
                return Ok(ExcelCell::Float(f.value()));
            }
            if ffi::PyLong_CheckExact(ptr) != 0 && ffi::PyBool_Check(ptr) == 0 {
                let val: XlsxInt = elem.extract()?;
                return Ok(ExcelCell::Int(val));
            }
            if ffi::PyUnicode_CheckExact(ptr) != 0 {
                let s: &Bound<'a, PyString> = elem.cast_unchecked();
                return Ok(ExcelCell::String(Cow::Borrowed(s.to_str()?)));
            }
            if elem.is_none() {
                return Ok(ExcelCell::Blank);
            }
            if ffi::PyBool_Check(ptr) != 0 {
                let b: &Bound<'a, PyBool> = elem.cast_unchecked();
                return Ok(ExcelCell::Bool(b.is_true()));
            }
        }

        if let Ok(dt) = elem.cast::<PyDateTime>() {
            return Ok(ExcelCell::DateTime(pydatetime_xlsx_format(dt)?));
        }
        if let Ok(d) = elem.cast::<PyDate>() {
            return Ok(ExcelCell::DateTime(pydate_xlsx_format(d)?));
        }
        if let Ok(t) = elem.cast::<PyTime>() {
            return Ok(ExcelCell::DateTime(pytime_xlsx_format(t)?));
        }
        if let Ok(t) = elem.cast::<PySequence>() {
            return Ok(ExcelCell::Sequence(pyone_dimensional_iter_xlsx_format(t)?));
        }
        if elem.get_type().name()? == "Decimal" {
            return Ok(ExcelCell::Float(pydecimal_xlsx_format(elem)?));
        }

        Err(PyNotImplementedError::new_err(format!(
            "Unsupported type for Excel export: {}",
            elem.get_type().name()?
        )))
    }

    pub fn from_py_hinted(elem: &'a Bound<'a, PyAny>, hint: Option<ColTypeHint>) -> PyResult<(Self, ColTypeHint)> {
        if let Some(h) = hint {
            if elem.is_none() {
                return Ok((ExcelCell::Blank, ColTypeHint::Blank));
            }
            unsafe {
                match h {
                    ColTypeHint::Float => {
                        let f: &Bound<'a, PyFloat> = elem.cast_unchecked();
                        return Ok((ExcelCell::Float(f.value()), ColTypeHint::Float));
                    }
                    ColTypeHint::Int => {
                        let i: XlsxInt = elem.extract()?;
                        return Ok((ExcelCell::Int(i), ColTypeHint::Int));
                    }
                    ColTypeHint::String => {
                        let s: &Bound<'a, PyString> = elem.cast_unchecked();
                        return Ok((ExcelCell::String(Cow::Borrowed(s.to_str()?)), ColTypeHint::String));
                    }
                    ColTypeHint::Bool => {
                        let b: &Bound<'a, PyBool> = elem.cast_unchecked();
                        return Ok((ExcelCell::Bool(b.is_true()), ColTypeHint::Bool));
                    }
                    _ => {}
                }
            }
        }

        let cell = Self::from_py(elem)?;
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
            (ExcelCell::DateTime(dt), None) => worksheet.write_datetime(row, col, dt),
            (ExcelCell::Sequence(s), None) => worksheet.write_string(row, col, s),
            
            (ExcelCell::String(s), Some(fmt)) => worksheet.write_string_with_format(row, col, s.as_ref(), fmt),
            (ExcelCell::Blank, Some(fmt)) => worksheet.write_blank(row, col, fmt),
            (ExcelCell::Int(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Float(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Bool(b), Some(fmt)) => worksheet.write_boolean_with_format(row, col, *b, fmt),
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

#[pyclass(from_py_object, weakref)]
#[derive(Clone)]
pub struct XIOWWorksheet {
    worksheet: *mut Worksheet,
    col_hints: Vec<ColTypeHint>,
    col_formats_setted: Vec<bool>,
    options: XIOWOptions,

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

    fn _write_cell_rs(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let colidx = col as usize;
        let cell: ExcelCell;
        if self.options.cache_typehints_write_optimization {
            let cached_hint = self.col_hints.get(colidx).copied();

            let hint: ColTypeHint;
            (cell, hint) = ExcelCell::from_py_hinted(value, cached_hint)?;
            if colidx >= self.col_hints.len() {
                self.col_hints.resize(colidx + 1, ColTypeHint::Unknown);
            }
            self.col_hints[colidx] = hint;
        } else {
            cell = ExcelCell::from_py(value)?;
        }
        
        let worksheet: &mut Worksheet;
        let mut cell_format: Option<&Format> = format;
        
        if self.options.cache_col_formats {
            if colidx >= self.col_formats_setted.len() {
                self.col_formats_setted.resize(colidx + 1, false);
            }

            let mut fmtflag = self.col_formats_setted[colidx];
            if !fmtflag && format.is_some() {
                fmtflag = true;
                self.col_formats_setted[colidx] = true;
                worksheet = self.worksheet_refmut();
                worksheet
                    .set_column_format(col, format.unwrap())
                    .map_err(|e| PyValueError::new_err(e.to_string()))?;
            } else {
                worksheet = self.worksheet_refmut();
            }
            if fmtflag {
                cell_format = None;
            }
        } else {
            worksheet = self.worksheet_refmut();
        }

        cell.write(worksheet, row, col, cell_format)
    }

    pub fn internal_new(worksheet: &mut Worksheet, name: String, options: XIOWOptions) -> Self {
        let _ = worksheet.set_name(name).map_err(|e| PyValueError::new_err(e.to_string()));
        Self {
            worksheet: worksheet as *mut Worksheet, 
            col_hints: Vec::new(),
            col_formats_setted: Vec::new(),
            options: options,
        }
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

        // PyList directly
        if let Ok(list) = value.cast::<PyList>() {
            for (offset, item) in list.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, &item, rs_format)?;
            }
            return Ok(());
        }

        // PyTuple directly
        if let Ok(tuple) = value.cast::<PyTuple>() {
            for (offset, item) in tuple.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, &item, rs_format)?;
            }
            return Ok(());
        }

        let iter_result= if let Ok(dict) = value.cast::<PyDict>() {
            dict.values().try_iter()
        } else {
            value.try_iter()
        }.map_err(|e| {
            PyValueError::new_err(format!("Cannot write row from invalid object: {}", e))
        })?;

        for (offset, item) in iter_result.enumerate() {
            extract_format_by_offset!(rs_formats, offset, rs_format);
            self._write_cell_rs(row, col + offset as ColNum, &item.unwrap(), rs_format)?;
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
        unpack_formats!(formats, rs_formats);
        
        let rows_iter = value.try_iter()?;
        for (r_offset, row_obj_res) in rows_iter.enumerate() {
            let row_obj = row_obj_res?;
            let current_row = row + r_offset as RowNum;

            for (c_offset, cell_obj_res) in row_obj.try_iter()?.enumerate() {
                let cell_obj = cell_obj_res?;
                let current_col = col + c_offset as ColNum;

                extract_format_by_offset!(rs_formats, c_offset, rs_format);
                self._write_cell_rs(current_row, current_col, &cell_obj, rs_format)?;
            }
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

        unpack_format!(format, rs_format);
        let iter = value.try_iter().map_err(|e| {
            PyValueError::new_err(format!("Cannot write column from invalid object: {}", e))
        })?;


        for (offset, elem_res) in iter.enumerate() {
            let elem = elem_res?;
            if elem.is_none() {
                continue;
            }

            self._write_cell_rs(row + offset as RowNum, col, value, rs_format)?;
        }

        Ok(())
    }

    pub fn write_columns<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PySequence>,
        formats: Option<&Bound<'py, PyList>>,
    ) -> PyResult<()> {
        todo!();
    }

    #[pyo3(signature = (first_row, first_col, last_row, last_col, value, format = None))]
    pub fn merge_range<'py>(
        &mut self,
        py: Python<'py>,
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

        py.detach(|| {
            worksheet.merge_range(first_row, first_col, last_row, last_col, rust_str, rs_format.unwrap_or(&DEFAULT_FORMAT))
        })
        .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }
    
    fn __repr__(&mut self) -> String {
        format!("<XIOWWorksheet \"{}\">", self.name())
    }
}

#[pyclass]
pub struct XIOWWorkbook {
    filepath: Option<String>,
    workbook: Workbook,
    worksheets: Vec<XIOWWorksheet>,
    pub options: XIOWOptions,
}

impl XIOWWorkbook {
    fn get_sheetnames_string(&mut self) -> String {
        let sheetnames = self.workbook.worksheets().iter()
            .map(|x| format!("\"{}\"", x.name()))
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
            filepath: filepath,
            workbook: Workbook::new(),
            worksheets: Vec::new(),
            options: options.unwrap_or(XIOWOptions::default()),
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

    #[pyo3(signature = (properties))]
    fn add_format<'py>(&self, properties: &Bound<'py, PyDict>) -> PyResult<XIOFormat> {
        XIOFormat::from_properties(properties)
    }

    #[pyo3(signature = (name, options = None))]
    fn add_worksheet(&mut self, name: String, options: Option<XIOWOptions>) -> PyResult<XIOWWorksheet> {
        let ws_options = options.unwrap_or(self.options.clone());
        let worksheet = if ws_options.constant_memory {
            self.workbook.add_worksheet_with_constant_memory()
        } else {
            self.workbook.add_worksheet()
        };

        let sheet = XIOWWorksheet::internal_new(worksheet, name, ws_options);
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
            .find(|ws| ws.name() == name) // Используем имя из сохраненной структуры
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
