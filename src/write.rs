use crate::pyconv::*;

use std::borrow::Cow;
use std::sync::LazyLock;

use pyo3::exceptions::{PyFileExistsError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyTuple, PyList, PySequence, PyString, PyTime};
use rust_xlsxwriter::{Workbook, Worksheet, Format, ExcelDateTime, RowNum, ColNum};

pub static DEFAULT_FORMAT: LazyLock<Format> = LazyLock::new(Format::default);

pub enum ExcelCell<'a> {
    String(Cow<'a, str>),
    Float(XlsxFloat),
    Int(XlsxInt),
    Bool(bool),
    DateTime(ExcelDateTime),
    Blank,
}

impl<'a> ExcelCell<'a> {
    pub fn from_py(elem: &'a Bound<'a, PyAny>) -> PyResult<Self> {
        if let Ok(s) = elem.downcast::<PyString>() {
            return Ok(ExcelCell::String(Cow::Borrowed(s.to_str()?)));
        }
        if elem.is_none() {
            return Ok(ExcelCell::Blank);
        }
        if let Ok(f) = elem.downcast::<PyFloat>() {
            return Ok(ExcelCell::Float(f.value()));
        }
        if let Ok(b) = elem.downcast::<PyBool>() {
            return Ok(ExcelCell::Bool(b.is_true()));
        }
        if let Ok(i) = elem.downcast::<PyInt>() {
            let val: XlsxInt = i.extract()?;
            return Ok(ExcelCell::Int(val));
        }
        if let Ok(dt) = elem.downcast::<PyDateTime>() {
            return Ok(ExcelCell::DateTime(pydatetime_xlsx_format(dt)?));
        }
        if let Ok(d) = elem.downcast::<PyDate>() {
            return Ok(ExcelCell::DateTime(pydate_xlsx_format(d)?));
        }
        if let Ok(t) = elem.downcast::<PyTime>() {
            return Ok(ExcelCell::DateTime(pytime_xlsx_format(t)?));
        }
        if elem.get_type().name()? == "Decimal" {
            return Ok(ExcelCell::Float(pydecimal_xlsx_format(elem)?));
        }
        // TODO: support write sequences

        Err(PyValueError::new_err(format!(
            "Unsupported type for Excel export: {}",
            elem.get_type().name()?
        )))
    }

    pub fn from_py_speculative(elem: &'a Bound<'a, PyAny>, hint: &ExcelCell) -> PyResult<Self> {
        if elem.is_none() {
            return Ok(ExcelCell::Blank);
        }

        match hint {
            ExcelCell::String(_) => {
                if let Ok(s) = elem.downcast::<PyString>() {
                    return Ok(ExcelCell::String(Cow::Borrowed(s.to_str()?)));
                }
            }
            ExcelCell::Int(_) => {
                if let Ok(f) = elem.extract::<XlsxInt>() {
                    if !elem.is_instance_of::<PyBool>() {
                        return Ok(ExcelCell::Int(f));
                    }
                }
            }
            ExcelCell::Float(_) => {
                if let Ok(f) = elem.extract::<XlsxFloat>() {
                    if !elem.is_instance_of::<PyBool>() {
                        return Ok(ExcelCell::Float(f));
                    }
                }
            }
            ExcelCell::Bool(_) => {
                if let Ok(b) = elem.downcast::<PyBool>() {
                    return Ok(ExcelCell::Bool(b.is_true()));
                }
            }
            ExcelCell::DateTime(_) => {
                if let Ok(dt) = elem.downcast::<PyDateTime>() {
                    return Ok(ExcelCell::DateTime(pydatetime_xlsx_format(dt)?));
                }
            }
            ExcelCell::Blank => {}
        }

        Self::from_py(elem)
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
            (ExcelCell::Bool(b), None) => worksheet.write_boolean(row, col, *b),
            (ExcelCell::DateTime(dt), None) => worksheet.write_datetime(row, col, dt),

            (ExcelCell::String(s), Some(fmt)) => worksheet.write_string_with_format(row, col, s.as_ref(), fmt),
            (ExcelCell::Blank, Some(fmt)) => worksheet.write_blank(row, col, fmt),
            (ExcelCell::Int(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Float(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Bool(b), Some(fmt)) => worksheet.write_boolean_with_format(row, col, *b, fmt),
            (ExcelCell::DateTime(dt), Some(fmt)) => worksheet.write_datetime_with_format(row, col, dt, fmt),
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
                        vec.push(Some(item.downcast::<XIOFormat>()?.borrow()));
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

#[pyclass(weakref)]
#[derive(Clone)]
pub struct XIOWorksheet {
    worksheet: *mut Worksheet,
    is_constant_memory: bool,
}
unsafe impl Send for XIOWorksheet {}
unsafe impl Sync for XIOWorksheet {}

impl XIOWorksheet {

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
        &self,
        row: RowNum,
        col: ColNum,
        cell: ExcelCell,
        format: Option<&Format>,
    ) -> PyResult<()> {
        cell.write(self.worksheet_refmut(), row, col, format)
    }

    fn _write_cell(
        &self,
        row: RowNum,
        col: ColNum,
        cell: ExcelCell,
        format: Option<&Bound<'_, XIOFormat>>,
    ) -> PyResult<()> {
        unpack_format!(format, rs_format);
        self._write_cell_rs(row, col, cell, rs_format)
    }

    pub fn internal_new(worksheet: &mut Worksheet, name: String, is_constant_memory: bool) -> Self {
        let _ = worksheet.set_name(name);
        Self {
            worksheet: worksheet as *mut Worksheet, is_constant_memory: is_constant_memory
        }
    }
}

#[pymethods]
impl XIOWorksheet {
    
    #[getter]
    pub fn constant_memory(&self) -> bool {
        self.is_constant_memory
    }

    #[getter]
    pub fn name(&self) -> String {
        self.worksheet_ref().name()
    }

    #[pyo3(signature = (row, col, value, format = None))]
    pub fn write_cell<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PyAny>,
        format: Option<&Bound<'py, XIOFormat>>,
    ) -> PyResult<()> {
        self._write_cell(row, col, ExcelCell::from_py(value)?, format)
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
        if let Ok(list) = value.downcast::<PyList>() {
            for (offset, item) in list.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, ExcelCell::from_py(&item)?, rs_format)?;
            }
            return Ok(());
        }

        // PyTuple directly
        if let Ok(tuple) = value.downcast::<PyTuple>() {
            for (offset, item) in tuple.iter().enumerate() {
                extract_format_by_offset!(rs_formats, offset, rs_format);
                self._write_cell_rs(row, col + offset as ColNum, ExcelCell::from_py(&item)?, rs_format)?;
            }
            return Ok(());
        }

        let iter_result= if let Ok(dict) = value.downcast::<PyDict>() {
            dict.values().try_iter()
        } else {
            value.try_iter()
        }.map_err(|e| {
            PyValueError::new_err(format!("Cannot write row from invalid object: {}", e))
        })?;

        for (offset, item) in iter_result.enumerate() {
            extract_format_by_offset!(rs_formats, offset, rs_format);
            self._write_cell_rs(row, col + offset as ColNum, ExcelCell::from_py(&item.unwrap())?, rs_format)?;
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
        
        let mut rows_iter = value.try_iter()?;
        let Some(first_row_res) = rows_iter.next() else {
            return Ok(());
        };
        let first_row_obj = first_row_res?;
        
        // Collect first row only to evaluate column count and hints
        let first_row: Vec<Bound<'py, PyAny>> = first_row_obj
            .try_iter()?
            .collect::<PyResult<_>>()?;

        let col_hints: Vec<Option<ExcelCell>> = first_row
            .iter()
            .map(|cell| ExcelCell::from_py(cell).ok())
            .collect();
    
        unpack_formats!(formats, rs_formats);
        let full_rows_iter = std::iter::once(Ok(first_row_obj)).chain(rows_iter);

        for (r_offset, row_obj_res) in full_rows_iter.enumerate() {
            let row_obj = row_obj_res?;
            let current_row = row + r_offset as RowNum;

            for (c_offset, cell_obj_res) in row_obj.try_iter()?.enumerate() {
                let cell_obj = cell_obj_res?;
                let current_col = col + c_offset as ColNum;

                extract_format_by_offset!(rs_formats, c_offset, rs_format);
                let hint = col_hints.get(c_offset).and_then(|h| h.as_ref());

                let cell = match hint {
                    Some(h) => ExcelCell::from_py_speculative(&cell_obj, h)?,
                    None => ExcelCell::from_py(&cell_obj)?,
                };
                self._write_cell_rs(current_row, current_col, cell, rs_format)?;
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
        if self.is_constant_memory {
            return Err(PyRuntimeError::new_err(
                "Cannot write columnar data in constant_memory worksheet mode",
            ));
        }
        if format.is_some() {
            return Err(PyValueError::new_err(
                "Formatted write not supported for write_column",
            ));
        }

        let mut iter = value.try_iter().map_err(|e| {
            PyValueError::new_err(format!("Cannot write column from invalid object: {}", e))
        })?;

        let Some(first_elem_res) = iter.next() else {
            return Ok(());
        };
        let first_elem = first_elem_res?;

        let hint = if !first_elem.is_none() {
            ExcelCell::from_py(&first_elem).ok()
        } else {
            None
        };

        let full_iter = std::iter::once(Ok(first_elem.clone())).chain(iter);
        for (offset, elem_res) in full_iter.enumerate() {
            let elem = elem_res?;
            if elem.is_none() {
                continue;
            }

            let cell = match &hint {
                Some(sample) => ExcelCell::from_py_speculative(&elem, sample)?,
                None => ExcelCell::from_py(&elem)?,
            };

            self._write_cell_rs(row + offset as RowNum, col, cell, None)?;
        }

        Ok(())
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

        py.allow_threads(|| {
            worksheet.merge_range(first_row, first_col, last_row, last_col, rust_str, rs_format.unwrap_or(&DEFAULT_FORMAT))
        })
        .map_err(|e| PyValueError::new_err(e.to_string()))?;

        Ok(())
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_matrix<'py>(
        &mut self,
        row: RowNum,
        col: ColNum,
        value: &Bound<'py, PyAny>,
    ) -> PyResult<()> {
        todo!();
    }

    fn __repr__(&mut self) -> String {
        format!("<XIOWorksheet \"{}\">", self.name())
    }
}

#[pyclass]
pub struct XIOWorkbook {
    filepath: Option<String>,
    workbook: Workbook,
    worksheets: Vec<XIOWorksheet>,
}

impl XIOWorkbook {
    fn get_sheetnames_string(&mut self) -> String {
        let sheetnames = self.workbook.worksheets().iter()
            .map(|x| format!("\"{}\"", x.name()))
            .collect::<Vec<String>>();
        format!("[{}]", sheetnames.join(", "))
    }
}

#[pymethods]
impl XIOWorkbook {

    #[new]
    #[pyo3(signature = (filepath = None))]
    fn new(filepath: Option<String>) -> Self {
        Self {
            filepath: filepath,
            workbook: Workbook::new(),
            worksheets: Vec::new()
        }
    }

    #[pyo3(signature = (properties))]
    fn add_format<'py>(&self, properties: &Bound<'py, PyDict>) -> PyResult<XIOFormat> {
        XIOFormat::from_properties(properties)
    }

    #[pyo3(signature = (name, constant_memory = false))]
    fn add_worksheet(&mut self, name: String, constant_memory: bool) -> PyResult<XIOWorksheet> {
        let worksheet = if constant_memory {
            self.workbook.add_worksheet_with_constant_memory()
        } else {
            self.workbook.add_worksheet()
        };

        worksheet
            .set_name(&name)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        let sheet = XIOWorksheet::internal_new(worksheet, name, constant_memory);
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
        py.allow_threads(|| {self.workbook.save(path)}).map_err(|e| PyFileExistsError::new_err(e.to_string()))
    }

    fn get_by_idx(&mut self, idx: usize) -> PyResult<XIOWorksheet> {
        let sheet = self
            .worksheets
            .get(idx)
            .ok_or_else(|| PyValueError::new_err(format!("Worksheet at index {} not found", idx)))?;

        Ok(sheet.clone())
    }

    fn get_by_name(&mut self, name: String) -> PyResult<XIOWorksheet> {
        let sheet = self
            .worksheets
            .iter()
            .find(|ws| ws.name() == name) // Используем имя из сохраненной структуры
            .ok_or_else(|| PyValueError::new_err(format!("Worksheet '{}' not found", name)))?;

        Ok(sheet.clone())
    }

    #[getter]
    fn sheetnames(&mut self) -> Vec<String> {
        self.worksheets
            .iter()
            .map(|ws| ws.name())
            .collect()
    }

    fn __repr__(&mut self) -> String {
        format!(
            "<XIOWorkbook(sheetnames={})>",
            self.get_sheetnames_string()
        )
    }
}
