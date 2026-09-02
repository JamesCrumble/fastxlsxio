use crate::pyconv::*;

use std::borrow::Cow;
use std::sync::LazyLock;

use pyo3::exceptions::{PyFileExistsError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyIterator, PyList, PySequence, PyString, PyTime};
use rust_xlsxwriter::{Workbook, Worksheet, Format, ExcelDateTime};

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
        row: u32,
        col: u16,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let res: Result<_, _> = match (self, format) {
            (ExcelCell::String(s), Some(fmt)) => worksheet.write_string_with_format(row, col, s.as_ref(), fmt),
            (ExcelCell::String(s), None) => worksheet.write_string(row, col, s.as_ref()),

            (ExcelCell::Int(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Int(n), None) => worksheet.write_number(row, col, *n),

            (ExcelCell::Float(n), Some(fmt)) => worksheet.write_number_with_format(row, col, *n, fmt),
            (ExcelCell::Float(n), None) => worksheet.write_number(row, col, *n),

            (ExcelCell::Bool(b), Some(fmt)) => worksheet.write_boolean_with_format(row, col, *b, fmt),
            (ExcelCell::Bool(b), None) => worksheet.write_boolean(row, col, *b),

            (ExcelCell::DateTime(dt), Some(fmt)) => worksheet.write_datetime_with_format(row, col, dt, fmt),
            (ExcelCell::DateTime(dt), None) => worksheet.write_datetime(row, col, dt),

            (ExcelCell::Blank, Some(fmt)) => worksheet.write_blank(row, col, fmt),
            (ExcelCell::Blank, None) => Ok(worksheet),
        };

        res.map_err(|e| PyValueError::new_err(e.to_string()))?;
        Ok(())
    }
}

pub fn write_cell_direct<'py>(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'py, PyAny>,
    format: Option<&Bound<'py, XIOFormat>>,
) -> PyResult<()> {
    let format_guard = format.map(|f| f.borrow());
    let rs_format = format_guard.as_ref().map(|g| &g.rs_format);
    
    let cell = ExcelCell::from_py(value)?;
    cell.write(worksheet, row, col, rs_format)
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

    /// Strict parser: accepts PyList containing XIOFormat objects or None.
    /// Raises PyTypeError immediately if any non-None item fails to downcast to XIOFormat.
    fn parse_column_formats<'py>(&self, formats: Option<&Bound<'py, PyList>>) -> PyResult<Vec<Option<PyRef<'py, XIOFormat>>>> {
        let Some(list) = formats else {
            return Ok(Vec::new());
        };

        list.iter()
            .map(|item| {
                if item.is_none() {
                    Ok(None)
                } else {
                    Ok(Some(item.downcast::<XIOFormat>()?.borrow()))
                }
            })
            .collect()
    }

    /// Map borrow guards to raw &Format references (supports None slots)
    fn extract_rs_format<'a>(&self, guards: &'a [Option<PyRef<'_, XIOFormat>>]) -> Vec<Option<&'a Format>> {
        guards
            .iter()
            .map(|g| g.as_ref().map(|f| &f.rs_format))
            .collect()
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
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
        format: Option<&Bound<'py, XIOFormat>>,
    ) -> PyResult<()> {
        write_cell_direct(self.worksheet_refmut(), row, col, value, format)
    }

    #[pyo3(signature = (row, col, value, formats = None))]
    pub fn write_row<'py>(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
        formats: Option<&Bound<'py, PyList>>,
    ) -> PyResult<()> {
        let iter_result: Result<Bound<'py, PyIterator>, PyErr>;
        if let Ok(dict) = value.downcast::<PyDict>() {
            iter_result = dict.values().into_any().try_iter()
        } else {
            iter_result = value.try_iter();
        }
        let iter = iter_result.map_err(|e| PyValueError::new_err(format!("Cannot write row from invalid object. {}", e)))?;

        match formats {
            Some(s) => {
                for (offset, element) in iter.enumerate() {
                    let item = s.get_item(offset).map_err(|e|PyValueError::new_err(format!("List of formats should be same length as row itself. {}", e)))?;
                    let format_item = item.downcast::<XIOFormat>()?;
                    self.write_cell(row, col + offset as u16, &element?, Some(format_item))?;
                }
            }
            None => {
                for (offset, element) in iter.enumerate() {
                    self.write_cell(row, col + offset as u16, &element?, None)?;
                }
            }
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, value, formats = None))]
    pub fn write_rows<'py>(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PySequence>,
        formats: Option<&Bound<'py, PyList>>,
    ) -> PyResult<()> {
        let format_guards = self.parse_column_formats(formats)?;
        let rs_formats = self.extract_rs_format(&format_guards);
        
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

        let worksheet = self.worksheet_refmut();
        let full_rows_iter = std::iter::once(Ok(first_row_obj)).chain(rows_iter);

        for (r_offset, row_obj_res) in full_rows_iter.enumerate() {
            let row_obj = row_obj_res?;
            let current_row = row + r_offset as u32;

            for (c_offset, cell_obj_res) in row_obj.try_iter()?.enumerate() {
                let cell_obj = cell_obj_res?;
                let current_col = col + c_offset as u16;

                let cell_format = rs_formats.get(c_offset).copied().flatten();
                let hint = col_hints.get(c_offset).and_then(|h| h.as_ref());

                let cell = match hint {
                    Some(h) => ExcelCell::from_py_speculative(&cell_obj, h)?,
                    None => ExcelCell::from_py(&cell_obj)?,
                };

                cell.write(worksheet, current_row, current_col, cell_format)?;
            }
        }

        Ok(())
    }

    #[pyo3(signature = (row, col, value, format = None))]
    pub fn write_column<'py>(
        &mut self,
        row: u32,
        col: u16,
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

        // Извлекаем первый элемент для спекуляции без полной сборки вектора
        let Some(first_elem_res) = iter.next() else {
            return Ok(());
        };
        let first_elem = first_elem_res?;

        let hint = if !first_elem.is_none() {
            ExcelCell::from_py(&first_elem).ok()
        } else {
            None
        };

        let worksheet = self.worksheet_refmut();
        let full_iter = std::iter::once(Ok(first_elem.clone())).chain(iter);

        // Стриминговая запись элемента за элементом
        for (offset, elem_res) in full_iter.enumerate() {
            let elem = elem_res?;
            if elem.is_none() {
                continue;
            }

            let cell = match &hint {
                Some(sample) => ExcelCell::from_py_speculative(&elem, sample)?,
                None => ExcelCell::from_py(&elem)?,
            };

            cell.write(worksheet, row + offset as u32, col, None)?;
        }

        Ok(())
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_matrix<'py>(
        &mut self,
        row: u32,
        col: u16,
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
    fn new() -> Self {
        Self {
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

        // Создаем Rust-структуру XIOWorksheet
        let sheet = XIOWorksheet::internal_new(worksheet, name, constant_memory);

        // Сохраняем копию в Rust-векторе (это не увеличивает Py_REFCNT в Python!)
        self.worksheets.push(sheet.clone());

        // Возвращаем экземпляр в Python
        Ok(sheet)
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

    fn save(&mut self, path: String) -> PyResult<()> {
        self.workbook.save(path).map_err(|e| PyFileExistsError::new_err(e.to_string()))
    }

    fn __repr__(&mut self) -> String {
        format!(
            "<XIOWorkbook(sheetnames={})>",
            self.get_sheetnames_string()
        )
    }
}
