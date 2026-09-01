use pyo3::exceptions::{PyFileExistsError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDate, PyDateTime, PyFloat, PyInt, PyList, PyTime, PyTuple};
use rust_xlsxwriter::{Workbook, Worksheet};
use crate::pyconv::*;

pub fn write_cell_direct<'py>(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'py, PyAny>
) -> PyResult<()> {
    if let Ok(s) = value.extract::<&str>() {
        worksheet
            .write_string(row, col, s)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyFloat>() {
        let val: f64 = value.extract()?;
        worksheet
            .write_number(row, col, val)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyBool>() {
        let val = pybool_xlsx_format(value)?;
        worksheet
            .write_string(row, col, val)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyInt>() {
        let val: f64 = value.extract()?;
        worksheet
            .write_number(row, col, val)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyDateTime>() {
        let excel_dt = pydatetime_xlsx_format(value)?;
        worksheet
            .write_datetime(row, col, &excel_dt)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyDate>() {
        let excel_dt = pydate_xlsx_format(value)?;
        worksheet
            .write_datetime(row, col, &excel_dt)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyTime>() {
        let excel_dt = pytime_xlsx_format(value)?;
        worksheet
            .write_datetime(row, col, &excel_dt)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.get_type().name()? == "Decimal" {
        let val = pydecimal_xlsx_format(value)?;
        worksheet
            .write_number(row, col, val)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else if value.is_instance_of::<PyList>() || value.is_instance_of::<PyTuple>() {
        let buffer = pyone_dimensional_iter_xlsx_format(value)?;
        worksheet
            .write_string(row, col, &buffer)
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

    } else {
        let s = value.str()?;
        worksheet
            .write_string(row, col, s.to_string_lossy())
            .map_err(|e| PyValueError::new_err(e.to_string()))?;
    }

    Ok(())
}


/// Write-only worksheet class
#[pyclass]
#[derive(Clone)]
pub struct XIOWorksheet {
    worksheet: *mut Worksheet,
    is_constant_memory: bool,
}
unsafe impl Send for XIOWorksheet {}
unsafe impl Sync for XIOWorksheet {}

impl XIOWorksheet {
    fn worksheet_deref(&mut self) -> &mut Worksheet {
        unsafe { &mut *self.worksheet }
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
    pub fn name(&mut self) -> String {
        self.worksheet_deref().name()
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_cell<'py>(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
    ) -> PyResult<()> {
        write_cell_direct(self.worksheet_deref(), row, col, value)
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_row<'py>(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
    ) -> PyResult<()> {
        let seq = value.try_iter().map_err(|e| PyValueError::new_err(format!("Cannot write row from invalid object. {}", e)))?;
        for (offset, element) in seq.enumerate() {
            self.write_cell(row, col + offset as u16, &element?)?;
        }
        Ok(())
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_rows<'py> (
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
    ) -> PyResult<()> {
        // TODO: Column type prediction
        let rows = value.try_iter().map_err(|e| {
            PyValueError::new_err(format!("Cannot write batched rows from invalid object. {}", e))
        })?;

        for (r_offset, row_obj) in rows.enumerate() {
            self.write_row(row + r_offset as u32, col, &row_obj?)?;
        }

        Ok(())
    }

    #[pyo3(signature = (row, col, value))]
    pub fn write_column<'py>(
        &mut self,
        row: u32,
        col: u16,
        value: &Bound<'py, PyAny>,
    ) -> PyResult<()> {
        // Columnar writes are incompatible with constant_memory mode in rust_xlsxwriter
        if self.is_constant_memory {
            return Err(PyRuntimeError::new_err(
                "Cannot write columnar data in constant_memory worksheet mode",
            ));
        }

        // Collect sequence references into a vector (zero-copy for underlying Python data)
        let seq: Vec<Bound<'py, PyAny>> = value
            .try_iter()
            .map_err(|e| {
                PyValueError::new_err(format!("Cannot write column from invalid object: {}", e))
            })?
            .collect::<PyResult<_>>()?;

        if seq.is_empty() {
            return Ok(());
        }

        // Inspect the first non-None element to speculate and prefetch column type
        let sample = seq.iter().find(|elem| !elem.is_none());
        let worksheet = self.worksheet_deref();

        if let Some(first) = sample {
            // Local macro generating a specialized, monomorphic fast loop
            macro_rules! run_fast_column_loop {
                ($elem_var:ident, $r_var:ident, $check:expr, $write_expr:expr) => {{
                    for (offset, elem) in seq.iter().enumerate() {
                        let $r_var = row + offset as u32;
                        if elem.is_none() {
                            continue;
                        }
                        let $elem_var = elem;
                        if $check {
                            let _ = $write_expr?;
                        } else {
                            // Fallback to full type resolution if an unexpected type appears mid-column
                            write_cell_direct(worksheet, $r_var, col, $elem_var)?;
                        }
                    }
                }};
            }

            // Speculative type dispatch with explicit return on match
            if first.extract::<&str>().is_ok() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.extract::<&str>().is_ok(),
                    worksheet
                        .write_string(r, col, elem.extract::<&str>()?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyFloat>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyFloat>(),
                    worksheet
                        .write_number(r, col, elem.extract::<f64>()?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyBool>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyBool>(),
                    worksheet
                        .write_string(r, col, pybool_xlsx_format(elem)?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyInt>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyInt>() && !elem.is_instance_of::<PyBool>(),
                    worksheet
                        .write_number(r, col, elem.extract::<f64>()?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyDateTime>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyDateTime>(),
                    worksheet
                        .write_datetime(r, col, &pydatetime_xlsx_format(elem)?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyDate>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyDate>() && !elem.is_instance_of::<PyDateTime>(),
                    worksheet
                        .write_datetime(r, col, &pydate_xlsx_format(elem)?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.is_instance_of::<PyTime>() {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.is_instance_of::<PyTime>(),
                    worksheet
                        .write_datetime(r, col, &pytime_xlsx_format(elem)?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            } else if first.get_type().name()? == "Decimal" {
                run_fast_column_loop!(
                    elem,
                    r,
                    elem.get_type()
                        .name()
                        .map(|n| n == "Decimal")
                        .unwrap_or(false),
                    worksheet
                        .write_number(r, col, pydecimal_xlsx_format(elem)?)
                        .map_err(|err| PyValueError::new_err(err.to_string()))
                );
                return Ok(());
            }
        }

        // Fallback loop: executes only if the column consists exclusively of `None` values
        // or contains unoptimized composite types (e.g., lists, tuples, custom Python objects)
        for (offset, elem) in seq.iter().enumerate() {
            write_cell_direct(worksheet, row + offset as u32, col, elem)?;
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

/// Write-only workbook class
#[pyclass]
pub struct XIOWorkbook {
    workbook: Workbook,
    worksheets: Vec<Py<XIOWorksheet>>,
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

    #[pyo3(signature = (name, constant_memory = false))]
    fn add_worksheet(&mut self, py: Python<'_>, name: String, constant_memory: bool) -> PyResult<Py<XIOWorksheet>> {
        let worksheet;
        if constant_memory {
            worksheet = self.workbook.add_worksheet_with_constant_memory();
        } else {
            worksheet = self.workbook.add_worksheet();
        }

        let sheet_inst = XIOWorksheet::internal_new(worksheet, name.clone(), constant_memory);
        let py_worksheet = Py::new(
            py,
            sheet_inst
        ).unwrap();
        self.worksheets.push(py_worksheet);
        self.worksheets.last().map(|ws| ws.clone_ref(py)).ok_or(PyValueError::new_err(format!(
            "Cannot create worksheet with name \"{}\". ",
            name,
        )))
    }

    fn get_by_idx(&self, py: Python<'_>, idx: usize) -> PyResult<Py<XIOWorksheet>> {
        self.worksheets
        .get(idx).map(|ws| ws.clone_ref(py))
        .ok_or(PyValueError::new_err(format!(
            "Worksheet at index {} not found. Total worksheets available: {}",
            idx,
            self.worksheets.len()
        )))
    }

    fn get_by_name(&mut self, py: Python<'_>, name: String) -> PyResult<Py<XIOWorksheet>> {
        for worksheet in self.worksheets.iter() {
            if worksheet.bind(py).borrow_mut().name() == name {
                return Ok(worksheet.clone_ref(py));
            }
        }

        Err(PyValueError::new_err(format!(
            "Worksheet with name \"{}\" not found. Available worksheets: {}",
            name,
            self.get_sheetnames_string()
        )))

    }

    fn save(&mut self, path: String) -> PyResult<()> {
        self.workbook.save(path).map_err(|e| PyFileExistsError::new_err(e.to_string()))
    }

    #[getter]
    fn sheetnames(&mut self) -> PyResult<Vec<String>> {
        Ok(self.workbook.worksheets().iter().map(|s| s.name()).collect())
    }

    fn __repr__(&mut self) -> String {
        format!(
            "<XIOWorkbook(sheetnames={})>",
            self.get_sheetnames_string()
        )
    }
}
