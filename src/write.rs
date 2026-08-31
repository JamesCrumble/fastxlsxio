use crate::types::{
    Array1Container, Array2Container, CalamineData, DType, ValueContainer, WrappedValue,
    WriteToSheet,
};
use crate::types::{CellAddr, DShape};
use chrono::{NaiveDate, NaiveDateTime};
use pyo3::exceptions::{PyFileExistsError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

/// Write-only worksheet class
#[pyclass]
#[derive(Clone)]
pub struct XIOWorksheet {
    worksheet: *mut Worksheet,
}
unsafe impl Send for XIOWorksheet {}
unsafe impl Sync for XIOWorksheet {}

impl XIOWorksheet {
    fn worksheet_deref(&mut self) -> &mut Worksheet {
        unsafe { &mut *self.worksheet }
    }

    pub fn internal_new(worksheet: &mut Worksheet, title: String) -> Self {
        let _ = worksheet.set_name(title);
        Self {
            worksheet: worksheet as *mut Worksheet,
        }
    }

    pub fn name(&mut self) -> String {
        self.worksheet_deref().name()
    }

    // pub fn write_to_self(
    //     &mut self,
    //     row: usize,
    //     col: usize,
    //     value: WrappedValue,
    //     is_column: bool,
    // ) -> PyResult<()> {
    //     let row_u32 =
    //         u32::try_from(row).map_err(|_| PyValueError::new_err("Row index out of range"))?;
    //     let col_u16 =
    //         u16::try_from(col).map_err(|_| PyValueError::new_err("Column index out of range"))?;
    //     self.data_to_write
    //         .push(((row_u32, col_u16), value, is_column));
    //     Ok(())
    // }
    // pub fn to_sheet(&self, sheet: &mut Worksheet) -> PyResult<()> {
    //     let _ = sheet
    //         .set_name(&self.title)
    //         .map_err(|e| PyFileExistsError::new_err(e.to_string()))?;
    //     self.data_to_write
    //         .iter()
    //         .try_for_each(|(pos, data, is_column)| {
    //             data.write_to_sheet(sheet, *pos, *is_column)
    //                 .map_err(|e| PyValueError::new_err(e.to_string()))
    //         })
    // }
}
#[pymethods]
impl XIOWorksheet {
    /**
        Write a value to a specific cell in the worksheet.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The cell address, either as a 0-based tuple of (row, col) or a string (e.g., "A1").
        value : Any
            The value to write to the cell.
        dtype : DType, default DType.Any
            The data type to enforce for the value. If None, try for each type automatically.
    */
    #[pyo3(signature = (cell_addr, value, *, dtype = Some(DType::Any)))]
    pub fn write_cell<'py>(
        &mut self,
        cell_addr: CellAddr,
        value: &Bound<'py, PyAny>,
        dtype: Option<DType>,
        // style: Option<f64>,
    ) -> PyResult<()> {
        let (row, col) = cell_addr.as_idx()?;
        let value = extract_scalar!(value, dtype)?;
        let row_u32 =
            u32::try_from(row).map_err(|_| PyValueError::new_err("Row index out of range"))?;
        let col_u16 =
            u16::try_from(col).map_err(|_| PyValueError::new_err("Column index out of range"))?;
        value.write_to_sheet(self.worksheet_deref(), (row_u32, col_u16), false).map_err(|e: XlsxError| pyo3::exceptions::PyValueError::new_err(e.to_string()))?;
        Ok(())
    }
    /**
        Write a row of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[Any]]
            The row of values to write. 1d array-like.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
    */
    #[pyo3(signature = (cell_addr, value, *, dtype = None))]
    pub fn write_row<'py>(
        &mut self,
        cell_addr: CellAddr,
        value: &Bound<'py, PyAny>,
        dtype: Option<DType>,
        // style: Option<f64>,
    ) -> PyResult<()> {
        todo!();
        // let value = extract_array1!(value, dtype)?;
        // let (row, col) = cell_addr.as_idx()?;
        // let data_shape = value.get_shape(false);
        // if !matches!(data_shape, DShape::Row { n_cols: _ }) {
        //     return Err(PyFileExistsError::new_err(
        //         "write_row only accepty 1d-array, please use write_cell for scalar or write_matrix for 2d-array",
        //     ));
        // }
        // self.write_to_self(row, col, value, false)
    }
    /**
        Write a column of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[Any]]
            The column of values to write.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
    */
    #[pyo3(signature = (cell_addr, value, *, dtype = None))]
    pub fn write_column<'py>(
        &mut self,
        cell_addr: CellAddr,
        value: &Bound<'py, PyAny>,
        dtype: Option<DType>,
        // style: Option<f64>,
    ) -> PyResult<()> {
        todo!();
        // let (row, col) = cell_addr.as_idx()?;
        // let value = extract_array1!(value, dtype)?;
        // let data_shape = value.get_shape(true);
        // if !matches!(data_shape, DShape::Column { n_rows: _ }) {
        //     return Err(PyFileExistsError::new_err(
        //         "write_column only accepty 1d-array, please use write_cell for scalar or write_matrix for matrix",
        //     ));
        // }
        // self.write_to_self(row, col, value, true)
    }
    /**
        Write a matrix of values starting from a specific cell.

        Parameters
        ----------
        cell_addr : Union[Tuple[int, int], str]
            The starting cell address, either as a tuple of (row, col) or a string (e.g., "A1").
        value : Union[np.ndarray, List[List[Any]]]
            The matrix of values to write.\n
            MUST be `np.ndarray` with correct dtype if `dtype` is one of [`DType.Bool`, `DType.Int`, `DType.Float`]
        dtype : Optional[DType], default None
            The data type to enforce for the values. If None, will try for every possible types.\n
            DType.Any allow each item have different type, but will also increase the time cost.
    */
    #[pyo3(signature = (cell_addr, value, *, dtype = None))]
    pub fn write_matrix<'py>(
        &mut self,
        cell_addr: CellAddr,
        value: &Bound<'py, PyAny>,
        dtype: Option<DType>,
        // style: Option<f64>,
    ) -> PyResult<()> {
        todo!();
        // let (row, col) = cell_addr.as_idx()?;
        // let value = extract_array2!(value, dtype)?;
        // let data_shape = value.get_shape(false);
        // if !matches!(
        //     data_shape,
        //     DShape::Matrix {
        //         n_rows: _,
        //         n_cols: _
        //     }
        // ) {
        //     return Err(PyFileExistsError::new_err(
        //         "write_matrix only accepty 2d-array, please use write_cell for scalar or write_row/write_column for 1d-array",
        //     ));
        // }
        // self.write_to_self(row, col, value, false)
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

    /**
        Create a new worksheet with the specified name.

        Parameters
        ----------
        name : str
            The name of the new worksheet.

        Returns
        -------
        XIOWorksheet
            The newly created worksheet.
    */
    fn create_sheet(&mut self, py: Python<'_>, title: String) -> PyResult<Py<XIOWorksheet>> {
        let worksheet = self.workbook.add_worksheet();
        let sheet_inst = XIOWorksheet::internal_new(worksheet, title.clone());
        let py_worksheet = Py::new(
            py,
            sheet_inst
        ).unwrap();
        self.worksheets.push(py_worksheet);
        self.worksheets.last().map(|ws| ws.clone_ref(py)).ok_or(PyValueError::new_err(format!(
            "Cannot create worksheet with title \"{}\". ",
            title,
        )))
    }

    /**
        Create a new worksheet with the specified name and constant memory option.

        Parameters
        ----------
        name : str
            The name of the new worksheet.

        Returns
        -------
        XIOWorksheet
            The newly created worksheet.
    */
    fn create_sheet_with_constant_memory(&mut self, py: Python<'_>, title: String) -> PyResult<Py<XIOWorksheet>> {
        let worksheet = self.workbook.add_worksheet_with_constant_memory();
        let sheet_inst = XIOWorksheet::internal_new(worksheet, title.clone());
        let py_worksheet = Py::new(
            py,
            sheet_inst
        ).unwrap();
        self.worksheets.push(py_worksheet);
        self.worksheets.last().map(|ws| ws.clone_ref(py)).ok_or(PyValueError::new_err(format!(
            "Cannot create worksheet with title \"{}\". ",
            title,
        )))
    }

    /**
        Get a worksheet by its index.

        Parameters
        ----------
        idx : int
            The 0-based index of the worksheet.

        Returns
        -------
        XIOWorksheet
            The worksheet at the specified index.
    */
    fn get_by_idx(&self, py: Python<'_>, idx: usize) -> PyResult<Py<XIOWorksheet>> {
        self.worksheets
        .get(idx).map(|ws| ws.clone_ref(py))
        .ok_or(PyValueError::new_err(format!(
            "Worksheet at index {} not found. Total worksheets available: {}",
            idx,
            self.worksheets.len()
        )))
    }
    /**
        Get a worksheet by its name.

        Parameters
        ----------
        name : str
            The name of the worksheet.

        Returns
        -------
        XIOWorksheet
            The worksheet with the specified name.
    */
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
    /**
        Save the workbook to the specified file path.

        Parameters
        ----------
        path : str
            The file path where the workbook will be saved.
    */
    fn save(&mut self, path: String) -> PyResult<()> {
        self.workbook.save(path).map_err(|e| PyFileExistsError::new_err(e.to_string()))
    }
    /**
        Get the names of all worksheets in the workbook.

            Returns
            -------
            List[str]
                A list of worksheet names.
    */
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
