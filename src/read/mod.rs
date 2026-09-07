pub mod fromcell;
pub mod types;
#[macro_use]
pub mod utils;
pub mod reader;
use pyo3::prelude::*;

/**
    Convert a cell address string (e.g., "A1") to a 0-based (row, col) index.

    Parameters
    ----------
    addr : str
        The cell address string (e.g., "A1").

    Returns
    -------
    Tuple[int, int]
        A tuple of (row, col) indices.
*/
#[pyfunction]
pub fn addr_to_idx(addr: String) -> PyResult<(usize, usize)> {
    types::CellAddr::Name(addr).as_idx()
}

/**
    Convert a 0-based (row, col) index to a cell address string (e.g., "A1").

    Parameters
    ----------
    row : int
        The 0-based row index.
    col : int
        The 0-based column index.

    Returns
    -------
    str
        The cell address string (e.g., "A1").
*/
#[pyfunction]
pub fn idx_to_addr(row: usize, col: usize) -> PyResult<String> {
    types::CellAddr::Idx((row, col)).as_addr()
}
