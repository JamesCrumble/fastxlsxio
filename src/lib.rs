mod read;
mod write;

use pyo3::prelude::*;

#[pyfunction]
pub fn version() -> PyResult<String> {
    Ok(env!("CARGO_PKG_VERSION").to_string())
}

#[pymodule]
fn fastxlsxio(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(version, m)?)?;

    m.add_class::<write::writer::XIOWWorkbook>()?;
    m.add_class::<write::writer::XIOWWorksheet>()?;

    m.add_class::<read::types::DType>()?;
    m.add_class::<read::types::DShape>()?;
    m.add_class::<read::types::RangeInfo>()?;
    m.add_class::<read::reader::XIORWorkbook>()?;
    m.add_class::<read::reader::XIORWorksheet>()?;
    m.add_function(wrap_pyfunction!(read::reader::read_many, m)?)?;
    m.add_function(wrap_pyfunction!(read::addr_to_idx, m)?)?;
    m.add_function(wrap_pyfunction!(read::idx_to_addr, m)?)?;
    Ok(())
}
