//! Sheet/Workbook API bindings

use pyo3::prelude::*;
use pyo3::types::PyModule;
use std::path::PathBuf;
use std::sync::Arc;

use crate::common::boxed_err_to_py_err;

/// Registers sheet types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;
    Ok(())
}

/// Excel workbook interface
///
/// Provides support for Excel workbooks in various formats (.xls, .xlsx, .xlsb).
/// The format is automatically detected when opening a file.
///
/// # Examples
///
/// ```python
/// from litchi_py import Workbook
///
/// # Open an Excel workbook
/// wb = Workbook.open("workbook.xlsx")
///
/// # Get worksheet count
/// print(f"Worksheets: {wb.worksheet_count()}")
///
/// # Access worksheet names
/// for name in wb.worksheet_names():
///     print(f"Sheet: {name}")
/// ```
#[pyclass(unsendable)]
pub struct Workbook {
    inner: Arc<litchi::sheet::Workbook>,
}

#[pymethods]
impl Workbook {
    /// Open an Excel workbook from a file path
    ///
    /// The file format (.xls, .xlsx, .xlsb, .ods, .numbers) is automatically detected.
    ///
    /// Args:
    ///     path: Path to the workbook file
    ///
    /// Returns:
    ///     Workbook instance
    ///
    /// Raises:
    ///     IOError: If the file cannot be read
    ///     ValueError: If the file format is invalid or unsupported
    #[staticmethod]
    fn open(path: PathBuf) -> PyResult<Self> {
        let wb = litchi::sheet::Workbook::open(path).map_err(boxed_err_to_py_err)?;
        Ok(Workbook {
            inner: Arc::new(wb),
        })
    }

    /// Get the number of worksheets in the workbook
    ///
    /// Returns:
    ///     Number of worksheets
    fn worksheet_count(&self) -> PyResult<usize> {
        self.inner.worksheet_count().map_err(boxed_err_to_py_err)
    }

    /// Get all worksheet names
    ///
    /// Returns:
    ///     List of worksheet names
    fn worksheet_names(&self) -> PyResult<Vec<String>> {
        self.inner.worksheet_names().map_err(boxed_err_to_py_err)
    }

    /// Extract all text from all worksheets
    ///
    /// Returns:
    ///     All text content as a single string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(boxed_err_to_py_err)
    }

    fn __repr__(&self) -> PyResult<String> {
        let count = self.worksheet_count().unwrap_or(0);
        Ok(format!("<Workbook: {} worksheets>", count))
    }
}

/// A worksheet in a workbook
///
/// Note: This is a placeholder for future worksheet-level API.
/// Currently, use Workbook.worksheet_names() and Workbook.text() for data access.
#[pyclass]
pub struct Worksheet {
    _private: (),
}

#[pymethods]
impl Worksheet {
    fn __repr__(&self) -> String {
        "<Worksheet>".to_string()
    }
}
