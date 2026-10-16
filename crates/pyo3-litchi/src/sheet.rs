//! Sheet/Workbook API bindings

use std::path::PathBuf;

use litchi::FileFormat as CoreFormat;
use pyo3::prelude::*;
use pyo3::types::PyModule;

use crate::errors::{Failure, Opener};
use crate::iwork;

/// Registers sheet types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;
    Ok(())
}

/// What a `Workbook` reads: a unified file of the workbook family, or a
/// Numbers document, which `litchi::File` opens without a family view.
enum Source {
    File(litchi::File),
    Numbers(litchi::File),
}

/// Excel workbook interface
///
/// Provides support for workbooks in various formats (.xls, .xlsx, .xlsb,
/// .ods, .numbers). The format is detected from the file's content when it
/// is opened.
///
/// A Workbook is immutable and can be shared between threads. Opening and
/// the reads release the GIL while they run.
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
#[pyclass(frozen, module = "litchi_py")]
pub struct Workbook {
    source: Source,
}

impl Workbook {
    fn read(path: &std::path::Path) -> Result<Source, Failure> {
        match litchi::open(path) {
            // `litchi::File` opens Numbers packages and package directories
            // (records 0296 and 0306), but its family views do not cover
            // them; this reader projects the package the file retains.
            Ok(file) if file.format() == Some(CoreFormat::Numbers) => Ok(Source::Numbers(file)),
            Ok(file) => match file.workbook() {
                Ok(_) => Ok(Source::File(file)),
                Err(error) => Err(Failure::open(&error, Opener::Workbook)),
            },
            Err(error) => Err(Failure::open(&error, Opener::Workbook)),
        }
    }

    fn count(&self) -> Result<usize, Failure> {
        match &self.source {
            Source::File(file) => file
                .workbook()
                .and_then(|workbook| workbook.worksheet_count())
                .map_err(|error| Failure::read(&error)),
            Source::Numbers(file) => Ok(iwork::numbers(file)?.sheets().len()),
        }
    }
}

#[pymethods]
impl Workbook {
    /// Open an Excel workbook from a file path
    ///
    /// The format (.xls, .xlsx, .xlsb, .ods, .numbers) is detected from the
    /// file's content, not from its name.
    ///
    /// Args:
    ///     path: Path to the workbook file
    ///
    /// Returns:
    ///     Workbook instance
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file cannot be read
    ///     UnsupportedError: (a ValueError) If the file is not a workbook,
    ///         is not a recognized format, or uses a variant that cannot be
    ///         read (flat OpenDocument)
    ///     CorruptFileError: (a ValueError) If the file is malformed
    ///     ProtectedError: (a ValueError) If the workbook is encrypted
    #[staticmethod]
    fn open(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let source = py.detach(|| Self::read(&path))?;
        Ok(Workbook { source })
    }

    /// Get the number of worksheets in the workbook
    ///
    /// Returns:
    ///     Number of worksheets
    fn worksheet_count(&self, py: Python<'_>) -> PyResult<usize> {
        py.detach(|| self.count()).map_err(PyErr::from)
    }

    /// Get all worksheet names
    ///
    /// Returns:
    ///     List of worksheet names
    fn worksheet_names(&self, py: Python<'_>) -> PyResult<Vec<String>> {
        py.detach(|| match &self.source {
            Source::File(file) => file
                .workbook()
                .and_then(|workbook| workbook.worksheet_names())
                .map_err(|error| Failure::read(&error)),
            Source::Numbers(file) => Ok(iwork::numbers(file)?
                .sheets()
                .iter()
                .map(|sheet| sheet.name().to_owned())
                .collect()),
        })
        .map_err(PyErr::from)
    }

    /// Extract all text from all worksheets
    ///
    /// Returns:
    ///     All text content as a single string
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file changed after it was opened
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        py.detach(|| match &self.source {
            Source::File(file) => file.text().map_err(|error| Failure::read(&error)),
            Source::Numbers(file) => file.text().map_err(|error| Failure::read(&error)),
        })
        .map_err(PyErr::from)
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let count = py.detach(|| self.count()).unwrap_or(0);
        Ok(format!("<Workbook: {} worksheets>", count))
    }
}

/// A worksheet in a workbook
///
/// Note: This is a placeholder for future worksheet-level API. No method
/// returns one yet; use Workbook.worksheet_names() and Workbook.text() for
/// data access.
#[pyclass(frozen, module = "litchi_py")]
pub struct Worksheet {
    _private: (),
}

#[pymethods]
impl Worksheet {
    fn __repr__(&self) -> String {
        "<Worksheet>".to_string()
    }
}
