//! Common types and utilities

use pyo3::exceptions::{PyException, PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyModule;
use std::path::PathBuf;

/// Registers common types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<FileFormat>()?;
    m.add_class::<RGBColor>()?;
    m.add_class::<Length>()?;
    m.add_function(wrap_pyfunction!(detect_file_format, m)?)?;
    m.add_function(wrap_pyfunction!(detect_file_format_from_bytes, m)?)?;
    Ok(())
}

/// Converts a Rust litchi::Error to a Python exception
pub fn to_py_err(err: litchi::Error) -> PyErr {
    match err {
        litchi::Error::Io(e) => PyIOError::new_err(e.to_string()),
        litchi::Error::InvalidFormat(msg) => PyValueError::new_err(msg),
        litchi::Error::Unsupported(msg) => PyValueError::new_err(msg),
        _ => PyException::new_err(err.to_string()),
    }
}

/// Converts a boxed error to a Python exception
pub fn boxed_err_to_py_err(err: Box<dyn std::error::Error>) -> PyErr {
    PyException::new_err(err.to_string())
}

/// File format enumeration
///
/// Represents the different Office file formats supported by Litchi.
#[pyclass]
#[derive(Clone, Copy, Debug)]
pub enum FileFormat {
    /// Microsoft Word 97-2003 (.doc)
    Doc,
    /// Microsoft Word 2007+ (.docx)
    Docx,
    /// Microsoft PowerPoint 97-2003 (.ppt)
    Ppt,
    /// Microsoft PowerPoint 2007+ (.pptx)
    Pptx,
    /// Microsoft Excel 97-2003 (.xls)
    Xls,
    /// Microsoft Excel 2007+ (.xlsx)
    Xlsx,
    /// Microsoft Excel Binary 2007+ (.xlsb)
    Xlsb,
    /// OpenDocument Text (.odt)
    Odt,
    /// OpenDocument Spreadsheet (.ods)
    Ods,
    /// OpenDocument Presentation (.odp)
    Odp,
    /// Apple Pages (.pages)
    Pages,
    /// Apple Keynote (.key)
    Keynote,
    /// Apple Numbers (.numbers)
    Numbers,
    /// Rich Text Format (.rtf)
    Rtf,
}

impl From<litchi::FileFormat> for FileFormat {
    fn from(fmt: litchi::FileFormat) -> Self {
        match fmt {
            litchi::FileFormat::Doc => FileFormat::Doc,
            litchi::FileFormat::Docx => FileFormat::Docx,
            litchi::FileFormat::Ppt => FileFormat::Ppt,
            litchi::FileFormat::Pptx => FileFormat::Pptx,
            litchi::FileFormat::Xls => FileFormat::Xls,
            litchi::FileFormat::Xlsx => FileFormat::Xlsx,
            litchi::FileFormat::Xlsb => FileFormat::Xlsb,
            litchi::FileFormat::Odt => FileFormat::Odt,
            litchi::FileFormat::Ods => FileFormat::Ods,
            litchi::FileFormat::Odp => FileFormat::Odp,
            litchi::FileFormat::Pages => FileFormat::Pages,
            litchi::FileFormat::Keynote => FileFormat::Keynote,
            litchi::FileFormat::Numbers => FileFormat::Numbers,
            litchi::FileFormat::Rtf => FileFormat::Rtf,
        }
    }
}

#[pymethods]
impl FileFormat {
    /// Returns the string representation of the format
    fn __str__(&self) -> &'static str {
        match self {
            FileFormat::Doc => "Doc",
            FileFormat::Docx => "Docx",
            FileFormat::Ppt => "Ppt",
            FileFormat::Pptx => "Pptx",
            FileFormat::Xls => "Xls",
            FileFormat::Xlsx => "Xlsx",
            FileFormat::Xlsb => "Xlsb",
            FileFormat::Odt => "Odt",
            FileFormat::Ods => "Ods",
            FileFormat::Odp => "Odp",
            FileFormat::Pages => "Pages",
            FileFormat::Keynote => "Keynote",
            FileFormat::Numbers => "Numbers",
            FileFormat::Rtf => "Rtf",
        }
    }

    fn __repr__(&self) -> String {
        format!("FileFormat.{}", self.__str__())
    }
}

/// RGB color representation
///
/// Represents a color in RGB format with values from 0-255.
#[pyclass]
#[derive(Clone, Debug)]
pub struct RGBColor {
    inner: litchi::RGBColor,
}

#[pymethods]
impl RGBColor {
    /// Create a new RGB color
    ///
    /// Args:
    ///     r: Red component (0-255)
    ///     g: Green component (0-255)
    ///     b: Blue component (0-255)
    #[new]
    fn new(r: u8, g: u8, b: u8) -> Self {
        RGBColor {
            inner: litchi::RGBColor::new(r, g, b),
        }
    }

    /// Red component (0-255)
    #[getter]
    fn r(&self) -> u8 {
        self.inner.r
    }

    /// Green component (0-255)
    #[getter]
    fn g(&self) -> u8 {
        self.inner.g
    }

    /// Blue component (0-255)
    #[getter]
    fn b(&self) -> u8 {
        self.inner.b
    }

    fn __str__(&self) -> String {
        format!("RGB({}, {}, {})", self.r(), self.g(), self.b())
    }

    fn __repr__(&self) -> String {
        format!("RGBColor({}, {}, {})", self.r(), self.g(), self.b())
    }
}

/// Length with units
///
/// Represents a measurement with associated units (EMUs, points, inches, etc.).
#[pyclass]
#[derive(Clone, Debug)]
pub struct Length {
    inner: litchi::Length,
}

#[pymethods]
impl Length {
    /// Create a length from EMUs (English Metric Units)
    ///
    /// Args:
    ///     emus: Length in EMUs (914400 EMUs = 1 inch)
    #[staticmethod]
    fn from_emus(emus: i64) -> Self {
        Length {
            inner: litchi::Length::from_emus(emus),
        }
    }

    /// Create a length from points
    ///
    /// Args:
    ///     points: Length in points (72 points = 1 inch)
    #[staticmethod]
    fn from_points(points: f64) -> Self {
        Length {
            inner: litchi::Length::from_inches(points / 72.0),
        }
    }

    /// Create a length from inches
    ///
    /// Args:
    ///     inches: Length in inches
    #[staticmethod]
    fn from_inches(inches: f64) -> Self {
        Length {
            inner: litchi::Length::from_inches(inches),
        }
    }

    /// Convert to EMUs
    fn to_emus(&self) -> i64 {
        self.inner.emus()
    }

    /// Convert to points
    fn to_points(&self) -> f64 {
        self.inner.points()
    }

    /// Convert to inches
    fn to_inches(&self) -> f64 {
        self.inner.inches()
    }

    fn __str__(&self) -> String {
        format!("{:.2} pt", self.to_points())
    }

    fn __repr__(&self) -> String {
        format!("Length({} EMUs)", self.to_emus())
    }
}

/// Detect file format from file path
///
/// Args:
///     path: Path to the file
///
/// Returns:
///     The detected FileFormat, or None if format cannot be determined
#[pyfunction]
fn detect_file_format(path: PathBuf) -> PyResult<Option<FileFormat>> {
    match litchi::detect_file_format(&path) {
        Some(fmt) => Ok(Some(FileFormat::from(fmt))),
        None => Ok(None),
    }
}

/// Detect file format from bytes
///
/// Args:
///     data: File content as bytes
///
/// Returns:
///     The detected FileFormat, or None if format cannot be determined
#[pyfunction]
fn detect_file_format_from_bytes(data: &[u8]) -> Option<FileFormat> {
    litchi::detect_file_format_from_bytes(data).map(FileFormat::from)
}
