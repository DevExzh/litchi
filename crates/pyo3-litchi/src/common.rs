//! Common types and utilities

use pyo3::prelude::*;
use pyo3::types::PyModule;
use std::borrow::Cow;
use std::path::PathBuf;

use crate::errors::Failure;

/// Registers common types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<FileFormat>()?;
    m.add_class::<RGBColor>()?;
    m.add_class::<Length>()?;
    m.add_function(wrap_pyfunction!(detect_file_format, m)?)?;
    m.add_function(wrap_pyfunction!(detect_file_format_from_bytes, m)?)?;
    Ok(())
}

/// File format enumeration
///
/// Represents the different Office file formats supported by Litchi: one
/// value per format detection can report.
#[pyclass(skip_from_py_object, module = "litchi_py")]
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
    /// OpenDocument Drawing (.odg, .otg)
    Odg,
    /// OpenDocument Chart (.odc, .otc)
    Odc,
    /// OpenDocument Formula (.odf, .otf)
    Odf,
    /// OpenDocument Image (.odi, .oti)
    Odi,
    /// OpenDocument Master Document (.odm)
    Odm,
    /// OpenDocument Web Document (.oth)
    Oth,
    /// Apple Pages (.pages)
    Pages,
    /// Apple Keynote (.key)
    Keynote,
    /// Apple Numbers (.numbers)
    Numbers,
    /// Rich Text Format (.rtf)
    Rtf,
    /// OpenDocument Database Front End (.odb). Last, so the integer values
    /// of the formats before it stay as they were.
    Odb,
}

impl FileFormat {
    /// The Python value of a detected format.
    ///
    /// Every format of this Litchi build has a Python value; a format a
    /// newer Litchi adds before these bindings learn it is reported as an
    /// internal error instead of being given another format's name.
    pub(crate) fn from_core(fmt: litchi::FileFormat) -> Result<Self, Failure> {
        Ok(match fmt {
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
            litchi::FileFormat::Odg => FileFormat::Odg,
            litchi::FileFormat::Odc => FileFormat::Odc,
            litchi::FileFormat::Odf => FileFormat::Odf,
            litchi::FileFormat::Odi => FileFormat::Odi,
            litchi::FileFormat::Odm => FileFormat::Odm,
            litchi::FileFormat::Oth => FileFormat::Oth,
            litchi::FileFormat::Odb => FileFormat::Odb,
            litchi::FileFormat::Pages => FileFormat::Pages,
            litchi::FileFormat::Keynote => FileFormat::Keynote,
            litchi::FileFormat::Numbers => FileFormat::Numbers,
            litchi::FileFormat::Rtf => FileFormat::Rtf,
            other => return Err(Failure::unknown_format(other)),
        })
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
            FileFormat::Odg => "Odg",
            FileFormat::Odc => "Odc",
            FileFormat::Odf => "Odf",
            FileFormat::Odi => "Odi",
            FileFormat::Odm => "Odm",
            FileFormat::Oth => "Oth",
            FileFormat::Odb => "Odb",
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
#[pyclass(frozen, skip_from_py_object, module = "litchi_py")]
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
#[pyclass(frozen, skip_from_py_object, module = "litchi_py")]
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
/// Detection reads the file's content; the file name is not consulted.
/// The GIL is released while the file is read.
///
/// Args:
///     path: Path to the file
///
/// Returns:
///     The detected FileFormat, or None if format cannot be determined
#[pyfunction]
fn detect_file_format(py: Python<'_>, path: PathBuf) -> PyResult<Option<FileFormat>> {
    py.detach(|| litchi::detect_file_format(&path))
        .map(FileFormat::from_core)
        .transpose()
        .map_err(PyErr::from)
}

/// Detect file format from bytes
///
/// The GIL is released while the bytes are inspected.
///
/// Args:
///     data: File content as bytes
///
/// Returns:
///     The detected FileFormat, or None if format cannot be determined
#[pyfunction]
fn detect_file_format_from_bytes(py: Python<'_>, data: &[u8]) -> PyResult<Option<FileFormat>> {
    py.detach(|| litchi::detect_file_format_from_bytes(data))
        .map(FileFormat::from_core)
        .transpose()
        .map_err(PyErr::from)
}

/// A `repr()` preview of `text`: the text itself when it has at most
/// `limit` characters, otherwise its first `limit` characters and `...`.
///
/// The cut is counted in characters, as Python slices a `str`, so it never
/// falls inside a multi-byte character.
pub(crate) fn preview(text: &str, limit: usize) -> Cow<'_, str> {
    match text.char_indices().nth(limit) {
        Some((end, _)) => Cow::Owned(format!("{}...", &text[..end])),
        None => Cow::Borrowed(text),
    }
}

#[cfg(test)]
mod tests {
    use super::preview;

    #[test]
    fn preview_cuts_at_a_character_boundary() {
        assert_eq!(preview("short", 30), "short");
        assert_eq!(preview(&"a".repeat(30), 30), "a".repeat(30));
        assert_eq!(
            preview(&"a".repeat(31), 30),
            format!("{}...", "a".repeat(30))
        );
        // Two-byte letters: a byte cut at 30 would split the 16th letter.
        let russian = "Привет мир, это длинный абзац текста";
        assert_eq!(
            preview(russian, 30),
            format!("{}...", russian.chars().take(30).collect::<String>())
        );
        assert_eq!(preview("😀😀😀", 2), "😀😀...");
    }
}
