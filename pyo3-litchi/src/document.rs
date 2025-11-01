//! Document API bindings

use pyo3::prelude::*;
use pyo3::types::PyModule;
use std::path::PathBuf;
use std::sync::Arc;

use crate::common::to_py_err;

/// Registers document types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Document>()?;
    m.add_class::<Paragraph>()?;
    m.add_class::<Run>()?;
    m.add_class::<Table>()?;
    m.add_class::<TableRow>()?;
    m.add_class::<TableCell>()?;
    Ok(())
}

/// Unified Word document interface
///
/// Provides format-agnostic interface for both .doc and .docx files.
/// The format is automatically detected when opening a file.
///
/// # Examples
///
/// ```python
/// from litchi_py import Document
///
/// # Open any Word document (.doc or .docx)
/// doc = Document.open("document.docx")
///
/// # Extract all text
/// text = doc.text()
/// print(text)
///
/// # Access paragraphs
/// for para in doc.paragraphs():
///     print(f"Paragraph: {para.text()}")
/// ```
#[pyclass(unsendable)]
pub struct Document {
    inner: Arc<litchi::Document>,
}

#[pymethods]
impl Document {
    /// Open a Word document from a file path
    ///
    /// The file format (.doc or .docx) is automatically detected.
    ///
    /// Args:
    ///     path: Path to the document file
    ///
    /// Returns:
    ///     Document instance
    ///
    /// Raises:
    ///     IOError: If the file cannot be read
    ///     ValueError: If the file format is invalid or unsupported
    #[staticmethod]
    fn open(path: PathBuf) -> PyResult<Self> {
        let doc = litchi::Document::open(path).map_err(to_py_err)?;
        Ok(Document {
            inner: Arc::new(doc),
        })
    }

    /// Extract all text from the document
    ///
    /// Returns:
    ///     All text content as a single string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    /// Get all paragraphs in the document
    ///
    /// Returns:
    ///     List of Paragraph objects
    fn paragraphs(&self) -> PyResult<Vec<Paragraph>> {
        let paras = self.inner.paragraphs().map_err(to_py_err)?;
        Ok(paras
            .into_iter()
            .map(|p| Paragraph { inner: Arc::new(p) })
            .collect())
    }

    /// Get all tables in the document
    ///
    /// Returns:
    ///     List of Table objects
    fn tables(&self) -> PyResult<Vec<Table>> {
        let tables = self.inner.tables().map_err(to_py_err)?;
        Ok(tables
            .into_iter()
            .map(|t| Table { inner: Arc::new(t) })
            .collect())
    }

    fn __repr__(&self) -> String {
        format!("<Document>")
    }
}

/// A paragraph in a document
///
/// Represents a single paragraph with text and formatting.
#[pyclass(unsendable)]
pub struct Paragraph {
    inner: Arc<litchi::document::Paragraph>,
}

#[pymethods]
impl Paragraph {
    /// Extract text from the paragraph
    ///
    /// Returns:
    ///     Paragraph text as a string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    /// Get all runs in the paragraph
    ///
    /// A run is a contiguous section of text with the same formatting.
    ///
    /// Returns:
    ///     List of Run objects
    fn runs(&self) -> PyResult<Vec<Run>> {
        let runs = self.inner.runs().map_err(to_py_err)?;
        Ok(runs
            .into_iter()
            .map(|r| Run { inner: Arc::new(r) })
            .collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        let preview = if text.len() > 50 {
            format!("{}...", &text[..50])
        } else {
            text
        };
        Ok(format!("<Paragraph: '{}'>", preview))
    }
}

/// A run of text with consistent formatting
///
/// Represents a contiguous section of text that shares the same formatting properties.
#[pyclass(unsendable)]
pub struct Run {
    inner: Arc<litchi::document::Run>,
}

#[pymethods]
impl Run {
    /// Extract text from the run
    ///
    /// Returns:
    ///     Run text as a string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    /// Check if the run is bold
    ///
    /// Returns:
    ///     True if bold, False if not bold, None if unspecified
    fn bold(&self) -> PyResult<Option<bool>> {
        self.inner.bold().map_err(to_py_err)
    }

    /// Check if the run is italic
    ///
    /// Returns:
    ///     True if italic, False if not italic, None if unspecified
    fn italic(&self) -> PyResult<Option<bool>> {
        self.inner.italic().map_err(to_py_err)
    }

    /// Check if the run is underlined
    ///
    /// Returns:
    ///     True if underlined, False if not underlined, None if unspecified
    ///     
    /// Note: This method may not be supported for all document formats.
    fn underline(&self) -> PyResult<Option<bool>> {
        // The underline() method may not exist on all Run types
        // For now, return None as a placeholder
        Ok(None)
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        let preview = if text.len() > 30 {
            format!("{}...", &text[..30])
        } else {
            text
        };
        Ok(format!("<Run: '{}'>", preview))
    }
}

/// A table in a document
///
/// Represents a table with rows and cells.
#[pyclass(unsendable)]
pub struct Table {
    inner: Arc<litchi::document::Table>,
}

#[pymethods]
impl Table {
    /// Get the number of rows in the table
    ///
    /// Returns:
    ///     Number of rows
    fn row_count(&self) -> PyResult<usize> {
        self.inner.row_count().map_err(to_py_err)
    }

    /// Get all rows in the table
    ///
    /// Returns:
    ///     List of TableRow objects
    fn rows(&self) -> PyResult<Vec<TableRow>> {
        let rows = self.inner.rows().map_err(to_py_err)?;
        Ok(rows
            .into_iter()
            .map(|r| TableRow { inner: Arc::new(r) })
            .collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let row_count = self.row_count().unwrap_or(0);
        Ok(format!("<Table: {} rows>", row_count))
    }
}

/// A row in a table
///
/// Represents a single row containing cells.
#[pyclass(unsendable)]
pub struct TableRow {
    inner: Arc<litchi::document::Row>,
}

#[pymethods]
impl TableRow {
    /// Get all cells in the row
    ///
    /// Returns:
    ///     List of TableCell objects
    fn cells(&self) -> PyResult<Vec<TableCell>> {
        let cells = self.inner.cells().map_err(to_py_err)?;
        Ok(cells
            .into_iter()
            .map(|c| TableCell { inner: Arc::new(c) })
            .collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let cell_count = self.cells().map(|c| c.len()).unwrap_or(0);
        Ok(format!("<TableRow: {} cells>", cell_count))
    }
}

/// A cell in a table
///
/// Represents a single cell containing text and possibly other content.
#[pyclass(unsendable)]
pub struct TableCell {
    inner: Arc<litchi::document::Cell>,
}

#[pymethods]
impl TableCell {
    /// Extract text from the cell
    ///
    /// Returns:
    ///     Cell text as a string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        let preview = if text.len() > 30 {
            format!("{}...", &text[..30])
        } else {
            text
        };
        Ok(format!("<TableCell: '{}'>", preview))
    }
}
