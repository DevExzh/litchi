//! Document API bindings

use std::path::PathBuf;

use litchi::FileFormat as CoreFormat;
use litchi::compat::word;
use pyo3::prelude::*;
use pyo3::types::PyModule;

use crate::common::preview;
use crate::errors::{Failure, Opener, to_py_err};
use crate::iwork;

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

/// What a `Document` reads: a unified file of the word-processing family,
/// or a Pages `File`, which has no family view (its leaf package is read).
enum Source {
    File(litchi::File),
    Pages(litchi::File),
}

/// Unified Word document interface
///
/// Provides a format-agnostic interface for word-processing documents:
/// .doc, .docx, .rtf, .odt and .pages. The format is detected from the
/// file's content when it is opened.
///
/// A Document is immutable and can be shared between threads. Opening and
/// the document-wide reads release the GIL while they run.
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
#[pyclass(frozen, module = "litchi_py")]
pub struct Document {
    source: Source,
}

impl Document {
    fn read(path: &std::path::Path) -> Result<Source, Failure> {
        match litchi::open(path) {
            // `litchi::File` opens Pages packages and package directories
            // (records 0296 and 0306), but its family views do not cover
            // them; this reader projects the package the file retains.
            Ok(file) if file.format() == Some(CoreFormat::Pages) => Ok(Source::Pages(file)),
            Ok(file) => match file.document() {
                Ok(_) => Ok(Source::File(file)),
                Err(error) => Err(Failure::open(&error, Opener::Document)),
            },
            Err(error) => Err(Failure::open(&error, Opener::Document)),
        }
    }
}

#[pymethods]
impl Document {
    /// Open a Word document from a file path
    ///
    /// The format (.doc, .docx, .rtf, .odt or .pages) is detected from the
    /// file's content, not from its name.
    ///
    /// Args:
    ///     path: Path to the document file
    ///
    /// Returns:
    ///     Document instance
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file cannot be read
    ///     UnsupportedError: (a ValueError) If the file is not a
    ///         word-processing document, is not a recognized format, or
    ///         uses a variant that cannot be read (flat OpenDocument, Word 6)
    ///     CorruptFileError: (a ValueError) If the file is malformed
    ///     ProtectedError: (a ValueError) If the document is encrypted
    #[staticmethod]
    fn open(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let source = py.detach(|| Self::read(&path))?;
        Ok(Document { source })
    }

    /// Extract all text from the document
    ///
    /// Returns:
    ///     All text content as a single string
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file changed after it was opened
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        py.detach(|| match &self.source {
            Source::File(file) | Source::Pages(file) => {
                file.text().map_err(|error| Failure::read(&error))
            },
        })
        .map_err(PyErr::from)
    }

    /// Get all paragraphs in the document
    ///
    /// Returns:
    ///     List of Paragraph objects
    fn paragraphs(&self, py: Python<'_>) -> PyResult<Vec<Paragraph>> {
        let paragraphs = py.detach(|| match &self.source {
            Source::File(file) => file
                .document()
                .and_then(|document| document.paragraphs())
                .map(|paragraphs| paragraphs.into_iter().map(ParagraphInner::Word).collect())
                .map_err(|error| Failure::read(&error)),
            Source::Pages(file) => Ok(iwork::pages_paragraphs(iwork::pages(file)?)
                .into_iter()
                .map(ParagraphInner::Pages)
                .collect::<Vec<_>>()),
        })?;
        Ok(paragraphs
            .into_iter()
            .map(|inner| Paragraph { inner })
            .collect())
    }

    /// Get all tables in the document
    ///
    /// Pages tables are not exposed; the list is empty for .pages files.
    ///
    /// Returns:
    ///     List of Table objects
    fn tables(&self, py: Python<'_>) -> PyResult<Vec<Table>> {
        let tables = py.detach(|| match &self.source {
            Source::File(file) => file
                .document()
                .and_then(|document| document.tables())
                .map_err(|error| Failure::read(&error)),
            Source::Pages(_) => Ok(Vec::new()),
        })?;
        Ok(tables.into_iter().map(|inner| Table { inner }).collect())
    }

    fn __repr__(&self) -> String {
        "<Document>".to_string()
    }
}

#[allow(
    clippy::large_enum_variant,
    reason = "each value is the payload of one Python object, and all but the iWork case hold the large variant, so boxing it would add an allocation per element and save nothing"
)]
enum ParagraphInner {
    Word(word::Paragraph),
    /// One Pages text block, which has no run formatting.
    Pages(String),
}

/// A paragraph in a document
///
/// Represents a single paragraph with text and formatting.
#[pyclass(frozen, module = "litchi_py")]
pub struct Paragraph {
    inner: ParagraphInner,
}

#[pymethods]
impl Paragraph {
    /// Extract text from the paragraph
    ///
    /// Returns:
    ///     Paragraph text as a string
    fn text(&self) -> PyResult<String> {
        match &self.inner {
            ParagraphInner::Word(paragraph) => paragraph.text().map_err(to_py_err),
            ParagraphInner::Pages(text) => Ok(text.clone()),
        }
    }

    /// Get all runs in the paragraph
    ///
    /// A run is a contiguous section of text with the same formatting. A
    /// Pages paragraph is one run without formatting.
    ///
    /// Returns:
    ///     List of Run objects
    fn runs(&self) -> PyResult<Vec<Run>> {
        Ok(match &self.inner {
            ParagraphInner::Word(paragraph) => paragraph
                .runs()
                .map_err(to_py_err)?
                .into_iter()
                .map(|run| Run {
                    inner: RunInner::Word(run),
                })
                .collect(),
            ParagraphInner::Pages(text) => vec![Run {
                inner: RunInner::Pages(text.clone()),
            }],
        })
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        Ok(format!("<Paragraph: '{}'>", preview(&text, 50)))
    }
}

#[allow(
    clippy::large_enum_variant,
    reason = "each value is the payload of one Python object, and all but the iWork case hold the large variant, so boxing it would add an allocation per element and save nothing"
)]
enum RunInner {
    Word(word::Run),
    /// The text of a Pages paragraph, which reports no formatting.
    Pages(String),
}

/// A run of text with consistent formatting
///
/// Represents a contiguous section of text that shares the same formatting properties.
#[pyclass(frozen, module = "litchi_py")]
pub struct Run {
    inner: RunInner,
}

#[pymethods]
impl Run {
    /// Extract text from the run
    ///
    /// Returns:
    ///     Run text as a string
    fn text(&self) -> PyResult<String> {
        match &self.inner {
            RunInner::Word(run) => run.text().map_err(to_py_err),
            RunInner::Pages(text) => Ok(text.clone()),
        }
    }

    /// Check if the run is bold
    ///
    /// Returns:
    ///     True if bold, False if not bold, None if unspecified
    fn bold(&self) -> PyResult<Option<bool>> {
        match &self.inner {
            RunInner::Word(run) => run.bold().map_err(to_py_err),
            RunInner::Pages(_) => Ok(None),
        }
    }

    /// Check if the run is italic
    ///
    /// Returns:
    ///     True if italic, False if not italic, None if unspecified
    fn italic(&self) -> PyResult<Option<bool>> {
        match &self.inner {
            RunInner::Word(run) => run.italic().map_err(to_py_err),
            RunInner::Pages(_) => Ok(None),
        }
    }

    /// Check if the run is underlined
    ///
    /// .docx reports the run's own underline (False for an explicit
    /// "none"); .doc reports True for any underline and None otherwise; .rtf
    /// always answers True or False. .odt and .pages runs report None.
    ///
    /// Returns:
    ///     True if underlined, False if not underlined, None if unspecified
    fn underline(&self) -> PyResult<Option<bool>> {
        match &self.inner {
            RunInner::Word(run) => run.underline().map_err(to_py_err),
            RunInner::Pages(_) => Ok(None),
        }
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        Ok(format!("<Run: '{}'>", preview(&text, 30)))
    }
}

/// A table in a document
///
/// Represents a table with rows and cells.
#[pyclass(frozen, module = "litchi_py")]
pub struct Table {
    inner: word::Table,
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
        Ok(rows.into_iter().map(|inner| TableRow { inner }).collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let row_count = self.row_count().unwrap_or(0);
        Ok(format!("<Table: {} rows>", row_count))
    }
}

/// A row in a table
///
/// Represents a single row containing cells.
#[pyclass(frozen, module = "litchi_py")]
pub struct TableRow {
    inner: word::Row,
}

#[pymethods]
impl TableRow {
    /// Get all cells in the row
    ///
    /// Returns:
    ///     List of TableCell objects
    fn cells(&self) -> PyResult<Vec<TableCell>> {
        let cells = self.inner.cells().map_err(to_py_err)?;
        Ok(cells.into_iter().map(|inner| TableCell { inner }).collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let cell_count = self.inner.cells().map(|cells| cells.len()).unwrap_or(0);
        Ok(format!("<TableRow: {} cells>", cell_count))
    }
}

/// A cell in a table
///
/// Represents a single cell containing text and possibly other content.
#[pyclass(frozen, module = "litchi_py")]
pub struct TableCell {
    inner: word::Cell,
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
        Ok(format!("<TableCell: '{}'>", preview(&text, 30)))
    }
}
