//! Python bindings for Litchi - High-performance Office file format parser
//!
//! This module provides Python bindings for the Litchi Rust library using PyO3.

use pyo3::prelude::*;
use pyo3::types::PyModule;

mod common;
mod document;
mod errors;
mod file;
mod iwork;
mod m4;
mod merge;
mod nodes;
mod presentation;
mod props;
mod pyde;
mod pyser;
mod raw;
mod refresh;
mod sheet;
mod write;

/// Litchi - High-performance Office file format parser
///
/// This module provides Python bindings for reading various Office file
/// formats including:
/// - Word documents (.doc, .docx)
/// - PowerPoint presentations (.ppt, .pptx)
/// - Excel workbooks (.xls, .xlsx, .xlsb)
/// - OpenDocument formats (.odt, .ods, .odp)
/// - Apple iWork formats (.pages, .key, .numbers)
/// - RTF documents (.rtf)
///
/// Errors are raised as `LitchiError` subclasses that carry the error's
/// `code`, `message` and `suggestion`; each family also derives from the
/// matching built-in exception (`OSError`, `ValueError` or `LookupError`).
/// Handles are immutable and can be used from any thread.
///
/// # Examples
///
/// ## The unified API
///
/// `open`, `open_bytes` and `create` return a `File`, which reads with
/// OfficeCLI's paths and selectors, returns OfficeCLI's node views as dicts,
/// and edits through `set`, `add`, `remove`, `move`, `swap`, `edit` and
/// OfficeCLI batches.
///
/// ```python
/// import litchi_py
///
/// with litchi_py.open("report.docx") as doc:
///     paragraph = doc.get("/body/p[1]")    # path, type, text, format, children
///     matches = doc.query("paragraph[style=Heading1]", limit=10)
///     print(doc.view("outline"))
///     print(doc.capabilities()["verbs"]["get"])
///
///     # Writes return a new File; the one they are called on is unchanged.
///     doc = doc.set("/body/p[1]/r[1]", {"font.bold": True})
///     doc, path = doc.add("/body", "paragraph", {"text": "Appendix"})
///     doc = doc.batch([{"command": "remove", "path": "/body/p[2]"}]).file
///     doc.save()
/// ```
///
/// ## Reading Word Documents
///
/// ```python
/// from litchi_py import Document
///
/// # Open a document (auto-detects format)
/// doc = Document.open("document.docx")
///
/// # Extract text
/// text = doc.text()
/// print(text)
///
/// # Access paragraphs
/// for para in doc.paragraphs():
///     print(f"Paragraph: {para.text()}")
///     
///     # Access runs with formatting
///     for run in para.runs():
///         print(f"  Text: {run.text()}")
///         if run.bold():
///             print("    (bold)")
/// ```
///
/// ## Reading PowerPoint Presentations
///
/// ```python
/// from litchi_py import Presentation
///
/// # Open a presentation
/// pres = Presentation.open("presentation.pptx")
///
/// # Extract text
/// text = pres.text()
/// print(text)
///
/// # Get slide count
/// print(f"Total slides: {pres.slide_count()}")
///
/// # Access individual slides
/// for i, slide in enumerate(pres.slides()):
///     print(f"Slide {i + 1}: {slide.text()}")
/// ```
#[pymodule]
fn litchi_py(m: &Bound<'_, PyModule>) -> PyResult<()> {
    // Register the exception classes
    errors::register(m)?;

    // Register common types
    common::register(m)?;

    // Register document types
    document::register(m)?;

    // Register presentation types
    presentation::register(m)?;

    // Register sheet types
    sheet::register(m)?;

    // Register the unified read API: `open`, `open_bytes` and `File`
    file::register(m)?;

    // Register the unified write API: `Commit` and `Edit`
    write::register(m)?;

    Ok(())
}
