//! Python bindings for Litchi - High-performance Office file format parser
//!
//! This module provides Python bindings for the Litchi Rust library using PyO3.

use pyo3::prelude::*;
use pyo3::types::PyModule;

mod common;
mod document;
mod presentation;
mod sheet;

/// Litchi - High-performance Office file format parser
///
/// This module provides Python bindings for parsing and manipulating various
/// Office file formats including:
/// - Word documents (.doc, .docx)
/// - PowerPoint presentations (.ppt, .pptx)
/// - Excel workbooks (.xls, .xlsx, .xlsb)
/// - OpenDocument formats (.odt, .ods, .odp)
/// - Apple iWork formats (.pages, .key, .numbers)
/// - RTF documents (.rtf)
///
/// # Examples
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
    // Register common types
    common::register(m)?;

    // Register document types
    document::register(m)?;

    // Register presentation types
    presentation::register(m)?;

    // Register sheet types
    sheet::register(m)?;

    Ok(())
}
