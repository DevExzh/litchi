//! Unified Word document module.
//!
//! This module provides a unified API for working with Word documents in both
//! legacy (.doc) and modern (.docx) formats. The format is automatically detected
//! and handled transparently.
//!
//! # Architecture
//!
//! The module provides a format-agnostic API following the python-docx design:
//! - `Document`: The main document API (auto-detects format)
//! - `Paragraph`: Paragraph with text runs
//! - `Run`: Text run with formatting
//! - `Table`: Table with rows and cells (for table-capable document formats)
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::Document;
//!
//! // Open any Word document (.doc or .docx) - format auto-detected
//! let doc = Document::open("document.doc")?;
//!
//! // Extract all text
//! let text = doc.text()?;
//! println!("Document text: {}", text);
//!
//! // Access paragraphs
//! for para in doc.paragraphs()? {
//!     println!("Paragraph: {}", para.text()?);
//!
//!     // Access runs in paragraph
//!     for run in para.runs()? {
//!         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
//!     }
//! }
//!
//! # Ok::<(), litchi::common::Error>(())
//! ```

// Submodule declarations
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod cell_merge;
mod doc;
mod element;
mod paragraph;
mod run;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod table;
mod types;

// Re-exports
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use cell_merge::CellMerge;
pub use doc::Document;
pub use element::DocumentElement;
pub use paragraph::Paragraph;
pub use run::Run;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use table::{Cell, Row, Table};
