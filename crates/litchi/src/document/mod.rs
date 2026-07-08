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
//! - `Table`: Table with rows and cells
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
//! // Access tables
//! for table in doc.tables()? {
//!     for row in table.rows()? {
//!         for cell in row.cells()? {
//!             println!("Cell: {}", cell.text()?);
//!         }
//!     }
//! }
//! # Ok::<(), litchi::common::Error>(())
//! ```

// Submodule declarations
mod doc;
mod element;
mod paragraph;
mod run;
mod table;
mod types;

// Re-exports
pub use doc::Document;
pub use element::DocumentElement;
pub use paragraph::Paragraph;
pub use run::Run;
pub use table::{Cell, Row, Table};
