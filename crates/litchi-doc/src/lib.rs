#![forbid(unsafe_code)]

//! Read, inspect, edit, and write legacy Word (`.doc`) documents.
//!
//! Compound-file and OfficeArt primitives live in `litchi-cfb`,
//! `litchi-ole-common`, and `litchi-odraw`. The public API is re-exported at
//! this crate's root so callers can use concise paths such as
//! [`Package`] and [`writer::DocWriter`].
//!
//! A Word binary document is an OLE2 structured storage containing a
//! `WordDocument` stream, a `0Table` or `1Table` stream, and optional data,
//! object, drawing, macro, and property streams. [`Package`] owns the
//! container-facing state; [`Document`] exposes typed content such as
//! [`Paragraph`], [`Run`], and [`Table`].
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_doc::Package;
//!
//! let mut package = Package::open("document.doc")?;
//! let document = package.document()?;
//!
//! for paragraph in document.paragraphs()? {
//!     println!("{}", paragraph.text()?);
//! }
//!
//! for table in document.tables()? {
//!     for row in table.rows()? {
//!         for cell in row.cells()? {
//!             println!("{}", cell.text()?);
//!         }
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

/// MTEF extractor for OLE documents (internal use only)
#[cfg(feature = "formula")]
mod mtef_extractor;

/// Shared SPRM (Single Property Modifier) parsing.
///
/// SPRM parsing logic shared between DOC and PPT formats.
/// Based on Apache POI's SPRM handling.
pub mod sprm;

/// SPRM operation constants and utilities.
///
/// Complete SPRM operation definitions based on Apache POI.
pub mod sprm_operations;

mod doc;

pub use doc::*;
