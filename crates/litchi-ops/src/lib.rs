//! Unified operations layer for Litchi.
//!
//! This crate is the home of the format-agnostic operations surface agreed in
//! `docs/UNIFIED_OPS_API_DESIGN.md` and ADR 0033: one way to open, read,
//! query, and edit any document Litchi understands, dispatched statically to
//! the existing per-format owners and their snapshot/Edit/Commit/Patch
//! transactions. It is a dispatch-and-coercion layer, not a second document
//! model. The `litchi` facade re-exports its entry points, so users see a
//! single `litchi::open()`.
//!
//! # Opening and reading
//!
//! [`open`] and [`open_bytes`] detect the format from the content in one
//! pass and return a [`File`]: a flat enum with one variant per format, plus
//! [`File::Locked`] for encrypted and [`File::Signed`] for signed inputs.
//! [`OpenOptions`] carries limits and policy; the common case needs none.
//!
//! ```no_run
//! use litchi_ops::File;
//!
//! # fn main() -> litchi_ops::Result<()> {
//! let file = litchi_ops::open("quarterly.xlsx")?;
//! println!("{:?}: {}", file.format(), file.text()?);
//! if let File::Xlsx(xlsx) = &file {
//!     // Format-specific reads live on the handle.
//!     let data = xlsx.sheet("dAtA")?;
//!     # let _ = data;
//! }
//! println!("{:?}", file.workbook()?.worksheet_names()?);
//! # Ok(())
//! # }
//! ```
//!
//! Every failure is one typed [`Error`] with a stable [`Code`].
//!
//! # Status
//!
//! M0a delivers the open and read surface: [`File`], the family views
//! ([`DocumentView`], [`PresentationView`], [`WorkbookView`]), the
//! [`ToleranceReport`], and [`detect`]. The element wrappers in [`compat`]
//! are scheduled for deletion in M1, when the node view replaces them. Edit
//! verbs, the property registry, batches and saving arrive in later
//! milestones.

#![forbid(unsafe_code)]

mod batch;
mod capabilities;
pub mod compat;
pub mod error;
mod file;
pub mod format;
mod handler;
mod locked;
mod node;
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
mod open;
mod options;
mod path;
mod registry;
mod signed;
mod tolerance;
mod view;

#[cfg(test)]
mod tests;

pub use error::{Code, Error, Result};
pub use file::{Family, File};
pub use locked::{EncryptionKind, Locked};
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
pub use open::{
    open, open_bytes, open_bytes_with, open_bytes_with_tail, open_with, open_with_tail,
};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use open::{open_prepared, open_prepared_with};
#[cfg(feature = "doc")]
pub use options::DocOptions;
pub use options::{OpenOptions, SourceMode};
pub use signed::{SignatureOrigin, Signed};
pub use tolerance::ToleranceReport;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use view::DocumentView;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
pub use view::PresentationView;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
pub use view::WorkbookView;

/// The password [`Locked::unlock`] consumes: owned, never `Clone`, redacted
/// in `Debug`, and zeroized on drop (ADR 0006).
///
/// ```
/// let password = litchi_ops::Password::new("pa55word".to_owned());
/// assert_eq!(format!("{password:?}"), "Password([REDACTED])");
/// ```
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml::Password;

/// Format detection: classify a path, reader, or byte buffer, optionally
/// retaining the parsed owner or the prepared OOXML package.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
pub mod detect;

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
#[doc(hidden)]
pub mod __migration;

// `detect` moved verbatim from the facade, which names its leaf owners
// through the re-export modules `crate::opc`, `crate::docx`, and so on. These
// private aliases keep those paths, so the move stays reviewable as a move.
#[cfg(feature = "docx")]
use litchi_docx as docx;
#[cfg(feature = "opc")]
#[allow(
    unused_imports,
    reason = "an `opc` build without an OPC-consuming format names no `crate::opc` path"
)]
use litchi_opc as opc;
#[cfg(feature = "ppt")]
use litchi_ppt as ppt;
#[cfg(feature = "pptx")]
use litchi_pptx as pptx;
#[cfg(feature = "xls")]
use litchi_xls as xls;
#[cfg(feature = "xlsb")]
use litchi_xlsb as xlsb;
#[cfg(feature = "xlsx")]
use litchi_xlsx as xlsx;
