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
//! `open` and `open_bytes` detect the format from the content in one
//! pass and return a [`File`]: a flat enum with one variant per format, plus
//! [`File::Locked`] for encrypted and [`File::Signed`] for signed inputs.
//! [`OpenOptions`] carries limits and policy; the common case needs none.
//!
//! ```no_run
//! use litchi_ops::File;
//!
//! # #[cfg(feature = "xlsx")]
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
//! # #[cfg(not(feature = "xlsx"))]
//! # fn main() {}
//! ```
//!
//! Every failure is one typed [`Error`] with a stable [`Code`].
//!
//! # Markdown
//!
//! With the `markdown` feature, [`File`], its document and presentation
//! views and the [`compat`] element wrappers implement
//! `litchi_markdown::ToMarkdown`, reporting this crate's typed [`Error`].
//!
//! ```no_run
//! # #[cfg(all(feature = "markdown", feature = "docx"))]
//! # {
//! use litchi_markdown::{MarkdownOptions, ToMarkdown};
//!
//! let file = litchi_ops::open("report.docx")?;
//! println!("{}", file.to_markdown()?);
//! let plain = MarkdownOptions::new().with_styles(false);
//! println!("{}", file.to_markdown_with_options(&plain)?);
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! - Documents (DOC, DOCX, RTF, ODT) convert paragraph by paragraph and
//!   table by table, in document order; DOCX adds hyperlinks, embedded
//!   images, footnotes and endnotes from its package.
//! - Presentations (PPT, PPTX, ODP) convert to one `# Slide N` section per
//!   slide, separated by thematic breaks.
//! - Workbooks are refused with [`Code::UnsupportedCapability`].
//! - A [`File::Locked`] is refused with [`Code::DocumentProtected`]; a
//!   [`File::Signed`] converts the file it wraps.
//!
//! The conversion refuses what Markdown cannot carry instead of dropping it:
//!
//! | Condition | Code |
//! |---|---|
//! | Content Markdown cannot represent faithfully: unresolved list numbering, active `altChunk` blocks, inline images in DOC, hyperlinks or notes it cannot place, formulas without a LaTeX form | [`Code::UnsupportedElement`] |
//! | Output that needs a cargo feature this build lacks (`formula` for formulas, `yaml` for front matter) | [`Code::UnsupportedCapability`], with a suggestion |
//! | Malformed content met while converting, such as a note the body references but the package lacks | [`Code::CorruptFile`] |
//! | A failed internal invariant of the conversion | [`Code::InternalError`] |
//! | A source-backed file that changed during or after the conversion | [`Code::IoError`] |
//!
//! Read failures keep the leaf's typed error as [`Error::leaf`], and every
//! error a document or presentation conversion reports names the format.
//!
//! # Status
//!
//! Every format reads and writes through [`File`]: paths, queries and views;
//! [`File::set`], [`File::add`], [`File::remove`], [`File::move_to`] and
//! [`File::swap`], each returning a new [`File`]; [`File::edit`] transactions
//! with exact inverse patches; [`batch`] payloads of [`BatchItem`]s; and
//! publication through [`File::save`], [`File::save_as`], [`File::to_bytes`]
//! and [`File::write_to`]. [`create`] starts a new file. Coverage differs per
//! element and format; [`File::capabilities`] reports the supported verbs and
//! properties and the typed gaps of each handle. The element wrappers in
//! [`compat`] are retained for the facade's family views.
#![forbid(unsafe_code)]

pub mod batch;
#[cfg(all(
    feature = "eval",
    any(feature = "xlsx", feature = "xlsb", feature = "xls")
))]
pub mod calc;
pub mod capabilities;
pub mod compat;
mod create;
pub mod error;
mod file;
mod find;
pub mod format;
mod handler;
pub mod import;
mod locked;
#[cfg(feature = "markdown")]
mod markdown;
pub mod merge;
pub mod new_node;
pub mod node;
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
    feature = "numbers",
    odf_objects
))]
mod open;
mod options;
pub mod path;
pub mod placement;
pub mod props;
pub mod query;
pub mod raw;
mod read;
pub mod refresh;
mod registry;
mod save;
pub mod schema;
mod signed;
pub mod target;
mod tolerance;
pub mod transaction;
pub mod value;
pub mod view;
mod write;

#[cfg(test)]
mod tests;

pub use batch::{Applied, BatchCommit, BatchItem, BatchMode, ItemResult};
pub use create::{Conformance, CreateOptions, Flavor, create, create_with};
pub use error::{Code, Error, Result};
pub use file::{Family, File};
pub use locked::{EncryptionKind, Locked};
pub use new_node::{Added, IntoNode, NewNode};
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
    feature = "numbers",
    odf_objects
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
pub use placement::{IntoPlacement, Placement};
pub use props::{IntoProps, Key, PropValue, Props};
pub use save::{Overwrite, SaveOptions, SaveTo, Saved};
pub use schema::schema;
pub use signed::{SignatureOrigin, Signed};
pub use target::{IntoTarget, TargetArg};
pub use tolerance::ToleranceReport;
pub use transaction::EditOptions;
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
    feature = "numbers",
    odf_objects
))]
pub mod detect;

// `detect` moved verbatim from the facade, which names its OPC and PPTX leaf
// owners through the re-export modules `crate::opc` and `crate::pptx`. These
// private aliases keep those paths.
#[cfg(feature = "opc")]
#[allow(
    unused_imports,
    reason = "an `opc` build without an OPC-consuming format names no `crate::opc` path"
)]
use litchi_opc as opc;
#[cfg(feature = "pptx")]
use litchi_pptx as pptx;

pub use path::{IntoPath, PathArg};
pub use query::{IntoQuery, QueryArg};
