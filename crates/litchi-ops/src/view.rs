//! Family views of a [`File`](crate::File): borrowed, format-neutral reads
//! for one family of formats (record 0003 §2.4).
//!
//! [`File::document`](crate::File::document),
//! [`File::presentation`](crate::File::presentation) and
//! [`File::workbook`](crate::File::workbook) return them, refusing a file of
//! another family with [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat).
//! The views carry the reads the facade's `Document`, `Presentation` and
//! `sheet::Workbook` types offered, with the same per-format behavior; the
//! element values they return are the [`crate::compat`] wrappers, which the
//! M1 node view replaces.

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod document;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
mod presentation;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
mod workbook;

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
pub use document::DocumentView;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
pub use presentation::PresentationView;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
pub use workbook::WorkbookView;
