//! Per-format handles held by the unified [`File`](crate::File) enum.
//!
//! Each handle is one `Arc` newtype over its format's state (record 0003 §2.2):
//! cheap to clone, `Send + Sync`, retaining the read owner plus the shared
//! source an editor can later be built from (§2.3). `crate::open` builds
//! them; the read methods here are the per-format surface, including the
//! format-specific reads the generic family views do not cover.
//!
//! The tier 1 and tier 2 handles also carry a signature disposition (record
//! 0013): their state sits behind `signature::Opened`, so a signed file and
//! the file [`Signed::unsign`](crate::Signed::unsign) returns share one
//! state and differ only in what their editor builders may do.

#[cfg(any(feature = "xlsx", feature = "xlsb"))]
mod cell_text;
#[cfg(any(feature = "pptx", feature = "xlsx", feature = "docx", feature = "xlsb"))]
pub(crate) mod chart;
#[cfg(feature = "doc")]
pub mod doc;
#[cfg(feature = "docx")]
pub mod docx;
#[cfg(feature = "iwork-adapter")]
pub mod iwork;
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) mod metadata;
#[cfg(any(feature = "ods", feature = "odp", odf_objects))]
pub(crate) mod odf_metadata;
#[cfg(feature = "odp")]
pub(crate) mod odf_odp;
#[cfg(odf_objects)]
pub mod odf_tree;
#[cfg(feature = "odp")]
pub mod odp;
#[cfg(feature = "ods")]
pub mod ods;
#[cfg(feature = "odt")]
pub mod odt;
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
pub(crate) mod ole_metadata;
#[cfg(feature = "ppt")]
pub mod ppt;
#[cfg(feature = "pptx")]
pub mod pptx;
#[cfg(feature = "rtf")]
pub mod rtf;
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb"
))]
pub(crate) mod signature;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
pub(crate) mod workbook;
#[cfg(feature = "xls")]
pub mod xls;
#[cfg(feature = "xlsb")]
pub mod xlsb;
#[cfg(feature = "xlsx")]
pub mod xlsx;

/// What a legacy binary handle made of an in-memory compound file.
///
/// A document that needs a password to open hands its parsed container back
/// unchanged, so the opener can keep it as the source of a
/// [`Locked`](crate::Locked) handle without copying or parsing it again.
#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
pub(crate) enum LegacyOpen<T> {
    /// The document opened.
    Opened(T),
    /// The document is encrypted and no password was supplied. Boxed: this
    /// is the rare arm, and the parsed container is large.
    PasswordRequired(Box<litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>>),
}
