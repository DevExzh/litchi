//! Per-format handles held by the unified [`File`](crate::File) enum.
//!
//! Each handle is one `Arc` newtype over its format's state (record 0003 §2.2):
//! cheap to clone, `Send + Sync`, retaining the read owner plus the shared
//! source an editor can later be built from (§2.3). [`crate::open`] builds
//! them; the read methods here are the per-format surface, including the
//! format-specific reads the generic family views do not cover.

#[cfg(any(feature = "xlsx", feature = "xlsb"))]
mod cell_text;
#[cfg(feature = "doc")]
pub mod doc;
#[cfg(feature = "docx")]
pub mod docx;
#[cfg(feature = "odp")]
pub mod odp;
#[cfg(feature = "ods")]
pub mod ods;
#[cfg(feature = "odt")]
pub mod odt;
#[cfg(feature = "ppt")]
pub mod ppt;
#[cfg(feature = "pptx")]
pub mod pptx;
#[cfg(feature = "rtf")]
pub mod rtf;
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
