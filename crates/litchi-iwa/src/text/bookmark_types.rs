//! IWA-facing names for the archive-free bookmark values.
//!
//! The semantic definitions live in `litchi-iwa-text`; this module keeps the
//! existing text-facade vocabulary available while native object lookup stays
//! in the private IWA adapter.

pub use litchi_iwa_text::bookmark::{
    Bookmark as TextBookmark, Id as TextBookmarkId, Name as TextBookmarkName,
    Settings as TextBookmarkSettings, Visibility as TextBookmarkVisibility,
};

impl From<litchi_iwa_text::bookmark::Error> for crate::Error {
    fn from(error: litchi_iwa_text::bookmark::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}
