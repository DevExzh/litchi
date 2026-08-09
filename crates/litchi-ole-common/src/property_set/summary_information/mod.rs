//! Typed [MS-OLEPS] 2.25.1 `SummaryInformation` properties.
//!
//! This owner projects the standard PIDSI vocabulary over the generic
//! [`Section`](super::Section) grammar.  Unknown properties and the complete
//! source ordering remain attached to every [`Snapshot`].  The transaction
//! validates code-page representability, bounded values, and the standard
//! PID types before publishing a source-checked reversible patch.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::bool_assert_comparison,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use model::Snapshot;
pub use model::{
    APP_NAME, AUTHOR, CHARACTER_COUNT, CODEPAGE, COMMENTS, CREATE_DTM, ClipboardTag, DOC_SECURITY,
    DocumentSecurity, EDIT_TIME, FileTime, ImageFormat, KEYWORDS, LAST_AUTHOR, LAST_PRINTED,
    LAST_SAVE_DTM, MAX_TEXT_BYTES, MAX_THUMBNAIL_BYTES, PAGE_COUNT, REVISION_NUMBER, SUBJECT,
    TEMPLATE, THUMBNAIL, TITLE, Thumbnail, ThumbnailRef, WORD_COUNT,
};
pub use transaction::{Commit, Edit, Patch, Transaction};
