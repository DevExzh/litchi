//! Typed MS-OSHARED Document Summary Information properties.
//!
//! This owner layers the fixed `PIDDSI` vocabulary over the generic
//! [`Section`](super::Section) model.  The complete section remains attached
//! to every [`Snapshot`], so unrecognized properties and unsupported wire
//! payloads are retained while known fields receive checked, contextual
//! accessors.  [`Transaction`] produces a source-checked [`Patch`] and never
//! activates links, signatures, or other OLE behaviors.

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
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use model::{
    BYTE_COUNT, CATEGORY, CHARACTER_COUNT_WITH_SPACES, CODEPAGE, COMPANY, CONTENT_STATUS,
    CONTENT_TYPE, DIGITAL_SIGNATURE, DOCUMENT_PARTS, DOCUMENT_VERSION, HEADING_PAIRS, HIDDEN_COUNT,
    HYPERLINKS, HYPERLINKS_CHANGED, LANGUAGE, LINE_COUNT, LINK_BASE, LINKS_DIRTY, MANAGER,
    MAX_TEXT_BYTES, MULTIMEDIA_CLIP_COUNT, NOTE_COUNT, PARAGRAPH_COUNT, PRESENTATION_FORMAT, SCALE,
    SHARED_DOCUMENT, SLIDE_COUNT, Snapshot, VERSION, Version,
};
pub use transaction::{Commit, Edit, Patch, Transaction};
