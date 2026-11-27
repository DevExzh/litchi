//! Contextual document queries grouped by semantic domain.
//!
//! The inherent [`Document`](super::super::state::Document) API remains flat
//! for callers, while its implementation is layered by content, fields,
//! annotations, layout, and binary-backed metadata.

mod basics;
mod body;
mod content_support;
mod field_support;
mod fields;
mod fonts;
mod layout;
mod links;
mod media;
mod metadata;
mod numbering;
mod prelude;
pub(in crate::document) mod stories;
mod subdocuments;

pub use body::BodyParagraphs;
pub use subdocuments::HeaderFooterSlot;
