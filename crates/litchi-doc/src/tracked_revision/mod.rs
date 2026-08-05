//! Layered tracked-revision editing for Word 97+ binary documents.
//!
//! The public owner preserves the historical facade while keeping typed
//! revision values, package mutation, and MS-DOC binary codecs separate.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use litchi_ole_common::object::Limits;
pub use model::{DocTrackedRevision, DocTrackedRevisionKind, DocTrackedRevisionMetadata};
pub use package::DocTrackedRevisionEditor;
