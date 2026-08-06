//! Layered tracked-revision editing for Word 97+ binary documents.
//!
//! The public owner preserves the historical facade while keeping typed
//! revision values, package mutation, and MS-DOC binary codecs separate.

mod codec;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use litchi_ole_common::object::Limits;
pub use model::{Revision, RevisionKind, RevisionMetadata};
pub use package::RevisionEditor;
pub use transaction::{Commit, Error, Patch, Snapshot, Transaction};
