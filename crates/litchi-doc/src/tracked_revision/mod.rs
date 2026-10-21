//! Layered tracked-revision editing for Word 97+ binary documents.
//!
//! The public owner preserves the historical facade while keeping typed
//! revision values, package mutation, and MS-DOC binary codecs separate.

mod codec;
mod family;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use litchi_ole_common::object::Limits;
pub use model::{Revision, RevisionKind, RevisionMetadata};
pub(crate) use package::RevisionEditor;
pub(crate) use package::{BookmarkRange, ParagraphCopy, PictureGraph};
