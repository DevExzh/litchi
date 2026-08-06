//! Source-checked, slide-owned OLE graph edits.
//!
//! The transaction owns only one PresentationML slide.  OLE payloads are
//! retained as inert bytes; this capability never opens, activates, or
//! instantiates them.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

pub use model::Definition;
pub use package::{apply_commit, apply_patch, load};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

#[cfg(test)]
mod tests;
