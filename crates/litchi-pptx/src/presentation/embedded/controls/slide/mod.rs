//! Source-checked ActiveX/OCX edits owned by one PresentationML slide.
//!
//! The slide owner keeps its XML, descriptor relationship, and opaque binary
//! state in one detached snapshot.  Transactions only touch the small set of
//! typed metadata attributes selected by the caller; all other XML and binary
//! bytes remain opaque and inert.

mod codec;
mod package;
mod transaction;
mod validation;

pub use package::{apply_commit, apply_patch, load};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

#[cfg(test)]
mod tests;
