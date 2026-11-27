//! Source-checked ActiveX/OCX edits owned by one PresentationML slide.
//!
//! The slide owner keeps its XML, descriptor relationship, and opaque binary
//! state in one detached snapshot.  Transactions only touch the small set of
//! typed metadata attributes selected by the caller; all other XML and binary
//! bytes remain opaque and inert.

mod codec;
mod edit;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
mod validation;

pub use package::load;
pub use transaction::{Draft, Revision, Snapshot};

#[cfg(test)]
mod determinism_tests;
#[cfg(test)]
mod tests;
