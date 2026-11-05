//! Source-checked, slide-owned OLE graph edits.
//!
//! The transaction owns only one PresentationML slide.  OLE payloads are
//! retained as inert bytes; this capability never opens, activates, or
//! instantiates them.

mod codec;
mod edit;
mod model;
mod package;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
mod validation;

pub use model::Definition;
pub use package::load;
pub use transaction::{Draft, Revision, Snapshot};

#[cfg(test)]
mod tests;
