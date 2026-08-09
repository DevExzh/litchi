//! Validated package ownership for this family.

mod snapshot;

pub(crate) use snapshot::MIMETYPE;
pub use snapshot::{Commit, NameChange, Patch, Snapshot, TextChange, Transaction};
