//! Validated package ownership for this family.

mod snapshot;

pub(crate) use snapshot::MIMETYPE;
pub use snapshot::{Commit, LayerChange, NameChange, Patch, Snapshot, TextChange, Transaction};
