//! Typed terminal structure of a legacy `PowerPoint` document container.

mod codec;
mod model;
#[allow(
    dead_code,
    reason = "the record-level owner the primary Package::edit drives; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
mod validation;

pub(crate) use codec::encode_record;
pub use model::{CustomTableStylesPlacement, DocumentStructure, Limits, Master, Slide};
#[allow(
    unused_imports,
    reason = "the record-level edit is reached by its unit tests and the slide-order root"
)]
pub(crate) use transaction::{Change, ChangeKind, Commit, Patch, Transaction};
pub use transaction::{Revision, Snapshot};

#[cfg(test)]
mod tests;
