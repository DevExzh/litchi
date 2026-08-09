//! Typed terminal structure of a legacy `PowerPoint` document container.

mod codec;
mod model;
mod transaction;
mod validation;

pub use model::{CustomTableStylesPlacement, DocumentStructure, Limits, Master, Slide};
pub use transaction::{Change, ChangeKind, Commit, Patch, Revision, Snapshot, Transaction};

#[cfg(test)]
mod tests;
