//! Contextual master-layout inventory and transactional record-tree editing.
//!
//! This owner deliberately stops at the PowerPoint record-tree boundary. A
//! parent presentation editor can use Inventory to locate main, title,
//! notes, and handout masters, then open one of those records as a Snapshot.
//! Edits are validated and re-encoded before they become visible, so records
//! that this owner does not understand remain opaque and lossless.

mod codec;
mod inventory;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use inventory::inventory;
pub use model::{Context, Entry, Inventory, Limits, Path, Snapshot};
pub use transaction::{Change, ChangeSet, Commit, Revision, Transaction, View};
