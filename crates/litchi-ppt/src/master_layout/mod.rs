//! Contextual master-layout inventory and package-level record-tree editing.
//!
//! This owner deliberately stops at the PowerPoint record-tree boundary.
//! [`inventory`](crate::master_layout::inventory) locates main, title,
//! notes, and handout masters in a parsed record tree and
//! [`Snapshot`](crate::master_layout::Snapshot) reads one of those records.
//! An [`Edit`] (from [`crate::Package::edit_master_layout`]) stages add / append /
//! remove / replace operations on one master record of a package; its
//! commit validates and re-encodes the candidate before publishing a new
//! package, so records that this owner does not understand remain opaque and
//! lossless.

mod codec;
mod edit;
mod inventory;
mod model;
#[allow(
    dead_code,
    reason = "the record-level owner behind the package-level Edit; its full surface stays exercised by the unit tests (record 0357)"
)]
mod transaction;
pub(crate) mod validation;

#[cfg(test)]
mod tests;

pub use edit::Edit;
pub use inventory::inventory;
pub use model::{Context, Entry, Inventory, Limits, Path, Snapshot};
pub use transaction::{Change, Revision, View};
pub(crate) use transaction::{RecordEdit, RecordPatch};

crate::scoped::family!("master-layout");
