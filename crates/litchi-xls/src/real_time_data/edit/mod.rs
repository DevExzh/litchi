//! Source-checked, source-preserving edits for workbook-global RTD records.
//!
//! The facade keeps the typed RTD collection separate from BIFF framing.  A
//! transaction stages semantic records detached from its source snapshot;
//! commit reparses the complete candidate before publishing a reversible patch.

#![allow(dead_code, unreachable_pub)]

mod model;
mod transaction;

pub use model::{Commit, Patch, Snapshot};
pub use transaction::Transaction;

/// Read and validate one complete workbook-global BIFF stream.
pub fn read(bytes: impl AsRef<[u8]>) -> crate::Result<Snapshot> {
    Snapshot::parse(bytes)
}

/// Apply a source-checked patch to an immutable snapshot.
pub fn apply(patch: &Patch, source: &Snapshot) -> crate::Result<Snapshot> {
    patch.apply(source)
}
