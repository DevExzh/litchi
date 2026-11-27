//! Snapshot transaction layers for DOC embedded-object edits.
//!
//! Opening, mutation, and persistence are separate so each operation can
//! validate a candidate snapshot before publishing it atomically.

#![allow(
    dead_code,
    reason = "the record-level embedded-object API is driven by the package family and unit tests; each uses a subset"
)]

mod commit;
mod edit;
mod inventory;
mod metadata;
mod mutate;
mod open;
mod patch;
mod snapshot;
mod storage;

pub(crate) use edit::{Transaction, TransactionError};
pub(crate) use patch::{Commit, Patch};
pub use snapshot::Snapshot;
