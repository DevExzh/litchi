//! Snapshot transaction layers for DOC embedded-object edits.
//!
//! Opening, mutation, and persistence are separate so each operation can
//! validate a candidate snapshot before publishing it atomically.

mod commit;
mod edit;
mod inventory;
mod metadata;
mod mutate;
mod open;
mod patch;
mod snapshot;
mod storage;

pub use edit::{Transaction, TransactionError};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
