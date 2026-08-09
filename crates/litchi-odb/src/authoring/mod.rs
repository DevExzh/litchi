//! Detached construction for new family packages.

mod builder;
mod transaction;

pub use builder::Builder;
pub(crate) use transaction::producer_extensions;
pub use transaction::{Change, ChangeKind, Commit, Edit, Patch, QueryChange};

/// Explicit budgeted undo/redo retention for immutable database snapshots.
pub type History = litchi_core::patch::History<crate::Database>;

/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;
