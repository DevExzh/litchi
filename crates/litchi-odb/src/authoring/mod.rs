//! Detached construction for new family packages.

mod builder;
mod composition;
mod durable;
mod policy;
mod transaction;

pub use builder::Builder;
pub use composition::{
    CompositionLimits, JoinError, JoinedEdits, Lineage, MergeChoice, MergePlan, MergePlanError,
    PreparedEdit,
};
pub use durable::{DurablePatch, SealedPatch};
pub use policy::{
    ActiveContentDisposition, DependencyDisposition, EditPolicy, EncryptionCapability,
    EncryptionPolicy, ProtectionCapabilities, ProtectionOperation, ProtectionStatus,
    ProtectionSupport, ProtectionTransition, SignatureCapability, SignaturePolicy,
};
pub(crate) use transaction::producer_extensions;
pub use transaction::{Change, ChangeAction, ChangeKind, Commit, Edit, Patch, QueryChange};

/// Explicit budgeted undo/redo retention for immutable database snapshots.
pub type History = litchi_core::patch::History<crate::Database>;

/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;
