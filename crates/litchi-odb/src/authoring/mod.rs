//! Detached construction for new family packages.

mod builder;
mod composition;
mod durable;
mod policy;
mod transaction;

pub use builder::Builder;
pub use composition::{
    CompositionLimits, JoinError, JoinedEdits, Lineage, MergeChoice, MergePlan, MergePlanError,
    SubEdit,
};
pub use durable::{DurablePatch, SealedPatch};
pub use policy::{
    ActiveContentDisposition, DependencyDisposition, EditPolicy, EncryptionCapability,
    EncryptionPolicy, ProtectionCapabilities, ProtectionDecision, ProtectionOperation,
    ProtectionRefusal, ProtectionStatus, ProtectionSupport, ProtectionTransition,
    SignatureCapability, SignaturePolicy,
};
pub(crate) use transaction::producer_extensions;
pub use transaction::{
    Change, ChangeAction, ChangeKind, Commit, Diagnostics, Edit, Patch, QueryChange,
};

/// Explicit budgeted undo/redo retention for immutable database packages.
pub type History = litchi_core::patch::History<crate::Package>;

/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;
