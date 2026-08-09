//! Validated package ownership for this family.

mod snapshot;

pub(crate) use snapshot::MIMETYPE;
pub use snapshot::{
    Change, Commit, GeometryChange, LayerChange, NameChange, Patch, ResourceChange, Snapshot,
    StructureChange, StyleChange, TextChange, Transaction,
};

/// Explicit bounded undo/redo history for immutable ODG snapshots.
pub type History = litchi_core::patch::History<Snapshot>;
pub use litchi_core::patch::HistoryLimits;
