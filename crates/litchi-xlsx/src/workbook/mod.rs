//! Immutable workbook snapshots and selector-first sheet lookup.
//!
//! The public surface is a semantic facade. Workbook and worksheet models are
//! kept in `model`, OPC relationship ownership in `package`, raw-to-typed
//! conversions in `codec`, and snapshot edits in `edit`. Physical package
//! identities do not leak into ordinary selector-based reads.

pub mod comments;
pub mod data_model;
pub mod edit;
pub mod source;
pub mod worksheet;

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use edit::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DefaultsEdit, DurablePatch, Edit,
    JoinError, JoinFailure, MergeChoice, MergeLimits, NewSheet, PackageChange, Patch, RowEdit,
    SealedPatch, State, TabEdit, ThreeWayPlan, WorksheetEdit,
};
/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;
pub(crate) use model::Inner;
pub use model::{DateSystem, Flavor, Selector, Visibility, Workbook, Worksheet, WorksheetKind};
pub use source::{SourceBackedWorkbook, SourceCell, SourceCellView, SourceWorksheet};

/// Explicit budgeted undo/redo retention for immutable workbook snapshots.
pub type History = litchi_core::patch::History<Workbook>;
