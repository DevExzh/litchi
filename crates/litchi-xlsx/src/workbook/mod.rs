//! The read-only workbook projection and selector-first sheet lookup.
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
mod drawing;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use drawing::{
    WorksheetContentPart, WorksheetContentPartOutboundRelationship, WorksheetContentPartPayload,
    WorksheetContentPartRelationship, WorksheetDrawing, WorksheetImageReference, WorksheetPicture,
    WorksheetSvgDescriptor, WorksheetSvgImage,
};
pub use edit::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, Diagnostics, DurablePatch, Edit, EditOptions,
    JoinError, JoinFailure, JoinNormalization, MergeChoice, MergeLimits, NewSheet, PackageChange,
    Patch, SealedPatch, State, ThreeWayPlan,
};
/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;
pub(crate) use model::Inner;
pub use model::{DateSystem, Flavor, Selector, Visibility, Workbook, Worksheet, WorksheetKind};
pub use source::{SourceBackedWorkbook, SourceCell, SourceCellView, SourceWorksheet};

/// Explicit budgeted undo/redo retention for immutable package snapshots.
pub type History = litchi_core::patch::History<crate::Package>;
