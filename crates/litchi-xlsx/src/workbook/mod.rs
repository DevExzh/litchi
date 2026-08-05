//! Immutable workbook snapshots and selector-first sheet lookup.
//!
//! The public surface is a semantic facade. Workbook and worksheet models are
//! kept in [`model`], OPC relationship ownership in [`package`], raw-to-typed
//! conversions in [`codec`], and snapshot edits in [`edit`]. Physical package
//! identities do not leak into ordinary selector-based reads.

pub mod comments;
pub mod data_model;
pub mod edit;
pub mod worksheet;

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use edit::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DefaultsEdit, Edit, JoinError,
    JoinFailure, NewSheet, PackageChange, Patch, RowEdit, State, TabEdit, WorksheetEdit,
};
pub use model::{DateSystem, Flavor, Selector, Visibility, Workbook, Worksheet, WorksheetKind};

pub(crate) use model::Inner;
