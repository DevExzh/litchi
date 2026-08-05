//! Transactional workbook editing entry points.
//!
//! The canonical edit implementation remains next to the immutable workbook
//! snapshot; this module exposes that owner as the standalone edit layer.

pub use crate::workbook::edit::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DefaultsEdit, Edit, JoinError,
    JoinFailure, NewSheet, PackageChange, Patch, RowEdit, State, TabEdit, WorksheetEdit,
};
