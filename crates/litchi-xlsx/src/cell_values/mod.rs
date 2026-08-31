//! Source-backed scalar-cell and formula edits for existing worksheet cells.
//!
//! This deliberately narrow capability changes or removes scalar cells already
//! stored in selected worksheets. Clearing retains the `<c>` owner and its local
//! style; removal deletes that complete owner while retaining its row and the
//! producer's conservative dimension. Cacheless scalar formulas and direct
//! dates are supported; every effective mutation invalidates workbook
//! calculation properties and removes a captured calculation chain atomically.
//! It never creates cells or mutates shared style/string tables. Workbooks
//! outside the statically provable closure are refused.

mod patch;
mod snapshot;
mod source;
mod validation;

pub use patch::Patch;
pub use patch::{Commit, Diagnostics, MultiCommit, MultiDiagnostics, MultiPatch};
pub use snapshot::{MAX_MULTI_WORKSHEET_BYTES, MAX_SHEET_OWNERS, MultiSnapshot, Snapshot};
pub use source::{
    CellValueEdit, MAX_BATCH_EDITS, MultiSourceEdit, SheetCellValueEdit, SourceBackedEditor,
    SourceEdit,
};
