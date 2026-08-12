//! Source-backed scalar-cell edits for existing worksheet cells.
//!
//! This deliberately narrow capability changes or removes scalar cells already
//! stored in a single worksheet. Clearing retains the `<c>` owner and its local
//! style; removal deletes that complete owner while retaining its row and the
//! producer's conservative dimension. It never creates cells, mutates shared
//! style tables or shared-string topology, authors formulas, or claims to
//! recalculate Excel dependencies. Workbooks outside the statically provable
//! subset are refused.

mod patch;
mod snapshot;
mod source;
mod validation;

pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{CellValueEdit, MAX_BATCH_EDITS, SourceBackedEditor, SourceEdit};
