//! Source-backed, value-only edits for existing worksheet cells.
//!
//! This deliberately narrow capability changes only scalar values already
//! stored in a single worksheet. It never creates cells, changes styles or
//! shared-string topology, authors formulas, or claims to recalculate Excel
//! dependencies. Workbooks outside the statically provable subset are refused.

mod patch;
mod snapshot;
mod source;
mod validation;

pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{CellValueEdit, MAX_BATCH_EDITS, SourceBackedEditor, SourceEdit};
