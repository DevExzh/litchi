//! Source-bound publication of the inert workbook defined-name catalog.

mod patch;
mod snapshot;
mod source;

pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{SourceBackedEditor, SourceEdit};
