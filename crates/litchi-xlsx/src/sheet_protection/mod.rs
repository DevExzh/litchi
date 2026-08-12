//! Immutable worksheet-protection metadata for `SpreadsheetML` worksheets.
//!
//! The owner is layered by semantic responsibility: typed protection models,
//! bounded XML codecs, and focused regression coverage.

mod codec;
mod model;
mod patch;
mod snapshot;
mod source;

#[cfg(test)]
mod tests;

pub use codec::{
    parse_protection, replace_protection, validate_metadata, write_core, write_extensions,
    write_protection,
};
pub use model::*;
pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;
pub use source::{SourceBackedEditor, SourceEdit};
