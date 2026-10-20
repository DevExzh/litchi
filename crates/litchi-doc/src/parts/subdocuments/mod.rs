//! Master-document subdocument metadata: the `PlcfWKB` subdocument directory
//! and the `SttbFnm` referenced-file name table (MS-DOC 2.8.34, 2.9.288,
//! 2.9.346, 2.9.92, and 2.9.93).
//!
//! The `PlcfWKB` lists where each subdocument begins in the main document and
//! references its file through an `FNPI`; the `SttbFnm` stores the full paths
//! of all external files the document references (subdocuments and mail merge
//! data sources) together with per-file `FNIF` metadata.
//!
//! Everything here is inert: file paths are stored verbatim and are never
//! opened, resolved, or followed, and no subdocument content is ever loaded.

mod codec;
mod family;
mod model;
mod package;
mod patch;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use family::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Result};
pub use model::{
    Collection, FileNameKey, FileNameKeyError, FileNameMetadata, Kind, Name, Reference,
};
pub use patch::{Error as PatchError, SourceContext, SourceRanges, TablePatch, TableRange};
pub use transaction::{FileNameSelector, ReferenceSelector, SelectionError};

/// The part-level subdocument codec: the semantic [`record::Snapshot`] of the
/// `SttbFnm`/`PlcfWKB` tables and its [`record::Edit`] staging, which
/// publishes a new semantic snapshot; package edits go through [`Edit`].
pub mod record {
    pub use super::transaction::{Commit, Diagnostics, Edit, Error, Patch, Snapshot};
}
