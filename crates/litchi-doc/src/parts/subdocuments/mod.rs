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
mod model;
mod package;
mod patch;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    Collection, FileNameKey, FileNameKeyError, FileNameMetadata, Kind, Name, Reference,
};
pub use package::{Editor, PackageCommit, PackageSnapshot};
pub use patch::{PatchError, SourceContext, SourceRanges, TablePatch, TableRange};
pub use transaction::{
    Commit, FileNameSelector, Patch, ReferenceSelector, SelectionError, Snapshot, Transaction,
    TransactionError,
};
