//! Inert PowerPoint 10 document-comparison metadata.
//!
//! The owner is split by responsibility: typed records live in [`model`],
//! binary record I/O in [`codec`], structural and resource checks in
//! [`validation`], and focused regression coverage in [`tests`]. Parsing
//! never compares presentations, opens external data, or executes embedded
//! content.

mod codec;
mod model;
pub(crate) mod package;
mod transaction;
mod validation;

pub use model::{
    DiffFlags, DiffNode, DiffRecordHeaders, DiffTree10, DiffType, DocDiffFlags, ElementType, Entry,
    Limits, MainMasterDiffFlags, POWERPOINT_DIFF_MAX_DEPTH, POWERPOINT_DIFF_MAX_RECORDS, Review,
    ReviewingToolbarStates, ShapeDiffFlags, SlideCreationEntry, SlideDiffFlags, SlideListTable10,
    TableDiffFlags, TextDiffFlags, Unknown,
};

pub use package::{
    Commit as PackageCommit, MAX_PACKAGE_BYTES, Patch as PackagePatch, Snapshot as PackageSnapshot,
};
pub use transaction::{Change, ChangeKind, Commit, Editor, Patch, Revision, Snapshot};

#[cfg(test)]
mod tests;
