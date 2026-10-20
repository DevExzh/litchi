//! Inert `PowerPoint` 10 document-comparison metadata.
//!
//! The owner is split by responsibility: typed records live in `model`,
//! binary record I/O in `codec`, structural and resource checks in
//! `validation`, and focused regression coverage in `tests`. Parsing
//! never compares presentations, opens external data, or executes embedded
//! content.
//!
//! [`crate::Package::edit_document_comparison`] stages edits of the inert
//! review metadata of the live document; see [`Edit`].

mod codec;
mod edit;
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

pub use edit::Edit;
pub use package::{MAX_PACKAGE_BYTES, Snapshot as PackageSnapshot};
pub use transaction::{Change, ChangeKind, Revision, Snapshot};

crate::scoped::family!("document-comparison");

#[cfg(test)]
mod tests;
