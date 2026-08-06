//! Classic SpreadsheetML comments owned by one worksheet.
//!
//! The owner is layered by responsibility: the semantic author/comment graph
//! lives in [`model`], bounded XML conversion in [`codec`], and worksheet OPC
//! relationship lifecycle in [`package`]. Immutable source snapshots and
//! failure-atomic semantic edits live in [`snapshot`], [`transaction`], and
//! [`patch`]. This module intentionally does not share models with
//! [`crate::threaded_comments`]. VML shape IDs are retained as inert note
//! metadata; this owner never allocates, interprets, or rewrites VML shapes.

pub mod codec;
pub mod model;
pub mod package;
pub mod patch;
pub mod snapshot;
pub mod transaction;
pub mod validation;

pub use codec::{parse_comments, validate_comments, write_comments};
pub use model::{Comment, Comments, Part};
pub use package::{
    COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, STRICT_COMMENTS_RELATIONSHIP_TYPE,
    load_from_worksheet, remove_from_worksheet, replace_on_worksheet, store_on_worksheet,
    validate_graph,
};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

/// Validate a complete classic-comments graph without mutating it.
pub fn validate(value: &Comments) -> crate::Result<()> {
    validation::comments(value)
}

pub(crate) const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_AUTHORS: usize = 100_000;
pub(crate) const MAX_COMMENTS: usize = 100_000;
pub(crate) const MAX_TEXT_BYTES: usize = 1_048_576;

#[cfg(test)]
mod tests;
