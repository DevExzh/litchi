//! Classic SpreadsheetML comments owned by one worksheet.
//!
//! The owner is layered by responsibility: the semantic author/comment graph
//! lives in [`model`], bounded XML conversion in [`codec`], and worksheet OPC
//! relationship lifecycle in [`package`]. This module intentionally does not
//! share models with [`crate::threaded_comments`].

mod codec;
mod model;
mod package;

pub use codec::{parse_comments, validate_comments, write_comments};
pub use model::{Comment, Comments, Part};
pub use package::{
    COMMENTS_CONTENT_TYPE, COMMENTS_RELATIONSHIP_TYPE, STRICT_COMMENTS_RELATIONSHIP_TYPE,
    load_from_worksheet, remove_from_worksheet, replace_on_worksheet, store_on_worksheet,
    validate_graph,
};

pub(crate) const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_AUTHORS: usize = 100_000;
pub(crate) const MAX_COMMENTS: usize = 100_000;
pub(crate) const MAX_TEXT_BYTES: usize = 1_048_576;
