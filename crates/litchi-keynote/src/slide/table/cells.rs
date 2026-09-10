//! Archive-free values returned by a focused Keynote slide-table read.
//!
//! The common table model owns the value and comment vocabulary.  This module
//! gives Keynote callers a discoverable path to those types without exposing
//! native table, tile, list, or comment-storage identifiers.

/// A neutral table-cell value.
pub use litchi_iwa_common::table::cell::value::Value;
/// A checked sparse table coordinate.
pub use litchi_iwa_common::table::coordinate::CellPosition;
/// A positioned semantic table-cell comment.
pub use litchi_iwa_common::table::read::CellComment as TableCellComment;
/// A semantic table-cell comment value.
pub use litchi_iwa_common::table::read::Comment as TableComment;
/// A semantic table-cell comment author.
pub use litchi_iwa_common::table::read::CommentAuthor as TableCommentAuthor;
/// A semantic table-cell comment reply.
pub use litchi_iwa_common::table::read::CommentReply as TableCommentReply;
/// A semantic table-cell comment timestamp.
pub use litchi_iwa_common::table::read::CommentTimestamp as TableCommentTimestamp;
/// A sparse semantic table read, including its positioned cell comments.
pub use litchi_iwa_common::table::read::TableRead;
