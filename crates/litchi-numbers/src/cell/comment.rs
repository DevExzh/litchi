//! Numbers cell comments: their archive-free values and the selector-first
//! exact-source transaction families.
//!
//! [`crate::Package::edit_table_cell_comment`] stages a cell's comment
//! ([`Edit`](crate::cell::comment::Edit)); [`crate::Package::edit_table_cell_comment_replies`] stages
//! direct replies ([`transaction::Edit`](crate::cell::comment::transaction::Edit)).

pub use crate::package::comments::{
    Comment, CommentAuthor, CommentReply, CommentReplyIndex, CommentTimestamp, Commit, Diagnostics,
    Edit, Error, LimitKind, Patch, Path, Refusal,
};

/// Selector-first exact-source transactions for direct comment replies.
pub mod transaction {
    pub use crate::package::comments::comments_reply::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Refusal,
    };
}
