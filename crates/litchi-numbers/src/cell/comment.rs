//! Archive-free Numbers cell-comment values.
//!
//! Comment reply text and reply ordinals are semantic values. The package
//! transaction namespace is re-exported here so callers do not need to know
//! that the exact-source owner lives under the package adapter.

pub use crate::package::comments::CommentReply;
pub use crate::package::comments::CommentReplyIndex;

/// Selector-first exact-source transactions for direct comment replies.
pub mod transaction {
    pub use crate::package::comments::{
        CommentReplyCommit, CommentReplyDiagnostics, CommentReplyEdit, CommentReplyError,
        CommentReplyLimitKind, CommentReplyPatch, CommentReplyPath,
    };
    pub use crate::package::comments::{
        CommentReplyCommit as Commit, CommentReplyDiagnostics as Diagnostics,
        CommentReplyEdit as Edit, CommentReplyError as Error, CommentReplyLimitKind as LimitKind,
        CommentReplyPatch as Patch, CommentReplyPath as Path,
    };
}
