//! Canonical layered owner for PowerPoint 2018 modern comments.

mod codec;
mod model;
mod package;

pub use model::{
    Anchor, AnchorKind, Author, AuthorPart, Authors, Comment, Graph, List, ModernComment,
    ModernCommentAnchor, ModernCommentAnchorKind, ModernCommentAuthor, ModernCommentAuthorList,
    ModernCommentAuthorPart, ModernCommentGraph, ModernCommentList,
    ModernCommentNamespaceDeclaration, ModernCommentPart, ModernCommentPosition,
    ModernCommentReply, ModernCommentStatus, NamespaceDeclaration, Part, Position, Progress, Reply,
    Status,
};
pub use package::{
    add_modern_comment, add_modern_comment_author, add_modern_comment_reply, find_modern_comment,
    find_modern_comment_author, find_modern_comment_reply, load_modern_comment_authors,
    load_modern_comment_graph, load_modern_comments, remove_modern_comment,
    remove_modern_comment_author, remove_modern_comment_reply, reorder_modern_comment_authors,
    reorder_modern_comments, replace_modern_comment, replace_modern_comment_author,
    replace_modern_comment_reply, store_modern_comment, store_modern_comment_authors,
    update_modern_comment, update_modern_comment_author, update_modern_comment_reply,
    validate_modern_comment_author_references,
};

pub const MODERN_COMMENT_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.comments+xml";
pub const MODERN_COMMENT_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/comments";
pub const MODERN_COMMENT_AUTHOR_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.authors+xml";
pub const MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/authors";

pub(crate) const P188: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
pub(crate) const PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
pub(crate) const AC: &str = "http://schemas.microsoft.com/office/drawing/2013/main/command";
pub(crate) const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
pub(crate) const MAX_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 192;
pub(crate) const MAX_NODES: usize = 250_000;
pub(crate) const MAX_COMMENTS: usize = 100_000;
pub(crate) const MAX_REPLIES: usize = 100_000;
pub(crate) const MAX_AUTHORS: usize = 65_536;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
