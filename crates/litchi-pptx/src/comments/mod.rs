//! Semantic layered owner for PresentationML legacy comments.
//!
//! \`model.rs\` contains contextual values, \`codec.rs\` handles bounded XML,
//! and \`package.rs\` owns OPC relationships and transactional CRUD.

mod codec;
mod model;
mod package;
pub mod part;

pub use codec::{
    parse_comment_authors, parse_slide_comments, write_comment_authors, write_slide_comments,
};
pub use model::{Author, Comment, Comments, Conformance, List};
pub use package::{
    add_presentation_comment, add_presentation_comment_author, find_presentation_comment,
    find_presentation_comment_author, load_presentation_comments, remove_presentation_comment,
    remove_presentation_comment_author, reorder_presentation_comment_authors,
    reorder_presentation_comments, replace_presentation_comment,
    replace_presentation_comment_author, store_presentation_comments, update_presentation_comment,
    update_presentation_comment_author,
};
pub use part::{AuthorListPart, ListPart};

pub(crate) const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
pub(crate) const STRICT_COMMENTS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
pub(crate) const AUTHORS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
pub(crate) const STRICT_AUTHORS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors";
pub(crate) const SLIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
pub(crate) const STRICT_SLIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
pub(crate) const AUTHORS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
pub(crate) const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
pub(crate) const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
pub(crate) const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_NODES: usize = 262_144;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_AUTHORS: usize = 4096;
pub(crate) const MAX_SLIDES: usize = 100_000;
pub(crate) const MAX_COMMENTS_PER_SLIDE: usize = 65_536;
pub(crate) const MAX_TOTAL_COMMENTS: usize = 1_000_000;

pub(crate) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}
