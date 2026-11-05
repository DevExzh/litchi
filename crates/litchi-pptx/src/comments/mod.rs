//! Semantic layered owner for `PresentationML` legacy comments.
//!
//! \`model.rs\` contains contextual values, \`codec.rs\` handles bounded XML,
//! and \`package.rs\` owns OPC relationships and transactional CRUD.

mod codec;
pub mod collaboration;
mod model;
mod package;
pub mod part;

pub use codec::{
    parse_comment_authors, parse_slide_comments, parse_slide_comments_strict,
    parse_slide_comments_with_tolerance, write_comment_authors, write_slide_comments,
};
pub(crate) use codec::{parse_comment_authors_processed, parse_slide_comments_processed};
pub use collaboration::{
    AuthorPresenceSnapshot, CommentThreadingSnapshot, ParentComment, PresenceInfo,
    Revision as CollaborationRevision, ThreadingInfo, load_presence, load_presence_snapshot,
    load_threading, load_threading_snapshot,
};
pub use model::{Author, Comment, Comments, Conformance, List};
pub(crate) use package::{
    add_presentation_comment_author_with_limits, add_presentation_comment_with_limits,
    load_presentation_comments_strict_with_limits,
    load_presentation_comments_with_tolerance_and_limits,
    reattribute_presentation_comment_with_limits, remove_presentation_comment_author_with_limits,
    remove_presentation_comment_with_limits, reorder_presentation_comment_authors_with_limits,
    reorder_presentation_comments_with_limits, store_presentation_comments_with_limits,
    update_presentation_comment_author_with_limits, update_presentation_comment_with_limits,
    validate_replay_part,
};
pub use package::{
    find_presentation_comment, find_presentation_comment_author, load_presentation_comments,
    load_presentation_comments_strict, load_presentation_comments_with_tolerance,
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

/// Resource bounds for one raw legacy-comment part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_nodes: usize,
    /// Raw bytes in one XML attribute value, charged before decoding.
    ///
    /// This is the physical source ceiling.  It is deliberately separate
    /// from `max_string_bytes`, which is the selected semantic aggregate and
    /// may be tightened for a replay recipient without rejecting inactive
    /// MCE branches at the physical admission boundary.
    pub max_attribute_bytes: usize,
    pub max_string_bytes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_input_bytes: MAX_PART_BYTES,
        max_output_bytes: MAX_PART_BYTES,
        max_depth: MAX_DEPTH,
        max_nodes: MAX_NODES,
        max_attribute_bytes: MAX_STRING_BYTES,
        max_string_bytes: MAX_STRING_BYTES,
    };

    #[must_use]
    pub const fn from_source(source: litchi_opc::SourceLimits) -> Self {
        Self::DEFAULT.constrained_by_source(source)
    }

    /// Intersect a selected semantic policy with the physical source ceiling.
    #[must_use]
    pub const fn constrained_by_source(self, source: litchi_opc::SourceLimits) -> Self {
        let part = cap_u64(source.max_part_bytes());
        Self {
            max_input_bytes: min_usize(self.max_input_bytes, part),
            max_output_bytes: min_usize(self.max_output_bytes, part),
            max_depth: min_usize(self.max_depth, source.max_depth()),
            max_nodes: min_usize(self.max_nodes, source.max_events()),
            max_attribute_bytes: min_usize(
                self.max_attribute_bytes,
                source.max_xml_attribute_bytes(),
            ),
            max_string_bytes: min_usize(self.max_string_bytes, source.max_xml_attribute_bytes()),
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

const fn cap_u64(value: u64) -> usize {
    if value > usize::MAX as u64 {
        usize::MAX
    } else {
        value as usize
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

mod edit;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
#[cfg(test)]
pub(crate) use package::store_presentation_comments;
