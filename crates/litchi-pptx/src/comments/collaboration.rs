//! Source-bound, inert `p15` collaboration metadata on legacy comments.
//!
//! The 2012 PresentationML collaboration extensions are persisted inside the
//! existing legacy comment-author and comment parts.  They identify users and
//! parent comments, but this crate never resolves identities, contacts a
//! provider, or performs collaboration.  The source-bound snapshots below
//! keep those values as ordinary bounded data and splice only the selected
//! extension when an edit is published.

mod codec;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use package::{
    apply_presence_commit, apply_presence_patch, apply_threading_commit, apply_threading_patch,
    load_presence, load_presence_snapshot, load_threading, load_threading_snapshot, put_presence,
    put_threading, remove_presence, remove_threading,
};
pub use transaction::{
    AuthorPresenceCommit, AuthorPresencePatch, AuthorPresenceSnapshot, AuthorPresenceTransaction,
    CommentThreadingCommit, CommentThreadingPatch, CommentThreadingSnapshot,
    CommentThreadingTransaction, Revision,
};

/// The namespace introduced by [MS-PPTX] 2.4 for these extensions.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";

/// The extension URI for `p15:presenceInfo` under `cmAuthor/extLst`.
pub const PRESENCE_EXTENSION_URI: &str = "{19B8F6BF-5375-455C-9EA6-DF929625EA0E}";

/// The extension URI for `p15:threadingInfo` under `cm/extLst`.
pub const THREADING_EXTENSION_URI: &str = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

/// A provider-issued identifier retained as inert document metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresenceInfo {
    pub user_id: String,
    pub provider_id: String,
}

impl PresenceInfo {
    /// Construct a bounded presence value.  The strings are inert: no lookup
    /// or provider access is performed.
    #[must_use]
    pub fn new(user_id: impl Into<String>, provider_id: impl Into<String>) -> Self {
        Self {
            user_id: user_id.into(),
            provider_id: provider_id.into(),
        }
    }

    /// Return the provider-issued user identifier.
    #[inline]
    #[must_use]
    pub fn user_id(&self) -> &str {
        &self.user_id
    }

    /// Return the inert identity-provider identifier.
    #[inline]
    #[must_use]
    pub fn provider_id(&self) -> &str {
        &self.provider_id
    }
}

/// The author/index pair identifying a parent legacy comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParentComment {
    pub author_id: Option<u32>,
    pub index: Option<u32>,
}

impl ParentComment {
    /// Construct an optional-field parent identifier exactly as represented by
    /// `CT_ParentCommentIdentifier`.
    #[must_use]
    pub const fn new(author_id: Option<u32>, index: Option<u32>) -> Self {
        Self { author_id, index }
    }

    /// Return the optional parent author identifier.
    #[inline]
    #[must_use]
    pub const fn author_id(&self) -> Option<u32> {
        self.author_id
    }

    /// Return the optional parent comment index.
    #[inline]
    #[must_use]
    pub const fn index(&self) -> Option<u32> {
        self.index
    }

    /// Alias using the XML attribute spelling.
    #[inline]
    #[must_use]
    pub const fn idx(&self) -> Option<u32> {
        self.index
    }
}

/// Typed, inert threading metadata attached to one legacy comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThreadingInfo {
    pub time_zone_bias: Option<i32>,
    pub parent: Option<ParentComment>,
}

impl ThreadingInfo {
    /// Construct threading metadata while retaining omission of both optional
    /// schema fields.
    #[must_use]
    pub const fn new(time_zone_bias: Option<i32>, parent: Option<ParentComment>) -> Self {
        Self {
            time_zone_bias,
            parent,
        }
    }

    /// Return the optional UTC bias in minutes.
    #[inline]
    #[must_use]
    pub const fn time_zone_bias(&self) -> Option<i32> {
        self.time_zone_bias
    }

    /// Return the optional parent-comment identifier.
    #[inline]
    #[must_use]
    pub fn parent(&self) -> Option<&ParentComment> {
        self.parent.as_ref()
    }

    /// Alias using the XML element spelling.
    #[inline]
    #[must_use]
    pub fn parent_comment(&self) -> Option<&ParentComment> {
        self.parent()
    }
}
