//! Strict public types for native ranged iWork text comments.

use crate::{Error, Result};
use litchi_iwa_common::comment::Uuid;

use litchi_iwa_text::position::TextRange;

/// Identifier of a native ranged text-comment annotation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextCommentId(u64);

impl TextCommentId {
    /// Construct an identifier obtained from a previously read comment.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork text-comment object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying package object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// Identifier of one direct reply in a ranged text-comment thread.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextCommentReplyId(u64);

impl TextCommentReplyId {
    /// Construct an identifier obtained from a previously read reply.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork text-comment reply identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying comment-storage object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// Nonempty Unicode body of a native ranged text comment.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextCommentBody(String);

impl TextCommentBody {
    /// Validate and own a comment body without cloning an owned input.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() {
            return Err(Error::ParseError(
                "iWork ranged text comments require nonempty text".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Borrow the comment text.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the wrapper and return its allocation.
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn from_native(value: String) -> Self {
        debug_assert!(!value.is_empty());
        Self(value)
    }
}

/// Nonempty Unicode body of a direct ranged-comment reply.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextCommentReplyBody(String);

impl TextCommentReplyBody {
    /// Validate and own a reply body without cloning an owned input.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() {
            return Err(Error::ParseError(
                "iWork ranged text-comment replies require nonempty text".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Borrow the reply text.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the wrapper and return its allocation.
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn from_native(value: String) -> Self {
        debug_assert!(!value.is_empty());
        Self(value)
    }
}

/// One native comment attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, PartialEq)]
pub struct TextComment {
    /// Stable annotation-object identifier used for update and deletion.
    pub id: TextCommentId,
    /// Half-open commented text range.
    pub range: TextRange,
    /// Root comment body.
    pub body: TextCommentBody,
    /// Seconds since Apple's 2001-01-01 reference date, when stored.
    pub creation_date_seconds: Option<f64>,
    /// Annotation-author object identifier, when stored.
    pub author_object_id: Option<u64>,
    /// Stable UUID of the owned root comment storage.
    pub storage_uuid: Uuid,
    /// Number of direct replies preserved by updates and reclaimed by deletion.
    pub reply_count: usize,
}

/// One direct reply owned by a native ranged text comment.
#[derive(Debug, Clone, PartialEq)]
pub struct TextCommentReply {
    /// Stable comment-storage identifier used for update and deletion.
    pub id: TextCommentReplyId,
    /// Parent ranged-comment annotation identifier.
    pub comment_id: TextCommentId,
    /// Reply body.
    pub body: TextCommentReplyBody,
    /// Seconds since Apple's 2001-01-01 reference date, when stored.
    pub creation_date_seconds: Option<f64>,
    /// Annotation-author object identifier, when stored.
    pub author_object_id: Option<u64>,
    /// Stable UUID of the reply comment storage.
    pub storage_uuid: Uuid,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identifiers_and_bodies_reject_empty_sentinels() {
        assert!(TextCommentId::from_object_id(0).is_err());
        assert!(TextCommentReplyId::from_object_id(0).is_err());
        assert!(TextCommentBody::new("").is_err());
        assert!(TextCommentReplyBody::new("").is_err());
        assert_eq!(TextCommentBody::new("note").unwrap().as_str(), "note");
        assert_eq!(
            TextCommentReplyBody::new("reply").unwrap().as_str(),
            "reply"
        );
    }
}
