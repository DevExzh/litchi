//! Archive-free semantic values for ranged iWork text comments.
//!
//! Native object traversal, comment-storage decoding, and transactional wire
//! updates remain in the owning format adapter. This leaf contains only the
//! bounded values that an adapter can exchange with Pages, Numbers, and
//! Keynote.

#![allow(
    clippy::module_name_repetitions,
    reason = "TextComment names identify the comment semantic value in this focused module"
)]

use std::num::NonZeroU64;

use litchi_iwa_common::comment::Uuid;

use crate::position::TextRange;

/// Maximum UTF-8 byte length retained by a ranged comment body.
pub const MAX_BODY_BYTES: usize = 64 * 1024;

/// Validation failures produced while constructing ranged-comment values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A root comment identifier was zero.
    ZeroCommentId,
    /// A direct reply identifier was zero.
    ZeroReplyId,
    /// A root comment body was empty.
    EmptyCommentBody,
    /// A root comment body exceeded [`MAX_BODY_BYTES`].
    CommentBodyTooLong,
    /// A direct reply body was empty.
    EmptyReplyBody,
    /// A direct reply body exceeded [`MAX_BODY_BYTES`].
    ReplyBodyTooLong,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::ZeroCommentId => "iWork text comment identifier must be non-zero",
            Self::ZeroReplyId => "iWork text comment reply identifier must be non-zero",
            Self::EmptyCommentBody => "iWork text comment body must not be empty",
            Self::CommentBodyTooLong => "iWork text comment body exceeds 65536 UTF-8 bytes",
            Self::EmptyReplyBody => "iWork text comment reply body must not be empty",
            Self::ReplyBodyTooLong => "iWork text comment reply body exceeds 65536 UTF-8 bytes",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for ranged-comment semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// A compact, non-zero identifier for a ranged text comment annotation.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextCommentId(NonZeroU64);

impl TextCommentId {
    /// Validate a native comment identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroCommentId`] for the native zero sentinel.
    pub fn new(identifier: u64) -> Result<Self> {
        NonZeroU64::new(identifier)
            .map(Self)
            .ok_or(Error::ZeroCommentId)
    }

    /// Construct an identifier obtained from a native object reference.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroCommentId`] for the native zero sentinel.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        Self::new(identifier)
    }

    /// Return the native object identifier.
    #[must_use]
    pub const fn object_id(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for TextCommentId {
    type Error = Error;

    fn try_from(identifier: u64) -> Result<Self> {
        Self::new(identifier)
    }
}

impl From<TextCommentId> for u64 {
    fn from(identifier: TextCommentId) -> Self {
        identifier.object_id()
    }
}

/// A compact, non-zero identifier for a direct reply in a text-comment thread.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextCommentReplyId(NonZeroU64);

impl TextCommentReplyId {
    /// Validate a native reply identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroReplyId`] for the native zero sentinel.
    pub fn new(identifier: u64) -> Result<Self> {
        NonZeroU64::new(identifier)
            .map(Self)
            .ok_or(Error::ZeroReplyId)
    }

    /// Construct an identifier obtained from a native reply reference.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroReplyId`] for the native zero sentinel.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        Self::new(identifier)
    }

    /// Return the native reply-storage object identifier.
    #[must_use]
    pub const fn object_id(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for TextCommentReplyId {
    type Error = Error;

    fn try_from(identifier: u64) -> Result<Self> {
        Self::new(identifier)
    }
}

impl From<TextCommentReplyId> for u64 {
    fn from(identifier: TextCommentReplyId) -> Self {
        identifier.object_id()
    }
}

/// A bounded, nonempty Unicode body of a ranged text comment.
#[repr(transparent)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextCommentBody(Box<str>);

impl TextCommentBody {
    /// Maximum UTF-8 byte length accepted for one root comment body.
    pub const MAX_BYTES: usize = MAX_BODY_BYTES;

    /// Validate and own a borrowed comment body.
    ///
    /// Validation precedes allocation for borrowed input.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `value` is empty or exceeds the bounded
    /// native body size.
    pub fn new(value: &str) -> Result<Self> {
        validate_body(value, false)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed body without another allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `value` is empty or exceeds the bounded
    /// native body size.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_body(&value, false)?;
        Ok(Self(value))
    }

    /// Borrow the exact comment body.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the wrapper and return its owned UTF-8 body.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for TextCommentBody {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for TextCommentBody {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_boxed(value.into_boxed_str())
    }
}

impl TryFrom<Box<str>> for TextCommentBody {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        Self::from_boxed(value)
    }
}

impl TryFrom<&str> for TextCommentBody {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// A bounded, nonempty Unicode body of a direct text-comment reply.
#[repr(transparent)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextCommentReplyBody(Box<str>);

impl TextCommentReplyBody {
    /// Maximum UTF-8 byte length accepted for one reply body.
    pub const MAX_BYTES: usize = MAX_BODY_BYTES;

    /// Validate and own a borrowed reply body.
    ///
    /// Validation precedes allocation for borrowed input.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `value` is empty or exceeds the bounded
    /// native body size.
    pub fn new(value: &str) -> Result<Self> {
        validate_body(value, true)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed body without another allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `value` is empty or exceeds the bounded
    /// native body size.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_body(&value, true)?;
        Ok(Self(value))
    }

    /// Borrow the exact reply body.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the wrapper and return its owned UTF-8 body.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for TextCommentReplyBody {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for TextCommentReplyBody {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_boxed(value.into_boxed_str())
    }
}

impl TryFrom<Box<str>> for TextCommentReplyBody {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        Self::from_boxed(value)
    }
}

impl TryFrom<&str> for TextCommentReplyBody {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// One ranged text comment with optional native metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct TextComment {
    /// Stable semantic comment identity.
    pub id: TextCommentId,
    /// Nonempty half-open UTF-16 range covered by the comment.
    pub range: TextRange,
    /// Root comment body.
    pub body: TextCommentBody,
    /// Seconds since Apple's 2001-01-01 reference date, when stored.
    pub creation_date_seconds: Option<f64>,
    /// Native annotation-author object identifier, when stored.
    pub author_object_id: Option<u64>,
    /// Stable UUID of the owned root comment storage.
    pub storage_uuid: Uuid,
    /// Number of direct replies in the thread.
    pub reply_count: u32,
}

/// One direct reply owned by a ranged text comment.
#[derive(Debug, Clone, PartialEq)]
pub struct TextCommentReply {
    /// Stable semantic reply identity.
    pub id: TextCommentReplyId,
    /// Parent ranged-comment identity.
    pub comment_id: TextCommentId,
    /// Reply body.
    pub body: TextCommentReplyBody,
    /// Seconds since Apple's 2001-01-01 reference date, when stored.
    pub creation_date_seconds: Option<f64>,
    /// Native annotation-author object identifier, when stored.
    pub author_object_id: Option<u64>,
    /// Stable UUID of the reply comment storage.
    pub storage_uuid: Uuid,
}

fn validate_body(value: &str, reply: bool) -> Result<()> {
    if value.is_empty() {
        return Err(if reply {
            Error::EmptyReplyBody
        } else {
            Error::EmptyCommentBody
        });
    }
    if value.len() > MAX_BODY_BYTES {
        return Err(if reply {
            Error::ReplyBodyTooLong
        } else {
            Error::CommentBodyTooLong
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use litchi_iwa_common::comment::Uuid;

    use super::{
        Error, MAX_BODY_BYTES, TextComment, TextCommentBody, TextCommentId, TextCommentReply,
        TextCommentReplyBody, TextCommentReplyId,
    };
    use crate::position::TextRange;

    #[test]
    fn identifiers_are_nonzero_and_compact() {
        assert_eq!(TextCommentId::new(0), Err(Error::ZeroCommentId));
        assert_eq!(TextCommentReplyId::new(0), Err(Error::ZeroReplyId));
        assert_eq!(TextCommentId::new(7).map(TextCommentId::object_id), Ok(7));
        assert_eq!(
            TextCommentReplyId::try_from(9).map(TextCommentReplyId::object_id),
            Ok(9)
        );
        assert_eq!(size_of::<TextCommentId>(), size_of::<u64>());
        assert_eq!(size_of::<TextCommentReplyId>(), size_of::<u64>());
    }

    #[test]
    fn bodies_are_nonempty_bounded_and_allocation_efficient() {
        assert_eq!(TextCommentBody::new(""), Err(Error::EmptyCommentBody));
        assert_eq!(TextCommentReplyBody::new(""), Err(Error::EmptyReplyBody));
        assert_eq!(
            TextCommentBody::new(&"x".repeat(MAX_BODY_BYTES + 1)),
            Err(Error::CommentBodyTooLong)
        );
        assert_eq!(
            TextCommentReplyBody::try_from("x".repeat(MAX_BODY_BYTES + 1)),
            Err(Error::ReplyBodyTooLong)
        );
        let max_body = "x".repeat(MAX_BODY_BYTES);
        assert_eq!(
            TextCommentBody::new(&max_body)
                .as_ref()
                .map(TextCommentBody::as_str),
            Ok(max_body.as_str())
        );

        let body = TextCommentBody::try_from("review 😀".to_owned());
        assert_eq!(body.as_ref().map(TextCommentBody::as_str), Ok("review 😀"));
        assert_eq!(
            body.map(TextCommentBody::into_string),
            Ok("review 😀".to_owned())
        );
        assert_eq!(
            TextCommentReplyBody::from_boxed("reply".to_owned().into_boxed_str())
                .map(|value| value.as_str().to_owned()),
            Ok("reply".to_owned())
        );
    }

    #[test]
    fn records_preserve_ranges_uuid_metadata_and_reply_counts() -> std::result::Result<(), String> {
        let comment = TextComment {
            id: TextCommentId::new(11).map_err(|error| error.to_string())?,
            range: TextRange::from_utf16_indexes(2, 7).map_err(|error| error.to_string())?,
            body: TextCommentBody::new("root").map_err(|error| error.to_string())?,
            creation_date_seconds: Some(42.5),
            author_object_id: Some(13),
            storage_uuid: Uuid::from_parts(14, 15).map_err(|error| error.to_string())?,
            reply_count: 3,
        };
        assert_eq!(comment.id.object_id(), 11);
        assert_eq!(comment.range.start().utf16_index(), 2);
        assert_eq!(comment.range.end().utf16_index(), 7);
        assert_eq!(comment.body.as_str(), "root");
        assert_eq!(comment.creation_date_seconds, Some(42.5));
        assert_eq!(comment.author_object_id, Some(13));
        assert_eq!(
            (comment.storage_uuid.lower(), comment.storage_uuid.upper()),
            (14, 15)
        );
        assert_eq!(comment.reply_count, 3);

        let reply = TextCommentReply {
            id: TextCommentReplyId::new(16).map_err(|error| error.to_string())?,
            comment_id: comment.id,
            body: TextCommentReplyBody::new("reply").map_err(|error| error.to_string())?,
            creation_date_seconds: None,
            author_object_id: None,
            storage_uuid: Uuid::from_parts(17, 18).map_err(|error| error.to_string())?,
        };
        assert_eq!(reply.id.object_id(), 16);
        assert_eq!(reply.comment_id, comment.id);
        assert_eq!(reply.body.as_str(), "reply");
        assert_eq!(reply.creation_date_seconds, None);
        assert_eq!(reply.author_object_id, None);
        assert_eq!(
            (reply.storage_uuid.lower(), reply.storage_uuid.upper()),
            (17, 18)
        );
        Ok(())
    }
}
