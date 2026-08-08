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
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Opaque IDs keep their raw adapter module adjacent to private representations."
)]

use std::num::NonZeroU64;

use litchi_iwa_common::comment::{AuthorId, Uuid};

use crate::date_time::Instant;
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

/// Explicit native-boundary conversions for the opaque comment handle.
#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Keep the explicit raw adapter boundary adjacent to its opaque handles."
)]
pub mod raw {
    use super::{Result, TextCommentId, TextCommentReplyId};

    /// Validate a native ranged-comment object identifier.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::ZeroCommentId`] for the native zero sentinel.
    pub const fn comment_id(identifier: u64) -> Result<TextCommentId> {
        TextCommentId::from_raw(identifier)
    }

    /// Recover a native ranged-comment object identifier inside an adapter.
    #[must_use]
    pub const fn comment_id_value(identifier: TextCommentId) -> u64 {
        identifier.into_raw()
    }

    /// Validate a native direct-reply object identifier.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::ZeroReplyId`] for the native zero sentinel.
    pub const fn reply_id(identifier: u64) -> Result<TextCommentReplyId> {
        TextCommentReplyId::from_raw(identifier)
    }

    /// Recover a native direct-reply object identifier inside an adapter.
    #[must_use]
    pub const fn reply_id_value(identifier: TextCommentReplyId) -> u64 {
        identifier.into_raw()
    }
}

impl TextCommentId {
    const fn from_raw(identifier: u64) -> Result<Self> {
        match NonZeroU64::new(identifier) {
            Some(non_zero) => Ok(Self(non_zero)),
            None => Err(Error::ZeroCommentId),
        }
    }

    const fn into_raw(self) -> u64 {
        self.0.get()
    }
}

/// A compact, non-zero identifier for a direct reply in a text-comment thread.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextCommentReplyId(NonZeroU64);

impl TextCommentReplyId {
    const fn from_raw(identifier: u64) -> Result<Self> {
        match NonZeroU64::new(identifier) {
            Some(non_zero) => Ok(Self(non_zero)),
            None => Err(Error::ZeroReplyId),
        }
    }

    const fn into_raw(self) -> u64 {
        self.0.get()
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

/// Metadata retained for a ranged text comment without exposing archive IDs.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Metadata {
    creation_date: Option<Instant>,
    author: Option<AuthorId>,
    storage_uuid: Uuid,
}

impl Metadata {
    /// Construct validated comment metadata.
    #[must_use]
    pub const fn new(
        creation_date: Option<Instant>,
        author: Option<AuthorId>,
        storage_uuid: Uuid,
    ) -> Self {
        Self {
            creation_date,
            author,
            storage_uuid,
        }
    }

    /// Return the optional Apple reference-date creation instant.
    #[must_use]
    pub const fn creation_date(self) -> Option<Instant> {
        self.creation_date
    }

    /// Return the optional typed author identity.
    #[must_use]
    pub const fn author(self) -> Option<AuthorId> {
        self.author
    }

    /// Return the stable semantic comment-storage UUID.
    #[must_use]
    pub const fn storage_uuid(self) -> Uuid {
        self.storage_uuid
    }
}

/// One ranged text comment with validated semantic metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct TextComment {
    /// Stable semantic comment identity.
    id: TextCommentId,
    /// Nonempty half-open UTF-16 range covered by the comment.
    range: TextRange,
    /// Root comment body.
    body: TextCommentBody,
    metadata: Metadata,
    reply_count: u32,
}

impl TextComment {
    /// Construct a comment from validated semantic values.
    #[must_use]
    pub const fn new(
        id: TextCommentId,
        range: TextRange,
        body: TextCommentBody,
        metadata: Metadata,
        reply_count: u32,
    ) -> Self {
        Self {
            id,
            range,
            body,
            metadata,
            reply_count,
        }
    }

    /// Return the opaque semantic comment identity.
    #[must_use]
    pub const fn id(&self) -> TextCommentId {
        self.id
    }

    /// Return the nonempty UTF-16 range covered by the comment.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

    /// Borrow the validated root-comment body.
    #[must_use]
    pub const fn body(&self) -> &TextCommentBody {
        &self.body
    }

    /// Return validated comment metadata.
    #[must_use]
    pub const fn metadata(&self) -> Metadata {
        self.metadata
    }

    /// Return the optional Apple reference-date creation instant.
    #[must_use]
    pub const fn creation_date(&self) -> Option<Instant> {
        self.metadata.creation_date()
    }

    /// Return the optional typed author identity.
    #[must_use]
    pub const fn author(&self) -> Option<AuthorId> {
        self.metadata.author()
    }

    /// Return the stable semantic comment-storage UUID.
    #[must_use]
    pub const fn storage_uuid(&self) -> Uuid {
        self.metadata.storage_uuid()
    }

    /// Return the number of direct replies in the thread.
    #[must_use]
    pub const fn reply_count(&self) -> u32 {
        self.reply_count
    }
}

/// One direct reply owned by a ranged text comment.
#[derive(Debug, Clone, PartialEq)]
pub struct TextCommentReply {
    /// Stable semantic reply identity.
    id: TextCommentReplyId,
    /// Parent ranged-comment identity.
    comment_id: TextCommentId,
    /// Reply body.
    body: TextCommentReplyBody,
    metadata: Metadata,
}

impl TextCommentReply {
    /// Construct a reply from validated semantic values.
    #[must_use]
    pub const fn new(
        id: TextCommentReplyId,
        comment_id: TextCommentId,
        body: TextCommentReplyBody,
        metadata: Metadata,
    ) -> Self {
        Self {
            id,
            comment_id,
            body,
            metadata,
        }
    }

    /// Return the opaque semantic reply identity.
    #[must_use]
    pub const fn id(&self) -> TextCommentReplyId {
        self.id
    }

    /// Return the parent ranged-comment identity.
    #[must_use]
    pub const fn comment_id(&self) -> TextCommentId {
        self.comment_id
    }

    /// Borrow the validated reply body.
    #[must_use]
    pub const fn body(&self) -> &TextCommentReplyBody {
        &self.body
    }

    /// Return validated reply metadata.
    #[must_use]
    pub const fn metadata(&self) -> Metadata {
        self.metadata
    }

    /// Return the optional Apple reference-date creation instant.
    #[must_use]
    pub const fn creation_date(&self) -> Option<Instant> {
        self.metadata.creation_date()
    }

    /// Return the optional typed author identity.
    #[must_use]
    pub const fn author(&self) -> Option<AuthorId> {
        self.metadata.author()
    }

    /// Return the stable semantic comment-storage UUID.
    #[must_use]
    pub const fn storage_uuid(&self) -> Uuid {
        self.metadata.storage_uuid()
    }
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
    use std::result::Result as TestResult;

    use litchi_iwa_common::comment::{AuthorId, Uuid};

    use super::{
        Error, MAX_BODY_BYTES, Metadata, TextComment, TextCommentBody, TextCommentId,
        TextCommentReply, TextCommentReplyBody, TextCommentReplyId, raw,
    };
    use crate::date_time::Instant;
    use crate::position::TextRange;

    #[test]
    fn identifiers_are_nonzero_and_compact() {
        assert_eq!(raw::comment_id(0), Err(Error::ZeroCommentId));
        assert_eq!(raw::reply_id(0), Err(Error::ZeroReplyId));
        assert_eq!(raw::comment_id_value(raw::comment_id(7).unwrap()), 7);
        assert_eq!(raw::reply_id_value(raw::reply_id(9).unwrap()), 9);
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
    fn records_preserve_ranges_uuid_metadata_and_reply_counts() -> TestResult<(), String> {
        let comment = TextComment::new(
            raw::comment_id(11).map_err(|error| error.to_string())?,
            TextRange::from_utf16_indexes(2, 7).map_err(|error| error.to_string())?,
            TextCommentBody::new("root").map_err(|error| error.to_string())?,
            Metadata::new(
                Some(Instant::from_reference_date_seconds(42.5).unwrap()),
                Some(AuthorId::new(13).unwrap()),
                Uuid::from_parts(14, 15).map_err(|error| error.to_string())?,
            ),
            3,
        );
        assert_eq!(raw::comment_id_value(comment.id()), 11);
        assert_eq!(comment.range().start().utf16_index(), 2);
        assert_eq!(comment.range().end().utf16_index(), 7);
        assert_eq!(comment.body().as_str(), "root");
        assert_eq!(
            comment.creation_date().map(Instant::reference_date_seconds),
            Some(42.5)
        );
        assert_eq!(comment.author().map(AuthorId::get), Some(13));
        assert_eq!(
            (
                comment.storage_uuid().lower(),
                comment.storage_uuid().upper()
            ),
            (14, 15)
        );
        assert_eq!(comment.reply_count(), 3);

        let reply = TextCommentReply::new(
            raw::reply_id(16).map_err(|error| error.to_string())?,
            comment.id(),
            TextCommentReplyBody::new("reply").map_err(|error| error.to_string())?,
            Metadata::new(
                None,
                None,
                Uuid::from_parts(17, 18).map_err(|error| error.to_string())?,
            ),
        );
        assert_eq!(raw::reply_id_value(reply.id()), 16);
        assert_eq!(reply.comment_id(), comment.id());
        assert_eq!(reply.body().as_str(), "reply");
        assert_eq!(reply.creation_date(), None);
        assert_eq!(reply.author(), None);
        assert_eq!(
            (reply.storage_uuid().lower(), reply.storage_uuid().upper()),
            (17, 18)
        );
        Ok(())
    }
}
