//! Archive-free values for direct comments attached to Keynote drawables.
//!
//! Native comment-storage identifiers, author objects, protobuf messages, and
//! archive locations remain private to the package adapter.  These values are
//! snapshots: a reply ordinal is supplied separately by the package API when
//! a caller wants to edit an existing thread.

use std::fmt;

/// Semantic kind of a drawable that can own a direct comment.
///
/// This is intentionally an application-level vocabulary.  Native message
/// types that are not recognized by the focused adapter remain opaque and are
/// omitted from inventory rather than being surfaced as an unstable catch-all.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum DrawableKind {
    /// A generic drawable envelope.
    Drawable,
    /// A shape or shape-info drawable.
    Shape,
    /// An image drawable.
    Image,
    /// A mask drawable.
    Mask,
    /// A movie or audio drawable.
    Movie,
    /// A group drawable.
    Group,
    /// A connection line drawable.
    ConnectionLine,
    /// A placeholder drawable.
    Placeholder,
    /// A chart drawable.
    Chart,
    /// A table drawable.
    Table,
    /// A word-processing table drawable.
    WordProcessingTable,
}

/// A checked finite comment timestamp measured in seconds from Apple's
/// 2001-01-01 reference date.
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct CommentTimestamp(f64);

impl CommentTimestamp {
    /// Construct a timestamp from Apple-epoch seconds.
    ///
    /// `None` is returned for NaN and infinite values.  Negative values are
    /// accepted because the native format permits historical dates before the
    /// reference epoch.
    #[must_use]
    pub const fn new(seconds: f64) -> Option<Self> {
        if seconds.is_finite() {
            Some(Self(seconds))
        } else {
            None
        }
    }

    /// Return Apple-epoch seconds.
    #[must_use]
    pub const fn as_f64(self) -> f64 {
        self.0
    }
}

impl fmt::Display for CommentTimestamp {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

/// Display metadata for a comment author.
///
/// The focused API intentionally carries display data only.  Native author
/// object identifiers and storage references never cross the package boundary.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CommentAuthor {
    display_name: Option<Box<str>>,
    public_id: Option<Box<str>>,
}

impl CommentAuthor {
    /// Construct display metadata with an optional display name and public
    /// author identifier.
    #[must_use]
    pub fn new(display_name: Option<Box<str>>, public_id: Option<Box<str>>) -> Self {
        Self {
            display_name,
            public_id,
        }
    }

    /// Return the optional display name.
    #[must_use]
    pub fn display_name(&self) -> Option<&str> {
        self.display_name.as_deref()
    }

    /// Return the optional semantic public author identifier.
    #[must_use]
    pub fn public_id(&self) -> Option<&str> {
        self.public_id.as_deref()
    }
}

/// An immutable semantic snapshot of one direct comment or reply.
#[derive(Clone, Debug, PartialEq)]
pub struct Comment {
    text: Box<str>,
    timestamp: Option<CommentTimestamp>,
    author: Option<CommentAuthor>,
}

impl Comment {
    /// Construct a comment snapshot with no optional metadata.
    #[must_use]
    pub fn new(text: impl Into<Box<str>>) -> Self {
        Self {
            text: text.into(),
            timestamp: None,
            author: None,
        }
    }

    /// Construct a comment snapshot from its semantic fields.
    #[must_use]
    pub fn with_metadata(
        text: impl Into<Box<str>>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Self {
        Self {
            text: text.into(),
            timestamp,
            author,
        }
    }

    /// Borrow the comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the optional creation timestamp.
    #[must_use]
    pub const fn timestamp(&self) -> Option<CommentTimestamp> {
        self.timestamp
    }

    /// Borrow the optional author display metadata.
    #[must_use]
    pub fn author(&self) -> Option<&CommentAuthor> {
        self.author.as_ref()
    }
}

/// An immutable semantic snapshot of one ordered direct reply.
#[derive(Clone, Debug, PartialEq)]
pub struct Reply {
    text: Box<str>,
    timestamp: Option<CommentTimestamp>,
    author: Option<CommentAuthor>,
}

impl Reply {
    /// Construct a reply snapshot with no optional metadata.
    #[must_use]
    pub fn new(text: impl Into<Box<str>>) -> Self {
        Self {
            text: text.into(),
            timestamp: None,
            author: None,
        }
    }

    /// Construct a reply snapshot from its semantic fields.
    #[must_use]
    pub fn with_metadata(
        text: impl Into<Box<str>>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Self {
        Self {
            text: text.into(),
            timestamp,
            author,
        }
    }

    /// Borrow the reply text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the optional creation timestamp.
    #[must_use]
    pub const fn timestamp(&self) -> Option<CommentTimestamp> {
        self.timestamp
    }

    /// Borrow the optional author display metadata.
    #[must_use]
    pub fn author(&self) -> Option<&CommentAuthor> {
        self.author.as_ref()
    }
}

#[cfg(test)]
mod tests {
    use super::{Comment, CommentAuthor, CommentTimestamp, Reply};

    #[test]
    fn timestamp_rejects_non_finite_values() {
        assert_eq!(CommentTimestamp::new(f64::NAN), None);
        assert_eq!(CommentTimestamp::new(f64::INFINITY), None);
        assert_eq!(
            CommentTimestamp::new(-1.0).map(CommentTimestamp::as_f64),
            Some(-1.0)
        );
    }

    #[test]
    fn snapshots_preserve_text_and_display_metadata_without_ids() {
        let author = CommentAuthor::new(Some("Ada".into()), Some("public-id".into()));
        let comment =
            Comment::with_metadata("root", CommentTimestamp::new(42.5), Some(author.clone()));
        let reply = Reply::with_metadata("reply", None, Some(author));

        assert_eq!(comment.text(), "root");
        assert_eq!(
            comment.timestamp().map(CommentTimestamp::as_f64),
            Some(42.5)
        );
        assert_eq!(
            comment.author().and_then(CommentAuthor::display_name),
            Some("Ada")
        );
        assert_eq!(reply.text(), "reply");
    }
}
