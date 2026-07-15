//! Tracked text revision input types for the legacy Word writer.

use crate::doc::CommentDateTime;

/// Metadata for an inserted or deleted text run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextRevision {
    /// Revision author name.
    pub author: String,
    /// Revision timestamp.
    pub timestamp: Option<CommentDateTime>,
    /// Optional revision identifier.
    pub revision_id: Option<u16>,
}

impl TextRevision {
    /// Create revision metadata for an author.
    pub fn new(author: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            revision_id: None,
        }
    }

    /// Set the revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }

    /// Set the revision identifier.
    pub fn with_id(mut self, revision_id: u16) -> Self {
        self.revision_id = Some(revision_id);
        self
    }
}
