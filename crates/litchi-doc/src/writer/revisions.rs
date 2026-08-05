//! Tracked text revision input types for the legacy Word writer.

use crate::{CommentDateTime, RevisionReason};

/// Metadata for an inserted or deleted text run.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextRevision {
    /// Revision author name.
    pub author: String,
    /// Revision timestamp.
    pub timestamp: Option<CommentDateTime>,
    /// Legacy raw edit-reason code.
    pub revision_id: Option<u16>,
    /// Structured edit reason.
    pub reason: Option<RevisionReason>,
    /// ECMA-376 single-session revision-save ID.
    pub revision_save_id: Option<u32>,
}

impl TextRevision {
    /// Create revision metadata for an author.
    pub fn new(author: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            revision_id: None,
            reason: None,
            revision_save_id: None,
        }
    }

    /// Set the revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }

    /// Set a legacy raw edit-reason code.
    ///
    /// Prefer [`Self::with_reason`]; invalid raw values are rejected when writing.
    pub fn with_id(mut self, revision_id: u16) -> Self {
        self.revision_id = Some(revision_id);
        self
    }

    /// Set the edit reason.
    pub fn with_reason(mut self, reason: RevisionReason) -> Self {
        self.reason = Some(reason);
        self
    }

    /// Set the single-session revision-save ID.
    pub fn with_revision_save_id(mut self, revision_save_id: u32) -> Self {
        self.revision_save_id = Some(revision_save_id);
        self
    }
}

/// Metadata for a tracked character-formatting change.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormattingRevision {
    /// Revision author name.
    pub author: String,
    /// Revision timestamp.
    pub timestamp: Option<CommentDateTime>,
    /// ECMA-376 single-session revision-save ID for the formatting change.
    pub revision_save_id: Option<u32>,
    /// Edit reason for the formatting change.
    pub reason: Option<RevisionReason>,
}

impl FormattingRevision {
    /// Create formatting-revision metadata for an author.
    pub fn new(author: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            revision_save_id: None,
            reason: None,
        }
    }

    /// Set the revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }

    /// Set the single-session revision-save ID.
    pub fn with_revision_save_id(mut self, revision_save_id: u32) -> Self {
        self.revision_save_id = Some(revision_save_id);
        self
    }

    /// Set the edit reason.
    pub fn with_reason(mut self, reason: RevisionReason) -> Self {
        self.reason = Some(reason);
        self
    }
}

/// Numbering state retained for a tracked paragraph numbering change.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingRevision {
    /// Revision author name.
    pub author: String,
    /// Revision timestamp.
    pub timestamp: Option<CommentDateTime>,
    /// Whether the paragraph was already numbered when tracking began.
    pub was_numbered: bool,
    /// Placeholder positions for the nine numbering levels.
    pub placeholder_positions: [u8; 9],
    /// MSONFC values for the nine numbering levels.
    pub number_formats: [u8; 9],
    /// Numeric values for the nine numbering levels.
    pub numbers: [u32; 9],
    /// Numbering format string.
    pub format_string: String,
}

impl NumberingRevision {
    /// Create numbering revision metadata for an author and format string.
    pub fn new(author: impl Into<String>, format_string: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            was_numbered: false,
            placeholder_positions: [0; 9],
            number_formats: [0; 9],
            numbers: [0; 9],
            format_string: format_string.into(),
        }
    }

    /// Set the revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }
}

/// Revision metadata for a LISTNUM display-field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DisplayFieldRevision {
    /// Revision author name.
    pub author: String,
    /// Revision timestamp.
    pub timestamp: Option<CommentDateTime>,
    /// Previous LISTNUM display-field result.
    pub previous_result: String,
}

impl DisplayFieldRevision {
    /// Create display-field revision metadata.
    pub fn new(author: impl Into<String>, previous_result: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            previous_result: previous_result.into(),
        }
    }

    /// Set the revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }
}
