//! Threaded comments module for XLSX files.
//!
//! This module provides structures and functions for reading and writing
//! threaded comments (modern Excel comment threads) in XLSX workbooks.
//!
//! Threaded comments are a modern feature introduced in Office 365 that
//! support conversation threads, @mentions, and richer collaboration features.

pub mod person;
pub mod reader;
pub mod writer;

pub use person::{Mention, Person, PersonList};
pub use reader::{read_persons, read_threaded_comments};
pub use writer::{write_persons, write_threaded_comments};

/// A threaded comment in an Excel worksheet.
///
/// Threaded comments support conversation-style threads with replies,
/// mentions, timestamps, and resolution status.
#[derive(Debug, Clone, Default)]
pub struct ThreadedComment {
    /// Cell reference (e.g., "A1")
    pub cell_ref: Option<String>,
    /// Unique identifier for this comment
    pub id: String,
    /// ID of the parent comment (for replies)
    pub parent_id: Option<String>,
    /// Person ID who authored this comment
    pub person_id: String,
    /// Comment text content
    pub text: Option<String>,
    /// Timestamp when comment was created/edited
    pub date_time: Option<String>,
    /// Whether this comment thread is marked as done/resolved
    pub done: Option<bool>,
    /// List of @mentions in the comment
    pub mentions: Vec<Mention>,
}

/// Collection of threaded comments for a worksheet.
#[derive(Debug, Clone, Default)]
pub struct ThreadedComments {
    /// List of threaded comments
    pub comments: Vec<ThreadedComment>,
}
