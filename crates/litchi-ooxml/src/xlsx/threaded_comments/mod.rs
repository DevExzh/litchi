//! Threaded comments module for XLSX files.
//!
//! This module provides structures and functions for reading and writing
//! threaded comments (modern Excel comment threads) in XLSX workbooks.
//!
//! Threaded comments are a modern feature introduced in Office 365 that
//! support conversation threads, @mentions, and richer collaboration features.

pub mod package;
pub mod person;
pub mod reader;
pub mod writer;

pub use package::{
    ThreadedCommentGraph, WorkbookPersonPart, WorksheetThreadedCommentPart, add_threaded_comment,
    add_threaded_comment_person, add_threaded_comment_reply, find_threaded_comment,
    find_threaded_comment_person, load_threaded_comment_graph, remove_threaded_comment,
    remove_threaded_comment_person, reorder_threaded_comment_persons, reorder_threaded_comments,
    replace_threaded_comment, replace_threaded_comment_person, update_threaded_comment,
    update_threaded_comment_person, validate_threaded_comment_graph,
};
pub use person::{Mention, Person, PersonList};
pub use reader::{read_persons, read_threaded_comments};
pub use writer::{write_persons, write_threaded_comments};

pub(crate) const MAX_THREADED_PART_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_THREADED_PERSONS: usize = 100_000;
pub(crate) const MAX_THREADED_COMMENTS: usize = 100_000;
pub(crate) const MAX_THREADED_MENTIONS: usize = 100_000;
pub(crate) const MAX_THREADED_TEXT_UTF16: usize = 1_048_576;
pub(crate) const MAX_THREADED_IDENTITY_BYTES: usize = 16_384;

pub(crate) fn validate_threaded_timestamp(value: Option<&str>) -> litchi_core::sheet::Result<()> {
    use chrono::{DateTime, NaiveDateTime};

    let Some(value) = value else {
        return Ok(());
    };
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(format!("invalid threaded-comment timestamp '{value}'").into())
    }
}

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
