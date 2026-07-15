use super::comment::CommentDateTime;
use super::package::{DocError, Result};

/// Kind of tracked text revision.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionKind {
    /// Text inserted while revision tracking was enabled.
    Insertion,
    /// Text deleted while revision tracking was enabled.
    Deletion,
    /// Character formatting changed while revision tracking was enabled.
    Formatting,
}

/// Tracked revision metadata attached to text or formatted document properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionMark {
    /// Revision kind.
    pub kind: RevisionKind,
    /// Index into the document revision-author table.
    pub author_index: u16,
    /// Resolved revision author name.
    pub author: String,
    /// Revision date and time, when the DTTM is not an ignored zero date.
    pub timestamp: Option<CommentDateTime>,
    /// Optional revision identifier.
    pub revision_id: Option<u16>,
}

/// Resolved numbering revision metadata for a paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingRevisionMark {
    /// Whether the paragraph was already numbered when tracking began.
    pub was_numbered: bool,
    /// Index into the document revision-author table.
    pub author_index: u16,
    /// Resolved revision author name.
    pub author: String,
    /// Revision date and time, when the packed DTTM is not an ignored zero date.
    pub timestamp: Option<CommentDateTime>,
    /// Placeholder positions for the nine numbering levels.
    pub placeholder_positions: [u8; 9],
    /// MSONFC values for the nine numbering levels.
    pub number_formats: [u8; 9],
    /// Numeric values for the nine numbering levels.
    pub numbers: [u32; 9],
    /// Numbering format string.
    pub format_string: String,
}

/// Resolved revision metadata for a LISTNUM display-field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DisplayFieldRevisionMark {
    /// Index into the document revision-author table.
    pub author_index: u16,
    /// Resolved revision author name.
    pub author: String,
    /// Revision date and time, when the packed DTTM is not an ignored zero date.
    pub timestamp: Option<CommentDateTime>,
    /// Previous LISTNUM field result.
    pub previous_result: String,
}

/// Resolved property revision metadata for a document section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionRevisionMark {
    /// First character position in the section.
    pub start: u32,
    /// Character position immediately after the section.
    pub end: u32,
    /// Index into the document revision-author table.
    pub author_index: u16,
    /// Resolved revision author name.
    pub author: String,
    /// Revision date and time, when the packed DTTM is not an ignored zero date.
    pub timestamp: Option<CommentDateTime>,
}

pub(crate) fn decode_dttm(value: u32) -> Result<Option<CommentDateTime>> {
    let minute = (value & 0x3F) as u8;
    let hour = ((value >> 6) & 0x1F) as u8;
    let day = ((value >> 11) & 0x1F) as u8;
    let month = ((value >> 16) & 0x0F) as u8;
    let year = ((value >> 20) & 0x01FF) as u16 + 1900;
    let weekday = ((value >> 29) & 0x07) as u8;
    if minute > 59 || hour > 23 || day > 31 || month > 12 || weekday > 6 {
        return Err(DocError::Corrupted(
            "revision contains an invalid DTTM".to_string(),
        ));
    }
    if day == 0 || month == 0 {
        return Ok(None);
    }
    Ok(Some(CommentDateTime {
        year,
        month,
        day,
        hour,
        minute,
        weekday,
    }))
}
