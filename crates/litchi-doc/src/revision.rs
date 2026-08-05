use super::comment::CommentDateTime;
use super::package::{Error as PackageError, Result};

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

/// Validated MS-DOC reason code for an inserted, modified, or deleted revision.
///
/// The binary format defines the contiguous values `0x0000..=0x002B`; use
/// [`Self::raw`] to distinguish the individual auto-formatting reasons.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionReason(u16);

impl RevisionReason {
    /// A normal user edit (`0x0000`).
    pub const NORMAL_EDIT: Self = Self(0);
    /// A style was applied (`0x0001`).
    pub const APPLIED_STYLE: Self = Self(1);
    /// The maximum reason value defined by MS-DOC.
    pub const MAX_VALUE: u16 = 0x002B;

    /// Construct a reason from its MS-DOC value.
    pub fn from_raw(value: u16) -> Option<Self> {
        (value <= Self::MAX_VALUE).then_some(Self(value))
    }

    /// Return the underlying MS-DOC reason value.
    pub fn raw(self) -> u16 {
        self.0
    }
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
    /// Structured edit reason, when an Idsl reason operand is present.
    pub reason: Option<RevisionReason>,
    /// Legacy raw alias for [`Self::reason`].
    ///
    /// This field is retained for source compatibility with the original API;
    /// it is a reason code, not a revision-save ID.
    pub revision_id: Option<u16>,
    /// ECMA-376 single-session revision-save ID.
    pub revision_save_id: Option<u32>,
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
        return Err(PackageError::Corrupted(
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_revision_reason_range() {
        assert_eq!(RevisionReason::NORMAL_EDIT.raw(), 0);
        assert_eq!(RevisionReason::APPLIED_STYLE.raw(), 1);
        assert_eq!(RevisionReason::from_raw(0x002B).unwrap().raw(), 0x002B);
        assert!(RevisionReason::from_raw(0x002C).is_none());
    }
}
