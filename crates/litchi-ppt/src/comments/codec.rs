//! Strict codecs for PowerPoint 10 presentation-comment records.

use super::model::{Author, Authors};
use crate::consts::PptRecordType;
use crate::package::{PptError, Result};
use crate::records::PptRecord;
use crate::slide::ParsedComment;

const AUTHOR_NAME_MAX_BYTES: usize = 104;
const COMMENT_TEXT_MAX_BYTES: usize = 64_000;

impl Authors {
    /// Parse comment-author records from `___PPT10` document extensions below
    /// `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut authors = Vec::new();
        for record in root.versioned_binary_tag_records(10)? {
            if record.record_type == PptRecordType::CommentIndex10 {
                authors.push(parse_author(&record)?);
            }
        }
        Ok(Self { authors })
    }
}

/// Parse the ordered `Comment2000` records from one slide's `___PPT10`
/// extension.
pub(crate) fn parse_slide_comments(root: &PptRecord) -> Result<Vec<ParsedComment>> {
    let mut comments = Vec::new();
    for record in root.versioned_binary_tag_records(10)? {
        if record.record_type == PptRecordType::Comment2000 {
            comments.push(parse_comment(&record)?);
        }
    }
    Ok(comments)
}

fn parse_author(record: &PptRecord) -> Result<Author> {
    if record.record_type != PptRecordType::CommentIndex10
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "CommentIndex10Container has an invalid record header".to_string(),
        ));
    }
    let children = PptRecord::parse_sequence_strict(&record.data, "comment author")?;
    let mut name = None;
    let mut color_index = None;
    let mut comment_index_seed = None;
    for child in children {
        match child.record_type {
            PptRecordType::CString if color_index.is_none() && name.is_none() => {
                name = Some(parse_string(
                    &child,
                    0,
                    AUTHOR_NAME_MAX_BYTES,
                    false,
                    "AuthorNameAtom",
                )?);
            },
            PptRecordType::CommentIndex10Atom
                if color_index.is_none() && comment_index_seed.is_none() =>
            {
                if child.version != 0 || child.instance != 0 || child.data.len() != 8 {
                    return Err(PptError::Corrupted(
                        "CommentIndex10Atom has an invalid record header or size".to_string(),
                    ));
                }
                let color = i32::from_le_bytes(child.data[0..4].try_into().map_err(|_| {
                    PptError::Corrupted("Comment color index is truncated".to_string())
                })?);
                let seed = i32::from_le_bytes(child.data[4..8].try_into().map_err(|_| {
                    PptError::Corrupted("Comment index seed is truncated".to_string())
                })?);
                if color < 0 || seed < 0 {
                    return Err(PptError::Corrupted(
                        "Comment author color index or seed is negative".to_string(),
                    ));
                }
                color_index = Some(color);
                comment_index_seed = Some(seed);
            },
            _ => {
                return Err(PptError::Corrupted(
                    "CommentIndex10Container has duplicate, out-of-order, or unexpected children"
                        .to_string(),
                ));
            },
        }
    }
    Ok(Author {
        name,
        color_index,
        comment_index_seed,
    })
}

fn parse_comment(record: &PptRecord) -> Result<ParsedComment> {
    if record.record_type != PptRecordType::Comment2000
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "Comment10Container has an invalid record header".to_string(),
        ));
    }
    let children = PptRecord::parse_sequence_strict(&record.data, "presentation comment")?;
    let mut comment = ParsedComment::default();
    let mut stage = 0u8;
    let mut has_atom = false;
    for child in children {
        match (child.record_type, child.instance) {
            (PptRecordType::CString, 0) if stage == 0 => {
                comment.author = parse_string(
                    &child,
                    0,
                    AUTHOR_NAME_MAX_BYTES,
                    false,
                    "Comment10AuthorAtom",
                )?;
                stage = 1;
            },
            (PptRecordType::CString, 1) if stage <= 1 => {
                comment.text =
                    parse_string(&child, 1, COMMENT_TEXT_MAX_BYTES, true, "Comment10TextAtom")?;
                stage = 2;
            },
            (PptRecordType::CString, 2) if stage <= 2 => {
                comment.initials = parse_string(
                    &child,
                    2,
                    AUTHOR_NAME_MAX_BYTES,
                    false,
                    "Comment10AuthorInitialAtom",
                )?;
                stage = 3;
            },
            (PptRecordType::Comment2000Atom, _) if !has_atom => {
                parse_comment_atom(&child, &mut comment)?;
                has_atom = true;
                stage = 4;
            },
            _ => {
                return Err(PptError::Corrupted(
                    "Comment10Container has duplicate, out-of-order, or unexpected children"
                        .to_string(),
                ));
            },
        }
    }
    if !has_atom || stage != 4 {
        return Err(PptError::Corrupted(
            "Comment10Container is missing Comment10Atom".to_string(),
        ));
    }
    Ok(comment)
}

fn parse_comment_atom(record: &PptRecord, comment: &mut ParsedComment) -> Result<()> {
    if record.record_type != PptRecordType::Comment2000Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 28
    {
        return Err(PptError::Corrupted(
            "Comment10Atom has an invalid record header or size".to_string(),
        ));
    }
    let data = &record.data;
    comment.index = i32::from_le_bytes(
        data[0..4]
            .try_into()
            .map_err(|_| PptError::Corrupted("Comment10Atom index is truncated".to_string()))?,
    );
    if comment.index < 0 {
        return Err(PptError::Corrupted(
            "Comment10Atom index is negative".to_string(),
        ));
    }
    comment.year = u16::from_le_bytes([data[4], data[5]]);
    comment.month = u16::from_le_bytes([data[6], data[7]]);
    comment.day_of_week = u16::from_le_bytes([data[8], data[9]]);
    comment.day = u16::from_le_bytes([data[10], data[11]]);
    comment.hour = u16::from_le_bytes([data[12], data[13]]);
    comment.minute = u16::from_le_bytes([data[14], data[15]]);
    comment.second = u16::from_le_bytes([data[16], data[17]]);
    comment.millisecond = u16::from_le_bytes([data[18], data[19]]);
    comment.x = i32::from_le_bytes(data[20..24].try_into().map_err(|_| {
        PptError::Corrupted("Comment10Atom horizontal anchor is truncated".to_string())
    })?);
    comment.y = i32::from_le_bytes(data[24..28].try_into().map_err(|_| {
        PptError::Corrupted("Comment10Atom vertical anchor is truncated".to_string())
    })?);
    Ok(())
}

fn parse_string(
    record: &PptRecord,
    instance: u16,
    max_len: usize,
    allow_tab_cr_lf: bool,
    name: &str,
) -> Result<String> {
    if record.record_type != PptRecordType::CString
        || record.version != 0
        || record.instance != instance
        || record.data.len() > max_len
        || record.data.len() & 1 != 0
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        let forbidden = if allow_tab_cr_lf {
            matches!(unit, 0x0001..=0x0008 | 0x000b..=0x000c | 0x000e..=0x001f | 0x007f..=0x009f)
        } else {
            matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f)
        };
        if forbidden {
            return Err(PptError::Corrupted(format!(
                "{name} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted(format!("{name} contains invalid UTF-16")))
}
