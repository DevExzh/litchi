//! Comment support for XLSB

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::decode_string;
use crate::xlsb::shared_strings::{SharedString, SharedStringRun};
use litchi_core::binary;
use litchi_xlsb::raw::{Record, Records, Writer, kind};
use std::collections::{HashMap, HashSet};
use std::io::Write;

/// Comment information
///
/// Represents a cell comment with author and text.
#[derive(Debug, Clone)]
pub struct Comment {
    /// Row (0-based)
    pub row: u32,
    /// Column (0-based)
    pub col: u32,
    /// Author of the comment
    pub author: String,
    /// Comment text
    pub text: String,
    /// Font runs within the comment text.
    pub runs: Vec<SharedStringRun>,
    /// Comment identifier used by shared-workbook metadata.
    pub guid: [u8; 16],
    /// Optional identifier carried by an alternate-content `BrtUid` record.
    pub alternate_guid: Option<[u8; 16]>,
    /// Whether comment is visible
    pub visible: bool,
}

impl Comment {
    /// Create a new comment
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_ooxml::xlsb::comments::Comment;
    ///
    /// let comment = Comment::new(0, 0, "John".to_string(), "This is a note".to_string());
    /// ```
    pub fn new(row: u32, col: u32, author: String, text: String) -> Self {
        Comment {
            row,
            col,
            author,
            text,
            runs: Vec::new(),
            guid: [0; 16],
            alternate_guid: None,
            visible: false,
        }
    }

    /// Set visibility
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}

pub(crate) fn read_comments(bytes: &[u8]) -> XlsbResult<Vec<Comment>> {
    let records = Records::new(bytes).collect::<Result<Vec<_>, _>>()?;
    let mut pos = 0;
    expect(&records, &mut pos, kind::BEGIN_COMMENTS)?;
    expect(&records, &mut pos, kind::BEGIN_COMMENT_AUTHORS)?;
    let mut authors = Vec::new();
    while records
        .get(pos)
        .is_some_and(|r| r.kind() == kind::COMMENT_AUTHOR)
    {
        let record = &records[pos];
        pos += 1;
        let (author, consumed) = decode_string(record.payload())?;
        let len = author.encode_utf16().count();
        if consumed != record.payload().len() || len > 54 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtCommentAuthor".to_string(),
                val: author,
            });
        }
        authors.push(author);
    }
    expect(&records, &mut pos, kind::END_COMMENT_AUTHORS)?;
    expect(&records, &mut pos, kind::BEGIN_COMMENT_LIST)?;
    let mut comments = Vec::new();
    let mut cells = HashSet::new();
    loop {
        let alternate_guid = if records.get(pos).is_some_and(|r| r.kind() == kind::AC_BEGIN) {
            pos += 1;
            let mut uid = None;
            while records
                .get(pos)
                .is_some_and(|record| record.kind() != kind::AC_END)
            {
                let record = &records[pos];
                if record.kind() == kind::UID {
                    if record.payload().len() != 16 || uid.is_some() {
                        return Err(XlsbError::Unrecognized {
                            typ: "ACUID BrtUid".to_string(),
                            val: "invalid or duplicate UID".to_string(),
                        });
                    }
                    let mut value = [0; 16];
                    value.copy_from_slice(record.payload());
                    uid = Some(value);
                }
                pos += 1;
            }
            expect(&records, &mut pos, kind::AC_END)?;
            uid
        } else {
            None
        };
        if records
            .get(pos)
            .is_none_or(|r| r.kind() != kind::BEGIN_COMMENT)
        {
            if alternate_guid.is_some() {
                return Err(XlsbError::Unrecognized {
                    typ: "ACUID".to_string(),
                    val: "not followed by BrtBeginComment".to_string(),
                });
            }
            break;
        }
        let begin = &records[pos];
        pos += 1;
        if begin.payload().len() != 36 {
            return Err(XlsbError::InvalidLength {
                expected: 36,
                found: begin.payload().len(),
            });
        }
        let author_raw = binary::read_u32_le_at(begin.payload(), 0)?;
        let row = binary::read_u32_le_at(begin.payload(), 4)?;
        let last_row = binary::read_u32_le_at(begin.payload(), 8)?;
        let col = binary::read_u32_le_at(begin.payload(), 12)?;
        let last_col = binary::read_u32_le_at(begin.payload(), 16)?;
        if author_raw > i32::MAX as u32
            || author_raw as usize >= authors.len()
            || row != last_row
            || col != last_col
            || row >= 1_048_576
            || col >= 16_384
            || !cells.insert((row, col))
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBeginComment".to_string(),
                val: format!("author={author_raw}, range={row}:{col}-{last_row}:{last_col}"),
            });
        }
        let rich = if records
            .get(pos)
            .is_some_and(|r| r.kind() == kind::COMMENT_TEXT)
        {
            let value = SharedString::parse(records[pos].payload())?;
            pos += 1;
            if records[pos - 1].payload()[0] & 1 == 0 || value.phonetic.is_some() {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtCommentText".to_string(),
                    val: "invalid RichStr flags".to_string(),
                });
            }
            value
        } else {
            SharedString {
                text: String::new(),
                runs: Vec::new(),
                phonetic: None,
            }
        };
        expect(&records, &mut pos, kind::END_COMMENT)?;
        let mut guid = [0; 16];
        guid.copy_from_slice(&begin.payload()[20..36]);
        comments.push(Comment {
            row,
            col,
            author: authors[author_raw as usize].clone(),
            text: rich.text,
            runs: rich.runs,
            guid,
            alternate_guid,
            visible: false,
        });
    }
    expect(&records, &mut pos, kind::END_COMMENT_LIST)?;
    while records
        .get(pos)
        .is_some_and(|record| record.kind() == kind::FRT_BEGIN)
    {
        if records[pos].payload().len() != 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
                found: records[pos].payload().len(),
            });
        }
        pos += 1;
        while records
            .get(pos)
            .is_some_and(|record| record.kind() != kind::FRT_END)
        {
            pos += 1;
        }
        expect(&records, &mut pos, kind::FRT_END)?;
    }
    expect(&records, &mut pos, kind::END_COMMENTS)?;
    if pos != records.len() {
        return Err(XlsbError::Unrecognized {
            typ: "Comments stream".to_string(),
            val: "trailing records".to_string(),
        });
    }
    Ok(comments)
}

fn expect(records: &[Record<'_>], pos: &mut usize, typ: litchi_xlsb::raw::Kind) -> XlsbResult<()> {
    let record = records
        .get(*pos)
        .ok_or(XlsbError::InvalidRecordType(typ.get()))?;
    if record.kind() != typ {
        return Err(XlsbError::InvalidRecordType(record.kind().get()));
    }
    if !record.payload().is_empty() {
        return Err(XlsbError::InvalidLength {
            expected: 0,
            found: record.payload().len(),
        });
    }
    *pos += 1;
    Ok(())
}

pub(crate) fn write_comments<W: Write>(
    writer: &mut Writer<W>,
    comments: &[Comment],
) -> XlsbResult<()> {
    writer.write_record(kind::BEGIN_COMMENTS, &[])?;
    writer.write_record(kind::BEGIN_COMMENT_AUTHORS, &[])?;
    let mut authors = Vec::<&str>::new();
    let mut author_ids = HashMap::<&str, u32>::new();
    for comment in comments {
        if !author_ids.contains_key(comment.author.as_str()) {
            author_ids.insert(comment.author.as_str(), authors.len() as u32);
            authors.push(&comment.author);
        }
    }
    for author in &authors {
        let len = author.encode_utf16().count();
        if len > 54 {
            return Err(XlsbError::Encoding(
                "comment author length must not exceed 54 characters".to_string(),
            ));
        }
        let mut data = Vec::new();
        data.extend_from_slice(&(len as u32).to_le_bytes());
        for unit in author.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        writer.write_record(kind::COMMENT_AUTHOR, &data)?;
    }
    writer.write_record(kind::END_COMMENT_AUTHORS, &[])?;
    writer.write_record(kind::BEGIN_COMMENT_LIST, &[])?;
    let mut cells = HashSet::new();
    for comment in comments {
        if comment.row >= 1_048_576
            || comment.col >= 16_384
            || !cells.insert((comment.row, comment.col))
        {
            return Err(XlsbError::Encoding(
                "invalid or duplicate comment cell".to_string(),
            ));
        }
        let author = author_ids[comment.author.as_str()];
        let mut begin = Vec::with_capacity(36);
        begin.extend_from_slice(&author.to_le_bytes());
        begin.extend_from_slice(&comment.row.to_le_bytes());
        begin.extend_from_slice(&comment.row.to_le_bytes());
        begin.extend_from_slice(&comment.col.to_le_bytes());
        begin.extend_from_slice(&comment.col.to_le_bytes());
        begin.extend_from_slice(&comment.guid);
        writer.write_record(kind::BEGIN_COMMENT, &begin)?;
        if !comment.text.is_empty() {
            let rich = SharedString {
                text: comment.text.clone(),
                runs: comment.runs.clone(),
                phonetic: None,
            };
            writer.write_record(kind::COMMENT_TEXT, &rich.to_comment_bytes()?)?;
        }
        writer.write_record(kind::END_COMMENT, &[])?;
    }
    writer.write_record(kind::END_COMMENT_LIST, &[])?;
    writer.write_record(kind::END_COMMENTS, &[])?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_comment_creation() {
        let comment = Comment::new(0, 0, "John".to_string(), "Note".to_string());
        assert_eq!(comment.row, 0);
        assert_eq!(comment.col, 0);
        assert_eq!(comment.author, "John");
        assert!(!comment.visible);
    }

    #[test]
    fn reads_alternate_uid_and_future_record_wrappers() {
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer.write_record(kind::BEGIN_COMMENTS, &[]).unwrap();
        writer
            .write_record(kind::BEGIN_COMMENT_AUTHORS, &[])
            .unwrap();
        writer
            .write_record(kind::COMMENT_AUTHOR, &0u32.to_le_bytes())
            .unwrap();
        writer.write_record(kind::END_COMMENT_AUTHORS, &[]).unwrap();
        writer.write_record(kind::BEGIN_COMMENT_LIST, &[]).unwrap();
        writer.write_record(kind::AC_BEGIN, &[0; 4]).unwrap();
        writer.write_record(kind::UID, &[9; 16]).unwrap();
        writer.write_record(kind::AC_END, &[]).unwrap();
        let mut begin = Vec::new();
        begin.extend_from_slice(&0u32.to_le_bytes());
        begin.extend_from_slice(&2u32.to_le_bytes());
        begin.extend_from_slice(&2u32.to_le_bytes());
        begin.extend_from_slice(&3u32.to_le_bytes());
        begin.extend_from_slice(&3u32.to_le_bytes());
        begin.extend_from_slice(&[7; 16]);
        writer.write_record(kind::BEGIN_COMMENT, &begin).unwrap();
        writer.write_record(kind::END_COMMENT, &[]).unwrap();
        writer.write_record(kind::END_COMMENT_LIST, &[]).unwrap();
        writer.write_record(kind::FRT_BEGIN, &[0; 4]).unwrap();
        writer.write_record(kind::UID, &[1; 16]).unwrap();
        writer.write_record(kind::FRT_END, &[]).unwrap();
        writer.write_record(kind::END_COMMENTS, &[]).unwrap();

        let comments = read_comments(bytes.as_slice()).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].alternate_guid, Some([9; 16]));
        assert_eq!(comments[0].guid, [7; 16]);
    }
}
