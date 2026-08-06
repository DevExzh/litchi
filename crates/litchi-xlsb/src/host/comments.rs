//! Compatibility adapter for the XLSB comments codec.
//!
//! The reusable comments stream and BIFF12 rich-string codec live in
//! [`crate::comments`]. This boundary retains the historical host
//! `Record` model, including its `SharedStringRun` type, while keeping
//! package and host error mapping in the owning `litchi_xlsb` layers.

use crate::package::error::{Error, Result};
use crate::package::shared_strings::SharedStringRun;
use crate::raw::Writer;
use std::io::Write;

/// Comment record information.
///
/// Represents a cell comment with author and text.
#[derive(Debug, Clone)]
pub struct Record {
    /// Row (0-based).
    pub row: u32,
    /// Column (0-based).
    pub col: u32,
    /// Author of the comment.
    pub author: String,
    /// Comment text.
    pub text: String,
    /// Font runs within the comment text.
    pub runs: Vec<SharedStringRun>,
    /// Comment identifier used by shared-workbook metadata.
    pub guid: [u8; 16],
    /// Optional identifier carried by an alternate-content `BrtUid` record.
    pub alternate_guid: Option<[u8; 16]>,
    /// Whether the comment is visible.
    pub visible: bool,
}

impl Record {
    /// Create a new comment.
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_xlsb::package::comments::Record;
    ///
    /// let comment = Record::new(0, 0, "John".to_string(), "This is a note".to_string());
    /// ```
    pub fn new(row: u32, col: u32, author: String, text: String) -> Self {
        Self {
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

    /// Set visibility.
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}

pub(crate) fn read_comments(bytes: &[u8]) -> Result<Vec<Record>> {
    let owner_comments = crate::comments::read(bytes).map_err(map_owner_error)?;
    let mut comments = Vec::new();
    comments
        .try_reserve(owner_comments.len())
        .map_err(|source| Error::Allocation {
            resource: "host comments",
            source,
        })?;
    for owner_comment in owner_comments {
        let crate::comments::Record {
            row,
            col,
            author,
            text,
            runs: owner_runs,
            guid,
            alternate_guid,
            visible,
        } = owner_comment;
        let mut runs = Vec::new();
        runs.try_reserve(owner_runs.len())
            .map_err(|source| Error::Allocation {
                resource: "host comment rich-string runs",
                source,
            })?;
        for run in owner_runs {
            runs.push(SharedStringRun {
                character_index: run.character_index,
                font_id: run.font_id,
            });
        }
        comments.push(Record {
            row,
            col,
            author,
            text,
            runs,
            guid,
            alternate_guid,
            visible,
        });
    }
    Ok(comments)
}

pub(crate) fn write_comments<W: Write>(writer: &mut Writer<W>, comments: &[Record]) -> Result<()> {
    let mut owner_comments = Vec::new();
    owner_comments
        .try_reserve(comments.len())
        .map_err(|source| Error::Allocation {
            resource: "owner comments",
            source,
        })?;
    for comment in comments {
        let mut runs = Vec::new();
        runs.try_reserve(comment.runs.len())
            .map_err(|source| Error::Allocation {
                resource: "owner comment rich-string runs",
                source,
            })?;
        for run in &comment.runs {
            runs.push(crate::comments::Run {
                character_index: run.character_index,
                font_id: run.font_id,
            });
        }
        owner_comments.push(crate::comments::Record {
            row: comment.row,
            col: comment.col,
            author: comment.author.clone(),
            text: comment.text.clone(),
            runs,
            guid: comment.guid,
            alternate_guid: comment.alternate_guid,
            visible: comment.visible,
        });
    }
    crate::comments::write(writer, &owner_comments).map_err(map_owner_error)
}

fn map_owner_error(error: crate::comments::Error) -> Error {
    match error {
        crate::comments::Error::Wire(error) => Error::Wire(error),
        crate::comments::Error::InvalidRecordType(record_type) => {
            Error::InvalidRecordType(record_type)
        },
        crate::comments::Error::InvalidLength { expected, found } => {
            Error::InvalidLength { expected, found }
        },
        crate::comments::Error::Unrecognized { typ, val } => Error::Unrecognized { typ, val },
        crate::comments::Error::Encoding(message) => Error::Encoding(message),
        crate::comments::Error::UnsupportedFeature(feature) => Error::UnsupportedFeature(feature),
        crate::comments::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::kind;

    #[test]
    fn test_comment_creation() {
        let comment = Record::new(0, 0, "John".to_string(), "Note".to_string());
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
