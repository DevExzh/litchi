//! Comment/Note parsing for XLS BIFF8 files.
//!
//! Parses NOTE records (0x001C) which associate a cell with a comment.
//! The actual comment text is stored in the TXO (Text Object) record
//! within the drawing layer (OBJ/MsoDrawing), referenced by `shape_id`.
//!
//! # Record Format (MS-XLS 2.4.179, based on Apache POI `NoteRecord`)
//!
//! ```text
//! Offset  Size  Field
//! 0       2     row - Row containing the comment (0-based)
//! 2       2     col - Column containing the comment (0-based)
//! 4       2     flags - 0x0000 = hidden, 0x0002 = visible
//! 6       2     shapeId - Object ID of the associated OBJ record
//! 8       2     cch - Length of author string in characters
//! 10      1     fHighByte - 0x00 = compressed, 0x01 = UTF-16LE
//! 11      var   author - Comment author name
//! ```

use crate::common::binary;
use crate::ole::xls::error::{XlsError, XlsResult};
use std::collections::HashMap;

/// NOTE record type identifier.
pub const RECORD_TYPE: u16 = 0x001C;

/// OBJ record type identifier.
pub const OBJ_TYPE: u16 = 0x005D;

/// TXO (Text Object) record type identifier.
pub const TXO_TYPE: u16 = 0x01B6;

/// CONTINUE record type identifier (carries TXO text/formatting data).
pub const CONTINUE_TYPE: u16 = 0x003C;

/// Comment visibility flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CommentVisibility {
    /// Comment is hidden (default).
    Hidden,
    /// Comment is always visible.
    Visible,
}

/// A parsed cell comment/note from a NOTE record.
#[derive(Debug, Clone)]
pub struct XlsComment {
    /// Row of the commented cell (0-based).
    pub row: u16,
    /// Column of the commented cell (0-based).
    pub col: u16,
    /// Whether the comment is visible or hidden.
    pub visibility: CommentVisibility,
    /// Object ID referencing the OBJ record that contains the comment text.
    pub shape_id: u16,
    /// Author of the comment.
    pub author: String,
    /// Comment text body (extracted from TXO + CONTINUE records).
    /// Empty until resolved via [`resolve_comment_texts`].
    pub text: String,
}

/// Parse a single NOTE record.
pub fn parse_note_record(data: &[u8]) -> XlsResult<XlsComment> {
    // Minimum size: 2+2+2+2+2+1 = 11 bytes (with empty author)
    if data.len() < 11 {
        return Err(XlsError::InvalidLength {
            expected: 11,
            found: data.len(),
        });
    }

    let row = binary::read_u16_le_at(data, 0)?;
    let col = binary::read_u16_le_at(data, 2)?;
    let flags = binary::read_u16_le_at(data, 4)?;
    let shape_id = binary::read_u16_le_at(data, 6)?;
    let cch = binary::read_u16_le_at(data, 8)? as usize;
    let f_high_byte = data[10];

    let visibility = if flags & 0x0002 != 0 {
        CommentVisibility::Visible
    } else {
        CommentVisibility::Hidden
    };

    let author = if cch == 0 {
        String::new()
    } else if f_high_byte != 0 {
        // UTF-16LE
        let byte_len = cch * 2;
        if 11 + byte_len > data.len() {
            return Err(XlsError::InvalidLength {
                expected: 11 + byte_len,
                found: data.len(),
            });
        }
        let slice = &data[11..11 + byte_len];
        let words: Vec<u16> = slice
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16(&words)
            .map_err(|e| XlsError::InvalidData(format!("Invalid UTF-16 in NOTE author: {}", e)))?
    } else {
        // Compressed (single-byte)
        if 11 + cch > data.len() {
            return Err(XlsError::InvalidLength {
                expected: 11 + cch,
                found: data.len(),
            });
        }
        data[11..11 + cch].iter().map(|&b| b as char).collect()
    };

    Ok(XlsComment {
        row,
        col,
        visibility,
        shape_id,
        author,
        text: String::new(),
    })
}

/// Extract the object ID from an OBJ record.
///
/// The first sub-record inside OBJ is `CommonObjectDataSubRecord` (ftCmo,
/// sub-record type 0x0015). Its layout is:
///
/// ```text
/// Offset  Size  Field
/// 0       2     sub-record type (0x0015)
/// 2       2     sub-record data size
/// 4       2     object type
/// 6       2     object ID  <-- this is what we want
/// 8       2     option flags
/// ...
/// ```
pub fn parse_obj_id(data: &[u8]) -> XlsResult<u16> {
    // Need at least 8 bytes for ftCmo header + object type + object ID
    if data.len() < 8 {
        return Err(XlsError::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }
    let sub_type = binary::read_u16_le_at(data, 0)?;
    if sub_type != 0x0015 {
        return Err(XlsError::InvalidData(format!(
            "Expected ftCmo sub-record (0x0015), got 0x{:04X}",
            sub_type
        )));
    }
    // object ID at offset 6
    Ok(binary::read_u16_le_at(data, 6)?)
}

/// State machine for collecting TXO text during worksheet record iteration.
///
/// Usage:
/// 1. Call [`TxoCollector::new`] before the record loop.
/// 2. For each OBJ record, call [`feed_obj`].
/// 3. For each TXO record, call [`feed_txo`].
/// 4. For each CONTINUE record, call [`feed_continue`].
/// 5. After the loop, call [`resolve_comment_texts`] to fill `XlsComment.text`.
#[derive(Debug)]
pub struct TxoCollector {
    /// Maps object_id → comment text.
    texts: HashMap<u16, String>,
    /// The object ID of the most recently seen OBJ record.
    last_obj_id: Option<u16>,
    /// Text length from the most recently seen TXO record.
    pending_text_len: Option<u16>,
    /// Whether we are waiting for the first CONTINUE (text data) after a TXO.
    awaiting_text_continue: bool,
}

impl Default for TxoCollector {
    fn default() -> Self {
        Self::new()
    }
}

impl TxoCollector {
    pub fn new() -> Self {
        Self {
            texts: HashMap::new(),
            last_obj_id: None,
            pending_text_len: None,
            awaiting_text_continue: false,
        }
    }

    /// Feed an OBJ record to extract the object ID.
    pub fn feed_obj(&mut self, data: &[u8]) {
        self.last_obj_id = parse_obj_id(data).ok();
        // Reset TXO state when a new OBJ appears
        self.pending_text_len = None;
        self.awaiting_text_continue = false;
    }

    /// Feed a TXO record. Extracts the text length from the header.
    ///
    /// TXO header layout (18 bytes):
    /// ```text
    ///  0  u16  options
    ///  2  u16  textOrientation
    ///  4  u16  reserved4
    ///  6  u16  reserved5
    ///  8  u16  reserved6
    /// 10  u16  cchText (text length in characters)
    /// 12  u16  cbFmtRuns (formatting run data length in bytes)
    /// 14  u32  reserved7
    /// ```
    pub fn feed_txo(&mut self, data: &[u8]) {
        if data.len() < 18 {
            return;
        }
        let cch_text = binary::read_u16_le_at(data, 10).unwrap_or(0);
        self.pending_text_len = Some(cch_text);
        self.awaiting_text_continue = cch_text > 0;
    }

    /// Feed a CONTINUE record that may carry TXO text data.
    ///
    /// The first CONTINUE after a TXO with non-zero text length contains:
    /// ```text
    ///  0   u8   compressByte (0 = compressed Latin-1, 1 = UTF-16LE)
    ///  1   var  string data
    /// ```
    pub fn feed_continue(&mut self, data: &[u8]) {
        if !self.awaiting_text_continue {
            return;
        }
        self.awaiting_text_continue = false;

        let cch = match self.pending_text_len.take() {
            Some(n) if n > 0 => n as usize,
            _ => return,
        };

        if data.is_empty() {
            return;
        }

        let compress_byte = data[0];
        let is_compressed = (compress_byte & 0x01) == 0;
        let str_data = &data[1..];

        let text = if is_compressed {
            // Latin-1: 1 byte per char
            if str_data.len() < cch {
                return;
            }
            str_data[..cch]
                .iter()
                .map(|&b| b as char)
                .collect::<String>()
        } else {
            // UTF-16LE: 2 bytes per char
            let byte_len = cch * 2;
            if str_data.len() < byte_len {
                return;
            }
            let words: Vec<u16> = str_data[..byte_len]
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            match String::from_utf16(&words) {
                Ok(s) => s,
                Err(_) => return,
            }
        };

        if let Some(obj_id) = self.last_obj_id {
            self.texts.insert(obj_id, text);
        }
    }

    /// Resolve comment texts: for each `XlsComment`, look up its `shape_id`
    /// in the collected TXO texts and populate the `text` field.
    pub fn resolve_comment_texts(&self, comments: &mut [XlsComment]) {
        for comment in comments.iter_mut() {
            if let Some(text) = self.texts.get(&comment.shape_id) {
                comment.text.clone_from(text);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_note_hidden() {
        let mut data = Vec::new();
        data.extend_from_slice(&5u16.to_le_bytes()); // row
        data.extend_from_slice(&3u16.to_le_bytes()); // col
        data.extend_from_slice(&0u16.to_le_bytes()); // flags = hidden
        data.extend_from_slice(&42u16.to_le_bytes()); // shapeId
        data.extend_from_slice(&4u16.to_le_bytes()); // cch
        data.push(0x00); // fHighByte = compressed
        data.extend_from_slice(b"User"); // author

        let note = parse_note_record(&data).unwrap();
        assert_eq!(note.row, 5);
        assert_eq!(note.col, 3);
        assert_eq!(note.visibility, CommentVisibility::Hidden);
        assert_eq!(note.shape_id, 42);
        assert_eq!(note.author, "User");
        assert!(note.text.is_empty());
    }

    #[test]
    fn test_parse_note_visible_unicode() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&0u16.to_le_bytes()); // col
        data.extend_from_slice(&2u16.to_le_bytes()); // flags = visible
        data.extend_from_slice(&1u16.to_le_bytes()); // shapeId
        // Author "AB" in UTF-16LE
        data.extend_from_slice(&2u16.to_le_bytes()); // cch = 2
        data.push(0x01); // fHighByte = UTF-16
        data.extend_from_slice(&[0x41, 0x00, 0x42, 0x00]); // "AB"

        let note = parse_note_record(&data).unwrap();
        assert_eq!(note.visibility, CommentVisibility::Visible);
        assert_eq!(note.author, "AB");
    }

    #[test]
    fn test_parse_obj_id() {
        let mut data = vec![0u8; 12];
        // ftCmo sub-record type
        data[0] = 0x15;
        data[1] = 0x00;
        // sub-record data size
        data[2] = 0x08;
        data[3] = 0x00;
        // object type = 0x19 (comment)
        data[4] = 0x19;
        data[5] = 0x00;
        // object ID = 42
        data[6] = 42;
        data[7] = 0x00;
        assert_eq!(parse_obj_id(&data).unwrap(), 42);
    }

    #[test]
    fn test_txo_collector_compressed() {
        let mut collector = TxoCollector::new();

        // Simulate OBJ record with object ID = 7
        let mut obj_data = vec![0u8; 12];
        obj_data[0..2].copy_from_slice(&0x0015u16.to_le_bytes());
        obj_data[2..4].copy_from_slice(&0x0008u16.to_le_bytes());
        obj_data[4..6].copy_from_slice(&0x0019u16.to_le_bytes()); // comment type
        obj_data[6..8].copy_from_slice(&7u16.to_le_bytes()); // object ID = 7
        collector.feed_obj(&obj_data);

        // Simulate TXO record with text length = 5
        let mut txo_data = vec![0u8; 18];
        txo_data[10..12].copy_from_slice(&5u16.to_le_bytes()); // cchText = 5
        collector.feed_txo(&txo_data);

        // Simulate CONTINUE with compressed text "Hello"
        let mut cont_data = vec![0u8]; // compress byte = 0 (compressed)
        cont_data.extend_from_slice(b"Hello");
        collector.feed_continue(&cont_data);

        // Resolve
        let mut comments = vec![XlsComment {
            row: 0,
            col: 0,
            visibility: CommentVisibility::Hidden,
            shape_id: 7,
            author: String::new(),
            text: String::new(),
        }];
        collector.resolve_comment_texts(&mut comments);
        assert_eq!(comments[0].text, "Hello");
    }

    #[test]
    fn test_txo_collector_utf16() {
        let mut collector = TxoCollector::new();

        // OBJ with ID=3
        let mut obj_data = vec![0u8; 12];
        obj_data[0..2].copy_from_slice(&0x0015u16.to_le_bytes());
        obj_data[2..4].copy_from_slice(&0x0008u16.to_le_bytes());
        obj_data[6..8].copy_from_slice(&3u16.to_le_bytes());
        collector.feed_obj(&obj_data);

        // TXO with text length = 2
        let mut txo_data = vec![0u8; 18];
        txo_data[10..12].copy_from_slice(&2u16.to_le_bytes());
        collector.feed_txo(&txo_data);

        // CONTINUE with UTF-16LE text "Hi"
        let mut cont_data = vec![0x01u8]; // compress byte = 1 (UTF-16)
        cont_data.extend_from_slice(&[0x48, 0x00, 0x69, 0x00]); // "Hi"
        collector.feed_continue(&cont_data);

        let mut comments = vec![XlsComment {
            row: 0,
            col: 0,
            visibility: CommentVisibility::Hidden,
            shape_id: 3,
            author: String::new(),
            text: String::new(),
        }];
        collector.resolve_comment_texts(&mut comments);
        assert_eq!(comments[0].text, "Hi");
    }
}
