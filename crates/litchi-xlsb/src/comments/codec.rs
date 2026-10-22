#![allow(
    clippy::cast_possible_truncation,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation to this codec boundary"
)]

//! BIFF12 framing and rich-string codec for XLSB comments.
//!
//! The record payloads are the `BrtBeginComment`/`BrtCommentText` records
//! from [MS-XLSB] sections 2.4.30 and 2.4.341, with collection boundaries
//! from sections 2.4.31, 2.4.32, 2.4.33, and 2.4.340/2.4.387--2.4.390.

use super::model::{Record, Run};
use crate::leniency::{Tolerance, ToleranceReport};
use crate::raw::{Cursor, Encoder, Record as RawRecord, Records, kind};
use std::collections::{HashMap, HashSet, TryReserveError};
use std::io::Write;
use thiserror::Error;

/// Maximum row index represented by an XLSB worksheet, plus one.
pub const MAX_ROWS: u32 = 1_048_576;
/// Maximum column index represented by an XLSB worksheet, plus one.
pub const MAX_COLUMNS: u32 = 16_384;
/// Maximum UTF-16 units in a `BrtCommentAuthor` value.
pub const MAX_AUTHOR_UNITS: usize = 54;
/// Maximum UTF-16 units in the rich string used by `BrtCommentText`.
pub const MAX_TEXT_UNITS: usize = 0x7FFF;
/// Maximum number of rich-string runs accepted by the existing host codec.
pub const MAX_TEXT_RUNS: usize = 0x7FFF;

/// Result type for the standalone XLSB comments codec.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the bounded comments codec.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A record header, payload, scalar, or string failed raw validation.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    /// A required record was absent or a different record was encountered.
    #[error("invalid record type: 0x{0:04X}")]
    InvalidRecordType(u16),
    /// A fixed-width payload or enclosing record boundary is malformed.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength {
        /// Required byte count.
        expected: usize,
        /// Observed byte count.
        found: usize,
    },
    /// A well-framed record contains an invalid invariant.
    #[error("unrecognized {typ}: {val}")]
    Unrecognized {
        /// Name of the record or field being rejected.
        typ: &'static str,
        /// Diagnostic value.
        val: &'static str,
    },
    /// A writer-side value cannot be represented by the codec.
    #[error("encoding error: {0}")]
    Encoding(&'static str),
    /// A valid value uses a feature that `BrtCommentText` does not permit.
    #[error("unsupported feature: {0}")]
    UnsupportedFeature(&'static str),
    /// A bounded collection could not reserve its required memory.
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        /// Collection being grown.
        resource: &'static str,
        /// Original allocator failure.
        source: TryReserveError,
    },
}

/// Read one complete XLSB comments stream.
///
/// This standalone function cannot report repairs, so it refuses every
/// defect an opened worksheet repairs and records (see [`crate::leniency`]).
pub fn read(bytes: &[u8], limits: crate::raw::Limits) -> Result<Vec<Record>> {
    read_with(bytes, limits, &mut Tolerance::refusing())
}

/// Read one comments stream, recording Excel-compatible repairs into `report`.
pub(crate) fn read_recording(
    bytes: &[u8],
    limits: crate::raw::Limits,
    report: &mut ToleranceReport,
) -> Result<Vec<Record>> {
    read_with(bytes, limits, &mut Tolerance::recording(report))
}

fn read_with(
    bytes: &[u8],
    limits: crate::raw::Limits,
    tolerance: &mut Tolerance<'_>,
) -> Result<Vec<Record>> {
    let records = collect_records(bytes, limits)?;
    let mut pos = 0;
    expect(&records, &mut pos, kind::BEGIN_COMMENTS)?;
    expect(&records, &mut pos, kind::BEGIN_COMMENT_AUTHORS)?;

    let mut authors = Vec::new();
    while records
        .get(pos)
        .is_some_and(|record| record.kind() == kind::COMMENT_AUTHOR)
    {
        let record = &records[pos];
        pos += 1;
        let (author, consumed) = decode_wide_string(
            record.payload(),
            limits,
            tolerance,
            kind::COMMENT_AUTHOR,
            record_location(record.offset()),
        )?;
        let length = author.encode_utf16().count();
        if length > MAX_AUTHOR_UNITS {
            return Err(Error::Unrecognized {
                typ: "BrtCommentAuthor",
                val: "unexpected value",
            });
        }
        let trailing = record.payload().len().saturating_sub(consumed);
        tolerance.tolerate_trailing(
            kind::COMMENT_AUTHOR,
            record_location(record.offset()),
            trailing,
            || Error::Unrecognized {
                typ: "BrtCommentAuthor",
                val: "unexpected value",
            },
        )?;
        reserve(&mut authors, 1, "comment authors")?;
        authors.push(author);
    }
    expect(&records, &mut pos, kind::END_COMMENT_AUTHORS)?;
    expect(&records, &mut pos, kind::BEGIN_COMMENT_LIST)?;

    let mut comments = Vec::new();
    let mut cells = HashSet::new();
    loop {
        let alternate_guid = if records
            .get(pos)
            .is_some_and(|record| record.kind() == kind::AC_BEGIN)
        {
            pos += 1;
            let mut uid = None;
            while records
                .get(pos)
                .is_some_and(|record| record.kind() != kind::AC_END)
            {
                let record = &records[pos];
                if record.kind() == kind::UID {
                    if record.payload().len() != 16 || uid.is_some() {
                        return Err(Error::Unrecognized {
                            typ: "ACUID BrtUid",
                            val: "invalid or duplicate UID",
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
            .is_none_or(|record| record.kind() != kind::BEGIN_COMMENT)
        {
            if alternate_guid.is_some() {
                return Err(Error::Unrecognized {
                    typ: "ACUID",
                    val: "not followed by BrtBeginComment",
                });
            }
            break;
        }

        let begin = &records[pos];
        pos += 1;
        if begin.payload().len() != 36 {
            return Err(Error::InvalidLength {
                expected: 36,
                found: begin.payload().len(),
            });
        }
        let author_raw = read_u32(begin.payload(), 0);
        let row = read_u32(begin.payload(), 4);
        let last_row = read_u32(begin.payload(), 8);
        let col = read_u32(begin.payload(), 12);
        let last_col = read_u32(begin.payload(), 16);
        if author_raw > i32::MAX as u32
            || author_raw as usize >= authors.len()
            || row != last_row
            || col != last_col
            || row >= MAX_ROWS
            || col >= MAX_COLUMNS
            || !cells.insert((row, col))
        {
            return Err(Error::Unrecognized {
                typ: "BrtBeginComment",
                val: ", range",
            });
        }

        let (text, runs) = if records
            .get(pos)
            .is_some_and(|record| record.kind() == kind::COMMENT_TEXT)
        {
            let record = &records[pos];
            pos += 1;
            let rich = RichString::parse_with(
                record.payload(),
                limits,
                record_location(record.offset()),
                tolerance,
            )?;
            if !rich.is_rich || rich.is_extended {
                return Err(Error::Unrecognized {
                    typ: "BrtCommentText",
                    val: "invalid RichStr flags",
                });
            }
            (rich.text, rich.runs)
        } else {
            (String::new(), Vec::new())
        };

        expect(&records, &mut pos, kind::END_COMMENT)?;
        let mut guid = [0; 16];
        guid.copy_from_slice(&begin.payload()[20..36]);
        reserve(&mut comments, 1, "comments")?;
        comments.push(Record {
            row,
            col,
            author: authors[author_raw as usize].clone(),
            text,
            runs,
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
            return Err(Error::InvalidLength {
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
        return Err(Error::Unrecognized {
            typ: "Comments stream",
            val: "trailing records",
        });
    }
    Ok(comments)
}

/// Write one complete XLSB comments stream.
pub fn write<W: Write>(writer: &mut Encoder<W>, comments: &[Record]) -> Result<()> {
    writer.write_record(kind::BEGIN_COMMENTS, &[])?;
    writer.write_record(kind::BEGIN_COMMENT_AUTHORS, &[])?;

    let mut authors = Vec::<&str>::new();
    let mut author_ids = HashMap::<&str, u32>::new();
    for comment in comments {
        if !author_ids.contains_key(comment.author.as_str()) {
            reserve(&mut authors, 1, "comment authors")?;
            reserve(&mut author_ids, 1, "comment author indexes")?;
            author_ids.insert(comment.author.as_str(), authors.len() as u32);
            authors.push(&comment.author);
        }
    }
    for author in &authors {
        writer.write_record(kind::COMMENT_AUTHOR, &encode_author(author)?)?;
    }
    writer.write_record(kind::END_COMMENT_AUTHORS, &[])?;
    writer.write_record(kind::BEGIN_COMMENT_LIST, &[])?;

    let mut cells = HashSet::new();
    for comment in comments {
        if comment.row >= MAX_ROWS
            || comment.col >= MAX_COLUMNS
            || !cells.insert((comment.row, comment.col))
        {
            return Err(Error::Encoding("invalid or duplicate comment cell"));
        }
        let author = author_ids[comment.author.as_str()];
        write_comment(writer, comment, author)?;
    }
    writer.write_record(kind::END_COMMENT_LIST, &[])?;
    writer.write_record(kind::END_COMMENTS, &[])?;
    Ok(())
}

/// Encode one `BrtCommentAuthor` payload.
pub(crate) fn encode_author(author: &str) -> Result<Vec<u8>> {
    let length = author.encode_utf16().count();
    if length > MAX_AUTHOR_UNITS {
        return Err(Error::Encoding(
            "comment author length must not exceed 54 characters",
        ));
    }
    let byte_length = length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4))
        .ok_or(Error::Encoding("comment author length overflow"))?;
    let mut data = Vec::new();
    data.try_reserve(byte_length)
        .map_err(|source| allocation("comment author payload", source))?;
    data.extend_from_slice(&(length as u32).to_le_bytes());
    for unit in author.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(data)
}

/// Write one comment (`BrtBeginComment`, its `BrtCommentText` when it has
/// text, and `BrtEndComment`) naming author index `author`.
pub(crate) fn write_comment<W: Write>(
    writer: &mut Encoder<W>,
    comment: &Record,
    author: u32,
) -> Result<()> {
    let mut begin = Vec::new();
    begin
        .try_reserve(36)
        .map_err(|source| allocation("comment header", source))?;
    begin.extend_from_slice(&author.to_le_bytes());
    begin.extend_from_slice(&comment.row.to_le_bytes());
    begin.extend_from_slice(&comment.row.to_le_bytes());
    begin.extend_from_slice(&comment.col.to_le_bytes());
    begin.extend_from_slice(&comment.col.to_le_bytes());
    begin.extend_from_slice(&comment.guid);
    writer.write_record(kind::BEGIN_COMMENT, &begin)?;
    if !comment.text.is_empty() {
        let rich = RichString {
            text: comment.text.clone(),
            runs: comment.runs.clone(),
            is_rich: true,
            is_extended: false,
        };
        let data = rich.to_comment_bytes()?;
        writer.write_record(kind::COMMENT_TEXT, &data)?;
    }
    writer.write_record(kind::END_COMMENT, &[])?;
    Ok(())
}

/// A record's byte offset in its part, saturated for the tolerance report.
fn record_location(offset: usize) -> u32 {
    u32::try_from(offset).unwrap_or(u32::MAX)
}

fn collect_records(bytes: &[u8], limits: crate::raw::Limits) -> Result<Vec<RawRecord<'_>>> {
    let mut records = Vec::new();
    for record in Records::new(bytes, limits) {
        reserve(&mut records, 1, "comments records")?;
        records.push(record?);
    }
    Ok(records)
}

fn expect(records: &[RawRecord<'_>], pos: &mut usize, expected: crate::raw::Kind) -> Result<()> {
    let record = records
        .get(*pos)
        .ok_or(Error::InvalidRecordType(expected.get()))?;
    if record.kind() != expected {
        return Err(Error::InvalidRecordType(record.kind().get()));
    }
    if !record.payload().is_empty() {
        return Err(Error::InvalidLength {
            expected: 0,
            found: record.payload().len(),
        });
    }
    *pos += 1;
    Ok(())
}

fn decode_wide_string(
    data: &[u8],
    limits: crate::raw::Limits,
    tolerance: &mut Tolerance<'_>,
    record_type: crate::raw::Kind,
    location: u32,
) -> Result<(String, usize)> {
    let mut cursor = Cursor::new(data, "XLWideString", limits);
    let value = tolerance.read_wide_string(&mut cursor, record_type, location)?;
    Ok((value, cursor.position()))
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

struct RichString {
    text: String,
    runs: Vec<Run>,
    is_rich: bool,
    is_extended: bool,
}

impl RichString {
    fn parse_with(
        data: &[u8],
        limits: crate::raw::Limits,
        location: u32,
        tolerance: &mut Tolerance<'_>,
    ) -> Result<Self> {
        if data.is_empty() {
            return Err(Error::InvalidLength {
                expected: 1,
                found: 0,
            });
        }
        let is_rich = data[0] & 1 != 0;
        let is_extended = data[0] & 2 != 0;
        let (text, consumed) =
            decode_wide_string(&data[1..], limits, tolerance, kind::COMMENT_TEXT, location)?;
        let text_len = text.encode_utf16().count();
        if text_len > MAX_TEXT_UNITS {
            return Err(Error::Unrecognized {
                typ: "RichStr text length",
                val: "unexpected value",
            });
        }
        let mut offset = 1usize
            .checked_add(consumed)
            .ok_or(Error::Encoding("RichStr offset overflow"))?;
        let mut runs = Vec::new();
        if is_rich {
            let count = read_count(data, &mut offset, "StrRun")?;
            let byte_count = count
                .checked_mul(4)
                .ok_or(Error::Encoding("StrRun byte count overflow"))?;
            let end = offset
                .checked_add(byte_count)
                .ok_or(Error::Encoding("StrRun offset overflow"))?;
            if end > data.len() {
                return Err(Error::InvalidLength {
                    expected: end,
                    found: data.len(),
                });
            }
            reserve(&mut runs, count, "comment rich-string runs")?;
            let mut previous = None;
            for chunk in data[offset..end].as_chunks::<4>().0.iter() {
                let character_index = u16::from_le_bytes([chunk[0], chunk[1]]);
                if usize::from(character_index) >= text_len
                    || previous.is_some_and(|value| character_index <= value)
                {
                    return Err(Error::Unrecognized {
                        typ: "StrRun ich",
                        val: "unexpected value",
                    });
                }
                previous = Some(character_index);
                runs.push(Run {
                    character_index,
                    font_id: u16::from_le_bytes([chunk[2], chunk[3]]),
                });
            }
            offset = end;
        }

        if is_extended {
            let (phonetic_text, consumed) = decode_wide_string(
                &data[offset..],
                limits,
                tolerance,
                kind::COMMENT_TEXT,
                location,
            )?;
            offset = offset
                .checked_add(consumed)
                .ok_or(Error::Encoding("phonetic text offset overflow"))?;
            let phonetic_len = phonetic_text.encode_utf16().count();
            let count = read_count(data, &mut offset, "PhRun")?;
            let byte_count = count
                .checked_mul(6)
                .ok_or(Error::Encoding("PhRun byte count overflow"))?;
            let runs_end = offset
                .checked_add(byte_count)
                .ok_or(Error::Encoding("PhRun offset overflow"))?;
            let end = runs_end
                .checked_add(4)
                .ok_or(Error::Encoding("phonetic settings offset overflow"))?;
            if end > data.len() {
                return Err(Error::InvalidLength {
                    expected: end,
                    found: data.len(),
                });
            }
            let mut previous_phonetic = None;
            let mut previous_base_end = None;
            for chunk in data[offset..runs_end].as_chunks::<6>().0.iter() {
                let phonetic_character_index = u16::from_le_bytes([chunk[0], chunk[1]]);
                let base_character_index = u16::from_le_bytes([chunk[2], chunk[3]]);
                let base_character_count = u16::from_le_bytes([chunk[4], chunk[5]]);
                let base_end = usize::from(base_character_index)
                    .checked_add(usize::from(base_character_count))
                    .ok_or(Error::Encoding("PhRun range overflow"))?;
                if usize::from(phonetic_character_index) >= phonetic_len
                    || usize::from(base_character_index) >= text_len
                    || base_end > text_len
                    || previous_phonetic.is_some_and(|value| phonetic_character_index <= value)
                    || previous_base_end
                        .is_some_and(|value| usize::from(base_character_index) < value)
                {
                    return Err(Error::Unrecognized {
                        typ: "PhRun index",
                        val: "/",
                    });
                }
                previous_phonetic = Some(phonetic_character_index);
                previous_base_end = Some(base_end);
            }
            offset = end;
        }

        let trailing = data.len().saturating_sub(offset);
        tolerance.tolerate_trailing(kind::COMMENT_TEXT, location, trailing, || {
            Error::Unrecognized {
                typ: "RichStr",
                val: "trailing bytes",
            }
        })?;
        Ok(Self {
            text,
            runs,
            is_rich,
            is_extended,
        })
    }

    fn to_comment_bytes(&self) -> Result<Vec<u8>> {
        let text_len = self.text.encode_utf16().count();
        if text_len > MAX_TEXT_UNITS {
            return Err(Error::Encoding("RichStr text exceeds 32,767 characters"));
        }
        if self.is_extended {
            return Err(Error::UnsupportedFeature(
                "phonetic metadata is not permitted in BrtCommentText",
            ));
        }
        let mut runs = self.runs.clone();
        if runs.is_empty() && text_len != 0 {
            reserve(&mut runs, 1, "comment rich-string runs")?;
            runs.push(Run {
                character_index: 0,
                font_id: 0,
            });
        }
        validate_runs(&runs, text_len)?;
        let text_bytes = text_len
            .checked_mul(2)
            .ok_or(Error::Encoding("RichStr text byte count overflow"))?;
        let run_bytes = runs
            .len()
            .checked_mul(4)
            .ok_or(Error::Encoding("RichStr run byte count overflow"))?;
        let capacity = 1usize
            .checked_add(4)
            .and_then(|value| value.checked_add(text_bytes))
            .and_then(|value| value.checked_add(4))
            .and_then(|value| value.checked_add(run_bytes))
            .ok_or(Error::Encoding("RichStr size overflow"))?;
        let mut data = Vec::new();
        data.try_reserve(capacity)
            .map_err(|source| allocation("comment rich-string payload", source))?;
        data.push(1);
        data.extend_from_slice(&(text_len as u32).to_le_bytes());
        for unit in self.text.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data.extend_from_slice(&(runs.len() as u32).to_le_bytes());
        for run in runs {
            data.extend_from_slice(&run.character_index.to_le_bytes());
            data.extend_from_slice(&run.font_id.to_le_bytes());
        }
        Ok(data)
    }
}

fn validate_runs(runs: &[Run], text_len: usize) -> Result<()> {
    if runs.len() > MAX_TEXT_RUNS {
        return Err(Error::Encoding("too many RichStr runs"));
    }
    let mut previous = None;
    for run in runs {
        if usize::from(run.character_index) >= text_len
            || previous.is_some_and(|value| run.character_index <= value)
        {
            return Err(Error::Encoding("invalid RichStr run ordering or index"));
        }
        previous = Some(run.character_index);
    }
    Ok(())
}

fn read_count(data: &[u8], offset: &mut usize, _context: &str) -> Result<usize> {
    let end = offset
        .checked_add(4)
        .ok_or(Error::Encoding("count offset overflow"))?;
    if end > data.len() {
        return Err(Error::InvalidLength {
            expected: end,
            found: data.len(),
        });
    }
    let count = read_u32(data, *offset) as usize;
    *offset = end;
    if count > MAX_TEXT_RUNS {
        return Err(Error::Unrecognized {
            typ: "count",
            val: "unexpected value",
        });
    }
    Ok(count)
}

fn reserve<T>(collection: &mut T, additional: usize, resource: &'static str) -> Result<()>
where
    T: TryReserve,
{
    collection
        .try_reserve(additional)
        .map_err(|source| allocation(resource, source))
}

trait TryReserve {
    fn try_reserve(&mut self, additional: usize) -> std::result::Result<(), TryReserveError>;
}

impl<T> TryReserve for Vec<T> {
    fn try_reserve(&mut self, additional: usize) -> std::result::Result<(), TryReserveError> {
        Vec::try_reserve(self, additional)
    }
}

impl<T> TryReserve for HashSet<T>
where
    T: Eq + std::hash::Hash,
{
    fn try_reserve(&mut self, additional: usize) -> std::result::Result<(), TryReserveError> {
        HashSet::try_reserve(self, additional)
    }
}

impl<K, V> TryReserve for HashMap<K, V>
where
    K: Eq + std::hash::Hash,
{
    fn try_reserve(&mut self, additional: usize) -> std::result::Result<(), TryReserveError> {
        HashMap::try_reserve(self, additional)
    }
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn comment_stream(comment: &Record) -> Vec<u8> {
        let mut bytes = Vec::new();
        write(
            &mut Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT),
            std::slice::from_ref(comment),
        )
        .unwrap();
        bytes
    }

    #[test]
    fn round_trips_rich_comment_text() {
        let mut comment = Record::new(2, 3, "Sven".to_string(), "A😀B".to_string());
        comment.runs = vec![
            Run {
                character_index: 0,
                font_id: 4,
            },
            Run {
                character_index: 3,
                font_id: 9,
            },
        ];
        comment.guid = [7; 16];

        let parsed = read(&comment_stream(&comment), crate::raw::Limits::DEFAULT).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].row, 2);
        assert_eq!(parsed[0].col, 3);
        assert_eq!(parsed[0].author, "Sven");
        assert_eq!(parsed[0].text, comment.text);
        assert_eq!(parsed[0].runs, comment.runs);
        assert_eq!(parsed[0].guid, [7; 16]);
        assert_eq!(parsed[0].alternate_guid, None);
        assert!(!parsed[0].visible);
    }

    #[test]
    fn required_raw_limits_bound_comment_record_payloads() {
        let comment = Record::new(2, 3, "Sven".to_string(), "text".to_string());
        let bytes = comment_stream(&comment);
        let limits = crate::raw::Limits::new(35, crate::raw::Limits::DEFAULT.string_units());
        assert!(matches!(
            read(&bytes, limits),
            Err(Error::Wire(crate::raw::Error::PayloadLimit {
                length: 36,
                limit: 35,
                ..
            }))
        ));
    }

    #[test]
    fn reads_alternate_uid_and_future_record_wrappers() {
        let mut bytes = Vec::new();
        let mut writer = Encoder::new(&mut bytes, crate::raw::Limits::DEFAULT);
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

        let comments = read(&bytes, crate::raw::Limits::DEFAULT).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].alternate_guid, Some([9; 16]));
        assert_eq!(comments[0].guid, [7; 16]);
    }

    #[test]
    fn rejects_invalid_comment_text_flags_and_duplicate_cells() {
        let comment = Record::new(0, 0, "A".to_string(), "text".to_string());
        let bytes = comment_stream(&comment);
        let mut duplicate = Vec::new();
        write(
            &mut Encoder::new(&mut duplicate, crate::raw::Limits::DEFAULT),
            &[comment.clone(), comment],
        )
        .unwrap_err();
        assert!(read(&bytes, crate::raw::Limits::DEFAULT).is_ok());

        let mut malformed = Vec::new();
        let mut writer = Encoder::new(&mut malformed, crate::raw::Limits::DEFAULT);
        writer.write_record(kind::BEGIN_COMMENTS, &[]).unwrap();
        writer
            .write_record(kind::BEGIN_COMMENT_AUTHORS, &[])
            .unwrap();
        writer
            .write_record(kind::COMMENT_AUTHOR, &[1, 0, 0, 0, b'A', 0])
            .unwrap();
        writer.write_record(kind::END_COMMENT_AUTHORS, &[]).unwrap();
        writer.write_record(kind::BEGIN_COMMENT_LIST, &[]).unwrap();
        let mut begin = vec![0; 36];
        begin[4..8].copy_from_slice(&0u32.to_le_bytes());
        begin[8..12].copy_from_slice(&0u32.to_le_bytes());
        begin[12..16].copy_from_slice(&0u32.to_le_bytes());
        begin[16..20].copy_from_slice(&0u32.to_le_bytes());
        writer.write_record(kind::BEGIN_COMMENT, &begin).unwrap();
        writer
            .write_record(kind::COMMENT_TEXT, &[0, 0, 0, 0, 0])
            .unwrap();
        writer.write_record(kind::END_COMMENT, &[]).unwrap();
        writer.write_record(kind::END_COMMENT_LIST, &[]).unwrap();
        writer.write_record(kind::END_COMMENTS, &[]).unwrap();
        assert!(read(&malformed, crate::raw::Limits::DEFAULT).is_err());
    }
}
