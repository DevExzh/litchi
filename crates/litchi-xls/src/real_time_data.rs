//! BIFF8 `RealTimeData` record (MS-XLS 2.4.214): real-time data (RTD)
//! topics in the workbook globals substream.
//!
//! There is one `RealTimeData` record per RTD topic; the `RTD` production
//! (MS-XLS 2.1) is `RealTimeData *ContinueFrt`, so the logical payload is the
//! record body followed by any `ContinueFrt` bodies concatenated. Each record
//! carries the topic as an `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298)
//! whose first sub-string is the RTD server ProgID and whose second is the
//! server name, the last value returned by the server as an `RTDOper`
//! variant (MS-XLS 2.5.224), and the cells subscribed to the topic as
//! `RTDEItem` entries (MS-XLS 2.5.223).
//!
//! Adjacent records share topic prefixes: `ichSamePrefix` counts the leading
//! characters this topic has in common with the previous record's topic, and
//! the stored string holds only the remainder. The fully reconstructed topic
//! is exposed as [`XlsRealTimeData::topic`]; pass the previous topic to
//! [`XlsRealTimeData::parse`] so the prefix can be re-applied.
//!
//! Everything in this module is INERT: ProgIDs, server names, and topics are
//! stored verbatim and no RTD server is ever located, launched, or queried.

use super::{XlsError, XlsResult};

/// Record type of the `RealTimeData` record (MS-XLS 2.4.214).
pub(crate) const REAL_TIME_DATA_RECORD_TYPE: u16 = 0x0813;
/// Record type of the `ContinueFrt` record (MS-XLS 2.4.60) that continues a
/// `RealTimeData` payload.
pub(crate) const CONTINUE_FRT_RECORD_TYPE: u16 = 0x0812;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of an `RTDEItem` structure (MS-XLS 2.5.223).
const RTD_E_ITEM_LEN: usize = 6;

// `RTDOper.grbit` variant-kind codes (MS-XLS 2.5.224).
const RTD_OPER_NUMBER: u32 = 0x0000_0001;
const RTD_OPER_SHORT_TEXT: u32 = 0x0000_0002;
const RTD_OPER_BOOLEAN: u32 = 0x0000_0004;
const RTD_OPER_ERROR: u32 = 0x0000_0010;
const RTD_OPER_INTEGER: u32 = 0x0000_0800;
const RTD_OPER_LONG_TEXT: u32 = 0x0000_1000;

/// `fHighByte` bit of a BIFF8 string option byte.
const HIGH_BYTE: u8 = 0x01;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: REAL_TIME_DATA_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

/// Decode `char_count` characters from `bytes` in compressed (1 byte/char)
/// or uncompressed UTF-16LE (2 bytes/char) form.
fn decode_chars(bytes: &[u8], wide: bool) -> XlsResult<String> {
    if wide {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units).map_err(|_| invalid("RTD string is not valid UTF-16LE"))
    } else {
        Ok(bytes.iter().map(|&byte| char::from(byte)).collect())
    }
}

/// The last value an RTD server returned for a topic (`RTDOper.rtdVt`).
#[derive(Debug, Clone, PartialEq)]
pub enum XlsRtdValue {
    /// A floating-point value (`Xnum`).
    Number(f64),
    /// A text value (`RTDOperStr`).
    Text(String),
    /// A Boolean value.
    Boolean(bool),
    /// A signed integer that indicates an error code.
    Error(i32),
    /// A signed integer used for purposes other than an error code.
    Integer(i32),
}

/// A cell subscribed to an RTD topic (`RTDEItem`, MS-XLS 2.5.223).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsRtdCell {
    /// Zero-based row index of the cell.
    pub row: u16,
    /// Zero-based column index of the cell.
    pub column: u8,
    /// Zero-based index of the sheet containing the cell (`TabIndex`).
    pub sheet_index: u16,
}

impl XlsRtdCell {
    /// Create a checked RTD subscriber cell from raw zero-based indices.
    pub fn new(row: u32, column: u16, sheet_index: usize) -> XlsResult<Self> {
        let invalid = || {
            XlsError::InvalidCellReference(format!(
                "RTD subscriber row {row}, column {column} is outside the BIFF8 grid"
            ))
        };
        let row = u16::try_from(row).map_err(|_| invalid())?;
        let column = u8::try_from(column).map_err(|_| invalid())?;
        let sheet_index = u16::try_from(sheet_index)
            .map_err(|_| XlsError::WorksheetNotFound(format!("Sheet {sheet_index}")))?;
        Ok(Self {
            row,
            column,
            sheet_index,
        })
    }
}

/// Typed `RealTimeData` record content (MS-XLS 2.4.214).
#[derive(Debug, Clone, PartialEq)]
pub struct XlsRealTimeData {
    /// Number of leading characters this topic shares with the previous
    /// record's topic (`ichSamePrefix`); always zero for the first record.
    pub common_prefix_len: u32,
    /// The topic sub-strings as stored (without the shared prefix). The
    /// first is the RTD server ProgID, the second the server name (empty for
    /// a local server), and the rest combine into the unique topic.
    pub topic_segments: Vec<String>,
    /// The fully reconstructed topic: the shared prefix of the previous
    /// topic followed by the stored sub-strings.
    pub topic: String,
    /// The last value returned by the RTD server (`rtdOper`).
    pub value: XlsRtdValue,
    /// The cells subscribed to this topic (`rgRTDE`).
    pub cells: Vec<XlsRtdCell>,
}

impl XlsRealTimeData {
    /// Parse one logical `RealTimeData` payload: the record body with any
    /// `ContinueFrt` bodies already appended.
    ///
    /// `previous_topic` is the reconstructed [`XlsRealTimeData::topic`] of
    /// the preceding `RealTimeData` record in the globals substream, needed
    /// to re-apply prefix compression; pass `None` for the first record.
    pub fn parse(data: &[u8], previous_topic: Option<&str>) -> XlsResult<Self> {
        if data.len() < FRT_HEADER_LEN + 4 {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_LEN + 4,
                found: data.len(),
            });
        }
        if read_u16(data, 0) != REAL_TIME_DATA_RECORD_TYPE {
            return Err(invalid("RealTimeData FrtHeader.rt mismatch"));
        }

        let common_prefix_len = read_u32(data, FRT_HEADER_LEN);
        let mut offset = FRT_HEADER_LEN + 4;

        // stTopic: XLUnicodeStringSegmentedRTD (MS-XLS 2.5.298).
        let (topic_segments, used) = parse_segmented_topic(&data[offset..])?;
        offset += used;

        // rtdOper: RTDOper (MS-XLS 2.5.224).
        let (value, used) = parse_rtd_oper(&data[offset..])?;
        offset += used;

        // rgRTDE: the rest of the payload in 6-byte RTDEItem entries.
        let remaining = &data[offset..];
        if !remaining.len().is_multiple_of(RTD_E_ITEM_LEN) {
            return Err(invalid("RealTimeData rgRTDE size is not a multiple of 6"));
        }
        let cells = remaining
            .chunks_exact(RTD_E_ITEM_LEN)
            .map(|chunk| {
                let column = u8::try_from(read_u16(chunk, 2))
                    .map_err(|_| invalid("RTD subscriber column exceeds the BIFF8 grid"))?;
                Ok(XlsRtdCell {
                    row: read_u16(chunk, 0),
                    column,
                    sheet_index: read_u16(chunk, 4),
                })
            })
            .collect::<XlsResult<Vec<_>>>()?;

        // Re-apply prefix compression against the previous topic.
        let stored: String = topic_segments.concat();
        let topic = if common_prefix_len == 0 {
            stored
        } else {
            let previous = previous_topic
                .ok_or_else(|| invalid("first RealTimeData record declares a shared prefix"))?;
            let prefix_len = usize::try_from(common_prefix_len)
                .map_err(|_| invalid("RealTimeData ichSamePrefix overflows"))?;
            if prefix_len > previous.chars().count() {
                return Err(invalid(
                    "RealTimeData ichSamePrefix exceeds the previous topic",
                ));
            }
            let mut topic = String::with_capacity(prefix_len + stored.len());
            topic.extend(previous.chars().take(prefix_len));
            topic.push_str(&stored);
            topic
        };

        Ok(XlsRealTimeData {
            common_prefix_len,
            topic_segments,
            topic,
            value,
            cells,
        })
    }

    /// Serialize back to a complete logical `RealTimeData` payload (the
    /// record body; the writer chunks it into `ContinueFrt` records when it
    /// exceeds the maximum record size).
    ///
    /// The stored topic sub-strings are written as-is, so a value parsed
    /// from a workbook round-trips exactly; `topic` is re-derived from
    /// `common_prefix_len` and the previous record on the next parse.
    pub(crate) fn to_payload(&self) -> XlsResult<Vec<u8>> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&REAL_TIME_DATA_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0u8; FRT_HEADER_LEN - 2]); // grbitFrt + reserved
        payload.extend_from_slice(&self.common_prefix_len.to_le_bytes());
        write_segmented_topic(&mut payload, &self.topic_segments)?;
        match &self.value {
            XlsRtdValue::Number(value) => {
                payload.extend_from_slice(&RTD_OPER_NUMBER.to_le_bytes());
                payload.extend_from_slice(&value.to_le_bytes());
            },
            XlsRtdValue::Text(text) => {
                let char_count = biff_char_count(text);
                let kind = if char_count < 256 {
                    RTD_OPER_SHORT_TEXT
                } else {
                    RTD_OPER_LONG_TEXT
                };
                payload.extend_from_slice(&kind.to_le_bytes());
                payload.extend_from_slice(&(char_count as u32).to_le_bytes());
                write_chars(&mut payload, text);
            },
            XlsRtdValue::Boolean(value) => {
                payload.extend_from_slice(&RTD_OPER_BOOLEAN.to_le_bytes());
                payload.extend_from_slice(&u32::from(*value).to_le_bytes());
            },
            XlsRtdValue::Error(value) => {
                payload.extend_from_slice(&RTD_OPER_ERROR.to_le_bytes());
                payload.extend_from_slice(&value.to_le_bytes());
            },
            XlsRtdValue::Integer(value) => {
                payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
                payload.extend_from_slice(&value.to_le_bytes());
            },
        }
        for cell in &self.cells {
            payload.extend_from_slice(&cell.row.to_le_bytes());
            payload.extend_from_slice(&u16::from(cell.column).to_le_bytes());
            payload.extend_from_slice(&cell.sheet_index.to_le_bytes());
        }
        Ok(payload)
    }
}

/// Whether a string can be stored compressed (every character in U+0000..=U+00FF).
fn is_compressible(text: &str) -> bool {
    text.chars().all(|ch| u32::from(ch) <= 0xFF)
}

/// Character count as BIFF stores it: byte count when the string is
/// compressible, UTF-16 code units otherwise.
fn biff_char_count(text: &str) -> usize {
    if is_compressible(text) {
        text.len()
    } else {
        text.encode_utf16().count()
    }
}

/// Append the option byte and characters of an `XLUnicodeStringNoCch`.
fn write_chars(out: &mut Vec<u8>, text: &str) {
    if is_compressible(text) {
        out.push(0u8); // fHighByte = 0
        out.extend(text.chars().map(|ch| ch as u8));
    } else {
        out.push(HIGH_BYTE);
        for unit in text.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

/// Serialize an `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298).
fn write_segmented_topic(out: &mut Vec<u8>, segments: &[String]) -> XlsResult<()> {
    let compressible = segments.iter().all(|segment| is_compressible(segment));
    // Compressed sub-string counts are one byte, wide counts two bytes
    // (MS-XLS 2.5.298); pick the narrowest encoding that fits every segment.
    let wide = !compressible
        || segments
            .iter()
            .any(|segment| segment.len() > usize::from(u8::MAX));
    if segments
        .iter()
        .any(|segment| biff_char_count(segment) > usize::from(u16::MAX))
    {
        return Err(XlsError::InvalidData(
            "RTD topic sub-string exceeds 65535 characters".to_string(),
        ));
    }
    let char_total: usize = segments
        .iter()
        .map(|segment| biff_char_count(segment))
        .sum();
    out.extend_from_slice(&(char_total as u32).to_le_bytes());
    out.push(if wide { HIGH_BYTE } else { 0u8 });
    for segment in segments {
        if wide {
            out.extend_from_slice(&(biff_char_count(segment) as u16).to_le_bytes());
            for unit in segment.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes());
            }
        } else {
            out.push(segment.len() as u8);
            out.extend(segment.chars().map(|ch| ch as u8));
        }
    }
    Ok(())
}

/// Parse an `XLUnicodeStringSegmentedRTD`; returns the sub-strings and the
/// number of bytes consumed.
fn parse_segmented_topic(data: &[u8]) -> XlsResult<(Vec<String>, usize)> {
    if data.len() < 5 {
        return Err(XlsError::InvalidLength {
            expected: 5,
            found: data.len(),
        });
    }
    let char_total = usize::try_from(read_u32(data, 0))
        .map_err(|_| invalid("RealTimeData stTopic.cch overflows"))?;
    let wide = data[4] & HIGH_BYTE != 0;
    let mut offset = 5;
    let mut segments = Vec::new();
    let mut chars_read = 0usize;
    while chars_read < char_total {
        // Each sub-string starts with a count: 1 byte compressed, 2 wide.
        let count_len = if wide { 2 } else { 1 };
        let count_bytes = data
            .get(offset..offset + count_len)
            .ok_or(XlsError::InvalidLength {
                expected: offset + count_len,
                found: data.len(),
            })?;
        let segment_chars = if wide {
            usize::from(read_u16(count_bytes, 0))
        } else {
            usize::from(count_bytes[0])
        };
        offset += count_len;
        if chars_read + segment_chars > char_total {
            return Err(invalid("RealTimeData stTopic sub-string overruns cch"));
        }
        let byte_len = if wide {
            segment_chars * 2
        } else {
            segment_chars
        };
        let bytes = data
            .get(offset..offset + byte_len)
            .ok_or(XlsError::InvalidLength {
                expected: offset + byte_len,
                found: data.len(),
            })?;
        segments.push(decode_chars(bytes, wide)?);
        offset += byte_len;
        chars_read += segment_chars;
    }
    Ok((segments, offset))
}

/// Parse an `RTDOper` variant; returns the value and the number of bytes
/// consumed.
fn parse_rtd_oper(data: &[u8]) -> XlsResult<(XlsRtdValue, usize)> {
    if data.len() < 4 {
        return Err(XlsError::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let kind = read_u32(data, 0);
    let body = &data[4..];
    match kind {
        RTD_OPER_NUMBER => {
            let bytes = body.get(..8).ok_or(XlsError::InvalidLength {
                expected: 12,
                found: data.len(),
            })?;
            let bytes = <[u8; 8]>::try_from(bytes).map_err(|_| XlsError::InvalidLength {
                expected: 12,
                found: data.len(),
            })?;
            Ok((XlsRtdValue::Number(f64::from_le_bytes(bytes)), 12))
        },
        RTD_OPER_SHORT_TEXT | RTD_OPER_LONG_TEXT => {
            let (text, used) = parse_rtd_oper_str(body)?;
            Ok((XlsRtdValue::Text(text), 4 + used))
        },
        RTD_OPER_BOOLEAN => {
            let raw = body.get(..4).ok_or(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let value = match read_u32(raw, 0) {
                0 => false,
                1 => true,
                other => {
                    return Err(invalid(format!("invalid RTD Boolean value {other}")));
                },
            };
            Ok((XlsRtdValue::Boolean(value), 8))
        },
        RTD_OPER_ERROR | RTD_OPER_INTEGER => {
            let raw = body.get(..4).ok_or(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let raw = <[u8; 4]>::try_from(raw).map_err(|_| XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let value = i32::from_le_bytes(raw);
            let variant = if kind == RTD_OPER_ERROR {
                XlsRtdValue::Error(value)
            } else {
                XlsRtdValue::Integer(value)
            };
            Ok((variant, 8))
        },
        other => Err(invalid(format!(
            "unknown RTDOper.grbit value 0x{other:08X}"
        ))),
    }
}

/// Parse an `RTDOperStr` (MS-XLS 2.5.225): a 4-byte character count followed
/// by an `XLUnicodeStringNoCch`.
fn parse_rtd_oper_str(data: &[u8]) -> XlsResult<(String, usize)> {
    if data.len() < 5 {
        return Err(XlsError::InvalidLength {
            expected: 5,
            found: data.len(),
        });
    }
    let char_count = usize::try_from(read_u32(data, 0))
        .map_err(|_| invalid("RTDOperStr.cchRTDOperStr overflows"))?;
    let wide = data[4] & HIGH_BYTE != 0;
    let byte_len = if wide { char_count * 2 } else { char_count };
    let bytes = data.get(5..5 + byte_len).ok_or(XlsError::InvalidLength {
        expected: 5 + byte_len,
        found: data.len(),
    })?;
    Ok((decode_chars(bytes, wide)?, 5 + byte_len))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build an FrtHeader for the RealTimeData record type.
    fn frt_header() -> Vec<u8> {
        let mut header = Vec::with_capacity(FRT_HEADER_LEN);
        header.extend_from_slice(&REAL_TIME_DATA_RECORD_TYPE.to_le_bytes());
        header.extend_from_slice(&[0u8; 10]); // grbitFrt + reserved
        header
    }

    /// Build a compressed XLUnicodeStringSegmentedRTD from sub-strings.
    fn segmented_topic(segments: &[&str]) -> Vec<u8> {
        let char_total: usize = segments.iter().map(|segment| segment.len()).sum();
        let mut out = Vec::new();
        out.extend_from_slice(&(char_total as u32).to_le_bytes());
        out.push(0u8); // fHighByte = 0
        for segment in segments {
            out.push(segment.len() as u8);
            out.extend_from_slice(segment.as_bytes());
        }
        out
    }

    fn rtd_cell(row: u16, column: u16, sheet_index: u16) -> [u8; 6] {
        let mut cell = [0u8; 6];
        cell[0..2].copy_from_slice(&row.to_le_bytes());
        cell[2..4].copy_from_slice(&column.to_le_bytes());
        cell[4..6].copy_from_slice(&sheet_index.to_le_bytes());
        cell
    }

    #[test]
    fn parses_text_topic_with_cells() {
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes()); // ichSamePrefix
        payload.extend_from_slice(&segmented_topic(&["PROG.ID", "", "STOCK", "MSFT"]));
        payload.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
        payload.extend_from_slice(&5u32.to_le_bytes()); // cchRTDOperStr
        payload.push(0u8); // compressed
        payload.extend_from_slice(b"58.25");
        payload.extend_from_slice(&rtd_cell(1, 2, 0));
        payload.extend_from_slice(&rtd_cell(3, 4, 1));

        let rtd = XlsRealTimeData::parse(&payload, None).expect("parse");
        assert_eq!(rtd.common_prefix_len, 0);
        assert_eq!(rtd.topic_segments, vec!["PROG.ID", "", "STOCK", "MSFT"]);
        assert_eq!(rtd.topic, "PROG.IDSTOCKMSFT");
        assert_eq!(rtd.value, XlsRtdValue::Text("58.25".to_string()));
        assert_eq!(
            rtd.cells,
            vec![
                XlsRtdCell {
                    row: 1,
                    column: 2,
                    sheet_index: 0
                },
                XlsRtdCell {
                    row: 3,
                    column: 4,
                    sheet_index: 1
                },
            ]
        );
    }

    #[test]
    fn parses_numeric_boolean_error_and_integer_values() {
        for (kind, body, expected) in [
            (
                RTD_OPER_NUMBER,
                42.5f64.to_le_bytes().to_vec(),
                XlsRtdValue::Number(42.5),
            ),
            (
                RTD_OPER_BOOLEAN,
                1u32.to_le_bytes().to_vec(),
                XlsRtdValue::Boolean(true),
            ),
            (
                RTD_OPER_ERROR,
                0x2Au32.to_le_bytes().to_vec(),
                XlsRtdValue::Error(0x2A),
            ),
            (
                RTD_OPER_INTEGER,
                (-7i32).to_le_bytes().to_vec(),
                XlsRtdValue::Integer(-7),
            ),
        ] {
            let mut payload = frt_header();
            payload.extend_from_slice(&0u32.to_le_bytes());
            payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
            payload.extend_from_slice(&kind.to_le_bytes());
            payload.extend_from_slice(&body);
            let rtd = XlsRealTimeData::parse(&payload, None).expect("parse");
            assert_eq!(rtd.value, expected);
            assert!(rtd.cells.is_empty());
        }
    }

    #[test]
    fn rejects_invalid_boolean() {
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_BOOLEAN.to_le_bytes());
        payload.extend_from_slice(&7u32.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());
    }

    #[test]
    fn reapplies_shared_prefix_from_previous_topic() {
        let mut first = frt_header();
        first.extend_from_slice(&0u32.to_le_bytes());
        first.extend_from_slice(&segmented_topic(&["PROG.ID", "", "STOCK"]));
        first.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        first.extend_from_slice(&1i32.to_le_bytes());
        let first = XlsRealTimeData::parse(&first, None).expect("parse first");
        assert_eq!(first.topic, "PROG.IDSTOCK");

        // Second record shares the "PROG.ID" prefix (7 characters) and only
        // stores the trailing sub-strings.
        let mut second = frt_header();
        second.extend_from_slice(&7u32.to_le_bytes()); // ichSamePrefix
        second.extend_from_slice(&segmented_topic(&["", "BOND"]));
        second.extend_from_slice(&RTD_OPER_LONG_TEXT.to_le_bytes());
        second.extend_from_slice(&3u32.to_le_bytes());
        second.push(0u8);
        second.extend_from_slice(b"102");
        let second = XlsRealTimeData::parse(&second, Some(&first.topic)).expect("parse second");
        assert_eq!(second.common_prefix_len, 7);
        assert_eq!(second.topic_segments, vec!["", "BOND"]);
        assert_eq!(second.topic, "PROG.IDBOND");
        assert_eq!(second.value, XlsRtdValue::Text("102".to_string()));
    }

    #[test]
    fn rejects_prefix_without_previous_topic() {
        let mut payload = frt_header();
        payload.extend_from_slice(&3u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());
    }

    #[test]
    fn rejects_prefix_longer_than_previous_topic() {
        let mut payload = frt_header();
        payload.extend_from_slice(&9u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, Some("short")).is_err());
    }

    #[test]
    fn parses_wide_strings() {
        let mut topic = Vec::new();
        topic.extend_from_slice(&3u32.to_le_bytes()); // cch
        topic.push(1u8); // fHighByte = 1
        // One wide sub-string of 3 characters: 'A', 'B', '€'.
        topic.extend_from_slice(&3u16.to_le_bytes());
        for unit in [u32::from('A') as u16, u32::from('B') as u16, 0x20AC] {
            topic.extend_from_slice(&unit.to_le_bytes());
        }

        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&topic);
        payload.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.push(1u8); // wide RTDOperStr
        payload.extend_from_slice(&0x20ACu16.to_le_bytes());

        let rtd = XlsRealTimeData::parse(&payload, None).expect("parse");
        assert_eq!(rtd.topic, "AB€");
        assert_eq!(rtd.value, XlsRtdValue::Text("€".to_string()));
    }

    #[test]
    fn rejects_frt_header_rt_mismatch() {
        let mut payload = frt_header();
        payload[0] = 0x12; // corrupt rt to 0x0912... -> mismatch
        payload[1] = 0x09;
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());
    }

    #[test]
    fn rejects_unknown_oper_kind_and_ragged_cells() {
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&0xDEADu32.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());

        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        payload.extend_from_slice(&[0u8; 5]); // not a multiple of 6
        assert!(XlsRealTimeData::parse(&payload, None).is_err());

        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        payload.extend_from_slice(&0i32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&256u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());
    }

    #[test]
    fn rejects_truncated_payloads() {
        assert!(XlsRealTimeData::parse(&[], None).is_err());
        assert!(XlsRealTimeData::parse(&frt_header(), None).is_err());
        let mut payload = frt_header();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
        // RTDOper kind present but the 4-byte body is missing.
        payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes());
        assert!(XlsRealTimeData::parse(&payload, None).is_err());
    }

    #[test]
    fn payload_round_trips() {
        let values = [
            XlsRealTimeData {
                common_prefix_len: 0,
                topic_segments: vec![
                    "PROG.ID".to_string(),
                    String::new(),
                    "STOCK".to_string(),
                    "MSFT".to_string(),
                ],
                topic: "PROG.IDSTOCKMSFT".to_string(),
                value: XlsRtdValue::Text("58.25".to_string()),
                cells: vec![XlsRtdCell {
                    row: 1,
                    column: 2,
                    sheet_index: 0,
                }],
            },
            XlsRealTimeData {
                common_prefix_len: 0,
                topic_segments: vec!["宽".to_string(), "server".to_string(), "€uro".to_string()],
                topic: "宽server€uro".to_string(),
                value: XlsRtdValue::Number(42.5),
                cells: Vec::new(),
            },
            XlsRealTimeData {
                common_prefix_len: 0,
                topic_segments: vec!["A".to_string(), "B".to_string(), "C".to_string()],
                topic: "ABC".to_string(),
                value: XlsRtdValue::Boolean(true),
                cells: vec![
                    XlsRtdCell {
                        row: 0,
                        column: 0,
                        sheet_index: 0,
                    },
                    XlsRtdCell {
                        row: 65535,
                        column: 255,
                        sheet_index: 3,
                    },
                ],
            },
        ];
        for value in values {
            let payload = value.to_payload().expect("serialize");
            let parsed = XlsRealTimeData::parse(&payload, None).expect("re-parse");
            assert_eq!(parsed, value);
        }
    }

    #[test]
    fn payload_round_trips_long_text_variant() {
        let long_text = "x".repeat(300);
        let value = XlsRealTimeData {
            common_prefix_len: 0,
            topic_segments: vec!["A".to_string(), "B".to_string(), "C".to_string()],
            topic: "ABC".to_string(),
            value: XlsRtdValue::Text(long_text),
            cells: Vec::new(),
        };
        let payload = value.to_payload().expect("serialize");
        // grbit 0x1000 selects the long-string RTDOperStr form.
        let kind_offset = payload.len() - 4 - 4 - 1 - 300;
        assert_eq!(read_u32(&payload, kind_offset), RTD_OPER_LONG_TEXT);
        let parsed = XlsRealTimeData::parse(&payload, None).expect("re-parse");
        assert_eq!(parsed, value);
    }

    #[test]
    fn serialize_promotes_long_compressed_segment_to_wide() {
        let value = XlsRealTimeData {
            common_prefix_len: 0,
            topic_segments: vec!["x".repeat(300)],
            topic: String::new(),
            value: XlsRtdValue::Integer(0),
            cells: Vec::new(),
        };
        // A 300-character compressed segment does not fit the 1-byte count,
        // but the wide encoding holds it.
        let payload = value.to_payload().expect("serialize");
        assert_eq!(payload[FRT_HEADER_LEN + 4 + 4], HIGH_BYTE);
    }
}
