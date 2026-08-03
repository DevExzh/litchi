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
/// Maximum logical `RealTimeData` payload after `ContinueFrt` reassembly.
const MAX_LOGICAL_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
/// Maximum retained topic/value character or encoded-unit count imposed before
/// allocation.
const MAX_STRING_CHARACTERS: usize = MAX_LOGICAL_PAYLOAD_BYTES;
/// Minimum and maximum number of segmented topic substrings allowed by
/// `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298).
const MIN_TOPIC_SEGMENTS: usize = 3;
const MIN_PREFIXED_TOPIC_SEGMENTS: usize = 2;
const MAX_TOPIC_SEGMENTS: usize = 39;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: REAL_TIME_DATA_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_bytes<const N: usize>(data: &[u8], offset: usize) -> XlsResult<&[u8]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| invalid("RealTimeData field offset overflows usize"))?;
    data.get(offset..end).ok_or(XlsError::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

fn read_u16(data: &[u8], offset: usize) -> XlsResult<u16> {
    let bytes = read_bytes::<2>(data, offset)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> XlsResult<u32> {
    let bytes = read_bytes::<4>(data, offset)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// Decode `char_count` characters from `bytes` in compressed (1 byte/char)
/// or uncompressed UTF-16LE (2 bytes/char) form.
fn decode_chars(bytes: &[u8], wide: bool) -> XlsResult<String> {
    if wide {
        if !bytes.len().is_multiple_of(2) {
            return Err(invalid("RTD wide string has an odd byte length"));
        }
        let mut value = String::new();
        value
            .try_reserve(bytes.len())
            .map_err(|_| XlsError::Allocation("decoding RTD UTF-16 text"))?;
        for result in char::decode_utf16(
            bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
        ) {
            value.push(result.map_err(|_| invalid("RTD string is not valid UTF-16LE"))?);
        }
        Ok(value)
    } else {
        let mut value = String::new();
        value
            .try_reserve(bytes.len())
            .map_err(|_| XlsError::Allocation("decoding RTD compressed text"))?;
        value.extend(bytes.iter().map(|&byte| char::from(byte)));
        Ok(value)
    }
}

fn join_segments(segments: &[String]) -> XlsResult<String> {
    let byte_len = segments.iter().try_fold(0usize, |total, segment| {
        total
            .checked_add(segment.len())
            .ok_or_else(|| invalid("RTD topic byte length overflows usize"))
    })?;
    let mut value = String::new();
    value
        .try_reserve(byte_len)
        .map_err(|_| XlsError::Allocation("reassembling RTD topic text"))?;
    for segment in segments {
        value.push_str(segment);
    }
    Ok(value)
}

/// Fallible bounded output buffer for a serialized logical RTD record.
struct Payload {
    bytes: Vec<u8>,
}

impl Payload {
    fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn push(&mut self, byte: u8) -> XlsResult<()> {
        if self.bytes.len() >= MAX_LOGICAL_PAYLOAD_BYTES {
            return Err(invalid(format!(
                "serialized RealTimeData payload exceeds {MAX_LOGICAL_PAYLOAD_BYTES} bytes"
            )));
        }
        if self.bytes.len() == self.bytes.capacity() {
            self.bytes
                .try_reserve(1)
                .map_err(|_| XlsError::Allocation("serializing RealTimeData payload"))?;
        }
        self.bytes.push(byte);
        Ok(())
    }

    fn extend_from_slice(&mut self, bytes: &[u8]) -> XlsResult<()> {
        let new_len = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| invalid("serialized RealTimeData payload length overflows"))?;
        if new_len > MAX_LOGICAL_PAYLOAD_BYTES {
            return Err(invalid(format!(
                "serialized RealTimeData payload exceeds {MAX_LOGICAL_PAYLOAD_BYTES} bytes"
            )));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|_| XlsError::Allocation("serializing RealTimeData payload"))?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    fn into_vec(self) -> Vec<u8> {
        self.bytes
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
        if data.len() > MAX_LOGICAL_PAYLOAD_BYTES {
            return Err(invalid(format!(
                "RealTimeData payload exceeds {MAX_LOGICAL_PAYLOAD_BYTES} bytes"
            )));
        }
        if data.len() < FRT_HEADER_LEN + 4 {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_LEN + 4,
                found: data.len(),
            });
        }
        if read_u16(data, 0)? != REAL_TIME_DATA_RECORD_TYPE {
            return Err(invalid("RealTimeData FrtHeader.rt mismatch"));
        }

        let common_prefix_len = read_u32(data, FRT_HEADER_LEN)?;
        let mut offset = FRT_HEADER_LEN + 4;

        // stTopic: XLUnicodeStringSegmentedRTD (MS-XLS 2.5.298).
        let (topic_segments, used) = parse_segmented_topic(
            data.get(offset..).ok_or(XlsError::InvalidLength {
                expected: offset,
                found: data.len(),
            })?,
            common_prefix_len != 0,
        )?;
        offset = offset
            .checked_add(used)
            .ok_or_else(|| invalid("RealTimeData topic offset overflows usize"))?;

        // rtdOper: RTDOper (MS-XLS 2.5.224).
        let (value, used) = parse_rtd_oper(data.get(offset..).ok_or(XlsError::InvalidLength {
            expected: offset,
            found: data.len(),
        })?)?;
        offset = offset
            .checked_add(used)
            .ok_or_else(|| invalid("RealTimeData value offset overflows usize"))?;

        // rgRTDE: the rest of the payload in 6-byte RTDEItem entries.
        let remaining = data.get(offset..).ok_or(XlsError::InvalidLength {
            expected: offset,
            found: data.len(),
        })?;
        if !remaining.len().is_multiple_of(RTD_E_ITEM_LEN) {
            return Err(invalid("RealTimeData rgRTDE size is not a multiple of 6"));
        }
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(remaining.len() / RTD_E_ITEM_LEN)
            .map_err(|_| XlsError::Allocation("retaining RTD subscriber cells"))?;
        for chunk in remaining.chunks_exact(RTD_E_ITEM_LEN) {
            let column = u8::try_from(read_u16(chunk, 2)?)
                .map_err(|_| invalid("RTD subscriber column exceeds the BIFF8 grid"))?;
            cells.push(XlsRtdCell {
                row: read_u16(chunk, 0)?,
                column,
                sheet_index: read_u16(chunk, 4)?,
            });
        }

        // Re-apply prefix compression against the previous topic.
        let stored = join_segments(&topic_segments)?;
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
            let capacity = prefix_len
                .checked_add(stored.len())
                .ok_or_else(|| invalid("reconstructed RealTimeData topic length overflows"))?;
            let mut topic = String::new();
            topic
                .try_reserve(capacity)
                .map_err(|_| XlsError::Allocation("reconstructing RealTimeData topic"))?;
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
        let common_prefix_len = usize::try_from(self.common_prefix_len)
            .map_err(|_| invalid("RTD common prefix does not fit in usize"))?;
        if common_prefix_len > MAX_STRING_CHARACTERS {
            return Err(invalid(
                "RTD common prefix exceeds the string resource limit",
            ));
        }
        let minimum_segments = if common_prefix_len == 0 {
            MIN_TOPIC_SEGMENTS
        } else {
            MIN_PREFIXED_TOPIC_SEGMENTS
        };
        if !(minimum_segments..=MAX_TOPIC_SEGMENTS).contains(&self.topic_segments.len()) {
            return Err(invalid(format!(
                "RTD topic must have {minimum_segments}..={MAX_TOPIC_SEGMENTS} segments"
            )));
        }
        if self.cells.len() > MAX_LOGICAL_PAYLOAD_BYTES / RTD_E_ITEM_LEN {
            return Err(invalid(
                "RTD subscriber cell count exceeds the resource limit",
            ));
        }
        let mut payload = Payload::new();
        payload.extend_from_slice(&REAL_TIME_DATA_RECORD_TYPE.to_le_bytes())?;
        payload.extend_from_slice(&[0u8; FRT_HEADER_LEN - 2])?; // grbitFrt + reserved
        payload.extend_from_slice(&self.common_prefix_len.to_le_bytes())?;
        write_segmented_topic(&mut payload, &self.topic_segments, minimum_segments)?;
        match &self.value {
            XlsRtdValue::Number(value) => {
                payload.extend_from_slice(&RTD_OPER_NUMBER.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
            XlsRtdValue::Text(text) => {
                let char_count = biff_char_count(text);
                if char_count > MAX_STRING_CHARACTERS {
                    return Err(invalid("RTD text exceeds the string resource limit"));
                }
                let char_count = u32::try_from(char_count)
                    .map_err(|_| invalid("RTD text character count overflows u32"))?;
                let kind = if char_count < 256 {
                    RTD_OPER_SHORT_TEXT
                } else {
                    RTD_OPER_LONG_TEXT
                };
                payload.extend_from_slice(&kind.to_le_bytes())?;
                payload.extend_from_slice(&char_count.to_le_bytes())?;
                write_chars(&mut payload, text)?;
            },
            XlsRtdValue::Boolean(value) => {
                payload.extend_from_slice(&RTD_OPER_BOOLEAN.to_le_bytes())?;
                payload.extend_from_slice(&u32::from(*value).to_le_bytes())?;
            },
            XlsRtdValue::Error(value) => {
                payload.extend_from_slice(&RTD_OPER_ERROR.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
            XlsRtdValue::Integer(value) => {
                payload.extend_from_slice(&RTD_OPER_INTEGER.to_le_bytes())?;
                payload.extend_from_slice(&value.to_le_bytes())?;
            },
        }
        for cell in &self.cells {
            payload.extend_from_slice(&cell.row.to_le_bytes())?;
            payload.extend_from_slice(&u16::from(cell.column).to_le_bytes())?;
            payload.extend_from_slice(&cell.sheet_index.to_le_bytes())?;
        }
        Ok(payload.into_vec())
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
        text.chars().count()
    } else {
        text.encode_utf16().count()
    }
}

/// Append the option byte and characters of an `XLUnicodeStringNoCch`.
fn write_chars(out: &mut Payload, text: &str) -> XlsResult<()> {
    if is_compressible(text) {
        out.push(0u8)?; // fHighByte = 0
        for ch in text.chars() {
            out.push(ch as u8)?;
        }
    } else {
        out.push(HIGH_BYTE)?;
        for unit in text.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes())?;
        }
    }
    Ok(())
}

/// Serialize an `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298).
fn write_segmented_topic(
    out: &mut Payload,
    segments: &[String],
    minimum_segments: usize,
) -> XlsResult<()> {
    if !(minimum_segments..=MAX_TOPIC_SEGMENTS).contains(&segments.len()) {
        return Err(invalid(format!(
            "RTD topic must have {minimum_segments}..={MAX_TOPIC_SEGMENTS} segments"
        )));
    }
    let compressible = segments.iter().all(|segment| is_compressible(segment));
    // Compressed sub-string counts are one byte, wide counts two bytes
    // (MS-XLS 2.5.298); pick the narrowest encoding that fits every segment.
    let wide = !compressible
        || segments
            .iter()
            .any(|segment| biff_char_count(segment) > usize::from(u8::MAX));
    // cch is the size of rgb in encoded units, including each sub-string's
    // one-unit count prefix (MS-XLS 2.5.298). Preflight all counts and lengths
    // before mutating the output so invalid input cannot leave a partial topic.
    let mut segment_counts = [0usize; MAX_TOPIC_SEGMENTS];
    let mut rgb_units = 0usize;
    for (index, segment) in segments.iter().enumerate() {
        let count = biff_char_count(segment);
        if count > usize::from(u16::MAX) {
            return Err(XlsError::InvalidData(
                "RTD topic sub-string exceeds 65535 characters".to_string(),
            ));
        }
        segment_counts[index] = count;
        rgb_units = rgb_units
            .checked_add(
                count
                    .checked_add(1)
                    .ok_or_else(|| invalid("RTD topic encoded-unit count overflows usize"))?,
            )
            .ok_or_else(|| invalid("RTD topic encoded-unit count overflows usize"))?;
    }
    if rgb_units > MAX_STRING_CHARACTERS {
        return Err(invalid("RTD topic exceeds the string resource limit"));
    }
    let encoded_unit_width = if wide { 2 } else { 1 };
    let rgb_bytes = rgb_units
        .checked_mul(encoded_unit_width)
        .ok_or_else(|| invalid("RTD topic byte length overflows usize"))?;
    if rgb_bytes > MAX_LOGICAL_PAYLOAD_BYTES {
        return Err(invalid("RTD topic exceeds the string resource limit"));
    }
    let cch = u32::try_from(rgb_units)
        .map_err(|_| invalid("RTD topic encoded-unit count overflows u32"))?;
    out.extend_from_slice(&cch.to_le_bytes())?;
    out.push(if wide { HIGH_BYTE } else { 0u8 })?;
    for (index, segment) in segments.iter().enumerate() {
        let count = segment_counts[index];
        if wide {
            let count = u16::try_from(count)
                .map_err(|_| invalid("RTD topic sub-string exceeds 65535 characters"))?;
            out.extend_from_slice(&count.to_le_bytes())?;
            for unit in segment.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes())?;
            }
        } else {
            let count = u8::try_from(count)
                .map_err(|_| invalid("RTD topic sub-string exceeds 255 characters"))?;
            out.push(count)?;
            for ch in segment.chars() {
                out.push(ch as u8)?;
            }
        }
    }
    Ok(())
}

/// Parse an `XLUnicodeStringSegmentedRTD`; returns the sub-strings and the
/// number of bytes consumed.
fn parse_segmented_topic(data: &[u8], prefixed: bool) -> XlsResult<(Vec<String>, usize)> {
    if data.len() < 5 {
        return Err(XlsError::InvalidLength {
            expected: 5,
            found: data.len(),
        });
    }
    let cch = usize::try_from(read_u32(data, 0)?)
        .map_err(|_| invalid("RealTimeData stTopic.cch overflows"))?;
    if cch > MAX_STRING_CHARACTERS {
        return Err(invalid(
            "RealTimeData stTopic exceeds the string resource limit",
        ));
    }
    let wide = data[4] & HIGH_BYTE != 0;
    let encoded_unit_width = if wide { 2 } else { 1 };
    let rgb_len = cch
        .checked_mul(encoded_unit_width)
        .ok_or_else(|| invalid("RealTimeData stTopic byte length overflows usize"))?;
    if rgb_len > MAX_LOGICAL_PAYLOAD_BYTES {
        return Err(invalid(
            "RealTimeData stTopic exceeds the string resource limit",
        ));
    }
    let rgb_end = 5usize
        .checked_add(rgb_len)
        .ok_or_else(|| invalid("RealTimeData stTopic byte offset overflows usize"))?;
    data.get(5..rgb_end).ok_or(XlsError::InvalidLength {
        expected: rgb_end,
        found: data.len(),
    })?;
    let mut offset = 5usize;
    let mut segments = Vec::new();
    let minimum_segments = if prefixed {
        MIN_PREFIXED_TOPIC_SEGMENTS
    } else {
        MIN_TOPIC_SEGMENTS
    };
    segments
        .try_reserve(minimum_segments)
        .map_err(|_| XlsError::Allocation("retaining RTD topic segments"))?;
    let mut units_read = 0usize;
    while units_read < cch {
        if segments.len() >= MAX_TOPIC_SEGMENTS {
            return Err(invalid(format!(
                "RealTimeData stTopic exceeds {MAX_TOPIC_SEGMENTS} segments"
            )));
        }
        // Each sub-string starts with a count occupying one encoded unit.
        // The count prefix is included in cch because cch measures rgb.
        let count_len = encoded_unit_width;
        let count_end = offset
            .checked_add(count_len)
            .ok_or_else(|| invalid("RealTimeData stTopic offset overflows usize"))?;
        let count_bytes = data.get(offset..count_end).ok_or(XlsError::InvalidLength {
            expected: count_end,
            found: data.len(),
        })?;
        let segment_chars = if wide {
            usize::from(read_u16(count_bytes, 0)?)
        } else {
            usize::from(count_bytes[0])
        };
        offset = count_end;
        units_read = units_read
            .checked_add(1)
            .ok_or_else(|| invalid("RealTimeData stTopic character count overflows usize"))?;
        let next_units = units_read
            .checked_add(segment_chars)
            .ok_or_else(|| invalid("RealTimeData stTopic character count overflows usize"))?;
        if next_units > cch {
            return Err(invalid("RealTimeData stTopic sub-string overruns cch"));
        }
        let byte_len = segment_chars
            .checked_mul(encoded_unit_width)
            .ok_or_else(|| invalid("RealTimeData stTopic byte length overflows usize"))?;
        let byte_end = offset
            .checked_add(byte_len)
            .ok_or_else(|| invalid("RealTimeData stTopic byte offset overflows usize"))?;
        if byte_end > rgb_end {
            return Err(invalid("RealTimeData stTopic sub-string overruns cch"));
        }
        let bytes = data.get(offset..byte_end).ok_or(XlsError::InvalidLength {
            expected: byte_end,
            found: data.len(),
        })?;
        segments
            .try_reserve(1)
            .map_err(|_| XlsError::Allocation("retaining RTD topic segments"))?;
        segments.push(decode_chars(bytes, wide)?);
        offset = byte_end;
        units_read = next_units;
    }
    if offset != rgb_end {
        return Err(invalid("RealTimeData stTopic does not consume cch bytes"));
    }
    if !(minimum_segments..=MAX_TOPIC_SEGMENTS).contains(&segments.len()) {
        return Err(invalid(format!(
            "RealTimeData stTopic must have {minimum_segments}..={MAX_TOPIC_SEGMENTS} segments"
        )));
    }
    Ok((segments, rgb_end))
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
    let kind = read_u32(data, 0)?;
    let body = data.get(4..).ok_or(XlsError::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
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
            let (text, used, char_count) = parse_rtd_oper_str(body)?;
            let is_long = kind == RTD_OPER_LONG_TEXT;
            if is_long != (char_count >= 256) {
                return Err(invalid(
                    "RTDOper string kind does not match its character count",
                ));
            }
            let total = 4usize
                .checked_add(used)
                .ok_or_else(|| invalid("RTDOper string length overflows usize"))?;
            Ok((XlsRtdValue::Text(text), total))
        },
        RTD_OPER_BOOLEAN => {
            let raw = body.get(..4).ok_or(XlsError::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let value = match read_u32(raw, 0)? {
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
fn parse_rtd_oper_str(data: &[u8]) -> XlsResult<(String, usize, usize)> {
    if data.len() < 5 {
        return Err(XlsError::InvalidLength {
            expected: 5,
            found: data.len(),
        });
    }
    let char_count = usize::try_from(read_u32(data, 0)?)
        .map_err(|_| invalid("RTDOperStr.cchRTDOperStr overflows"))?;
    if char_count > MAX_STRING_CHARACTERS {
        return Err(invalid("RTDOperStr exceeds the string resource limit"));
    }
    let wide = data[4] & HIGH_BYTE != 0;
    let byte_len = if wide {
        char_count
            .checked_mul(2)
            .ok_or_else(|| invalid("RTDOperStr byte length overflows usize"))?
    } else {
        char_count
    };
    let end = 5usize
        .checked_add(byte_len)
        .ok_or_else(|| invalid("RTDOperStr byte offset overflows usize"))?;
    let bytes = data.get(5..end).ok_or(XlsError::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    Ok((decode_chars(bytes, wide)?, end, char_count))
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
        // cch is the size of the complete compressed rgb field, including
        // the one-byte count prefix for every sub-string.
        let cch: usize = segments.iter().map(|segment| 1 + segment.len()).sum();
        let mut out = Vec::new();
        out.extend_from_slice(&(cch as u32).to_le_bytes());
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
    fn rejects_mismatched_rtd_string_kind() {
        for (kind, char_count) in [(RTD_OPER_SHORT_TEXT, 256u32), (RTD_OPER_LONG_TEXT, 5)] {
            let mut payload = frt_header();
            payload.extend_from_slice(&0u32.to_le_bytes());
            payload.extend_from_slice(&segmented_topic(&["A", "B", "C"]));
            payload.extend_from_slice(&kind.to_le_bytes());
            payload.extend_from_slice(&char_count.to_le_bytes());
            payload.push(0);
            payload.extend(std::iter::repeat_n(b'x', char_count as usize));
            assert!(XlsRealTimeData::parse(&payload, None).is_err());
        }
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
        second.extend_from_slice(&segmented_topic(&["", "BOND", "X"]));
        second.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
        second.extend_from_slice(&3u32.to_le_bytes());
        second.push(0u8);
        second.extend_from_slice(b"102");
        let second = XlsRealTimeData::parse(&second, Some(&first.topic)).expect("parse second");
        assert_eq!(second.common_prefix_len, 7);
        assert_eq!(second.topic_segments, vec!["", "BOND", "X"]);
        assert_eq!(second.topic, "PROG.IDBONDX");
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
        topic.extend_from_slice(&6u32.to_le_bytes()); // cch includes 3 two-byte count prefixes
        topic.push(1u8); // fHighByte = 1
        // Three wide substrings: 'A', 'B', and '€'.
        for unit in [u32::from('A') as u16, u32::from('B') as u16, 0x20AC] {
            topic.extend_from_slice(&1u16.to_le_bytes());
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
    fn rejects_overflowing_and_odd_wire_lengths() {
        assert!(read_u16(&[], usize::MAX).is_err());
        assert!(read_u32(&[], usize::MAX).is_err());

        let mut topic = frt_header();
        topic.extend_from_slice(&0u32.to_le_bytes());
        topic.extend_from_slice(&1u32.to_le_bytes());
        topic.push(HIGH_BYTE);
        topic.extend_from_slice(&1u16.to_le_bytes());
        topic.push(0); // Missing the second byte of the UTF-16 code unit.
        assert!(XlsRealTimeData::parse(&topic, None).is_err());

        let mut value = frt_header();
        value.extend_from_slice(&0u32.to_le_bytes());
        value.extend_from_slice(&segmented_topic(&[]));
        value.extend_from_slice(&RTD_OPER_SHORT_TEXT.to_le_bytes());
        value.extend_from_slice(&1u32.to_le_bytes());
        value.push(HIGH_BYTE);
        value.push(0); // Missing the second byte of the UTF-16 code unit.
        assert!(XlsRealTimeData::parse(&value, None).is_err());

        let mut huge_topic = frt_header();
        huge_topic.extend_from_slice(&0u32.to_le_bytes());
        huge_topic.extend_from_slice(&u32::MAX.to_le_bytes());
        huge_topic.push(0);
        assert!(XlsRealTimeData::parse(&huge_topic, None).is_err());
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
    fn compressed_latin1_uses_character_count_not_utf8_byte_count() {
        let value = XlsRealTimeData {
            common_prefix_len: 0,
            topic_segments: vec!["PROG".to_string(), "server".to_string(), "é".to_string()],
            topic: "PROGserveré".to_string(),
            value: XlsRtdValue::Text("é".to_string()),
            cells: Vec::new(),
        };

        let payload = value.to_payload().expect("serialize");
        let parsed = XlsRealTimeData::parse(&payload, None).expect("re-parse");
        assert_eq!(parsed, value);
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
        assert_eq!(read_u32(&payload, kind_offset).unwrap(), RTD_OPER_LONG_TEXT);
        let parsed = XlsRealTimeData::parse(&payload, None).expect("re-parse");
        assert_eq!(parsed, value);
    }

    #[test]
    fn serialize_promotes_long_compressed_segment_to_wide() {
        let value = XlsRealTimeData {
            common_prefix_len: 0,
            topic_segments: vec!["A".to_string(), "B".to_string(), "x".repeat(300)],
            topic: String::new(),
            value: XlsRtdValue::Integer(0),
            cells: Vec::new(),
        };
        // A 300-character compressed segment does not fit the 1-byte count,
        // but the wide encoding holds it.
        let payload = value.to_payload().expect("serialize");
        assert_eq!(payload[FRT_HEADER_LEN + 4 + 4], HIGH_BYTE);
    }

    #[test]
    fn segmented_topic_cch_covers_count_prefixes_and_empty_segments() {
        let value = XlsRealTimeData {
            common_prefix_len: 0,
            topic_segments: vec![
                "PROG".to_string(),
                String::new(),
                "A".to_string(),
                String::new(),
            ],
            topic: "PROGA".to_string(),
            value: XlsRtdValue::Integer(7),
            cells: Vec::new(),
        };

        let payload = value.to_payload().expect("serialize");
        let topic_offset = FRT_HEADER_LEN + 4;
        // rgb is [4, PROG, 0, 1, A, 0], so cch is 9 encoded bytes.
        assert_eq!(read_u32(&payload, topic_offset).unwrap(), 9);
        assert_eq!(
            &payload[topic_offset + 5..topic_offset + 5 + 9],
            b"\x04PROG\x00\x01A\x00"
        );
        let parsed = XlsRealTimeData::parse(&payload, None).expect("parse");
        assert_eq!(parsed, value);
        assert_eq!(parsed.to_payload().unwrap(), payload);
    }
}
