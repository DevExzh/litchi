//! Bounded BIFF8 wire codecs for RealTimeData records.

use super::model::Value;
use crate::error::{Error, Result};

/// Record type of the `RealTimeData` record (MS-XLS 2.4.214).
pub(crate) const REAL_TIME_DATA_RECORD_TYPE: u16 = 0x0813;
/// Record type of the `ContinueFrt` record (MS-XLS 2.4.60) that continues a
/// `RealTimeData` payload.
pub(crate) const CONTINUE_FRT_RECORD_TYPE: u16 = 0x0812;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
pub(super) const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of an `RTDEItem` structure (MS-XLS 2.5.223).
pub(super) const RTD_E_ITEM_LEN: usize = 6;

// `RTDOper.grbit` variant-kind codes (MS-XLS 2.5.224).
pub(super) const RTD_OPER_NUMBER: u32 = 0x0000_0001;
pub(super) const RTD_OPER_SHORT_TEXT: u32 = 0x0000_0002;
pub(super) const RTD_OPER_BOOLEAN: u32 = 0x0000_0004;
pub(super) const RTD_OPER_ERROR: u32 = 0x0000_0010;
pub(super) const RTD_OPER_INTEGER: u32 = 0x0000_0800;
pub(super) const RTD_OPER_LONG_TEXT: u32 = 0x0000_1000;

/// `fHighByte` bit of a BIFF8 string option byte.
pub(super) const HIGH_BYTE: u8 = 0x01;
/// Maximum logical `RealTimeData` payload after `ContinueFrt` reassembly.
pub(super) const MAX_LOGICAL_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
/// Maximum retained topic/value character or encoded-unit count imposed before
/// allocation.
pub(super) const MAX_STRING_CHARACTERS: usize = MAX_LOGICAL_PAYLOAD_BYTES;
/// Minimum and maximum number of segmented topic substrings allowed by
/// `XLUnicodeStringSegmentedRTD` (MS-XLS 2.5.298).
pub(super) const MIN_TOPIC_SEGMENTS: usize = 3;
pub(super) const MIN_PREFIXED_TOPIC_SEGMENTS: usize = 2;
pub(super) const MAX_TOPIC_SEGMENTS: usize = 39;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: REAL_TIME_DATA_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_bytes<const N: usize>(data: &[u8], offset: usize) -> Result<&[u8]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| invalid("RealTimeData field offset overflows usize"))?;
    data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = read_bytes::<2>(data, offset)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = read_bytes::<4>(data, offset)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// Decode `char_count` characters from `bytes` in compressed (1 byte/char)
/// or uncompressed UTF-16LE (2 bytes/char) form.
fn decode_chars(bytes: &[u8], wide: bool) -> Result<String> {
    if wide {
        if !bytes.len().is_multiple_of(2) {
            return Err(invalid("RTD wide string has an odd byte length"));
        }
        let mut value = String::new();
        value
            .try_reserve(bytes.len())
            .map_err(|_| Error::Allocation("decoding RTD UTF-16 text"))?;
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
            .map_err(|_| Error::Allocation("decoding RTD compressed text"))?;
        value.extend(bytes.iter().map(|&byte| char::from(byte)));
        Ok(value)
    }
}

pub(super) fn join_segments(segments: &[String]) -> Result<String> {
    let byte_len = segments.iter().try_fold(0usize, |total, segment| {
        total
            .checked_add(segment.len())
            .ok_or_else(|| invalid("RTD topic byte length overflows usize"))
    })?;
    let mut value = String::new();
    value
        .try_reserve(byte_len)
        .map_err(|_| Error::Allocation("reassembling RTD topic text"))?;
    for segment in segments {
        value.push_str(segment);
    }
    Ok(value)
}

/// Fallible bounded output buffer for a serialized logical RTD record.
pub(super) struct Payload {
    bytes: Vec<u8>,
}

impl Payload {
    pub(super) fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    pub(super) fn push(&mut self, byte: u8) -> Result<()> {
        if self.bytes.len() >= MAX_LOGICAL_PAYLOAD_BYTES {
            return Err(invalid(format!(
                "serialized RealTimeData payload exceeds {MAX_LOGICAL_PAYLOAD_BYTES} bytes"
            )));
        }
        if self.bytes.len() == self.bytes.capacity() {
            self.bytes
                .try_reserve(1)
                .map_err(|_| Error::Allocation("serializing RealTimeData payload"))?;
        }
        self.bytes.push(byte);
        Ok(())
    }

    pub(super) fn extend_from_slice(&mut self, bytes: &[u8]) -> Result<()> {
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
            .map_err(|_| Error::Allocation("serializing RealTimeData payload"))?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    pub(super) fn into_vec(self) -> Vec<u8> {
        self.bytes
    }
}
/// Whether a string can be stored compressed (every character in U+0000..=U+00FF).
pub(super) fn is_compressible(text: &str) -> bool {
    text.chars().all(|ch| u32::from(ch) <= 0xFF)
}

/// Character count as BIFF stores it: byte count when the string is
/// compressible, UTF-16 code units otherwise.
pub(super) fn biff_char_count(text: &str) -> usize {
    if is_compressible(text) {
        text.chars().count()
    } else {
        text.encode_utf16().count()
    }
}

/// Append the option byte and characters of an `XLUnicodeStringNoCch`.
pub(super) fn write_chars(out: &mut Payload, text: &str) -> Result<()> {
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
pub(super) fn write_segmented_topic(
    out: &mut Payload,
    segments: &[String],
    minimum_segments: usize,
) -> Result<()> {
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
            return Err(Error::InvalidData(
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
pub(super) fn parse_segmented_topic(data: &[u8], prefixed: bool) -> Result<(Vec<String>, usize)> {
    if data.len() < 5 {
        return Err(Error::InvalidLength {
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
    data.get(5..rgb_end).ok_or(Error::InvalidLength {
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
        .map_err(|_| Error::Allocation("retaining RTD topic segments"))?;
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
        let count_bytes = data.get(offset..count_end).ok_or(Error::InvalidLength {
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
        let bytes = data.get(offset..byte_end).ok_or(Error::InvalidLength {
            expected: byte_end,
            found: data.len(),
        })?;
        segments
            .try_reserve(1)
            .map_err(|_| Error::Allocation("retaining RTD topic segments"))?;
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
pub(super) fn parse_rtd_oper(data: &[u8]) -> Result<(Value, usize)> {
    if data.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let kind = read_u32(data, 0)?;
    let body = data.get(4..).ok_or(Error::InvalidLength {
        expected: 4,
        found: data.len(),
    })?;
    match kind {
        RTD_OPER_NUMBER => {
            let bytes = body.get(..8).ok_or(Error::InvalidLength {
                expected: 12,
                found: data.len(),
            })?;
            let bytes = <[u8; 8]>::try_from(bytes).map_err(|_| Error::InvalidLength {
                expected: 12,
                found: data.len(),
            })?;
            Ok((Value::Number(f64::from_le_bytes(bytes)), 12))
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
            Ok((Value::Text(text), total))
        },
        RTD_OPER_BOOLEAN => {
            let raw = body.get(..4).ok_or(Error::InvalidLength {
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
            Ok((Value::Boolean(value), 8))
        },
        RTD_OPER_ERROR | RTD_OPER_INTEGER => {
            let raw = body.get(..4).ok_or(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let raw = <[u8; 4]>::try_from(raw).map_err(|_| Error::InvalidLength {
                expected: 8,
                found: data.len(),
            })?;
            let value = i32::from_le_bytes(raw);
            let variant = if kind == RTD_OPER_ERROR {
                Value::Error(value)
            } else {
                Value::Integer(value)
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
fn parse_rtd_oper_str(data: &[u8]) -> Result<(String, usize, usize)> {
    if data.len() < 5 {
        return Err(Error::InvalidLength {
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
    let bytes = data.get(5..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    Ok((decode_chars(bytes, wide)?, end, char_count))
}
