//! Binary data parsing utilities shared across formats.
//!
//! This module provides common functions for reading binary data in little-endian format
//! and parsing strings (UTF-16LE, Windows-1252) used in Office file formats.

use zerocopy::{F64, FromBytes, I16, I32, LE, U16, U32};

/// Binary parsing error type
#[derive(Debug, Clone)]
pub enum BinaryError {
    /// Not enough data to read the requested type
    InsufficientData { expected: usize, available: usize },
    /// Failed to parse the data
    ParseError(String),
}

impl std::fmt::Display for BinaryError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            BinaryError::InsufficientData {
                expected,
                available,
            } => {
                write!(
                    f,
                    "Insufficient data: expected {}, got {}",
                    expected, available
                )
            },
            BinaryError::ParseError(msg) => write!(f, "Parse error: {}", msg),
        }
    }
}

impl std::error::Error for BinaryError {}

/// Result type for binary operations
pub type BinaryResult<T> = Result<T, BinaryError>;

/// Read a little-endian u16 from a byte slice at the given offset.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::read_u16_le;
/// let data = [0x34, 0x12, 0x78, 0x56];
/// assert_eq!(read_u16_le(&data, 0).unwrap(), 0x1234);
/// assert_eq!(read_u16_le(&data, 2).unwrap(), 0x5678);
/// ```
#[inline]
pub fn read_u16_le(data: &[u8], offset: usize) -> BinaryResult<u16> {
    if offset + 2 > data.len() {
        return Err(BinaryError::InsufficientData {
            expected: offset + 2,
            available: data.len(),
        });
    }
    U16::<LE>::read_from_bytes(&data[offset..offset + 2])
        .map(|v| v.get())
        .map_err(|_| BinaryError::ParseError("Failed to read u16".to_string()))
}

/// Read a little-endian u16 from a byte slice starting at offset 0.
#[inline]
pub fn read_u16_le_at(data: &[u8], offset: usize) -> BinaryResult<u16> {
    read_u16_le(data, offset)
}

/// Read a little-endian i16 from a byte slice at the given offset.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::read_i16_le;
/// let data = [0xFF, 0xFF];
/// assert_eq!(read_i16_le(&data, 0).unwrap(), -1i16);
/// ```
#[inline]
pub fn read_i16_le(data: &[u8], offset: usize) -> BinaryResult<i16> {
    if offset + 2 > data.len() {
        return Err(BinaryError::InsufficientData {
            expected: offset + 2,
            available: data.len(),
        });
    }
    I16::<LE>::read_from_bytes(&data[offset..offset + 2])
        .map(|v| v.get())
        .map_err(|_| BinaryError::ParseError("Failed to read i16".to_string()))
}

/// Read a little-endian u32 from a byte slice at the given offset.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::read_u32_le;
/// let data = [0x78, 0x56, 0x34, 0x12];
/// assert_eq!(read_u32_le(&data, 0).unwrap(), 0x12345678);
/// ```
#[inline]
pub fn read_u32_le(data: &[u8], offset: usize) -> BinaryResult<u32> {
    if offset + 4 > data.len() {
        return Err(BinaryError::InsufficientData {
            expected: offset + 4,
            available: data.len(),
        });
    }
    U32::<LE>::read_from_bytes(&data[offset..offset + 4])
        .map(|v| v.get())
        .map_err(|_| BinaryError::ParseError("Failed to read u32".to_string()))
}

/// Read a little-endian u32 from a byte slice starting at offset 0.
#[inline]
pub fn read_u32_le_at(data: &[u8], offset: usize) -> BinaryResult<u32> {
    read_u32_le(data, offset)
}

/// Read a little-endian i32 from a byte slice at the given offset.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::read_i32_le;
/// let data = [0xFF, 0xFF, 0xFF, 0xFF];
/// assert_eq!(read_i32_le(&data, 0).unwrap(), -1i32);
/// ```
#[inline]
pub fn read_i32_le(data: &[u8], offset: usize) -> BinaryResult<i32> {
    if offset + 4 > data.len() {
        return Err(BinaryError::InsufficientData {
            expected: offset + 4,
            available: data.len(),
        });
    }
    I32::<LE>::read_from_bytes(&data[offset..offset + 4])
        .map(|v| v.get())
        .map_err(|_| BinaryError::ParseError("Failed to read i32".to_string()))
}

/// Read a little-endian f64 from a byte slice at the given offset.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::read_f64_le;
/// let data = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F];
/// assert!((read_f64_le(&data, 0).unwrap() - 1.0).abs() < f64::EPSILON);
/// ```
#[inline]
pub fn read_f64_le(data: &[u8], offset: usize) -> BinaryResult<f64> {
    if offset + 8 > data.len() {
        return Err(BinaryError::InsufficientData {
            expected: offset + 8,
            available: data.len(),
        });
    }
    F64::<LE>::read_from_bytes(&data[offset..offset + 8])
        .map(|v| v.get())
        .map_err(|_| BinaryError::ParseError("Failed to read f64".to_string()))
}

/// Read a little-endian f64 from a byte slice starting at offset 0.
#[inline]
pub fn read_f64_le_at(data: &[u8], offset: usize) -> BinaryResult<f64> {
    read_f64_le(data, offset)
}

/// Parse UTF-16LE string from binary data with null terminator handling.
///
/// Based on Apache POI's StringUtil.getFromUnicodeLE.
/// Optimized for performance with minimal allocations.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::parse_utf16le_string;
/// let data = vec![0x48, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00, 0x00, 0x00];
/// assert_eq!(parse_utf16le_string(&data), "Hello");
/// ```
pub fn parse_utf16le_string(data: &[u8]) -> String {
    if data.is_empty() || data.len() < 2 {
        return String::new();
    }

    let estimated_chars = data.len() / 2;
    let mut result = String::with_capacity(estimated_chars);

    let mut i = 0;
    while i + 1 < data.len() {
        let code_unit = U16::<LE>::read_from_bytes(&data[i..i + 2])
            .map(|v| v.get())
            .unwrap_or(0);
        i += 2;

        // Stop at null terminator
        if code_unit == 0 {
            break;
        }

        // Convert to char and add to result
        if let Some(ch) = char::from_u32(code_unit as u32) {
            result.push(ch);
        }
    }

    result.shrink_to_fit();
    result
}

/// Parse UTF-16LE string with specified length (in characters, not bytes).
///
/// Based on Apache POI's StringUtil.getFromUnicodeLE(byte[], int, int).
///
/// # Examples
///
/// ```
/// use litchi::common::binary::parse_utf16le_string_len;
/// let data = vec![0x48, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00];
/// assert_eq!(parse_utf16le_string_len(&data, 0, 5), "Hello");
/// assert_eq!(parse_utf16le_string_len(&data, 0, 3), "Hel");
/// ```
pub fn parse_utf16le_string_len(data: &[u8], offset: usize, char_count: usize) -> String {
    let byte_count = char_count * 2;
    if offset + byte_count > data.len() {
        return String::new();
    }

    let mut result = String::with_capacity(char_count);
    let mut pos = offset;
    let end = offset + byte_count;

    while pos + 1 < end {
        let code_unit = U16::<LE>::read_from_bytes(&data[pos..pos + 2])
            .map(|v| v.get())
            .unwrap_or(0);
        pos += 2;

        if let Some(ch) = char::from_u32(code_unit as u32) {
            result.push(ch);
        }
    }

    result
}

/// Parse Windows-1252 (compressed Unicode) string from binary data.
///
/// Based on Apache POI's StringUtil.getFromCompressedUnicode.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::parse_windows1252_string;
/// let data = b"Hello\x93World\x94\0";
/// let result = parse_windows1252_string(data);
/// assert!(result.starts_with("Hello"));
/// ```
pub fn parse_windows1252_string(data: &[u8]) -> String {
    data.iter()
        .take_while(|&&b| b != 0)
        .map(|&b| windows_1252_to_char(b))
        .collect()
}

/// Parse Windows-1252 string with specified length.
///
/// # Examples
///
/// ```
/// use litchi::common::binary::parse_windows1252_string_len;
/// let data = b"Hello World";
/// assert_eq!(parse_windows1252_string_len(data, 0, 5), "Hello");
/// ```
pub fn parse_windows1252_string_len(data: &[u8], offset: usize, length: usize) -> String {
    if offset + length > data.len() {
        return String::new();
    }

    data[offset..offset + length]
        .iter()
        .map(|&b| windows_1252_to_char(b))
        .collect()
}

/// Convert a Windows-1252 byte to a Unicode character.
///
/// Windows-1252 is mostly compatible with ISO-8859-1, but has additional
/// printable characters in the 0x80-0x9F range.
#[inline]
fn windows_1252_to_char(byte: u8) -> char {
    match byte {
        0x80 => '€',
        0x82 => '‚',
        0x83 => 'ƒ',
        0x84 => '„',
        0x85 => '…',
        0x86 => '†',
        0x87 => '‡',
        0x88 => 'ˆ',
        0x89 => '‰',
        0x8A => 'Š',
        0x8B => '‹',
        0x8C => 'Œ',
        0x8E => 'Ž',
        0x91 => '\u{2018}',
        0x92 => '\u{2019}',
        0x93 => '"',
        0x94 => '"',
        0x95 => '•',
        0x96 => '–',
        0x97 => '—',
        0x98 => '˜',
        0x99 => '™',
        0x9A => 'š',
        0x9B => '›',
        0x9C => 'œ',
        0x9E => 'ž',
        0x9F => 'Ÿ',
        _ => byte as char,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_u16_le() {
        let data = [0x34, 0x12, 0x78, 0x56];
        assert!(read_u16_le(&data, 0).is_ok_and(|v| v == 0x1234));
        assert!(read_u16_le(&data, 2).is_ok_and(|v| v == 0x5678));
        assert!(read_u16_le(&data, 3).is_err());
    }

    #[test]
    fn test_read_u32_le() {
        let data = [0x78, 0x56, 0x34, 0x12];
        assert!(read_u32_le(&data, 0).is_ok_and(|v| v == 0x12345678));
        assert!(read_u32_le(&data, 1).is_err());
    }

    #[test]
    fn test_parse_utf16le() {
        let data = vec![
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
            0x00, 0x00, // null terminator
        ];
        assert_eq!(parse_utf16le_string(&data), "Hello");
    }

    #[test]
    fn test_parse_windows1252() {
        let data = b"Hello\x93World\x94\0";
        let result = parse_windows1252_string(data);
        assert!(result.starts_with("Hello"));
        assert!(result.contains('"'));
    }
}
