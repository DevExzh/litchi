use crate::ole::OleError;
use bytes::Bytes;
use zerocopy::{FromBytes, LE, U16, U32, I16, I32, F64};

/// Read a little-endian u16 from a byte slice at the given offset.
///
#[inline]
pub fn read_u16_le(data: &[u8], offset: usize) -> Result<u16, OleError> {
    if offset + 2 > data.len() {
        return Err(OleError::InvalidData("Not enough data for u16".to_string()));
    }
    U16::<LE>::read_from_bytes(&data[offset..offset + 2])
        .map(|v| v.get())
        .or_else(|_| Err(OleError::InvalidData("Failed to read u16".to_string())))
}

/// Read a little-endian u16 from a byte slice starting at offset 0.
#[inline]
pub fn read_u16_le_at(data: &[u8], offset: usize) -> Result<u16, OleError> {
    read_u16_le(data, offset)
}

/// Read a little-endian i16 from a byte slice at the given offset.
#[inline]
pub fn read_i16_le(data: &[u8], offset: usize) -> Result<i16, OleError> {
    if offset + 2 > data.len() {
        return Err(OleError::InvalidData("Not enough data for i16".to_string()));
    }
    I16::<LE>::read_from_bytes(&data[offset..offset + 2])
        .map(|v| v.get())
        .or_else(|_| Err(OleError::InvalidData("Failed to read i16".to_string())))
}

/// Read a little-endian u32 from a byte slice at the given offset.
#[inline]
pub fn read_u32_le(data: &[u8], offset: usize) -> Result<u32, OleError> {
    if offset + 4 > data.len() {
        return Err(OleError::InvalidData("Not enough data for u32".to_string()));
    }
    U32::<LE>::read_from_bytes(&data[offset..offset + 4])
        .map(|v| v.get())
        .or_else(|_| Err(OleError::InvalidData("Failed to read u32".to_string())))
}

/// Read a little-endian u32 from a byte slice starting at offset 0.
#[inline]
pub fn read_u32_le_at(data: &[u8], offset: usize) -> Result<u32, OleError> {
    read_u32_le(data, offset)
}

/// Read a little-endian i32 from a byte slice at the given offset.
#[inline]
pub fn read_i32_le(data: &[u8], offset: usize) -> Result<i32, OleError> {
    if offset + 4 > data.len() {
        return Err(OleError::InvalidData("Not enough data for i32".to_string()));
    }
    I32::<LE>::read_from_bytes(&data[offset..offset + 4])
        .map(|v| v.get())
        .or_else(|_| Err(OleError::InvalidData("Failed to read i32".to_string())))
}

/// Read a little-endian f64 from a byte slice at the given offset.
#[inline]
pub fn read_f64_le(data: &[u8], offset: usize) -> Result<f64, OleError> {
    if offset + 8 > data.len() {
        return Err(OleError::InvalidData("Not enough data for f64".to_string()));
    }
    F64::<LE>::read_from_bytes(&data[offset..offset + 8])
        .map(|v| v.get())
        .or_else(|_| Err(OleError::InvalidData("Failed to read f64".to_string())))
}

/// Read a little-endian f64 from a byte slice starting at offset 0.
#[inline]
pub fn read_f64_le_at(data: &[u8], offset: usize) -> Result<f64, OleError> {
    read_f64_le(data, offset)
}

/// Parse UTF-16LE string from binary data with null terminator handling.
///
/// Based on Apache POI's StringUtil.getFromUnicodeLE.
/// Optimized for performance with minimal allocations.
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
pub fn parse_windows1252_string(data: &[u8]) -> String {
    data.iter()
        .take_while(|&&b| b != 0)
        .map(|&b| windows_1252_to_char(b))
        .collect()
}

/// Parse Windows-1252 string with specified length.
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

/// Property List with Character Positions (PLCF) parser.
///
/// Based on Apache POI's PlexOfCps. PLCF is a common structure in Office files
/// that maps character positions to properties or data.
pub struct PlcfParser {
    /// Character positions (CP array)
    positions: Vec<u32>,
    /// Property data buffer containing all property elements
    properties_data: Bytes,
    /// Offsets into properties_data for each property element
    properties_offsets: Vec<(usize, usize)>, // (offset, length) pairs
}

impl PlcfParser {
    /// Parse a PLCF structure from binary data.
    ///
    /// # Arguments
    ///
    /// * `data` - The binary data containing the PLCF
    /// * `element_size` - Size in bytes of each property element
    ///
    /// # Format
    ///
    /// PLCF format:
    /// - n+1 character positions (4 bytes each)
    /// - n property elements (element_size bytes each)
    pub fn parse(data: &[u8], element_size: usize) -> Option<Self> {
        if data.len() < 4 {
            return None;
        }

        // Calculate number of elements
        // Formula: (data_length) / (4 + element_size) = n
        // So: n+1 CPs (4 bytes each) + n elements (element_size each)
        let n = if element_size > 0 {
            (data.len() - 4) / (4 + element_size)
        } else {
            return None;
        };

        if n == 0 {
            return Some(Self {
                positions: Vec::new(),
                properties_data: Bytes::new(),
                properties_offsets: Vec::new(),
            });
        }

        // Read character positions
        let mut positions = Vec::with_capacity(n + 1);
        for i in 0..=n {
            let offset = i * 4;
            if let Ok(cp) = read_u32_le(data, offset) {
                positions.push(cp);
            } else {
                return None;
            }
        }

        // Read property data into a single Bytes buffer
        let props_start = (n + 1) * 4;
        let props_end = props_start + (n * element_size);
        if props_end > data.len() {
            return None;
        }

        let properties_data = Bytes::copy_from_slice(&data[props_start..props_end]);
        let mut properties_offsets = Vec::with_capacity(n);

        for i in 0..n {
            let offset = i * element_size;
            properties_offsets.push((offset, element_size));
        }

        Some(Self {
            positions,
            properties_data,
            properties_offsets,
        })
    }

    /// Get the number of elements in the PLCF.
    #[inline]
    pub fn count(&self) -> usize {
        self.properties_offsets.len()
    }

    /// Get character position at index.
    #[inline]
    pub fn position(&self, index: usize) -> Option<u32> {
        self.positions.get(index).copied()
    }

    /// Get property data at index.
    #[inline]
    pub fn property(&self, index: usize) -> Option<&[u8]> {
        self.properties_offsets.get(index).map(|(offset, len)| {
            &self.properties_data[*offset..*offset + *len]
        })
    }

    /// Get character range for element at index.
    ///
    /// Returns (start_cp, end_cp) tuple.
    pub fn range(&self, index: usize) -> Option<(u32, u32)> {
        if index >= self.properties_offsets.len() {
            return None;
        }
        Some((self.positions[index], self.positions[index + 1]))
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

    #[test]
    fn test_plcf_parser() {
        // Create a simple PLCF with 2 elements, element_size = 2
        // CPs: 0, 10, 20
        // Props: [1, 2], [3, 4]
        let data = vec![
            0x00, 0x00, 0x00, 0x00, // CP 0
            0x0A, 0x00, 0x00, 0x00, // CP 10
            0x14, 0x00, 0x00, 0x00, // CP 20
            0x01, 0x02, // Property 1
            0x03, 0x04, // Property 2
        ];

        let plcf = PlcfParser::parse(&data, 2).unwrap();
        assert_eq!(plcf.count(), 2);
        assert_eq!(plcf.position(0), Some(0));
        assert_eq!(plcf.position(1), Some(10));
        assert_eq!(plcf.position(2), Some(20));
        assert_eq!(plcf.range(0), Some((0, 10)));
        assert_eq!(plcf.range(1), Some((10, 20)));
    }
}

