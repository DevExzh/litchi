//! Text extraction utilities for PPT records.
//!
//! This module provides functions to extract text from various PPT record types,
//! including TextCharsAtom (UTF-16LE), TextBytesAtom (ISO-8859-1), and CString.

use crate::ole::ppt::package::Result;
use zerocopy::{
    FromBytes,
    byteorder::{LittleEndian, U16},
};

/// Actions for processing UTF-16LE text characters.
#[derive(Debug, Clone, Copy)]
pub(crate) enum TextCharAction {
    /// Add the character to the result
    Add(char),
    /// Stop processing (null terminator found)
    Stop,
    /// Skip this character (invalid)
    Skip,
}

impl TextCharAction {
    /// Process a UTF-16LE code unit and determine the appropriate action.
    pub(crate) fn process_utf16_char(code_unit: u16) -> Self {
        match code_unit {
            // Null terminator - stop processing
            0 => TextCharAction::Stop,
            // ASCII range (0x00-0x7F) - add as character
            0x01..=0x7F => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    TextCharAction::Add(ch)
                } else {
                    TextCharAction::Skip
                }
            },
            // Unicode range (0x80 and above) - try to decode as Unicode
            0x80.. => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    TextCharAction::Add(ch)
                } else {
                    TextCharAction::Skip
                }
            },
        }
    }
}

/// Parse TextCharsAtom record (UTF-16LE text content).
/// Based on POI's TextCharsAtom.getText() method.
pub fn parse_text_chars_atom(data: &[u8]) -> Result<String> {
    if data.is_empty() {
        return Ok(String::new());
    }

    // TextCharsAtom contains UTF-16LE encoded text (little-endian)
    // Use the same logic as POI's StringUtil.getFromUnicodeLE
    let text = from_utf16le_lossy(data);

    // POI strips the trailing return character and null terminator if present
    let text = text
        .trim_end_matches('\r')
        .trim_end_matches('\u{0}')
        .to_string();

    Ok(text)
}

/// Convert UTF-16LE bytes to String (lossy conversion).
/// This follows POI's StringUtil.getFromUnicodeLE logic.
/// Optimized for performance with minimal allocations.
pub fn from_utf16le_lossy(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return String::new();
    }

    // Pre-calculate capacity for the result string
    let estimated_chars = bytes.len() / 2;
    let mut result = String::with_capacity(estimated_chars);

    // Process in chunks of 2 bytes for better performance
    let mut i = 0;
    while i + 1 < bytes.len() {
        let code_unit = U16::<LittleEndian>::read_from_bytes(&bytes[i..i + 2])
            .map(|v| v.get())
            .unwrap_or(0);
        i += 2;

        // Use match expression for cleaner character processing
        match TextCharAction::process_utf16_char(code_unit) {
            TextCharAction::Add(ch) => result.push(ch),
            TextCharAction::Stop => break,
            TextCharAction::Skip => continue,
        }
    }

    // Shrink to fit if we over-allocated
    result.shrink_to_fit();
    result
}

/// Parse TextBytesAtom record (byte text content).
/// Based on POI's TextBytesAtom.getText() method.
pub fn parse_text_bytes_atom(data: &[u8]) -> Result<String> {
    if data.is_empty() {
        return Ok(String::new());
    }

    // TextBytesAtom contains text in "compressed unicode" format (ISO-8859-1)
    // This follows POI's StringUtil.getFromCompressedUnicode logic
    let text = data.iter().map(|&b| b as char).collect::<String>();

    // POI strips the trailing return character and null terminator if present
    let text = text
        .trim_end_matches('\r')
        .trim_end_matches('\u{0}')
        .to_string();

    Ok(text)
}

/// Parse CString record (null-terminated string).
pub fn parse_cstring(data: &[u8]) -> Result<String> {
    // CString contains null-terminated ASCII text
    let null_pos = data.iter().position(|&b| b == 0).unwrap_or(data.len());
    let text = String::from_utf8_lossy(&data[..null_pos]).to_string();

    // POI strips the trailing return character if present
    let text = text.trim_end_matches('\r').to_string();

    // Filter out known garbage strings (from POI's QuickButCruddyTextExtractor)
    if text == "___PPT10" || text == "Default Design" || text.is_empty() {
        return Ok(String::new());
    }

    // Filter out non-printable/binary data - if more than 20% of characters are non-printable, skip it
    let printable_count = text
        .chars()
        .filter(|c| c.is_alphanumeric() || c.is_whitespace() || c.is_ascii_punctuation())
        .count();
    let total_count = text.chars().count();
    if total_count > 0 && (printable_count as f32 / total_count as f32) < 0.8 {
        return Ok(String::new());
    }

    Ok(text)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_chars_atom_parsing() {
        // Test TextCharsAtom parsing with UTF-16LE
        let text_data = vec![
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
            0x00, 0x00, // null terminator
        ];

        let text = parse_text_chars_atom(&text_data).unwrap();
        assert_eq!(text, "Hello");
    }

    #[test]
    fn test_text_bytes_atom_parsing() {
        let text_data = b"Hello World";
        let text = parse_text_bytes_atom(text_data).unwrap();
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn test_cstring_filtering() {
        // Should filter out ___PPT10
        let text = parse_cstring(b"___PPT10\0").unwrap();
        assert_eq!(text, "");

        // Should filter out Default Design
        let text = parse_cstring(b"Default Design\0").unwrap();
        assert_eq!(text, "");

        // Should keep normal text
        let text = parse_cstring(b"Normal Text\0").unwrap();
        assert_eq!(text, "Normal Text");
    }
}
