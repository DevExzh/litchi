//! Text extraction utilities for PPT records.
//!
//! This module provides functions to extract text from various PPT record types,
//! including `TextCharsAtom` (UTF-16LE), `TextBytesAtom` (ISO-8859-1), and `CString`.

use crate::package::Result;
/// Parse `TextCharsAtom` record (UTF-16LE text content).
/// Based on POI's `TextCharsAtom.getText()` method.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_text_chars_atom(data: &[u8]) -> Result<String> {
    if data.is_empty() {
        return Ok(String::new());
    }

    // TextCharsAtom contains UTF-16LE encoded text (little-endian)
    // Use the same logic as POI's StringUtil.getFromUnicodeLE
    let text = from_utf16le_lossy(data);

    // POI strips the trailing return character and null terminator if present
    let trimmed = text.trim_end_matches(['\r', '\u{0}']).to_string();

    Ok(trimmed)
}

/// Convert UTF-16LE bytes to String (lossy conversion).
/// This follows POI's StringUtil.getFromUnicodeLE logic.
/// Optimized for performance with minimal allocations.
#[must_use]
pub fn from_utf16le_lossy(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return String::new();
    }

    let code_units = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .take_while(|&code_unit| code_unit != 0)
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&code_units)
}

/// Decode the low bytes of UTF-16 characters stored by `TextBytesAtom`.
pub(crate) fn decode_text_bytes(data: &[u8]) -> String {
    data.iter().map(|&byte| char::from(byte)).collect()
}

/// Parse `TextBytesAtom` record (byte text content).
/// Based on POI's `TextBytesAtom.getText()` method.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_text_bytes_atom(data: &[u8]) -> Result<String> {
    if data.is_empty() {
        return Ok(String::new());
    }

    // Each byte is the low byte of a UTF-16 character whose high byte is zero.
    let text = decode_text_bytes(data);

    // POI strips the trailing return character and null terminator if present
    let trimmed = text.trim_end_matches(['\r', '\u{0}']).to_string();

    Ok(trimmed)
}

/// Parse a `CString` record containing a UTF-16LE string.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_cstring(data: &[u8]) -> Result<String> {
    let text = from_utf16le_lossy(data);

    // Strip a non-conforming trailing return for compatibility with old producers.
    let trimmed = text.trim_end_matches('\r').to_string();

    Ok(trimmed)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
    fn text_chars_atom_decodes_utf16_surrogate_pairs() {
        let text = parse_text_chars_atom(&[0x3D, 0xD8, 0x00, 0xDE, b'A', 0]).unwrap();

        assert_eq!(text, "😀A");
    }

    #[test]
    fn test_text_bytes_atom_parsing() {
        let text_data = b"Hello World";
        let text = parse_text_bytes_atom(text_data).unwrap();
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn text_bytes_atom_maps_bytes_directly_to_unicode() {
        let text = parse_text_bytes_atom(&[0x80, 0x91, 0xE9]).unwrap();

        assert_eq!(text, "\u{80}\u{91}é");
    }

    #[test]
    fn cstring_decodes_all_utf16_content_without_policy_filtering() {
        fn utf16(text: &str) -> Vec<u8> {
            text.encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>()
        }

        // Record parsing preserves internal tag names; callers decide whether to expose them.
        let tag_text = parse_cstring(&utf16("___PPT10")).unwrap();
        assert_eq!(tag_text, "___PPT10");

        let design_text = parse_cstring(&utf16("Default Design")).unwrap();
        assert_eq!(design_text, "Default Design");

        // Should preserve the complete UTF-16 string, including surrogate pairs.
        let unicode_text = parse_cstring(&utf16("普通の😀")).unwrap();
        assert_eq!(unicode_text, "普通の😀");
    }
}
