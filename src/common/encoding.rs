//! Character encoding utilities for Microsoft Office file formats.
//!
//! This module provides unified codepage decoding functionality for various Microsoft Office
//! formats that use codepage-based text encoding (OLE, RTF, etc.). It supports common Windows
//! codepages and provides efficient conversion to UTF-8.

use encoding_rs::Encoding;

/// Map Windows codepage identifier to encoding_rs Encoding.
///
/// This function provides a mapping from Windows codepage identifiers to
/// the corresponding `encoding_rs` encodings. It supports the most common
/// codepages used in Microsoft Office documents.
///
/// # Arguments
/// * `codepage` - Windows codepage number (e.g., 936 for Simplified Chinese, 1252 for Western European)
///
/// # Returns
/// Returns `Some(&'static Encoding)` if the codepage is supported, `None` otherwise.
///
/// # Performance
/// This function uses a match expression which compiles to an efficient jump table.
/// The returned encoding references are static, so no allocation occurs.
///
/// # Examples
/// ```
/// use litchi::common::encoding::codepage_to_encoding;
///
/// let encoding = codepage_to_encoding(936).unwrap();
/// assert_eq!(encoding.name(), "GBK");
/// ```
#[inline]
pub fn codepage_to_encoding(codepage: u32) -> Option<&'static Encoding> {
    match codepage {
        // DOS codepages
        437 => Some(encoding_rs::IBM866), // IBM866 (close approximation to CP437)
        850 => Some(encoding_rs::IBM866), // DOS Latin 1 (approximation)

        // Windows codepages (Western scripts)
        874 => Some(encoding_rs::WINDOWS_874),   // Thai
        1250 => Some(encoding_rs::WINDOWS_1250), // Central European
        1251 => Some(encoding_rs::WINDOWS_1251), // Cyrillic
        1252 => Some(encoding_rs::WINDOWS_1252), // Western European (most common, default ANSI)
        1253 => Some(encoding_rs::WINDOWS_1253), // Greek
        1254 => Some(encoding_rs::WINDOWS_1254), // Turkish
        1255 => Some(encoding_rs::WINDOWS_1255), // Hebrew
        1256 => Some(encoding_rs::WINDOWS_1256), // Arabic
        1257 => Some(encoding_rs::WINDOWS_1257), // Baltic
        1258 => Some(encoding_rs::WINDOWS_1258), // Vietnamese

        // East Asian codepages
        932 => Some(encoding_rs::SHIFT_JIS), // Japanese Shift-JIS
        936 => Some(encoding_rs::GBK),       // Simplified Chinese (GB2312/GBK)
        949 => Some(encoding_rs::EUC_KR),    // Korean
        950 => Some(encoding_rs::BIG5),      // Traditional Chinese (Big5)
        1041 => Some(encoding_rs::SHIFT_JIS), // Japanese (alternate)
        1042 => Some(encoding_rs::EUC_KR),   // Korean (alternate)
        2052 => Some(encoding_rs::GBK),      // Chinese PRC
        4100 => Some(encoding_rs::GBK),      // Chinese Singapore
        20936 => Some(encoding_rs::GBK),     // GB2312
        54936 => Some(encoding_rs::GB18030), // Chinese GB18030 (superset of GBK)
        1028 => Some(encoding_rs::BIG5),     // Chinese Taiwan
        3076 => Some(encoding_rs::BIG5),     // Chinese Hong Kong
        5124 => Some(encoding_rs::BIG5),     // Chinese Macau
        20932 => Some(encoding_rs::EUC_JP),  // Japanese EUC-JP

        // ISO 8859 series (Latin and others)
        28591 => Some(encoding_rs::WINDOWS_1252), // ISO-8859-1 approximation
        28592 => Some(encoding_rs::ISO_8859_2),   // Latin 2 (Central European)
        28593 => Some(encoding_rs::ISO_8859_3),   // Latin 3 (South European)
        28594 => Some(encoding_rs::ISO_8859_4),   // Latin 4 (North European)
        28595 => Some(encoding_rs::ISO_8859_5),   // Cyrillic
        28596 => Some(encoding_rs::ISO_8859_6),   // Arabic
        28597 => Some(encoding_rs::ISO_8859_7),   // Greek
        28598 => Some(encoding_rs::ISO_8859_8),   // Hebrew
        28603 => Some(encoding_rs::ISO_8859_13),  // Baltic
        28605 => Some(encoding_rs::ISO_8859_15),  // Latin 9 (Western European with Euro)

        // KOI8 series
        20866 => Some(encoding_rs::KOI8_R), // KOI8-R (Russian)
        21866 => Some(encoding_rs::KOI8_U), // KOI8-U (Ukrainian)

        // Macintosh
        10000 => Some(encoding_rs::MACINTOSH), // Macintosh Roman
        10001 => Some(encoding_rs::SHIFT_JIS), // Mac Japanese
        10004 => Some(encoding_rs::ISO_8859_6), // Mac Arabic (approximation)
        10005 => Some(encoding_rs::WINDOWS_1255), // Mac Hebrew (approximation)
        10007 => Some(encoding_rs::KOI8_R),    // Mac Cyrillic (approximation)

        // Unicode
        1200 => Some(encoding_rs::UTF_16LE), // UTF-16 Little Endian
        1201 => Some(encoding_rs::UTF_16BE), // UTF-16 Big Endian
        65000 => Some(encoding_rs::UTF_8),   // UTF-7 (use UTF-8 as fallback)
        65001 => Some(encoding_rs::UTF_8),   // UTF-8

        // Unsupported codepage
        _ => None,
    }
}

/// Decode UTF-16 LE bytes to a String.
///
/// This function efficiently decodes UTF-16 Little Endian byte sequences
/// into Rust strings, handling null terminators and invalid sequences.
///
/// # Arguments
/// * `bytes` - The byte sequence containing UTF-16LE encoded text
///
/// # Returns
/// Returns a String with invalid sequences replaced by U+FFFD (lossy conversion).
///
/// # Examples
/// ```
/// use litchi::common::encoding::decode_utf16le;
///
/// let bytes = b"H\x00e\x00l\x00l\x00o\x00";
/// let text = decode_utf16le(bytes);
/// assert_eq!(text, "Hello");
/// ```
#[inline]
pub fn decode_utf16le(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return String::new();
    }

    // Ensure we have complete UTF-16 code units (pairs of bytes)
    let byte_len = bytes.len() & !1; // Round down to even number
    let bytes = &bytes[..byte_len];

    // Convert to u16 slice
    let utf16_units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .take_while(|&c| c != 0) // Stop at null terminator
        .collect();

    // Decode UTF-16 to String (lossy - replaces invalid sequences)
    String::from_utf16_lossy(&utf16_units)
}

/// Strip null terminators from the end of a byte slice.
///
/// This is a zero-copy operation that returns a slice view.
#[inline]
pub fn strip_null_terminators(bytes: &[u8]) -> &[u8] {
    // Find the position of the first null terminator
    let end = bytes.iter().position(|&b| b == 0).unwrap_or(bytes.len());
    &bytes[..end]
}

/// Decode bytes using the specified Windows codepage
///
/// This function converts byte sequences encoded with various Windows codepages
/// to UTF-8 strings. It handles null terminators and supports a wide range of
/// legacy codepages commonly used in Office documents.
///
/// # Arguments
///
/// * `bytes` - The byte sequence to decode
/// * `codepage` - Optional Windows codepage identifier (e.g., 1252 for Western European)
///
/// # Returns
///
/// Returns `Some(String)` if the codepage is supported and decoding succeeds,
/// `None` if the codepage is not supported or decoding fails.
///
/// # Examples
///
/// ```
/// use litchi::common::encoding::decode_bytes;
///
/// // Decode Windows-1252 (Western European) text
/// let bytes = b"Hello, World!";
/// let text = decode_bytes(bytes, Some(1252));
/// assert_eq!(text, Some("Hello, World!".to_string()));
///
/// // Unsupported codepage returns None
/// let text = decode_bytes(bytes, Some(99999));
/// assert_eq!(text, None);
/// ```
///
/// # Supported Codepages
///
/// See the [Microsoft codepage documentation](https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers)
/// for a complete list of Windows codepage identifiers.
#[inline]
pub fn decode_bytes(bytes: &[u8], codepage: Option<u32>) -> Option<String> {
    // Remove null terminators efficiently
    let bytes = strip_null_terminators(bytes);

    // Return empty string for empty input
    if bytes.is_empty() {
        return Some(String::new());
    }

    // Determine encoding from codepage
    let encoding = codepage_to_encoding(codepage?)?;

    // Decode using the determined encoding
    // encoding_rs guarantees valid UTF-8 output
    Some(encoding.decode(bytes).0.into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_codepage_to_encoding_common() {
        assert!(codepage_to_encoding(1252).is_some()); // Windows-1252
        assert!(codepage_to_encoding(932).is_some()); // Shift-JIS
        assert!(codepage_to_encoding(936).is_some()); // GBK
        assert!(codepage_to_encoding(950).is_some()); // Big5
        assert!(codepage_to_encoding(65001).is_some()); // UTF-8
    }

    #[test]
    fn test_codepage_to_encoding_chinese() {
        assert_eq!(codepage_to_encoding(936).unwrap().name(), "GBK");
        assert_eq!(codepage_to_encoding(950).unwrap().name(), "Big5");
    }

    #[test]
    fn test_codepage_to_encoding_unsupported() {
        assert!(codepage_to_encoding(99999).is_none());
    }

    #[test]
    fn test_decode_utf16le() {
        let bytes = b"H\x00e\x00l\x00l\x00o\x00";
        let result = decode_utf16le(bytes);
        assert_eq!(result, "Hello");
    }

    #[test]
    fn test_decode_utf16le_with_null() {
        let bytes = b"H\x00e\x00l\x00l\x00o\x00\x00\x00W\x00o\x00r\x00l\x00d\x00";
        let result = decode_utf16le(bytes);
        assert_eq!(result, "Hello");
    }

    #[test]
    fn test_decode_utf16le_empty() {
        let bytes = b"";
        let result = decode_utf16le(bytes);
        assert_eq!(result, "");
    }

    #[test]
    fn test_strip_null_terminators() {
        let bytes = b"Hello\x00World";
        let result = strip_null_terminators(bytes);
        assert_eq!(result, b"Hello");
    }

    #[test]
    fn test_strip_null_terminators_no_null() {
        let bytes = b"Hello";
        let result = strip_null_terminators(bytes);
        assert_eq!(result, b"Hello");
    }
}
