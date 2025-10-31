//! Character encoding utilities for Microsoft Office file formats.
//!
//! This module provides unified codepage decoding functionality for various Microsoft Office
//! formats that use codepage-based text encoding (OLE, RTF, etc.). It supports common Windows
//! codepages and provides efficient conversion to UTF-8.

use crate::common::{Error, Result};
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

/// Decode hex-encoded string to bytes with SIMD acceleration.
///
/// This function efficiently decodes hex-encoded strings into byte vectors using
/// SIMD instructions when available (AVX2/SSE on x86_64, NEON on aarch64).
/// It automatically strips whitespace and validates the input.
///
/// # Arguments
///
/// * `hex_str` - The hex-encoded string to decode (e.g., "48656C6C6F" for "Hello")
///
/// # Returns
///
/// Returns a `Vec<u8>` with the decoded bytes.
///
/// # Errors
///
/// Returns an error if:
/// - The hex string has an odd number of characters (after removing whitespace)
/// - The string contains invalid hex characters
///
/// # Examples
///
/// ```
/// use litchi::common::encoding::decode_hex_data;
///
/// let hex = "48656C6C6F"; // "Hello"
/// let decoded = decode_hex_data(hex).unwrap();
/// assert_eq!(decoded, b"Hello");
///
/// // Whitespace is automatically stripped
/// let hex_with_spaces = "48 65 6C 6C 6F";
/// let decoded = decode_hex_data(hex_with_spaces).unwrap();
/// assert_eq!(decoded, b"Hello");
/// ```
///
/// # Performance
///
/// This function uses SIMD instructions when available:
/// - **x86_64**: AVX2 (32 bytes/iteration) or SSE4.1 (16 bytes/iteration)
/// - **aarch64**: NEON (16 bytes/iteration)
/// - **Fallback**: Optimized scalar implementation
///
/// Typical performance is 2-4x faster than naive hex decoding on modern CPUs.
#[inline]
pub fn decode_hex_data(hex_str: &str) -> Result<Vec<u8>> {
    // Remove whitespace efficiently
    let hex_clean: String = hex_str.chars().filter(|c| !c.is_whitespace()).collect();

    // Validate even length
    if !hex_clean.len().is_multiple_of(2) {
        return Err(Error::ParseError(
            "Hex data must have even number of characters".to_string(),
        ));
    }

    if hex_clean.is_empty() {
        return Ok(Vec::new());
    }

    // Dispatch to SIMD implementation if available
    #[cfg(all(target_arch = "x86_64", target_feature = "avx2"))]
    {
        decode_hex_avx2(hex_clean.as_bytes())
    }

    #[cfg(all(
        target_arch = "x86_64",
        not(target_feature = "avx2"),
        target_feature = "sse4.1"
    ))]
    {
        decode_hex_sse41(hex_clean.as_bytes())
    }

    #[cfg(target_arch = "aarch64")]
    {
        // SAFETY: aarch64 always has NEON support
        unsafe { decode_hex_neon(hex_clean.as_bytes()) }
    }

    #[cfg(not(any(
        all(
            target_arch = "x86_64",
            any(target_feature = "avx2", target_feature = "sse4.1")
        ),
        target_arch = "aarch64"
    )))]
    {
        decode_hex_scalar(hex_clean.as_bytes())
    }
}

/// Scalar fallback for hex decoding (optimized for small inputs).
#[inline(always)]
fn decode_hex_scalar(hex_bytes: &[u8]) -> Result<Vec<u8>> {
    let mut result = Vec::with_capacity(hex_bytes.len() / 2);

    for chunk in hex_bytes.chunks_exact(2) {
        let hi = hex_char_to_nibble(chunk[0])?;
        let lo = hex_char_to_nibble(chunk[1])?;
        result.push((hi << 4) | lo);
    }

    Ok(result)
}

/// Convert a hex character to its nibble value (0-15).
#[inline(always)]
fn hex_char_to_nibble(c: u8) -> Result<u8> {
    match c {
        b'0'..=b'9' => Ok(c - b'0'),
        b'a'..=b'f' => Ok(c - b'a' + 10),
        b'A'..=b'F' => Ok(c - b'A' + 10),
        _ => Err(Error::ParseError(format!(
            "Invalid hex character: '{}'",
            c as char
        ))),
    }
}

/// AVX2-accelerated hex decoding for x86_64.
#[cfg(all(target_arch = "x86_64", target_feature = "avx2"))]
#[inline]
fn decode_hex_avx2(hex_bytes: &[u8]) -> Result<Vec<u8>> {
    #[cfg(target_arch = "x86_64")]
    use std::arch::x86_64::*;

    let mut result = Vec::with_capacity(hex_bytes.len() / 2);
    let len = hex_bytes.len();

    // Process 64 hex chars (32 bytes output) at a time with AVX2
    let mut pos = 0;

    unsafe {
        // Constants for hex validation and conversion
        let ascii_zero = _mm256_set1_epi8(b'0' as i8);
        let ascii_nine = _mm256_set1_epi8(b'9' as i8);
        let ascii_a_lower = _mm256_set1_epi8(b'a' as i8);
        let ascii_f_lower = _mm256_set1_epi8(b'f' as i8);
        let ascii_a_upper = _mm256_set1_epi8(b'A' as i8);
        let ascii_f_upper = _mm256_set1_epi8(b'F' as i8);
        let offset_lower = _mm256_set1_epi8((b'a' - 10) as i8);
        let offset_upper = _mm256_set1_epi8((b'A' - 10) as i8);

        while pos + 64 <= len {
            // Load 64 hex characters
            let hex_chars = _mm256_loadu_si256(hex_bytes.as_ptr().add(pos) as *const __m256i);
            let hex_chars2 = _mm256_loadu_si256(hex_bytes.as_ptr().add(pos + 32) as *const __m256i);

            // Validate and convert both vectors
            let nibbles1 = convert_hex_to_nibbles_avx2(
                hex_chars,
                ascii_zero,
                ascii_nine,
                ascii_a_lower,
                ascii_f_lower,
                ascii_a_upper,
                ascii_f_upper,
                offset_lower,
                offset_upper,
            )?;

            let nibbles2 = convert_hex_to_nibbles_avx2(
                hex_chars2,
                ascii_zero,
                ascii_nine,
                ascii_a_lower,
                ascii_f_lower,
                ascii_a_upper,
                ascii_f_upper,
                offset_lower,
                offset_upper,
            )?;

            // Combine nibbles into bytes
            let bytes1 = combine_nibbles_avx2(nibbles1);
            let bytes2 = combine_nibbles_avx2(nibbles2);

            // Store to temporary buffer to avoid uninitialized memory
            let mut temp: [u8; 32] = [0; 32];
            _mm_storeu_si128(temp.as_mut_ptr() as *mut __m128i, bytes1);
            _mm_storeu_si128(temp.as_mut_ptr().add(16) as *mut __m128i, bytes2);
            result.extend_from_slice(&temp);

            pos += 64;
        }
    }

    // Handle remaining bytes with scalar fallback
    if pos < len {
        let remaining = decode_hex_scalar(&hex_bytes[pos..])?;
        result.extend_from_slice(&remaining);
    }

    Ok(result)
}

#[cfg(all(target_arch = "x86_64", target_feature = "avx2"))]
#[inline(always)]
unsafe fn convert_hex_to_nibbles_avx2(
    hex_chars: std::arch::x86_64::__m256i,
    ascii_zero: std::arch::x86_64::__m256i,
    ascii_nine: std::arch::x86_64::__m256i,
    ascii_a_lower: std::arch::x86_64::__m256i,
    ascii_f_lower: std::arch::x86_64::__m256i,
    ascii_a_upper: std::arch::x86_64::__m256i,
    ascii_f_upper: std::arch::x86_64::__m256i,
    offset_lower: std::arch::x86_64::__m256i,
    offset_upper: std::arch::x86_64::__m256i,
) -> Result<std::arch::x86_64::__m256i> {
    use std::arch::x86_64::*;

    // Check if digits (0-9)
    let is_digit = _mm256_and_si256(
        _mm256_cmpgt_epi8(hex_chars, _mm256_sub_epi8(ascii_zero, _mm256_set1_epi8(1))),
        _mm256_cmpgt_epi8(_mm256_add_epi8(ascii_nine, _mm256_set1_epi8(1)), hex_chars),
    );

    // Check if lowercase (a-f)
    let is_lower = _mm256_and_si256(
        _mm256_cmpgt_epi8(
            hex_chars,
            _mm256_sub_epi8(ascii_a_lower, _mm256_set1_epi8(1)),
        ),
        _mm256_cmpgt_epi8(
            _mm256_add_epi8(ascii_f_lower, _mm256_set1_epi8(1)),
            hex_chars,
        ),
    );

    // Check if uppercase (A-F)
    let is_upper = _mm256_and_si256(
        _mm256_cmpgt_epi8(
            hex_chars,
            _mm256_sub_epi8(ascii_a_upper, _mm256_set1_epi8(1)),
        ),
        _mm256_cmpgt_epi8(
            _mm256_add_epi8(ascii_f_upper, _mm256_set1_epi8(1)),
            hex_chars,
        ),
    );

    // Validate: must be digit OR lower OR upper
    let is_valid = _mm256_or_si256(_mm256_or_si256(is_digit, is_lower), is_upper);
    let valid_mask = _mm256_movemask_epi8(is_valid);

    if valid_mask != -1 {
        return Err(Error::ParseError(
            "Invalid hex character in input".to_string(),
        ));
    }

    // Convert to nibbles: digits subtract '0', lowercase subtract 'a'-10, uppercase subtract 'A'-10
    let from_digits = _mm256_sub_epi8(hex_chars, ascii_zero);
    let from_lower = _mm256_sub_epi8(hex_chars, offset_lower);
    let from_upper = _mm256_sub_epi8(hex_chars, offset_upper);

    // Select correct conversion based on character type
    let nibbles = _mm256_blendv_epi8(
        _mm256_blendv_epi8(from_digits, from_lower, is_lower),
        from_upper,
        is_upper,
    );

    Ok(nibbles)
}

#[cfg(all(target_arch = "x86_64", target_feature = "avx2"))]
#[inline(always)]
unsafe fn combine_nibbles_avx2(nibbles: std::arch::x86_64::__m256i) -> std::arch::x86_64::__m128i {
    use std::arch::x86_64::*;

    // nibbles contains 32 nibble values where pairs need to be combined
    // [h0, l0, h1, l1, ...] -> [(h0<<4)|l0, (h1<<4)|l1, ...]

    // Extract high nibbles (even positions) and low nibbles (odd positions)
    let shuffle_hi = _mm256_setr_epi8(
        0, 2, 4, 6, 8, 10, 12, 14, -1, -1, -1, -1, -1, -1, -1, -1, 0, 2, 4, 6, 8, 10, 12, 14, -1,
        -1, -1, -1, -1, -1, -1, -1,
    );
    let shuffle_lo = _mm256_setr_epi8(
        1, 3, 5, 7, 9, 11, 13, 15, -1, -1, -1, -1, -1, -1, -1, -1, 1, 3, 5, 7, 9, 11, 13, 15, -1,
        -1, -1, -1, -1, -1, -1, -1,
    );

    let hi_nibbles = _mm256_shuffle_epi8(nibbles, shuffle_hi);
    let lo_nibbles = _mm256_shuffle_epi8(nibbles, shuffle_lo);

    // Shift high nibbles left by 4 bits (multiply by 16)
    let sixteen = _mm256_set1_epi8(16);
    let hi_shifted = _mm256_mullo_epi16(hi_nibbles, sixteen);
    let combined = _mm256_or_si256(hi_shifted, lo_nibbles);

    // Extract first 8 bytes from each 128-bit lane
    let low128 = _mm256_castsi256_si128(combined);
    let high128 = _mm256_extracti128_si256(combined, 1);

    // Combine both 64-bit results into single 128-bit register
    _mm_unpacklo_epi64(low128, high128)
}

/// SSE4.1-accelerated hex decoding for x86_64.
#[cfg(all(target_arch = "x86_64", target_feature = "sse4.1"))]
#[inline]
fn decode_hex_sse41(hex_bytes: &[u8]) -> Result<Vec<u8>> {
    #[cfg(target_arch = "x86_64")]
    use std::arch::x86_64::*;

    let mut result = Vec::with_capacity(hex_bytes.len() / 2);
    let len = hex_bytes.len();
    let mut pos = 0;

    unsafe {
        // Process 32 hex chars (16 bytes output) at a time
        while pos + 32 <= len {
            let hex_chars = _mm_loadu_si128(hex_bytes.as_ptr().add(pos) as *const __m128i);
            let hex_chars2 = _mm_loadu_si128(hex_bytes.as_ptr().add(pos + 16) as *const __m128i);

            let nibbles1 = convert_hex_to_nibbles_sse(hex_chars)?;
            let nibbles2 = convert_hex_to_nibbles_sse(hex_chars2)?;

            let bytes = combine_nibbles_sse(nibbles1, nibbles2);

            // Store to temporary buffer to avoid uninitialized memory
            let mut temp: [u8; 16] = [0; 16];
            _mm_storeu_si128(temp.as_mut_ptr() as *mut __m128i, bytes);
            result.extend_from_slice(&temp);

            pos += 32;
        }
    }

    // Handle remaining bytes with scalar fallback
    if pos < len {
        let remaining = decode_hex_scalar(&hex_bytes[pos..])?;
        result.extend_from_slice(&remaining);
    }

    Ok(result)
}

#[cfg(all(target_arch = "x86_64", target_feature = "sse4.1"))]
#[inline(always)]
unsafe fn convert_hex_to_nibbles_sse(
    hex_chars: std::arch::x86_64::__m128i,
) -> Result<std::arch::x86_64::__m128i> {
    use std::arch::x86_64::*;

    let ascii_zero = _mm_set1_epi8(b'0' as i8);
    let nine = _mm_set1_epi8(9);
    let ascii_a_lower = _mm_set1_epi8((b'a' - 10) as i8);
    let ascii_a_upper = _mm_set1_epi8((b'A' - 10) as i8);

    // Subtract '0' from all characters
    let normalized = _mm_sub_epi8(hex_chars, ascii_zero);

    // Check if it's a digit (0-9)
    let is_digit = _mm_cmpgt_epi8(_mm_set1_epi8(10), normalized);

    // For letters, subtract additional offset
    let letter_offset = _mm_blendv_epi8(
        _mm_sub_epi8(hex_chars, ascii_a_upper),
        _mm_sub_epi8(hex_chars, ascii_a_lower),
        _mm_cmpgt_epi8(hex_chars, _mm_set1_epi8(b'Z' as i8)),
    );

    let nibbles = _mm_blendv_epi8(letter_offset, normalized, is_digit);

    // Validate range (0-15)
    let is_valid = _mm_cmpgt_epi8(_mm_set1_epi8(16), nibbles);
    let valid_mask = _mm_movemask_epi8(is_valid);

    if valid_mask != 0xFFFF {
        return Err(Error::ParseError(
            "Invalid hex character in input".to_string(),
        ));
    }

    Ok(nibbles)
}

#[cfg(all(target_arch = "x86_64", target_feature = "sse4.1"))]
#[inline(always)]
unsafe fn combine_nibbles_sse(
    nibbles1: std::arch::x86_64::__m128i,
    nibbles2: std::arch::x86_64::__m128i,
) -> std::arch::x86_64::__m128i {
    use std::arch::x86_64::*;

    // Each nibbles vector contains 16 values: [h0, l0, h1, l1, ...]
    // Extract high nibbles (even positions) and low nibbles (odd positions)
    let shuffle_hi = _mm_setr_epi8(0, 2, 4, 6, 8, 10, 12, 14, -1, -1, -1, -1, -1, -1, -1, -1);
    let shuffle_lo = _mm_setr_epi8(1, 3, 5, 7, 9, 11, 13, 15, -1, -1, -1, -1, -1, -1, -1, -1);

    let hi1 = _mm_shuffle_epi8(nibbles1, shuffle_hi);
    let lo1 = _mm_shuffle_epi8(nibbles1, shuffle_lo);
    let hi2 = _mm_shuffle_epi8(nibbles2, shuffle_hi);
    let lo2 = _mm_shuffle_epi8(nibbles2, shuffle_lo);

    // Shift high nibbles left by 4 bits (multiply by 16)
    let sixteen = _mm_set1_epi8(16);
    let bytes1 = _mm_or_si128(_mm_mullo_epi16(hi1, sixteen), lo1);
    let bytes2 = _mm_or_si128(_mm_mullo_epi16(hi2, sixteen), lo2);

    // Pack into single register (take first 8 bytes from each)
    _mm_unpacklo_epi64(bytes1, bytes2)
}

/// NEON-accelerated hex decoding for aarch64.
#[cfg(target_arch = "aarch64")]
#[target_feature(enable = "neon")]
#[inline]
unsafe fn decode_hex_neon(hex_bytes: &[u8]) -> Result<Vec<u8>> {
    use std::arch::aarch64::*;

    let mut result: Vec<u8> = Vec::with_capacity(hex_bytes.len() / 2);
    let len = hex_bytes.len();
    let mut pos = 0;

    // SAFETY: All NEON intrinsics require the neon feature which is enabled by target_feature
    unsafe {
        let ascii_zero = vdupq_n_u8(b'0');
        let ascii_nine = vdupq_n_u8(b'9');
        let ascii_a_lower = vdupq_n_u8(b'a');
        let ascii_a_upper = vdupq_n_u8(b'A');
        let offset_digit = vdupq_n_u8(b'0');
        let offset_lower = vdupq_n_u8(b'a' - 10);
        let offset_upper = vdupq_n_u8(b'A' - 10);

        // Process 32 hex chars (16 bytes output) at a time
        while pos + 32 <= len {
            let hex_chars1 = vld1q_u8(hex_bytes.as_ptr().add(pos));
            let hex_chars2 = vld1q_u8(hex_bytes.as_ptr().add(pos + 16));

            let nibbles1 = convert_hex_to_nibbles_neon(
                hex_chars1,
                ascii_zero,
                ascii_nine,
                ascii_a_lower,
                ascii_a_upper,
                offset_digit,
                offset_lower,
                offset_upper,
            )?;

            let nibbles2 = convert_hex_to_nibbles_neon(
                hex_chars2,
                ascii_zero,
                ascii_nine,
                ascii_a_lower,
                ascii_a_upper,
                offset_digit,
                offset_lower,
                offset_upper,
            )?;

            let bytes = combine_nibbles_neon(nibbles1, nibbles2);

            // Store to temporary buffer to avoid uninitialized memory
            let mut temp: [u8; 16] = [0; 16];
            vst1q_u8(temp.as_mut_ptr(), bytes);
            result.extend_from_slice(&temp);

            pos += 32;
        }
    }

    // Handle remaining bytes with scalar fallback
    if pos < len {
        let remaining = decode_hex_scalar(&hex_bytes[pos..])?;
        result.extend_from_slice(&remaining);
    }

    Ok(result)
}

#[cfg(target_arch = "aarch64")]
#[target_feature(enable = "neon")]
#[allow(clippy::too_many_arguments)] // SIMD functions naturally have many parameters
unsafe fn convert_hex_to_nibbles_neon(
    hex_chars: std::arch::aarch64::uint8x16_t,
    ascii_zero: std::arch::aarch64::uint8x16_t,
    ascii_nine: std::arch::aarch64::uint8x16_t,
    ascii_a_lower: std::arch::aarch64::uint8x16_t,
    ascii_a_upper: std::arch::aarch64::uint8x16_t,
    offset_digit: std::arch::aarch64::uint8x16_t,
    offset_lower: std::arch::aarch64::uint8x16_t,
    offset_upper: std::arch::aarch64::uint8x16_t,
) -> Result<std::arch::aarch64::uint8x16_t> {
    use std::arch::aarch64::*;

    // Check if digits (0-9)
    let is_digit = vandq_u8(
        vcgeq_u8(hex_chars, ascii_zero),
        vcleq_u8(hex_chars, ascii_nine),
    );

    // Check if lowercase (a-f)
    let is_lower = vandq_u8(
        vcgeq_u8(hex_chars, ascii_a_lower),
        vcleq_u8(hex_chars, vdupq_n_u8(b'f')),
    );

    // Check if uppercase (A-F)
    let is_upper = vandq_u8(
        vcgeq_u8(hex_chars, ascii_a_upper),
        vcleq_u8(hex_chars, vdupq_n_u8(b'F')),
    );

    // Validate: must be digit OR lower OR upper
    let is_valid = vorrq_u8(vorrq_u8(is_digit, is_lower), is_upper);

    // Check if all lanes are valid (all bits set)
    let valid_check = vminvq_u8(is_valid);
    if valid_check == 0 {
        return Err(Error::ParseError(
            "Invalid hex character in input".to_string(),
        ));
    }

    // Convert to nibbles
    let from_digits = vsubq_u8(hex_chars, offset_digit);
    let from_lower = vsubq_u8(hex_chars, offset_lower);
    let from_upper = vsubq_u8(hex_chars, offset_upper);

    // Select correct conversion
    let nibbles = vbslq_u8(
        is_digit,
        from_digits,
        vbslq_u8(is_lower, from_lower, from_upper),
    );

    Ok(nibbles)
}

#[cfg(target_arch = "aarch64")]
#[target_feature(enable = "neon")]
unsafe fn combine_nibbles_neon(
    nibbles1: std::arch::aarch64::uint8x16_t,
    nibbles2: std::arch::aarch64::uint8x16_t,
) -> std::arch::aarch64::uint8x16_t {
    use std::arch::aarch64::*;

    // nibbles1 contains 16 nibble values: [h0, l0, h1, l1, h2, l2, ...]
    // We need to combine them as: [(h0<<4)|l0, (h1<<4)|l1, ...]

    // Extract even-indexed nibbles (high nibbles) from both vectors
    let hi = vuzp1q_u8(nibbles1, nibbles2);
    // Extract odd-indexed nibbles (low nibbles) from both vectors
    let lo = vuzp2q_u8(nibbles1, nibbles2);

    // Shift high nibbles left by 4 and OR with low nibbles
    vorrq_u8(vshlq_n_u8(hi, 4), lo)
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

    #[test]
    fn test_decode_hex_data_basic() {
        let hex = "48656C6C6F"; // "Hello"
        let decoded = decode_hex_data(hex).unwrap();
        assert_eq!(decoded, b"Hello");
    }

    #[test]
    fn test_decode_hex_data_with_whitespace() {
        let hex = "48 65 6C 6C 6F"; // "Hello" with spaces
        let decoded = decode_hex_data(hex).unwrap();
        assert_eq!(decoded, b"Hello");
    }

    #[test]
    fn test_decode_hex_data_lowercase() {
        let hex = "48656c6c6f"; // "Hello" with lowercase hex
        let decoded = decode_hex_data(hex).unwrap();
        assert_eq!(decoded, b"Hello");
    }

    #[test]
    fn test_decode_hex_data_mixed_case() {
        let hex = "48656C6c6F"; // Mixed case
        let decoded = decode_hex_data(hex).unwrap();
        assert_eq!(decoded, b"Hello");
    }

    #[test]
    fn test_decode_hex_data_empty() {
        let hex = "";
        let decoded = decode_hex_data(hex).unwrap();
        assert_eq!(decoded, b"");
    }

    #[test]
    fn test_decode_hex_data_invalid_length() {
        let hex = "48656C6C6"; // Odd number of characters
        assert!(decode_hex_data(hex).is_err());
    }

    #[test]
    fn test_decode_hex_data_invalid_char() {
        let hex = "48656C6C6Z"; // Invalid character 'Z'
        assert!(decode_hex_data(hex).is_err());
    }

    #[test]
    fn test_decode_hex_data_large() {
        // Test with larger data to trigger SIMD paths
        let hex = "48656C6C6F576F726C64".repeat(100); // "HelloWorld" repeated
        let decoded = decode_hex_data(&hex).unwrap();
        let expected = b"HelloWorld".repeat(100);
        assert_eq!(decoded, expected);
    }

    #[test]
    fn test_hex_char_to_nibble() {
        assert_eq!(hex_char_to_nibble(b'0').unwrap(), 0);
        assert_eq!(hex_char_to_nibble(b'9').unwrap(), 9);
        assert_eq!(hex_char_to_nibble(b'a').unwrap(), 10);
        assert_eq!(hex_char_to_nibble(b'f').unwrap(), 15);
        assert_eq!(hex_char_to_nibble(b'A').unwrap(), 10);
        assert_eq!(hex_char_to_nibble(b'F').unwrap(), 15);
        assert!(hex_char_to_nibble(b'G').is_err());
        assert!(hex_char_to_nibble(b'g').is_err());
    }
}
