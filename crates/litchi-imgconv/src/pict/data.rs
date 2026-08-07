//! Data manipulation and compression utilities
//!
//! Provides functions for decompressing and manipulating binary data,
//! particularly for PICT format processing.

use litchi_core::error::{Error, Result};

/// UnpackBits decompression algorithm
///
/// Decompresses PackBits-compressed data as used in PICT files.
/// This is a run-length encoding scheme where:
/// - Positive values (0-127) indicate literal bytes to copy
/// - Negative values (-1 to -127) indicate run-length encoding
/// - -128 is a no-op (ignored)
///
/// # Arguments
/// * `compressed` - The compressed input data
/// * `expected_size` - The expected size of the decompressed output
///
/// # Returns
/// Decompressed data as a Vec<u8>
///
/// # Performance Notes
/// - Uses SIMD-friendly operations where possible
/// - Avoids unnecessary allocations during decompression
pub(super) fn unpack_bits(compressed: &[u8], expected_size: usize) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(expected_size);
    let mut input_pos = 0;
    let mut bytes_done = 0;

    while bytes_done < expected_size && input_pos < compressed.len() {
        let code = compressed[input_pos] as i8;
        input_pos += 1;

        if code == -128 {
            // No-op, skip
            continue;
        } else if code < 0 {
            // Run-length encoded: repeat next byte (1 - code) times
            let run_length = (1i32 - code as i32) as usize;
            if input_pos >= compressed.len() {
                return Err(Error::ParseError(
                    "Invalid PackBits data: unexpected end of input".into(),
                ));
            }
            let byte = compressed[input_pos];
            input_pos += 1;

            // Extend output with repeated byte
            if bytes_done + run_length > expected_size {
                return Err(Error::ParseError(
                    "PackBits decompression exceeded expected size".into(),
                ));
            }
            output.extend(std::iter::repeat_n(byte, run_length));
            bytes_done += run_length;
        } else {
            // Literal bytes: copy (code + 1) bytes directly
            let literal_count = (code as usize) + 1;
            if input_pos + literal_count > compressed.len() {
                return Err(Error::ParseError(
                    "Invalid PackBits data: not enough literal bytes".into(),
                ));
            }
            if bytes_done + literal_count > expected_size {
                return Err(Error::ParseError(
                    "PackBits decompression exceeded expected size".into(),
                ));
            }
            output.extend_from_slice(&compressed[input_pos..input_pos + literal_count]);
            input_pos += literal_count;
            bytes_done += literal_count;
        }
    }

    if bytes_done != expected_size {
        return Err(Error::ParseError(format!(
            "PackBits decompression size mismatch: expected {}, got {}",
            expected_size, bytes_done
        )));
    }

    Ok(output)
}

/// Get a pixel value from a 1-bit bitmap
///
/// Extracts a single bit from a packed bitmap and converts it to an RGBA color.
/// In PICT format, 1 = black (0xFF000000), 0 = white (0xFFFFFFFF).
///
/// # Arguments
/// * `bitmap` - The packed bitmap data
/// * `bounds` - The bitmap bounds rectangle
/// * `x` - X coordinate relative to bounds
/// * `y` - Y coordinate relative to bounds
///
/// # Returns
/// RGBA color value as u32 (0xAARRGGBB format)
#[inline(always)]
pub(super) fn get_bitmap_pixel(
    bitmap: &[u8],
    bounds: &super::types::PictRect,
    x: i32,
    y: i32,
) -> u32 {
    let width = i32::from(bounds.right) - i32::from(bounds.left);
    let height = i32::from(bounds.bottom) - i32::from(bounds.top);

    // Check bounds
    if x < 0 || y < 0 || x >= width || y >= height {
        return 0xFFFFFFFF; // White for out of bounds
    }

    let stride = i64::from(width);
    let bit_offset = 7 - (x % 8);
    let byte_pos = i64::from(y)
        .checked_mul(stride)
        .and_then(|bits| bits.checked_div(8))
        .and_then(|bytes| bytes.checked_add(i64::from(x / 8)))
        .and_then(|position| usize::try_from(position).ok());

    let Some(byte_pos) = byte_pos.filter(|&position| position < bitmap.len()) else {
        return 0xFFFFFFFF; // White for invalid position
    };

    let byte = bitmap[byte_pos];
    let bit_set = (byte & (1 << bit_offset)) != 0;

    // PICT format: 1 = black, 0 = white
    if bit_set {
        0xFF000000 // Black
    } else {
        0xFFFFFFFF // White
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_unpack_bits_literal() {
        // Test literal bytes (code = 2 means 3 literal bytes)
        let compressed = vec![2, 0xAA, 0xBB, 0xCC];
        let result = unpack_bits(&compressed, 3).unwrap();
        assert_eq!(result, vec![0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn test_unpack_bits_run() {
        // Test run-length encoding (code = -2 means 3 repetitions of next byte)
        let compressed = vec![0xFE, 0xDD]; // -2, then 0xDD
        let result = unpack_bits(&compressed, 3).unwrap();
        assert_eq!(result, vec![0xDD, 0xDD, 0xDD]);
    }

    #[test]
    fn test_unpack_bits_noop() {
        // Test no-op code (-128)
        let compressed = vec![0x80, 0, 0xEE]; // -128, then one literal byte
        let result = unpack_bits(&compressed, 1).unwrap();
        assert_eq!(result, vec![0xEE]);
    }

    #[test]
    fn test_unpack_bits_mixed() {
        // Test mixed literal and run-length
        let compressed = vec![1, 0x11, 0x22, 0xFE, 0x33]; // 2 literals, then 3 repeats
        let result = unpack_bits(&compressed, 5).unwrap();
        assert_eq!(result, vec![0x11, 0x22, 0x33, 0x33, 0x33]);
    }

    #[test]
    fn test_unpack_bits_error() {
        // Test error case - insufficient data
        let compressed = vec![2, 0xAA]; // Code says 3 bytes but only 1 provided
        assert!(unpack_bits(&compressed, 3).is_err());
    }
}
