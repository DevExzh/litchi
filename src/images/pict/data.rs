//! Data manipulation and compression utilities
//!
//! Provides functions for decompressing and manipulating binary data,
//! particularly for PICT format processing.

use crate::common::error::{Error, Result};

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
pub fn unpack_bits(compressed: &[u8], expected_size: usize) -> Result<Vec<u8>> {
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
pub fn get_bitmap_pixel(bitmap: &[u8], bounds: &super::types::PictRect, x: i32, y: i32) -> u32 {
    let width = bounds.right - bounds.left;
    let height = bounds.bottom - bounds.top;

    // Check bounds
    if x < 0 || y < 0 || x >= width as i32 || y >= height as i32 {
        return 0xFFFFFFFF; // White for out of bounds
    }

    let stride = width as i32;
    let bit_offset = 7 - (x % 8);
    let byte_pos = (y * stride / 8) + (x / 8);

    if byte_pos >= bitmap.len() as i32 {
        return 0xFFFFFFFF; // White for invalid position
    }

    let byte = bitmap[byte_pos as usize];
    let bit_set = (byte & (1 << bit_offset)) != 0;

    // PICT format: 1 = black, 0 = white
    if bit_set {
        0xFF000000 // Black
    } else {
        0xFFFFFFFF // White
    }
}

/// Stretch coordinates from destination rectangle to source rectangle
///
/// Used for scaling bitmap data when blitting from source to destination rectangles.
/// This implements bilinear-like coordinate mapping for bitmap scaling.
///
/// # Arguments
/// * `dst_rect` - Destination rectangle
/// * `src_rect` - Source rectangle
/// * `dst_x` - Destination X coordinate
/// * `dst_y` - Destination Y coordinate
/// * `src_x` - Output source X coordinate
/// * `src_y` - Output source Y coordinate
#[inline(always)]
pub fn stretch_coordinates(
    dst_rect: &super::types::PictRect,
    src_rect: &super::types::PictRect,
    dst_x: i32,
    dst_y: i32,
    src_x: &mut i32,
    src_y: &mut i32,
) {
    let dst_width = dst_rect.right - dst_rect.left;
    let dst_height = dst_rect.bottom - dst_rect.top;
    let src_width = src_rect.right - src_rect.left;
    let src_height = src_rect.bottom - src_rect.top;

    if dst_width == 0 || dst_height == 0 || src_width == 0 || src_height == 0 {
        *src_x = 0;
        *src_y = 0;
        return;
    }

    // Convert destination coordinates to relative positions
    let x_rel = dst_x - dst_rect.left as i32;
    let y_rel = dst_y - dst_rect.top as i32;

    // Calculate ratios and scale
    let x_ratio = src_width as f64 / dst_width as f64;
    let y_ratio = src_height as f64 / dst_height as f64;

    let x_scaled = x_rel as f64 * x_ratio;
    let y_scaled = y_rel as f64 * y_ratio;

    *src_x = src_rect.left as i32 + x_scaled as i32;
    *src_y = src_rect.top as i32 + y_scaled as i32;
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
        let compressed = vec![0x80, 1, 0xEE]; // -128, then literal
        let result = unpack_bits(&compressed, 1).unwrap();
        assert_eq!(result, vec![0xEE]);
    }

    #[test]
    fn test_unpack_bits_mixed() {
        // Test mixed literal and run-length
        let compressed = vec![1, 0x11, 0x22, 0xFD, 0x33]; // 2 literals, then 3 repeats
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
