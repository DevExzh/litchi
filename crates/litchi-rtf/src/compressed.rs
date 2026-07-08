//! Compressed RTF support.
//!
//! This module implements the RTF compression algorithm as specified in:
//! https://msdn.microsoft.com/en-us/library/cc463890(v=exchg.80).aspx
//!
//! Compressed RTF is commonly used in email attachments and other scenarios
//! where file size reduction is important.

use super::error::{RtfError, RtfResult};
use std::io::{Cursor, Read, Write};
use zerocopy::{FromBytes, IntoBytes};
use zerocopy_derive::{
    FromBytes as DeriveFromBytes, Immutable, IntoBytes as DeriveIntoBytes, KnownLayout,
};

/// Magic signature for compressed RTF
const COMPRESSED_SIGNATURE: &[u8; 4] = b"LZFu";

/// Magic signature for uncompressed RTF (stored with compression header)
const UNCOMPRESSED_SIGNATURE: &[u8; 4] = b"MELA";

/// Initial dictionary for compression/decompression
const INIT_DICT: &[u8] = b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}\
{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial\
Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\
\\b\\i\\u\\tab\\tx";

/// Size of initial dictionary
const INIT_DICT_SIZE: usize = 207;

/// Maximum dictionary size
const MAX_DICT_SIZE: usize = 4096;

/// Compressed RTF header (16 bytes)
#[repr(C)]
#[derive(Debug, Clone, Copy, DeriveIntoBytes, DeriveFromBytes, Immutable, KnownLayout)]
struct CompressedRtfHeader {
    /// Total size of compressed data including header (little-endian)
    compressed_size: [u8; 4],
    /// Size of uncompressed data (little-endian)
    raw_size: [u8; 4],
    /// Compression type signature
    compression_type: [u8; 4],
    /// CRC32 checksum (little-endian)
    crc32: [u8; 4],
}

impl CompressedRtfHeader {
    /// Get raw size as u32
    #[inline]
    fn get_raw_size(&self) -> u32 {
        u32::from_le_bytes(self.raw_size)
    }

    /// Get CRC32 as u32
    #[inline]
    fn get_crc32(&self) -> u32 {
        u32::from_le_bytes(self.crc32)
    }

    /// Create a new header
    fn new(compressed_size: u32, raw_size: u32, compression_type: [u8; 4], crc32: u32) -> Self {
        Self {
            compressed_size: compressed_size.to_le_bytes(),
            raw_size: raw_size.to_le_bytes(),
            compression_type,
            crc32: crc32.to_le_bytes(),
        }
    }

    /// Check if this is a compressed RTF signature
    fn is_compressed(&self) -> bool {
        &self.compression_type == COMPRESSED_SIGNATURE
    }

    /// Check if this is an uncompressed RTF signature
    fn is_uncompressed(&self) -> bool {
        &self.compression_type == UNCOMPRESSED_SIGNATURE
    }
}

/// Detect if data is compressed RTF
pub fn is_compressed_rtf(data: &[u8]) -> bool {
    if data.len() < 16 {
        return false;
    }

    // Check for LZFu or MELA signature
    let signature = &data[8..12];
    signature == COMPRESSED_SIGNATURE || signature == UNCOMPRESSED_SIGNATURE
}

/// Decompress RTF data
///
/// # Arguments
///
/// * `data` - Compressed RTF data with header
///
/// # Returns
///
/// Decompressed RTF data as bytes
///
/// # Errors
///
/// Returns error if:
/// - Data is too small
/// - CRC check fails
/// - Unknown compression type
pub fn decompress(data: &[u8]) -> RtfResult<Vec<u8>> {
    // Parse header using zerocopy
    if data.len() < 16 {
        return Err(RtfError::InvalidStructure(
            "Compressed RTF header must be at least 16 bytes".to_string(),
        ));
    }

    let header = <CompressedRtfHeader as FromBytes>::ref_from_bytes(&data[..16]).map_err(|_| {
        RtfError::InvalidStructure("Failed to parse compressed RTF header".to_string())
    })?;
    let header = *header;

    // Get compressed data (excluding 16-byte header)
    let compressed_data = &data[16..];

    if header.is_compressed() {
        decompress_lzfu(compressed_data, &header)
    } else if header.is_uncompressed() {
        decompress_uncompressed(compressed_data, &header)
    } else {
        Err(RtfError::InvalidStructure(format!(
            "Unknown compression type: {:?}",
            header.compression_type
        )))
    }
}

/// Decompress LZFu compressed data
fn decompress_lzfu(data: &[u8], header: &CompressedRtfHeader) -> RtfResult<Vec<u8>> {
    // Verify CRC32
    let calculated_crc = crc_fast::checksum(crc_fast::CrcAlgorithm::Crc32IsoHdlc, data) as u32;

    if calculated_crc != header.get_crc32() {
        return Err(RtfError::InvalidStructure(format!(
            "CRC32 mismatch: expected {:#x}, got {:#x}",
            header.get_crc32(),
            calculated_crc
        )));
    }

    // Initialize dictionary
    let mut dict = vec![0u8; MAX_DICT_SIZE];
    dict[..INIT_DICT_SIZE].copy_from_slice(INIT_DICT);
    dict[INIT_DICT_SIZE..].fill(b' ');

    let mut write_offset = INIT_DICT_SIZE;
    let mut output = Vec::with_capacity(header.get_raw_size() as usize);
    let mut cursor = Cursor::new(data);

    loop {
        // Read control byte
        let mut control_byte = [0u8; 1];
        if cursor.read_exact(&mut control_byte).is_err() {
            break;
        }
        let control = control_byte[0];

        // Process each bit in control byte (LSB to MSB)
        for bit in 0..8 {
            if (control & (1 << bit)) != 0 {
                // Bit is 1: token is a reference (16-bit)
                let mut token_bytes = [0u8; 2];
                if cursor.read_exact(&mut token_bytes).is_err() {
                    break;
                }
                let token = u16::from_be_bytes(token_bytes);

                // Extract offset (12 bits) and length (4 bits)
                let offset = ((token >> 4) & 0x0FFF) as usize;
                let length = (token & 0x0F) as usize;

                // Check for end indicator
                if write_offset == offset {
                    return Ok(output);
                }

                // Copy from dictionary
                let actual_length = length + 2;
                for step in 0..actual_length {
                    let read_offset = (offset + step) % MAX_DICT_SIZE;
                    let byte = dict[read_offset];
                    output.push(byte);
                    dict[write_offset] = byte;
                    write_offset = (write_offset + 1) % MAX_DICT_SIZE;
                }
            } else {
                // Bit is 0: token is a literal (8-bit)
                let mut literal = [0u8; 1];
                if cursor.read_exact(&mut literal).is_err() {
                    break;
                }
                output.push(literal[0]);
                dict[write_offset] = literal[0];
                write_offset = (write_offset + 1) % MAX_DICT_SIZE;
            }
        }
    }

    Ok(output)
}

/// Decompress uncompressed RTF data
fn decompress_uncompressed(data: &[u8], header: &CompressedRtfHeader) -> RtfResult<Vec<u8>> {
    // Verify CRC32 is zero for uncompressed
    if header.get_crc32() != 0x00000000 {
        return Err(RtfError::InvalidStructure(
            "CRC32 must be 0x00000000 for uncompressed RTF".to_string(),
        ));
    }

    // Return raw data up to raw_size
    let size = (header.get_raw_size() as usize).min(data.len());
    Ok(data[..size].to_vec())
}

/// Compress RTF data
///
/// # Arguments
///
/// * `data` - Uncompressed RTF data
/// * `compress` - If true, use LZFu compression; if false, store uncompressed
///
/// # Returns
///
/// Compressed RTF data with header
pub fn compress(data: &[u8], compress: bool) -> RtfResult<Vec<u8>> {
    if compress {
        compress_lzfu(data)
    } else {
        compress_uncompressed(data)
    }
}

/// Compress data using LZFu algorithm
fn compress_lzfu(data: &[u8]) -> RtfResult<Vec<u8>> {
    // Initialize dictionary
    let mut dict = vec![0u8; MAX_DICT_SIZE];
    dict[..INIT_DICT_SIZE].copy_from_slice(INIT_DICT);
    dict[INIT_DICT_SIZE..].fill(b' ');

    let mut write_offset = INIT_DICT_SIZE;
    let mut output = Vec::new();
    let mut cursor = Cursor::new(data);

    let mut control_byte: u8 = 0;
    let mut control_bit = 0;
    let mut token_buffer = Vec::new();

    loop {
        // Find longest match
        let (dict_offset, longest_match) = find_longest_match(&mut dict, &mut cursor, write_offset);

        let mut read_buf = vec![0u8; if longest_match > 1 { longest_match } else { 1 }];
        let bytes_read = cursor.read(&mut read_buf).unwrap_or(0);

        // EOF
        if bytes_read == 0 {
            // Add end marker
            control_byte |= 1 << control_bit;

            let dict_ref = ((write_offset & 0xFFF) << 4) as u16;
            token_buffer.write_all(&dict_ref.to_be_bytes()).unwrap();

            // Flush final control byte and tokens
            output.push(control_byte);
            output.write_all(&token_buffer).unwrap();
            break;
        }

        if longest_match > 1 {
            // Dictionary reference
            control_byte |= 1 << control_bit;
            control_bit += 1;

            let dict_ref = (((dict_offset & 0xFFF) << 4) | ((longest_match - 2) & 0xF)) as u16;
            token_buffer.write_all(&dict_ref.to_be_bytes()).unwrap();
        } else {
            // Literal
            if longest_match == 0 {
                dict[write_offset] = read_buf[0];
                write_offset = (write_offset + 1) % MAX_DICT_SIZE;
            }

            control_byte |= 0 << control_bit;
            control_bit += 1;
            token_buffer.push(read_buf[0]);
        }

        // Flush when control byte is full
        if control_bit >= 8 {
            output.push(control_byte);
            output.write_all(&token_buffer).unwrap();
            control_byte = 0;
            control_bit = 0;
            token_buffer.clear();
        }
    }

    // Calculate CRC32
    let crc32 = crc_fast::checksum(crc_fast::CrcAlgorithm::Crc32IsoHdlc, &output) as u32;

    // Build header
    let header = CompressedRtfHeader::new(
        (output.len() + 12) as u32,
        data.len() as u32,
        *COMPRESSED_SIGNATURE,
        crc32,
    );

    // Combine header and compressed data
    let mut result = Vec::new();
    result.write_all(IntoBytes::as_bytes(&header)).unwrap();
    result.write_all(&output).unwrap();

    Ok(result)
}

/// Compress data without compression (just add header)
fn compress_uncompressed(data: &[u8]) -> RtfResult<Vec<u8>> {
    let header = CompressedRtfHeader::new(
        (data.len() + 12) as u32,
        data.len() as u32,
        *UNCOMPRESSED_SIGNATURE,
        0x00000000,
    );

    let mut result = Vec::new();
    result.write_all(IntoBytes::as_bytes(&header)).unwrap();
    result.write_all(data).unwrap();

    Ok(result)
}

/// Find the longest match in the dictionary
fn find_longest_match(
    dict: &mut [u8],
    stream: &mut Cursor<&[u8]>,
    write_offset: usize,
) -> (usize, usize) {
    let start_pos = stream.position();

    // Read first byte
    let mut byte_buf = [0u8; 1];
    if stream.read_exact(&mut byte_buf).is_err() {
        stream.set_position(start_pos);
        return (0, 0);
    }

    let first_byte = byte_buf[0];
    let prev_write_offset = write_offset;
    let mut dict_index = 0;
    let mut match_len = 0;
    let mut longest_match_len = 0;
    let mut dict_offset = 0;

    // Search dictionary
    loop {
        if dict[dict_index % MAX_DICT_SIZE] == first_byte {
            match_len += 1;

            if longest_match_len < match_len && match_len <= 17 {
                dict_offset = dict_index - match_len + 1;
                dict[write_offset + match_len - 1] = first_byte;
                longest_match_len = match_len;
            }

            // Try to extend match
            if stream.read_exact(&mut byte_buf).is_ok()
                && dict[(dict_index + 1) % MAX_DICT_SIZE] == byte_buf[0]
            {
                dict_index += 1;
                continue;
            }

            // No more bytes or no match
            stream.set_position(start_pos + longest_match_len as u64);
            return (dict_offset, longest_match_len);
        }

        // Reset match and try next position
        stream.set_position(start_pos);
        if stream.read_exact(&mut byte_buf).is_err() {
            break;
        }
        match_len = 0;

        dict_index += 1;
        if dict_index >= prev_write_offset + longest_match_len {
            break;
        }
    }

    stream.set_position(start_pos + longest_match_len as u64);
    (dict_offset, longest_match_len)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_compressed_rtf() {
        // Compressed signature
        let mut data = vec![0u8; 16];
        data[8..12].copy_from_slice(b"LZFu");
        assert!(is_compressed_rtf(&data));

        // Uncompressed signature
        let mut data = vec![0u8; 16];
        data[8..12].copy_from_slice(b"MELA");
        assert!(is_compressed_rtf(&data));

        // Not compressed RTF
        let data = vec![0u8; 16];
        assert!(!is_compressed_rtf(&data));

        // Too small
        let data = vec![0u8; 8];
        assert!(!is_compressed_rtf(&data));
    }

    #[test]
    fn test_round_trip_uncompressed() {
        let original = b"{\\rtf1\\ansi Hello World!\\par}";
        let compressed = compress(original, false).unwrap();
        let decompressed = decompress(&compressed).unwrap();
        assert_eq!(original, decompressed.as_slice());
    }
}
