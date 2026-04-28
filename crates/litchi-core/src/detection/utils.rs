//! Utility functions and constants for file format detection.

// Magic number signatures
pub const OLE2_SIGNATURE: &[u8] = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
pub const ZIP_SIGNATURE: &[u8] = &[0x50, 0x4B, 0x03, 0x04];

/// Helper function to find a pattern in a buffer efficiently.
/// Only scans the necessary portion without full buffer traversal.
pub fn find_in_buffer(buffer: &[u8], pattern: &[u8]) -> bool {
    buffer
        .windows(pattern.len())
        .any(|window| window == pattern)
}

// Note: read_zip_file helper removed - use soapberry_zip::office::ArchiveReader::read() instead
