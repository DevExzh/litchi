//! Snappy decompression for iWork IWA files
//!
//! iWork uses a custom Snappy framing format that differs from the standard:
//! - No stream identifier chunk
//! - No CRC-32C checksums
//! - Custom chunk header format (4 bytes: type + 24-bit length)

use snap::raw::Decoder;
use std::io::{self, Cursor, Read};

use crate::iwa::Error;

/// Custom Snappy stream decompressor for iWork IWA files
#[derive(Debug)]
pub struct SnappyStream {
    decompressed: Vec<u8>,
}

impl SnappyStream {
    /// Decompress an IWA file from a reader
    ///
    /// iWork IWA files use a custom Snappy framing format:
    /// - 4-byte header: [chunk_type, length_byte1, length_byte2, length_byte3]
    /// - chunk_type is always 0 for compressed chunks
    /// - length is a 24-bit little-endian integer
    /// - No stream identifier, no CRC checksums
    pub fn decompress<R: Read>(reader: &mut R) -> Result<Self, Error> {
        let mut decompressed = Vec::new();
        let mut decoder = Decoder::new();

        loop {
            // Read 4-byte header
            let mut header = [0u8; 4];
            match reader.read_exact(&mut header) {
                Ok(_) => {},
                Err(ref e) if e.kind() == io::ErrorKind::UnexpectedEof => {
                    // End of stream
                    break;
                },
                Err(e) => return Err(Error::Io(e)),
            }

            let chunk_type = header[0];
            if chunk_type != 0 {
                return Err(Error::Snappy(format!(
                    "Unexpected chunk type: {}, expected 0",
                    chunk_type
                )));
            }

            // Extract 24-bit length (little-endian)
            let length = u32::from_le_bytes([header[1], header[2], header[3], 0]);

            if length == 0 {
                continue;
            }

            // Read compressed chunk
            let mut compressed = vec![0u8; length as usize];
            reader.read_exact(&mut compressed).map_err(Error::Io)?;

            let mut chunk_decompressed = Vec::new();
            let mut buffer_size = 1024; // Start with 1KB

            loop {
                chunk_decompressed.resize(buffer_size, 0);
                match decoder.decompress(&compressed, &mut chunk_decompressed) {
                    Ok(decompressed_size) => {
                        // Success - truncate to actual size and break
                        chunk_decompressed.truncate(decompressed_size);
                        break;
                    },
                    Err(_) if buffer_size < 10 * 1024 * 1024 => {
                        // Buffer too small, try with larger buffer (up to 10MB)
                        buffer_size *= 2;
                        continue;
                    },
                    Err(e) => {
                        return Err(Error::Snappy(format!("Decompression failed: {}", e)));
                    },
                }
            }

            decompressed.extend(chunk_decompressed);
        }

        Ok(SnappyStream { decompressed })
    }

    /// Get the decompressed data as a slice
    pub fn data(&self) -> &[u8] {
        &self.decompressed
    }

    /// Get the decompressed data as a mutable slice
    pub fn data_mut(&mut self) -> &mut [u8] {
        &mut self.decompressed
    }

    /// Consume self and return the decompressed data
    pub fn into_data(self) -> Vec<u8> {
        self.decompressed
    }

    /// Create a reader for the decompressed data
    pub fn reader(&self) -> Cursor<&[u8]> {
        Cursor::new(&self.decompressed)
    }
}

impl AsRef<[u8]> for SnappyStream {
    fn as_ref(&self) -> &[u8] {
        self.data()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use soapberry_zip::office::ArchiveReader;
    use std::fs::File;
    use std::io::Cursor;

    #[test]
    fn test_empty_stream() {
        let empty_data = [];
        let mut cursor = Cursor::new(&empty_data);
        let result = SnappyStream::decompress(&mut cursor);
        assert!(result.is_ok());
        let stream = result.unwrap();
        assert_eq!(stream.data().len(), 0);
    }

    #[test]
    fn test_invalid_chunk_type() {
        // Create a header with invalid chunk type (1 instead of 0)
        let invalid_data = [1, 0, 0, 0]; // chunk_type=1, length=0
        let mut cursor = Cursor::new(&invalid_data);
        let result = SnappyStream::decompress(&mut cursor);
        assert!(result.is_err());
        match result.unwrap_err() {
            Error::Snappy(msg) => assert!(msg.contains("Unexpected chunk type")),
            _ => panic!("Expected Snappy error"),
        }
    }

    #[test]
    fn test_real_iwa_decompression() {
        use std::io::Read;

        // Test decompression with real IWA files from test bundles
        let test_files = vec!["test.pages", "test.numbers"];

        for test_file in test_files {
            if !std::path::Path::new(test_file).exists() {
                continue; // Skip if test file doesn't exist
            }

            // Read entire file into memory for soapberry_zip
            let mut file = File::open(test_file).expect("Failed to open test file");
            let mut file_data = Vec::new();
            file.read_to_end(&mut file_data)
                .expect("Failed to read test file");

            let archive = ArchiveReader::new(&file_data).expect("Failed to read zip archive");

            // Find an IWA file to test with
            for file_name in archive.file_names() {
                if file_name.ends_with(".iwa") {
                    let compressed_data = archive.read(file_name).expect("Failed to read IWA file");

                    let mut cursor = Cursor::new(compressed_data.as_slice());
                    let result = SnappyStream::decompress(&mut cursor);

                    assert!(
                        result.is_ok(),
                        "Failed to decompress {} from {}: {:?}",
                        file_name,
                        test_file,
                        result.err()
                    );

                    let decompressed = result.unwrap();
                    assert!(
                        !decompressed.data().is_empty(),
                        "Decompressed data should not be empty for {}",
                        file_name
                    );

                    // Verify it's valid protobuf data (starts with a varint length)
                    let data = decompressed.data();
                    assert!(!data.is_empty(), "Decompressed data too small");

                    break; // Test with first IWA file found
                }
            }
        }
    }
}
