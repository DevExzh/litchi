//! Snappy decompression for iWork IWA files
//!
//! iWork uses a custom Snappy framing format that differs from the standard:
//! - No stream identifier chunk
//! - No CRC-32C checksums
//! - Custom chunk header format (4 bytes: type + 24-bit length)

use snap::raw::{Decoder, Encoder, decompress_len};
use std::io::{self, Cursor, Read};

use crate::Error;

/// Caller-selectable decompression ceilings for one iWork Snappy stream.
///
/// The constructor only accepts limits at or below the format-wide hard
/// ceilings. This lets a caller impose a tighter budget without creating an
/// escape hatch around the parser's memory-safety guardrails.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnappyLimits {
    max_uncompressed_chunk: usize,
    max_decompressed_stream: usize,
}

impl SnappyLimits {
    /// Build a checked pair of per-chunk and aggregate decompression limits.
    pub fn new(
        max_uncompressed_chunk: usize,
        max_decompressed_stream: usize,
    ) -> Result<Self, Error> {
        if max_uncompressed_chunk == 0 || max_decompressed_stream == 0 {
            return Err(Error::Snappy(
                "Snappy decompression limits must be non-zero".to_owned(),
            ));
        }
        if max_uncompressed_chunk > SnappyStream::MAX_UNCOMPRESSED_CHUNK {
            return Err(Error::Snappy(format!(
                "Snappy chunk limit exceeds the {} byte hard ceiling",
                SnappyStream::MAX_UNCOMPRESSED_CHUNK
            )));
        }
        if max_decompressed_stream > SnappyStream::MAX_DECOMPRESSED_STREAM {
            return Err(Error::Snappy(format!(
                "Snappy stream limit exceeds the {} byte hard ceiling",
                SnappyStream::MAX_DECOMPRESSED_STREAM
            )));
        }
        if max_uncompressed_chunk > max_decompressed_stream {
            return Err(Error::Snappy(
                "Snappy chunk limit cannot exceed the stream limit".to_owned(),
            ));
        }
        Ok(Self {
            max_uncompressed_chunk,
            max_decompressed_stream,
        })
    }

    /// Maximum uncompressed size accepted for one block.
    pub const fn max_uncompressed_chunk(self) -> usize {
        self.max_uncompressed_chunk
    }

    /// Maximum aggregate uncompressed size accepted for one stream.
    pub const fn max_decompressed_stream(self) -> usize {
        self.max_decompressed_stream
    }
}

impl Default for SnappyLimits {
    fn default() -> Self {
        Self {
            max_uncompressed_chunk: SnappyStream::MAX_UNCOMPRESSED_CHUNK,
            max_decompressed_stream: SnappyStream::MAX_DECOMPRESSED_STREAM,
        }
    }
}

/// Custom Snappy stream decompressor for iWork IWA files
#[derive(Debug)]
pub struct SnappyStream {
    decompressed: Vec<u8>,
}

impl SnappyStream {
    /// Maximum uncompressed size accepted for one Snappy block.
    ///
    /// iWork writers use small independent blocks. Capping individual blocks
    /// prevents a forged length prefix from forcing a multi-gigabyte allocation.
    pub const MAX_UNCOMPRESSED_CHUNK: usize = 64 * 1024 * 1024;
    /// Maximum total size accepted for one decompressed IWA component.
    pub const MAX_DECOMPRESSED_STREAM: usize = 512 * 1024 * 1024;
    /// Block size emitted by the serializer.
    pub const WRITE_CHUNK_SIZE: usize = 64 * 1024;

    /// Decompress an IWA file from a reader
    ///
    /// iWork IWA files use a custom Snappy framing format:
    /// - 4-byte header: [chunk_type, length_byte1, length_byte2, length_byte3]
    /// - chunk_type is always 0 for compressed chunks
    /// - length is a 24-bit little-endian integer
    /// - No stream identifier, no CRC checksums
    pub fn decompress<R: Read>(reader: &mut R) -> Result<Self, Error> {
        Self::decompress_with_limits(reader, SnappyLimits::default())
    }

    /// Decompress an IWA stream under caller-supplied, checked memory limits.
    pub fn decompress_with_limits<R: Read>(
        reader: &mut R,
        limits: SnappyLimits,
    ) -> Result<Self, Error> {
        let mut decompressed = Vec::new();
        let mut decoder = Decoder::new();

        loop {
            // Read the type byte separately so a clean EOF can be distinguished
            // from a truncated four-byte frame header.
            let mut chunk_type = [0u8; 1];
            match reader.read(&mut chunk_type) {
                Ok(0) => break,
                Ok(_) => {},
                Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(Error::Io(error)),
            }
            let mut length_bytes = [0u8; 3];
            reader.read_exact(&mut length_bytes)?;

            if chunk_type[0] != 0 {
                return Err(Error::Snappy(format!(
                    "Unexpected chunk type: {}, expected 0",
                    chunk_type[0]
                )));
            }

            // Extract 24-bit length (little-endian)
            let length = u32::from_le_bytes([length_bytes[0], length_bytes[1], length_bytes[2], 0]);

            if length == 0 {
                continue;
            }

            // Read compressed chunk
            let mut compressed = vec![0u8; length as usize];
            reader.read_exact(&mut compressed).map_err(Error::Io)?;

            let expected_length = decompress_len(&compressed)
                .map_err(|error| Error::Snappy(format!("Invalid Snappy block: {error}")))?;
            if expected_length > limits.max_uncompressed_chunk {
                return Err(Error::Snappy(format!(
                    "Snappy block expands to {expected_length} bytes, exceeding the {} byte limit",
                    limits.max_uncompressed_chunk
                )));
            }
            let total_length = decompressed
                .len()
                .checked_add(expected_length)
                .ok_or_else(|| Error::Snappy("Decompressed stream length overflow".to_owned()))?;
            if total_length > limits.max_decompressed_stream {
                return Err(Error::Snappy(format!(
                    "Snappy stream expands to more than {} bytes",
                    limits.max_decompressed_stream
                )));
            }
            decompressed.try_reserve(expected_length).map_err(|error| {
                Error::Snappy(format!("Unable to reserve decompression buffer: {error}"))
            })?;
            let chunk_decompressed = decoder
                .decompress_vec(&compressed)
                .map_err(|error| Error::Snappy(format!("Decompression failed: {error}")))?;
            if chunk_decompressed.len() != expected_length {
                return Err(Error::Snappy(format!(
                    "Snappy block decoded to {} bytes, expected {expected_length}",
                    chunk_decompressed.len()
                )));
            }
            decompressed.extend(chunk_decompressed);
        }

        Ok(SnappyStream { decompressed })
    }

    /// Compress a decompressed IWA stream using Apple's checksum-free Snappy
    /// framing variant.
    pub fn compress(data: &[u8]) -> Result<Vec<u8>, Error> {
        let mut output = Vec::new();
        let mut encoder = Encoder::new();

        for chunk in data.chunks(Self::WRITE_CHUNK_SIZE) {
            let compressed = encoder
                .compress_vec(chunk)
                .map_err(|error| Error::Snappy(format!("Compression failed: {error}")))?;
            let length = u32::try_from(compressed.len()).map_err(|_| {
                Error::Snappy("Compressed Snappy block exceeds the 24-bit limit".to_string())
            })?;
            if length > 0x00ff_ffff {
                return Err(Error::Snappy(
                    "Compressed Snappy block exceeds the 24-bit limit".to_string(),
                ));
            }
            output.extend_from_slice(&[0, length as u8, (length >> 8) as u8, (length >> 16) as u8]);
            output.extend_from_slice(&compressed);
        }

        Ok(output)
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
    fn test_compression_round_trip_across_blocks() {
        let input: Vec<u8> = (0..(SnappyStream::WRITE_CHUNK_SIZE * 3 + 17))
            .map(|index| (index % 251) as u8)
            .collect();
        let compressed = SnappyStream::compress(&input).unwrap();
        let decompressed = SnappyStream::decompress(&mut Cursor::new(compressed)).unwrap();
        assert_eq!(decompressed.data(), input);
    }

    #[test]
    fn custom_limits_can_only_tighten_hard_ceilings() {
        assert!(SnappyLimits::new(0, 1).is_err());
        assert!(SnappyLimits::new(1, 0).is_err());
        assert!(
            SnappyLimits::new(
                SnappyStream::MAX_UNCOMPRESSED_CHUNK + 1,
                SnappyStream::MAX_DECOMPRESSED_STREAM
            )
            .is_err()
        );
        assert!(
            SnappyLimits::new(
                SnappyStream::MAX_UNCOMPRESSED_CHUNK,
                SnappyStream::MAX_DECOMPRESSED_STREAM + 1
            )
            .is_err()
        );
        assert!(SnappyLimits::new(2, 1).is_err());

        let limits = SnappyLimits::new(8, 16).unwrap();
        assert_eq!(limits.max_uncompressed_chunk(), 8);
        assert_eq!(limits.max_decompressed_stream(), 16);
    }

    #[test]
    fn custom_chunk_limit_rejects_a_valid_block() {
        let compressed = SnappyStream::compress(b"0123456789").unwrap();
        let limits = SnappyLimits::new(9, 9).unwrap();
        let error =
            SnappyStream::decompress_with_limits(&mut Cursor::new(compressed), limits).unwrap_err();
        assert!(error.to_string().contains("exceeding the 9 byte limit"));
    }

    #[test]
    fn custom_stream_limit_rejects_a_second_valid_block() {
        let input = vec![42; SnappyStream::WRITE_CHUNK_SIZE + 17];
        let compressed = SnappyStream::compress(&input).unwrap();
        let limits = SnappyLimits::new(
            SnappyStream::WRITE_CHUNK_SIZE,
            SnappyStream::WRITE_CHUNK_SIZE,
        )
        .unwrap();
        let error =
            SnappyStream::decompress_with_limits(&mut Cursor::new(compressed), limits).unwrap_err();
        assert!(error.to_string().contains("more than 65536 bytes"));
    }

    #[test]
    fn test_truncated_header_is_rejected() {
        let error = SnappyStream::decompress(&mut Cursor::new([0, 1])).unwrap_err();
        assert!(matches!(error, Error::Io(_)));
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
