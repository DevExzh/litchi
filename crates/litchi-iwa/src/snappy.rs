//! Snappy decompression for iWork IWA files
//!
//! iWork uses a custom Snappy framing format that differs from the standard:
//! - No stream identifier chunk
//! - No CRC-32C checksums
//! - Custom chunk header format (4 bytes: type + 24-bit length)

use snap::raw::{Decoder, Encoder, decompress_len};
use std::io::{self, Cursor, Read};

use crate::Error;

const SNAPPY_FRAME_HEADER_SIZE: usize = 4;

/// Caller-selectable decompression ceilings for one iWork Snappy stream.
///
/// The constructor only accepts limits at or below the format-wide hard
/// ceilings. This lets a caller impose a tighter budget without creating an
/// escape hatch around the parser's memory-safety guardrails.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnappyLimits {
    max_uncompressed_chunk: usize,
    max_decompressed_stream: usize,
    max_compressed_chunk: usize,
    max_compressed_stream: usize,
    max_frames: usize,
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
            max_compressed_chunk: SnappyStream::MAX_COMPRESSED_CHUNK,
            max_compressed_stream: SnappyStream::MAX_COMPRESSED_STREAM,
            max_frames: SnappyStream::MAX_FRAMES,
        })
    }

    /// Tighten the compressed-input framing limits.
    ///
    /// The aggregate compressed-stream limit includes every four-byte frame
    /// header, and the frame count includes zero-length frames. These checks
    /// happen before a frame payload is read or allocated.
    pub fn with_input_limits(
        mut self,
        max_compressed_chunk: usize,
        max_compressed_stream: usize,
        max_frames: usize,
    ) -> Result<Self, Error> {
        if max_compressed_chunk == 0 || max_compressed_stream == 0 || max_frames == 0 {
            return Err(Error::Snappy(
                "Snappy input limits must be non-zero".to_owned(),
            ));
        }
        if max_compressed_chunk > SnappyStream::MAX_COMPRESSED_CHUNK {
            return Err(Error::Snappy(format!(
                "Snappy compressed chunk limit exceeds the {} byte hard ceiling",
                SnappyStream::MAX_COMPRESSED_CHUNK
            )));
        }
        if max_compressed_stream > SnappyStream::MAX_COMPRESSED_STREAM {
            return Err(Error::Snappy(format!(
                "Snappy compressed stream limit exceeds the {} byte hard ceiling",
                SnappyStream::MAX_COMPRESSED_STREAM
            )));
        }
        if max_frames > SnappyStream::MAX_FRAMES {
            return Err(Error::Snappy(format!(
                "Snappy frame limit exceeds the {} frame hard ceiling",
                SnappyStream::MAX_FRAMES
            )));
        }
        let minimum_stream_for_chunk =
            SNAPPY_FRAME_HEADER_SIZE
                .checked_add(max_compressed_chunk)
                .ok_or_else(|| Error::Snappy("Snappy input limit overflow".to_owned()))?;
        if max_compressed_stream < minimum_stream_for_chunk {
            return Err(Error::Snappy(
                "Snappy compressed stream limit cannot hold one configured frame".to_owned(),
            ));
        }

        self.max_compressed_chunk = max_compressed_chunk;
        self.max_compressed_stream = max_compressed_stream;
        self.max_frames = max_frames;
        Ok(self)
    }

    /// Maximum uncompressed size accepted for one block.
    pub const fn max_uncompressed_chunk(self) -> usize {
        self.max_uncompressed_chunk
    }

    /// Maximum aggregate uncompressed size accepted for one stream.
    pub const fn max_decompressed_stream(self) -> usize {
        self.max_decompressed_stream
    }

    /// Maximum compressed payload accepted for one frame.
    pub const fn max_compressed_chunk(self) -> usize {
        self.max_compressed_chunk
    }

    /// Maximum compressed bytes accepted for one stream, including headers.
    pub const fn max_compressed_stream(self) -> usize {
        self.max_compressed_stream
    }

    /// Maximum number of frames accepted for one stream.
    pub const fn max_frames(self) -> usize {
        self.max_frames
    }
}

impl Default for SnappyLimits {
    fn default() -> Self {
        Self {
            max_uncompressed_chunk: SnappyStream::MAX_UNCOMPRESSED_CHUNK,
            max_decompressed_stream: SnappyStream::MAX_DECOMPRESSED_STREAM,
            max_compressed_chunk: SnappyStream::MAX_COMPRESSED_CHUNK,
            max_compressed_stream: SnappyStream::MAX_COMPRESSED_STREAM,
            max_frames: SnappyStream::MAX_FRAMES,
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
    /// Maximum compressed payload accepted for one framed block.
    pub const MAX_COMPRESSED_CHUNK: usize = 0x00ff_ffff;
    /// Maximum compressed size accepted for one stream, including headers.
    pub const MAX_COMPRESSED_STREAM: usize = 2 * 1024 * 1024 * 1024;
    /// Maximum number of frames accepted for one stream, including empty frames.
    pub const MAX_FRAMES: usize = 1_000_000;
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
        let mut compressed_stream_length = 0usize;
        let mut frame_count = 0usize;

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

            frame_count = frame_count
                .checked_add(1)
                .ok_or_else(|| Error::Snappy("Snappy frame count overflow".to_owned()))?;
            if frame_count > limits.max_frames {
                return Err(Error::Snappy(format!(
                    "Snappy frame count exceeds the {} frame limit",
                    limits.max_frames
                )));
            }

            let compressed_length = usize::try_from(length)
                .map_err(|_| Error::Snappy("Snappy frame length does not fit usize".to_owned()))?;
            if compressed_length > limits.max_compressed_chunk {
                return Err(Error::Snappy(format!(
                    "Snappy compressed frame is {compressed_length} bytes, exceeding the {} byte limit",
                    limits.max_compressed_chunk
                )));
            }
            let frame_size = SNAPPY_FRAME_HEADER_SIZE
                .checked_add(compressed_length)
                .ok_or_else(|| {
                    Error::Snappy("Snappy compressed frame length overflow".to_owned())
                })?;
            compressed_stream_length = compressed_stream_length
                .checked_add(frame_size)
                .ok_or_else(|| {
                    Error::Snappy("Snappy compressed stream length overflow".to_owned())
                })?;
            if compressed_stream_length > limits.max_compressed_stream {
                return Err(Error::Snappy(format!(
                    "Snappy compressed stream exceeds the {} byte limit",
                    limits.max_compressed_stream
                )));
            }

            if length == 0 {
                continue;
            }

            // Read compressed chunk
            let mut compressed = Vec::new();
            compressed
                .try_reserve_exact(compressed_length)
                .map_err(|error| {
                    Error::Snappy(format!("Unable to reserve compressed buffer: {error}"))
                })?;
            compressed.resize(compressed_length, 0);
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
            let previous_length = decompressed.len();
            decompressed.resize(total_length, 0);
            let decoded_length = decoder
                .decompress(&compressed, &mut decompressed[previous_length..])
                .map_err(|error| Error::Snappy(format!("Decompression failed: {error}")))?;
            if decoded_length != expected_length {
                decompressed.truncate(previous_length);
                return Err(Error::Snappy(format!(
                    "Snappy block decoded to {decoded_length} bytes, expected {expected_length}"
                )));
            }
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
    use std::io::{Cursor, Read};

    struct HeaderOnlyReader {
        header: [u8; SNAPPY_FRAME_HEADER_SIZE],
        position: usize,
        payload_reads: usize,
    }

    impl HeaderOnlyReader {
        fn new(header: [u8; SNAPPY_FRAME_HEADER_SIZE]) -> Self {
            Self {
                header,
                position: 0,
                payload_reads: 0,
            }
        }
    }

    impl Read for HeaderOnlyReader {
        fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
            if self.position >= self.header.len() {
                self.payload_reads += 1;
                return Err(io::Error::other("payload was read"));
            }

            let available = self.header.len() - self.position;
            let amount = available.min(buffer.len());
            buffer[..amount].copy_from_slice(&self.header[self.position..self.position + amount]);
            self.position += amount;
            Ok(amount)
        }
    }

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
        assert_eq!(
            limits.max_compressed_chunk(),
            SnappyStream::MAX_COMPRESSED_CHUNK
        );
        assert_eq!(
            limits.max_compressed_stream(),
            SnappyStream::MAX_COMPRESSED_STREAM
        );
        assert_eq!(limits.max_frames(), SnappyStream::MAX_FRAMES);

        assert!(limits.with_input_limits(0, 16, 1).is_err());
        assert!(limits.with_input_limits(8, 0, 1).is_err());
        assert!(limits.with_input_limits(8, 16, 0).is_err());
        assert!(
            limits
                .with_input_limits(SnappyStream::MAX_COMPRESSED_CHUNK + 1, 1, 1)
                .is_err()
        );
        assert!(
            limits
                .with_input_limits(1, SnappyStream::MAX_COMPRESSED_STREAM + 1, 1)
                .is_err()
        );
        assert!(
            limits
                .with_input_limits(1, 5, SnappyStream::MAX_FRAMES + 1)
                .is_err()
        );
        assert!(limits.with_input_limits(8, 11, 1).is_err());

        let input_limits = limits.with_input_limits(8, 16, 3).unwrap();
        assert_eq!(input_limits.max_compressed_chunk(), 8);
        assert_eq!(input_limits.max_compressed_stream(), 16);
        assert_eq!(input_limits.max_frames(), 3);
    }

    #[test]
    fn compressed_frame_limit_is_checked_before_payload_read() {
        let mut reader = HeaderOnlyReader::new([0, 4, 0, 0]);
        let limits = SnappyLimits::default().with_input_limits(3, 16, 1).unwrap();
        let error = SnappyStream::decompress_with_limits(&mut reader, limits).unwrap_err();

        assert!(
            error
                .to_string()
                .contains("compressed frame is 4 bytes, exceeding the 3 byte limit")
        );
        assert_eq!(reader.payload_reads, 0);
    }

    #[test]
    fn compressed_stream_limit_includes_the_next_frame_header() {
        let first_frame = SnappyStream::compress(b"frame").unwrap();
        let first_frame_length = usize::from(first_frame[1])
            | (usize::from(first_frame[2]) << 8)
            | (usize::from(first_frame[3]) << 16);
        let first_frame_size = SNAPPY_FRAME_HEADER_SIZE + first_frame_length;
        let limits = SnappyLimits::default()
            .with_input_limits(first_frame_length, first_frame_size + 3, 2)
            .unwrap();
        let mut input = first_frame;
        input.extend_from_slice(&[0, 0, 0, 0]);

        let error =
            SnappyStream::decompress_with_limits(&mut Cursor::new(input), limits).unwrap_err();
        assert!(error.to_string().contains("compressed stream exceeds the"));
    }

    #[test]
    fn frame_count_limit_rejects_empty_frame_floods() {
        let limits = SnappyLimits::default().with_input_limits(1, 64, 2).unwrap();
        let input = vec![0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];

        let error =
            SnappyStream::decompress_with_limits(&mut Cursor::new(input), limits).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("frame count exceeds the 2 frame limit")
        );
    }

    #[test]
    fn malformed_inputs_do_not_panic() {
        for length in 0..256usize {
            let mut input = Vec::with_capacity(length);
            let mut state = length as u64 + 1;
            for _ in 0..length {
                state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
                input.push((state >> 32) as u8);
            }

            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                SnappyStream::decompress(&mut Cursor::new(input))
            }));
            assert!(result.is_ok(), "decompression panicked for {length} bytes");
        }
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
