use snap::raw::{Decoder, Encoder, decompress_len};

use crate::error::{Error, LimitKind, Result};

const FRAME_HEADER_BYTES: usize = 4;

/// Caller-selectable bounds for Apple's checksum-free IWA Snappy framing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes the resource-budget fields self-documenting."
)]
pub struct SnappyLimits {
    max_uncompressed_chunk: usize,
    max_decompressed_stream: usize,
    max_compressed_chunk: usize,
    max_compressed_stream: usize,
    max_frames: usize,
}

impl SnappyLimits {
    /// Build limits for decoded chunks and the decoded stream.
    ///
    /// # Errors
    ///
    /// Returns an error when either budget is zero, exceeds its hard ceiling,
    /// or when the chunk budget exceeds the stream budget.
    pub fn new(max_uncompressed_chunk: usize, max_decompressed_stream: usize) -> Result<Self> {
        check(
            LimitKind::SnappyChunkBytes,
            max_uncompressed_chunk,
            SnappyStream::MAX_UNCOMPRESSED_CHUNK,
        )?;
        check(
            LimitKind::SnappyStreamBytes,
            max_decompressed_stream,
            SnappyStream::MAX_DECOMPRESSED_STREAM,
        )?;
        if max_uncompressed_chunk > max_decompressed_stream {
            return Err(Error::invalid_limits(
                "Snappy chunk limit exceeds stream limit",
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

    /// Tighten compressed-input framing limits.
    ///
    /// # Errors
    ///
    /// Returns an error when a budget is zero, exceeds its hard ceiling, or
    /// cannot hold one complete configured frame.
    pub fn with_input_limits(
        mut self,
        max_compressed_chunk: usize,
        max_compressed_stream: usize,
        max_frames: usize,
    ) -> Result<Self> {
        check(
            LimitKind::SnappyCompressedChunkBytes,
            max_compressed_chunk,
            SnappyStream::MAX_COMPRESSED_CHUNK,
        )?;
        check(
            LimitKind::SnappyCompressedStreamBytes,
            max_compressed_stream,
            SnappyStream::MAX_COMPRESSED_STREAM,
        )?;
        check(
            LimitKind::SnappyFrames,
            max_frames,
            SnappyStream::MAX_FRAMES,
        )?;
        let minimum_stream = FRAME_HEADER_BYTES
            .checked_add(max_compressed_chunk)
            .ok_or_else(|| Error::invalid_limits("Snappy input limit overflow"))?;
        if max_compressed_stream < minimum_stream {
            return Err(Error::invalid_limits(
                "Snappy compressed stream cannot hold one configured frame",
            ));
        }
        self.max_compressed_chunk = max_compressed_chunk;
        self.max_compressed_stream = max_compressed_stream;
        self.max_frames = max_frames;
        Ok(self)
    }

    #[must_use]
    pub const fn max_uncompressed_chunk(self) -> usize {
        self.max_uncompressed_chunk
    }

    #[must_use]
    pub const fn max_decompressed_stream(self) -> usize {
        self.max_decompressed_stream
    }

    #[must_use]
    pub const fn max_compressed_chunk(self) -> usize {
        self.max_compressed_chunk
    }

    #[must_use]
    pub const fn max_compressed_stream(self) -> usize {
        self.max_compressed_stream
    }

    #[must_use]
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

/// A decoded Apple IWA Snappy stream.
///
/// IWA uses four-byte frame headers with a one-byte compressed-chunk marker and
/// a 24-bit little-endian payload length. It deliberately omits the standard
/// Snappy stream identifier and CRC chunks. The decoded bytes are contiguous
/// so the archive layer can parse them without another copy.
#[derive(Debug)]
pub struct SnappyStream {
    bytes: Vec<u8>,
}

impl SnappyStream {
    /// Hard ceiling for one decoded frame.
    pub const MAX_UNCOMPRESSED_CHUNK: usize = 64 * 1024 * 1024;
    /// Hard ceiling for one decoded stream.
    pub const MAX_DECOMPRESSED_STREAM: usize = 512 * 1024 * 1024;
    /// Hard ceiling for one compressed frame payload.
    pub const MAX_COMPRESSED_CHUNK: usize = 0x00ff_ffff;
    /// Hard ceiling for compressed input, including frame headers.
    pub const MAX_COMPRESSED_STREAM: usize = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for the number of frames, including empty frames.
    pub const MAX_FRAMES: usize = 1_000_000;
    /// Size of independently compressed frames emitted by [`compress`].
    pub const WRITE_CHUNK_SIZE: usize = 64 * 1024;

    /// Decode one complete Apple IWA Snappy stream.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed framing, unsupported chunk types,
    /// decompression failures, or resource-limit violations.
    pub fn decompress(data: &[u8]) -> Result<Self> {
        Self::decompress_with_limits(data, SnappyLimits::default())
    }

    /// Decode one stream under caller-selected, checked resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed framing, unsupported chunk types,
    /// decompression failures, or resource-limit violations.
    pub fn decompress_with_limits(data: &[u8], limits: SnappyLimits) -> Result<Self> {
        if data.len() > limits.max_compressed_stream {
            return Err(Error::limit(
                LimitKind::SnappyCompressedStreamBytes,
                data.len(),
                limits.max_compressed_stream,
            ));
        }

        let mut decoded = Vec::new();
        let mut snappy_decoder = Decoder::new();
        let mut cursor = 0usize;
        let mut frame_count = 0usize;
        let mut compressed_bytes = 0usize;

        while cursor < data.len() {
            let header_start = cursor;
            let header_end = cursor
                .checked_add(FRAME_HEADER_BYTES)
                .ok_or_else(|| Error::snappy("Snappy frame header offset overflow"))?;
            let header = data.get(cursor..header_end).ok_or_else(|| {
                Error::invalid_archive(header_start, "truncated Snappy frame header")
            })?;
            cursor = header_end;

            frame_count = frame_count
                .checked_add(1)
                .ok_or_else(|| Error::snappy("Snappy frame count overflow"))?;
            if frame_count > limits.max_frames {
                return Err(Error::limit(
                    LimitKind::SnappyFrames,
                    frame_count,
                    limits.max_frames,
                ));
            }

            let compressed_length = usize::from(header[1])
                | (usize::from(header[2]) << 8)
                | (usize::from(header[3]) << 16);
            if compressed_length > limits.max_compressed_chunk {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedChunkBytes,
                    compressed_length,
                    limits.max_compressed_chunk,
                ));
            }
            let frame_size = FRAME_HEADER_BYTES
                .checked_add(compressed_length)
                .ok_or_else(|| Error::snappy("Snappy frame length overflow"))?;
            compressed_bytes = compressed_bytes
                .checked_add(frame_size)
                .ok_or_else(|| Error::snappy("Snappy stream length overflow"))?;
            if compressed_bytes > limits.max_compressed_stream {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedStreamBytes,
                    compressed_bytes,
                    limits.max_compressed_stream,
                ));
            }

            if header[0] != 0 {
                return Err(Error::invalid_archive(
                    header_start,
                    "unsupported Snappy chunk type",
                ));
            }
            let payload_end = cursor
                .checked_add(compressed_length)
                .ok_or_else(|| Error::snappy("Snappy frame payload offset overflow"))?;
            let compressed = data.get(cursor..payload_end).ok_or_else(|| {
                Error::invalid_archive(header_start, "truncated Snappy frame payload")
            })?;
            cursor = payload_end;

            if compressed.is_empty() {
                continue;
            }

            let expected_length = decompress_len(compressed)
                .map_err(|error| Error::snappy(format!("invalid Snappy block: {error}")))?;
            if expected_length > limits.max_uncompressed_chunk {
                return Err(Error::limit(
                    LimitKind::SnappyChunkBytes,
                    expected_length,
                    limits.max_uncompressed_chunk,
                ));
            }
            let decoded_length = decoded
                .len()
                .checked_add(expected_length)
                .ok_or_else(|| Error::snappy("decoded Snappy stream length overflow"))?;
            if decoded_length > limits.max_decompressed_stream {
                return Err(Error::limit(
                    LimitKind::SnappyStreamBytes,
                    decoded_length,
                    limits.max_decompressed_stream,
                ));
            }
            decoded
                .try_reserve(expected_length)
                .map_err(|_allocation_error| {
                    Error::allocation("decoded Snappy stream", expected_length)
                })?;
            let previous_length = decoded.len();
            decoded.resize(decoded_length, 0);
            let actual_length = snappy_decoder
                .decompress(compressed, &mut decoded[previous_length..])
                .map_err(|error| Error::snappy(format!("Snappy decompression failed: {error}")))?;
            if actual_length != expected_length {
                decoded.truncate(previous_length);
                return Err(Error::snappy(format!(
                    "Snappy block decoded to {actual_length} bytes, expected {expected_length}"
                )));
            }
        }

        Ok(Self { bytes: decoded })
    }

    /// Encode bytes using Apple's checksum-free IWA Snappy framing.
    ///
    /// # Errors
    ///
    /// Returns an error when the input or encoded stream exceeds a hard
    /// ceiling, compression fails, or a checked allocation cannot be made.
    pub fn compress(data: &[u8]) -> Result<Vec<u8>> {
        if data.len() > Self::MAX_DECOMPRESSED_STREAM {
            return Err(Error::limit(
                LimitKind::SnappyStreamBytes,
                data.len(),
                Self::MAX_DECOMPRESSED_STREAM,
            ));
        }

        let mut output = Vec::new();
        let mut encoder = Encoder::new();
        for chunk in data.chunks(Self::WRITE_CHUNK_SIZE) {
            let compressed = encoder
                .compress_vec(chunk)
                .map_err(|error| Error::snappy(format!("Snappy compression failed: {error}")))?;
            if compressed.len() > Self::MAX_COMPRESSED_CHUNK {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedChunkBytes,
                    compressed.len(),
                    Self::MAX_COMPRESSED_CHUNK,
                ));
            }
            let frame_size = FRAME_HEADER_BYTES
                .checked_add(compressed.len())
                .ok_or_else(|| Error::snappy("Snappy frame length overflow"))?;
            let output_length = output
                .len()
                .checked_add(frame_size)
                .ok_or_else(|| Error::snappy("Snappy output length overflow"))?;
            if output_length > Self::MAX_COMPRESSED_STREAM {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedStreamBytes,
                    output_length,
                    Self::MAX_COMPRESSED_STREAM,
                ));
            }
            output
                .try_reserve(frame_size)
                .map_err(|_allocation_error| {
                    Error::allocation("compressed Snappy stream", frame_size)
                })?;
            let compressed_length =
                u32::try_from(compressed.len()).map_err(|_conversion_error| {
                    Error::snappy("Snappy compressed frame length exceeds u32")
                })?;
            let compressed_length_bytes = compressed_length.to_le_bytes();
            output.extend_from_slice(&[
                0,
                compressed_length_bytes[0],
                compressed_length_bytes[1],
                compressed_length_bytes[2],
            ]);
            output.extend_from_slice(&compressed);
        }
        Ok(output)
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

impl AsRef<[u8]> for SnappyStream {
    fn as_ref(&self) -> &[u8] {
        self.as_bytes()
    }
}

fn check(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 {
        return Err(Error::invalid_limits("Snappy limits must be non-zero"));
    }
    if value > maximum {
        return Err(Error::limit(kind, value, maximum));
    }
    Ok(())
}
