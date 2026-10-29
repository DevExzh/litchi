use std::{fmt, mem::size_of};

use snap::raw::{Decoder, Encoder, decompress_len, max_compress_len};

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

    /// Set the total compressed-byte ceiling without requiring room for a
    /// maximum-sized frame. An empty stream remains valid under any ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when the ceiling is zero or exceeds the hard maximum.
    pub fn with_compressed_stream_bytes(mut self, maximum: usize) -> Result<Self> {
        check(
            LimitKind::SnappyCompressedStreamBytes,
            maximum,
            SnappyStream::MAX_COMPRESSED_STREAM,
        )?;
        self.max_compressed_stream = maximum;
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

/// Exact input-independent requirements for one Snappy compression route.
///
/// `output_capacity_bytes` is the checked upper bound already used by the
/// encoder.  `temporary_frame_capacity_bytes` is a source-derived upper
/// bound for the fallback `Encoder::compress_vec` buffers.  The latter is
/// deliberately reported as a bound: whether a particular frame takes the
/// fallback path depends on the compressed length of earlier frames.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnappyCompressionRequirements {
    input_bytes: usize,
    output_capacity_bytes: usize,
    frame_count: usize,
    temporary_frame_capacity_bytes: usize,
    encoder_workspace_bytes: usize,
    encoder_table_reset_work_bytes: usize,
    work_bytes: usize,
    allocation_events: usize,
}

/// One allocation-free traversal admitted while preparing Snappy geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnappyCompressionPreparationChunk {
    kind: SnappyCompressionPreparationKind,
    elements: usize,
    traversal_bytes: usize,
}

impl SnappyCompressionPreparationChunk {
    #[must_use]
    pub const fn kind(self) -> SnappyCompressionPreparationKind {
        self.kind
    }

    #[must_use]
    pub const fn elements(self) -> usize {
        self.elements
    }

    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }
}

/// Geometry phase admitted before the Snappy backend is inspected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SnappyCompressionPreparationKind {
    StreamBounds,
    FrameGeometry,
}

/// Typed visitor for Snappy preparation work admission.
pub trait SnappyCompressionPreparationVisitor {
    type Error;

    fn admit(
        &mut self,
        chunk: SnappyCompressionPreparationChunk,
    ) -> std::result::Result<(), Self::Error>;
}

impl<F, E> SnappyCompressionPreparationVisitor for F
where
    F: FnMut(SnappyCompressionPreparationChunk) -> std::result::Result<(), E>,
{
    type Error = E;

    fn admit(&mut self, chunk: SnappyCompressionPreparationChunk) -> std::result::Result<(), E> {
        self(chunk)
    }
}

/// Failure from Snappy preparation work admission.
#[derive(Debug)]
#[non_exhaustive]
pub enum SnappyCompressionPreparationError<E> {
    Admission(E),
    ArithmeticOverflow { resource: &'static str },
    Core(Error),
}

impl<E> From<Error> for SnappyCompressionPreparationError<E> {
    fn from(error: Error) -> Self {
        Self::Core(error)
    }
}

pub type SnappyCompressionPreparationResult<T, E> =
    std::result::Result<T, SnappyCompressionPreparationError<E>>;

impl<E: fmt::Display> fmt::Display for SnappyCompressionPreparationError<E> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Admission(error) => write!(formatter, "Snappy admission failed: {error}"),
            Self::ArithmeticOverflow { resource } => {
                write!(formatter, "Snappy preparation {resource} overflow")
            },
            Self::Core(error) => write!(formatter, "Snappy preparation failed: {error}"),
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for SnappyCompressionPreparationError<E> {}

impl SnappyCompressionRequirements {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn output_capacity_bytes(self) -> usize {
        self.output_capacity_bytes
    }

    #[must_use]
    pub const fn frame_count(self) -> usize {
        self.frame_count
    }

    /// Conservative total capacity of possible temporary raw Snappy frames.
    /// The encoder may use less when earlier frames leave enough room in the
    /// reserved output buffer.
    #[must_use]
    pub const fn temporary_frame_capacity_bytes(self) -> usize {
        self.temporary_frame_capacity_bytes
    }

    /// Heap state allocated by the raw Snappy encoder when a chunk exceeds
    /// its small stack table.  The encoder reuses this table across frames.
    #[must_use]
    pub const fn encoder_workspace_bytes(self) -> usize {
        self.encoder_workspace_bytes
    }

    /// Bytes written while resetting the raw encoder's reusable hash table
    /// across frames. This is work, not an additional retained allocation.
    #[must_use]
    pub const fn encoder_table_reset_work_bytes(self) -> usize {
        self.encoder_table_reset_work_bytes
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// One output allocation plus one possible temporary allocation per frame.
    /// This is an upper bound for the actual number of allocations.
    #[must_use]
    pub const fn allocation_events(self) -> usize {
        self.allocation_events
    }
}

/// Borrow-free prepared Snappy compression admission.
///
/// The plan binds the input length and limit profile.  `execute` therefore
/// cannot accidentally replay a plan against a different length or a
/// different budget, while the source bytes remain borrowed only for the
/// actual encoding pass.  It binds the length, not the identity or contents
/// of the borrowed slice.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SnappyCompressionPlan {
    input_len: usize,
    limits: SnappyLimits,
    requirements: SnappyCompressionRequirements,
}

impl SnappyCompressionPlan {
    #[must_use]
    pub const fn requirements(self) -> SnappyCompressionRequirements {
        self.requirements
    }

    /// Execute the admitted compression against the exact input length used
    /// during preparation.
    pub fn execute(self, data: &[u8]) -> Result<Vec<u8>> {
        if data.len() != self.input_len {
            return Err(Error::invalid_archive(
                0,
                "prepared Snappy input length changed",
            ));
        }
        let maximum_output = SnappyStream::maximum_compressed_len(data.len(), self.limits)?;
        debug_assert_eq!(maximum_output, self.requirements.output_capacity_bytes);
        SnappyStream::compress_inner(data, self.limits, maximum_output)
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
    /// Size of independently compressed frames emitted by [`Self::compress`].
    pub const WRITE_CHUNK_SIZE: usize = 64 * 1024;

    /// Return a checked upper bound for [`Self::compress`] output.
    ///
    /// The bound includes one four-byte IWA frame header per independently
    /// compressed write chunk and uses Snappy's conservative raw-block bound
    /// for each chunk, bounded by the supplied encoded-size ceilings. No
    /// input or output buffer is allocated. The returned
    /// value is guaranteed to be at least the size of every successful
    /// [`Self::compress`] result for the same input length.
    pub fn maximum_compressed_len(input_len: usize, limits: SnappyLimits) -> Result<usize> {
        if input_len > limits.max_decompressed_stream {
            return Err(Error::limit(
                LimitKind::SnappyStreamBytes,
                input_len,
                limits.max_decompressed_stream,
            ));
        }

        let chunk_size = Self::WRITE_CHUNK_SIZE.min(limits.max_uncompressed_chunk);
        // Keep the ordinary 64-KiB path in constant-divisor arithmetic. A
        // smaller caller-selected chunk still uses its exact frame geometry.
        let (full_frames, remainder) = if limits.max_uncompressed_chunk >= Self::WRITE_CHUNK_SIZE {
            (
                input_len / Self::WRITE_CHUNK_SIZE,
                input_len % Self::WRITE_CHUNK_SIZE,
            )
        } else {
            (input_len / chunk_size, input_len % chunk_size)
        };
        let frame_count = full_frames + usize::from(remainder != 0);
        if frame_count > limits.max_frames {
            return Err(Error::limit(
                LimitKind::SnappyFrames,
                frame_count,
                limits.max_frames,
            ));
        }

        let mut maximum = if full_frames == 0 {
            0
        } else {
            let full_frame_length = FRAME_HEADER_BYTES
                .checked_add(max_compress_len(chunk_size).min(limits.max_compressed_chunk))
                .ok_or_else(|| Error::snappy("Snappy maximum frame length overflow"))?;
            full_frames
                .checked_mul(full_frame_length)
                .ok_or_else(|| Error::snappy("Snappy maximum stream length overflow"))?
        };
        if remainder != 0 {
            let final_frame_length = FRAME_HEADER_BYTES
                .checked_add(max_compress_len(remainder).min(limits.max_compressed_chunk))
                .ok_or_else(|| Error::snappy("Snappy maximum frame length overflow"))?;
            maximum = maximum
                .checked_add(final_frame_length)
                .ok_or_else(|| Error::snappy("Snappy maximum stream length overflow"))?;
        }
        Ok(maximum.min(limits.max_compressed_stream))
    }

    /// Prepare a bounded compression route without allocating or retaining
    /// any source bytes. `admit_work` is called with one phase unit before
    /// frame-bound computation and once before each frame's backend geometry
    /// calculation. The returned requirements describe execution-only work.
    pub fn prepare_compression<V>(
        input_len: usize,
        limits: SnappyLimits,
        visitor: &mut V,
    ) -> SnappyCompressionPreparationResult<SnappyCompressionPlan, V::Error>
    where
        V: SnappyCompressionPreparationVisitor,
    {
        visitor
            .admit(SnappyCompressionPreparationChunk {
                kind: SnappyCompressionPreparationKind::StreamBounds,
                elements: usize::from(input_len != 0),
                traversal_bytes: input_len,
            })
            .map_err(SnappyCompressionPreparationError::Admission)?;
        let output_capacity_bytes = Self::maximum_compressed_len(input_len, limits)?;
        let chunk_size = Self::WRITE_CHUNK_SIZE.min(limits.max_uncompressed_chunk);
        let frame_count = if input_len == 0 {
            0
        } else {
            input_len
                .checked_add(chunk_size - 1)
                .ok_or_else(|| Error::snappy("Snappy frame count overflow"))?
                / chunk_size
        };
        let mut temporary_frame_capacity_bytes = 0usize;
        let mut encoder_table_reset_work_bytes = 0usize;
        let mut remaining = input_len;
        let mut uses_encoder_workspace = false;
        while remaining != 0 {
            let chunk_len = remaining.min(chunk_size);
            visitor
                .admit(SnappyCompressionPreparationChunk {
                    kind: SnappyCompressionPreparationKind::FrameGeometry,
                    elements: 1,
                    traversal_bytes: chunk_len,
                })
                .map_err(SnappyCompressionPreparationError::Admission)?;
            if chunk_len >= 17 {
                let mut table_size = 256usize;
                while table_size < (1usize << 14) && table_size < chunk_len {
                    table_size *= 2;
                }
                uses_encoder_workspace |= table_size > 1_024;
                encoder_table_reset_work_bytes = encoder_table_reset_work_bytes
                    .checked_add(table_size * size_of::<u16>())
                    .ok_or_else(|| Error::snappy("Snappy table reset work overflow"))?;
            }
            temporary_frame_capacity_bytes = temporary_frame_capacity_bytes
                .checked_add(max_compress_len(chunk_len))
                .ok_or_else(|| Error::snappy("Snappy temporary capacity overflow"))?;
            remaining -= chunk_len;
        }
        let encoder_workspace_bytes = if uses_encoder_workspace {
            // snap::raw::Encoder uses one 2^14-entry u16 hash table for a
            // block larger than its 2^10-entry stack table.
            (1usize << 14) * size_of::<u16>()
        } else {
            0
        };
        let work_bytes = input_len
            .checked_add(output_capacity_bytes)
            .and_then(|value| value.checked_add(temporary_frame_capacity_bytes))
            .and_then(|value| value.checked_add(encoder_workspace_bytes))
            .and_then(|value| value.checked_add(encoder_table_reset_work_bytes))
            .ok_or_else(|| Error::snappy("Snappy compression work overflow"))?;
        let allocation_events = frame_count
            .checked_add(usize::from(encoder_workspace_bytes != 0))
            .and_then(|value| value.checked_add(usize::from(output_capacity_bytes != 0)))
            .ok_or_else(|| Error::snappy("Snappy allocation count overflow"))?;
        Ok(SnappyCompressionPlan {
            input_len,
            limits,
            requirements: SnappyCompressionRequirements {
                input_bytes: input_len,
                output_capacity_bytes,
                frame_count,
                temporary_frame_capacity_bytes,
                encoder_workspace_bytes,
                encoder_table_reset_work_bytes,
                work_bytes,
                allocation_events,
            },
        })
    }

    /// Return the exact decoded length of one Apple IWA Snappy stream without
    /// allocating its output buffer.
    ///
    /// The scan applies the same framing, compressed-size, frame-count, and
    /// decoded-chunk ceilings as [`Self::decompress`]. The block
    /// bytes are length-validated but not copied, so a caller can charge an
    /// aggregate ledger before retaining the decoded stream.
    pub fn decompressed_len(data: &[u8], limits: SnappyLimits) -> Result<usize> {
        if data.len() > limits.max_compressed_stream {
            return Err(Error::limit(
                LimitKind::SnappyCompressedStreamBytes,
                data.len(),
                limits.max_compressed_stream,
            ));
        }

        let mut cursor = 0usize;
        let mut frame_count = 0usize;
        let mut decoded_length = 0usize;
        while cursor < data.len() {
            let header_start = cursor;
            // Private, checked limits cap the input at 2 GiB. Adding the
            // four-byte header fits usize even on a 32-bit target.
            let header_end = cursor + FRAME_HEADER_BYTES;
            let header = data.get(cursor..header_end).ok_or_else(|| {
                Error::invalid_archive(header_start, "truncated Snappy frame header")
            })?;
            cursor = header_end;

            // The previous iteration admitted at most MAX_FRAMES (1,000,000).
            frame_count += 1;
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
            // cursor follows an admitted header within the 2-GiB source;
            // this 24-bit length can add at most 16 MiB. The sum fits usize
            // and is also the complete encoded byte count through this frame.
            let payload_end = cursor + compressed_length;
            if payload_end > limits.max_compressed_stream {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedStreamBytes,
                    payload_end,
                    limits.max_compressed_stream,
                ));
            }
            if header[0] != 0 {
                return Err(Error::invalid_archive(
                    header_start,
                    "unsupported Snappy chunk type",
                ));
            }
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
            // The admitted prefix is at most 512 MiB and the admitted next
            // frame at most 64 MiB, so their sum fits even a 32-bit usize.
            decoded_length += expected_length;
            if decoded_length > limits.max_decompressed_stream {
                return Err(Error::limit(
                    LimitKind::SnappyStreamBytes,
                    decoded_length,
                    limits.max_decompressed_stream,
                ));
            }
        }
        Ok(decoded_length)
    }

    /// Decode one stream under caller-selected, checked resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed framing, unsupported chunk types,
    /// decompression failures, or resource-limit violations.
    pub fn decompress(data: &[u8], limits: SnappyLimits) -> Result<Self> {
        let decoded_length = Self::decompressed_len(data, limits)?;
        let mut decoded = Vec::new();
        decoded
            .try_reserve_exact(decoded_length)
            .map_err(|_allocation_error| {
                Error::allocation("decoded Snappy stream", decoded_length)
            })?;
        let mut snappy_decoder = Decoder::new();
        let mut cursor = 0usize;

        // The allocation-free preflight above has admitted every frame range,
        // marker and resource bound, including the exact total output size.
        // The borrowed input cannot change between that scan and decoding.
        // Only raw Snappy payload validity remains to be checked here.
        while cursor < data.len() {
            let header = &data[cursor..cursor + FRAME_HEADER_BYTES];
            cursor += FRAME_HEADER_BYTES;
            let compressed_length = usize::from(header[1])
                | (usize::from(header[2]) << 8)
                | (usize::from(header[3]) << 16);
            let payload_end = cursor + compressed_length;
            let compressed = &data[cursor..payload_end];
            cursor = payload_end;
            if compressed.is_empty() {
                continue;
            }

            let expected_length = decompress_len(compressed)
                .map_err(|error| Error::snappy(format!("invalid Snappy block: {error}")))?;
            let previous_length = decoded.len();
            // Initialize each frame immediately before decoding, preserving
            // locality without another capacity reservation or limit scan.
            decoded.resize(previous_length + expected_length, 0);
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
    /// Returns an error when the input or encoded stream exceeds a configured
    /// ceiling, compression fails, or a checked allocation cannot be made.
    pub fn compress(data: &[u8], limits: SnappyLimits) -> Result<Vec<u8>> {
        let maximum_output = Self::maximum_compressed_len(data.len(), limits)?;
        Self::compress_inner(data, limits, maximum_output)
    }

    fn compress_inner(data: &[u8], limits: SnappyLimits, maximum_output: usize) -> Result<Vec<u8>> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(maximum_output)
            .map_err(|_allocation_error| {
                Error::allocation("compressed Snappy stream", maximum_output)
            })?;
        let mut encoder = Encoder::new();
        let chunk_size = Self::WRITE_CHUNK_SIZE.min(limits.max_uncompressed_chunk);
        for chunk in data.chunks(chunk_size) {
            let previous_length = output.len();
            // An admitted prefix is at most 2 GiB; the raw bound of a
            // write chunk (at most 64 KiB) keeps both sums within 32 bits.
            let payload_start = previous_length + FRAME_HEADER_BYTES;
            let scratch_end = payload_start + max_compress_len(chunk.len());
            // Use the reserved output as the raw encoder's workspace when
            // its complete conservative frame bound fits. Tight encoded
            // ceilings can admit a compressible result smaller than that
            // workspace, so retain a temporary frame for that case.
            let temporary = if scratch_end <= maximum_output {
                output.resize(scratch_end, 0);
                None
            } else {
                Some(encoder.compress_vec(chunk).map_err(|error| {
                    Error::snappy(format!("Snappy compression failed: {error}"))
                })?)
            };
            let compressed_length = match &temporary {
                Some(compressed) => compressed.len(),
                None => encoder
                    .compress(chunk, &mut output[payload_start..])
                    .map_err(|error| {
                        Error::snappy(format!("Snappy compression failed: {error}"))
                    })?,
            };
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
            let output_length = previous_length
                .checked_add(frame_size)
                .ok_or_else(|| Error::snappy("Snappy output length overflow"))?;
            if output_length > limits.max_compressed_stream {
                return Err(Error::limit(
                    LimitKind::SnappyCompressedStreamBytes,
                    output_length,
                    limits.max_compressed_stream,
                ));
            }
            let compressed_length =
                u32::try_from(compressed_length).map_err(|_conversion_error| {
                    Error::snappy("Snappy compressed frame length exceeds u32")
                })?;
            let compressed_length_bytes = compressed_length.to_le_bytes();
            let header = [
                0,
                compressed_length_bytes[0],
                compressed_length_bytes[1],
                compressed_length_bytes[2],
            ];
            if let Some(compressed) = temporary {
                // Admission above and the complete output bound prove that
                // appending this accepted frame cannot allocate again.
                output.extend_from_slice(&header);
                output.extend_from_slice(&compressed);
            } else {
                output.truncate(output_length);
                output[previous_length..payload_start].copy_from_slice(&header);
            }
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

#[cfg(test)]
mod tests {
    use snap::raw::max_compress_len;

    use super::{
        FRAME_HEADER_BYTES, SnappyCompressionPlan, SnappyCompressionPreparationChunk,
        SnappyCompressionPreparationError, SnappyStream,
    };
    use crate::{Error, LimitKind, SnappyLimits};

    fn prepared_compression(
        input_len: usize,
        limits: SnappyLimits,
    ) -> crate::Result<SnappyCompressionPlan> {
        let mut visitor = |_chunk: SnappyCompressionPreparationChunk| Ok::<(), Error>(());
        SnappyStream::prepare_compression(input_len, limits, &mut visitor).map_err(|error| {
            match error {
                SnappyCompressionPreparationError::Core(error)
                | SnappyCompressionPreparationError::Admission(error) => error,
                SnappyCompressionPreparationError::ArithmeticOverflow { resource } => {
                    Error::invalid_archive(0, resource)
                },
            }
        })
    }

    #[test]
    fn admitted_frame_still_validates_raw_snappy_payload() -> crate::Result<()> {
        // The frame and declared decoded length are valid, but a copy with
        // offset zero cannot produce the declared byte.
        let malformed = [0, 3, 0, 0, 1, 1, 0];
        assert_eq!(
            SnappyStream::decompressed_len(&malformed, SnappyLimits::default())?,
            1
        );
        assert!(matches!(
            SnappyStream::decompress(&malformed, SnappyLimits::default()),
            Err(Error::Snappy { .. })
        ));
        Ok(())
    }

    #[test]
    fn frame_admission_keeps_aggregate_counts_and_error_precedence() -> crate::Result<()> {
        // One empty frame, then an unsupported marker declaring a missing
        // nine-byte payload. Resource admission precedes marker/payload errors.
        let mut source = [0, 0, 0, 0, 7, 9, 0, 0];
        let stream = SnappyLimits::default().with_compressed_stream_bytes(8)?;
        let frames = stream
            .with_input_limits(
                SnappyStream::MAX_COMPRESSED_CHUNK,
                SnappyStream::MAX_COMPRESSED_STREAM,
                1,
            )?
            .with_compressed_stream_bytes(8)?;
        let chunk = SnappyLimits::default()
            .with_input_limits(8, 12, 2)?
            .with_compressed_stream_bytes(8)?;
        for (limits, expected_kind, expected_observed, expected_maximum) in [
            (stream, LimitKind::SnappyCompressedStreamBytes, 17, 8),
            (frames, LimitKind::SnappyFrames, 2, 1),
            (chunk, LimitKind::SnappyCompressedChunkBytes, 9, 8),
        ] {
            assert!(matches!(SnappyStream::decompressed_len(&source, limits),
                Err(Error::Limit { kind, observed, maximum })
                    if (kind, observed, maximum) == (expected_kind, expected_observed, expected_maximum)));
        }
        let roomy = SnappyLimits::default().with_compressed_stream_bytes(32)?;
        assert!(matches!(
            SnappyStream::decompressed_len(&source, roomy),
            Err(Error::InvalidArchive {
                offset: 4,
                reason: "unsupported Snappy chunk type"
            })
        ));
        source[4] = 0;
        assert!(matches!(
            SnappyStream::decompressed_len(&source, roomy),
            Err(Error::InvalidArchive {
                offset: 4,
                reason: "truncated Snappy frame payload"
            })
        ));
        Ok(())
    }

    #[test]
    fn maximum_compressed_len_bounds_every_emitted_stream() -> crate::Result<()> {
        for length in [
            0,
            1,
            15,
            16,
            SnappyStream::WRITE_CHUNK_SIZE - 1,
            SnappyStream::WRITE_CHUNK_SIZE,
            SnappyStream::WRITE_CHUNK_SIZE + 1,
            SnappyStream::WRITE_CHUNK_SIZE * 2 + 17,
        ] {
            let input = (0..length)
                .map(|index| (index as u8).wrapping_mul(31))
                .collect::<Vec<_>>();
            let maximum = SnappyStream::maximum_compressed_len(length, Default::default())?;
            let compressed = SnappyStream::compress(&input, Default::default())?;
            assert!(
                compressed.len() <= maximum,
                "compressed {} bytes to {}, bound was {}",
                input.len(),
                compressed.len(),
                maximum
            );
        }
        Ok(())
    }

    #[test]
    fn direct_and_tight_output_paths_emit_identical_frames() -> crate::Result<()> {
        let input = (0..3 * 4096 + 17)
            .map(|index| (index % 31) as u8)
            .collect::<Vec<_>>();
        let limits = SnappyLimits::new(4096, input.len())?;
        let direct = SnappyStream::compress(&input, limits)?;
        assert!(direct.len() < FRAME_HEADER_BYTES + max_compress_len(4096));
        let exact = limits.with_compressed_stream_bytes(direct.len())?;
        let tight = SnappyStream::compress(&input, exact)?;
        assert_eq!(direct, tight);
        assert_eq!(SnappyStream::decompress(&tight, exact)?.as_bytes(), input);

        let too_small = limits.with_compressed_stream_bytes(direct.len() - 1)?;
        assert!(matches!(SnappyStream::compress(&input, too_small),
            Err(Error::Limit {
                kind: LimitKind::SnappyCompressedStreamBytes,
                observed,
                maximum,
            }) if observed == direct.len() && maximum == direct.len() - 1));
        Ok(())
    }

    #[test]
    fn maximum_compressed_len_matches_checked_chunk_boundaries() -> crate::Result<()> {
        for length in [
            0,
            1,
            SnappyStream::WRITE_CHUNK_SIZE - 1,
            SnappyStream::WRITE_CHUNK_SIZE,
            SnappyStream::WRITE_CHUNK_SIZE + 1,
        ] {
            let mut expected = 0usize;
            let mut remaining = length;
            while remaining != 0 {
                let chunk = remaining.min(SnappyStream::WRITE_CHUNK_SIZE);
                expected = expected
                    .checked_add(FRAME_HEADER_BYTES + max_compress_len(chunk))
                    .ok_or_else(|| Error::snappy("test bound overflow"))?;
                remaining -= chunk;
            }
            assert_eq!(
                SnappyStream::maximum_compressed_len(length, Default::default())?,
                expected
            );
        }
        Ok(())
    }

    #[test]
    fn maximum_compressed_len_enforces_hard_input_and_output_ceilings() {
        let maximum = SnappyStream::maximum_compressed_len(
            SnappyStream::MAX_DECOMPRESSED_STREAM,
            Default::default(),
        )
        .expect("hard maximum should have a finite framed bound");
        assert!(maximum <= SnappyStream::MAX_COMPRESSED_STREAM);

        let error = SnappyStream::maximum_compressed_len(
            SnappyStream::MAX_DECOMPRESSED_STREAM + 1,
            Default::default(),
        )
        .expect_err("input above the hard stream ceiling must fail");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::SnappyStreamBytes,
                observed,
                maximum: SnappyStream::MAX_DECOMPRESSED_STREAM,
            } if observed == SnappyStream::MAX_DECOMPRESSED_STREAM + 1
        ));

        let error = SnappyStream::maximum_compressed_len(usize::MAX, Default::default())
            .expect_err("usize overflow input must fail closed");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::SnappyStreamBytes,
                observed: usize::MAX,
                maximum: SnappyStream::MAX_DECOMPRESSED_STREAM,
            }
        ));
    }

    #[test]
    fn tiny_compressed_stream_ceiling_accepts_only_empty_output() -> crate::Result<()> {
        for maximum in [1, 4] {
            let limits = SnappyLimits::default().with_compressed_stream_bytes(maximum)?;
            assert_eq!(SnappyStream::maximum_compressed_len(0, limits)?, 0);
            assert!(SnappyStream::compress(&[], limits)?.is_empty());
            assert!(SnappyStream::decompress(&[], limits)?.as_bytes().is_empty());
            assert!(
                matches!(SnappyStream::compress(&[1], limits), Err(Error::Limit {
                kind: LimitKind::SnappyCompressedStreamBytes, maximum: observed_maximum, ..
            }) if observed_maximum == maximum)
            );
        }
        Ok(())
    }

    #[test]
    fn compression_honors_chunk_and_frame_budgets() -> crate::Result<()> {
        let limits = SnappyLimits::new(8, 64)?.with_input_limits(32, 128, 4)?;
        let input = [7; 25];
        let encoded = SnappyStream::compress(&input, limits)?;
        assert_eq!(
            SnappyStream::decompress(&encoded, limits)?.as_bytes(),
            input
        );
        assert!(encoded.len() <= SnappyStream::maximum_compressed_len(input.len(), limits)?);
        assert!(matches!(
            SnappyStream::compress(&[7; 33], limits),
            Err(Error::Limit {
                kind: LimitKind::SnappyFrames,
                observed: 5,
                maximum: 4,
            })
        ));
        Ok(())
    }

    #[test]
    fn prepared_compression_matches_legacy_and_binds_input_length() -> crate::Result<()> {
        let empty = prepared_compression(0, SnappyLimits::default())?;
        assert_eq!(empty.requirements().allocation_events(), 0);
        assert_eq!(empty.requirements().encoder_workspace_bytes(), 0);
        assert!(empty.execute(&[])?.is_empty());

        let input = (0..(SnappyStream::WRITE_CHUNK_SIZE + 17))
            .map(|index| (index as u8).wrapping_mul(13))
            .collect::<Vec<_>>();
        let limits = SnappyLimits::new(4096, input.len())?;
        let plan = prepared_compression(input.len(), limits)?;
        let requirements = plan.requirements();
        assert_eq!(requirements.input_bytes(), input.len());
        assert_eq!(requirements.frame_count(), 17);
        assert!(
            requirements.output_capacity_bytes() >= requirements.temporary_frame_capacity_bytes()
        );
        assert!(requirements.work_bytes() >= input.len());
        assert_eq!(
            plan.execute(&input)?,
            SnappyStream::compress(&input, limits)?
        );

        assert!(matches!(
            prepared_compression(input.len(), limits)?.execute(&input[..input.len() - 1]),
            Err(Error::InvalidArchive {
                reason: "prepared Snappy input length changed",
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn prepared_compression_preserves_tight_fallback_output() -> crate::Result<()> {
        let input = vec![7u8; 3 * 4096 + 17];
        let base_limits = SnappyLimits::new(4096, input.len())?;
        let legacy = SnappyStream::compress(&input, base_limits)?;
        let tight_limits = base_limits.with_compressed_stream_bytes(legacy.len())?;
        let plan = prepared_compression(input.len(), tight_limits)?;
        assert!(plan.requirements().temporary_frame_capacity_bytes() > 0);
        let prepared = plan.execute(&input)?;
        assert_eq!(prepared, SnappyStream::compress(&input, tight_limits)?);
        assert_eq!(
            SnappyStream::decompress(&prepared, tight_limits)?.as_bytes(),
            input
        );
        Ok(())
    }

    #[test]
    fn prepared_compression_honors_work_admission_before_frame_scan() {
        let error = SnappyStream::prepare_compression(
            32,
            SnappyLimits::default(),
            &mut |chunk: SnappyCompressionPreparationChunk| {
                Err(Error::allocation(
                    "test preparation ledger",
                    chunk.traversal_bytes(),
                ))
            },
        )
        .expect_err("preparation work refusal");
        assert!(matches!(
            error,
            SnappyCompressionPreparationError::Admission(Error::Allocation { .. })
        ));
    }

    #[test]
    fn compression_reports_configured_encoded_and_decoded_ceilings() -> crate::Result<()> {
        let decoded = SnappyLimits::new(16, 16)?;
        assert!(matches!(
            SnappyStream::compress(&[0; 17], decoded),
            Err(Error::Limit {
                kind: LimitKind::SnappyStreamBytes,
                observed: 17,
                maximum: 16,
            })
        ));
        let chunk = SnappyLimits::new(32, 64)?.with_input_limits(1, 64, 4)?;
        assert!(matches!(
            SnappyStream::compress(&[7], chunk),
            Err(Error::Limit {
                kind: LimitKind::SnappyCompressedChunkBytes,
                maximum: 1,
                ..
            })
        ));
        let stream = SnappyLimits::new(32, 64)?.with_input_limits(64, 68, 4)?;
        let input = (0..64).collect::<Vec<u8>>();
        assert!(matches!(
            SnappyStream::compress(&input, stream),
            Err(Error::Limit {
                kind: LimitKind::SnappyCompressedStreamBytes,
                maximum: 68,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn conservative_bound_keeps_compressible_input_within_budget() -> crate::Result<()> {
        let limits = SnappyLimits::new(64, 128)?.with_input_limits(16, 24, 4)?;
        let input = [0; 128];
        let maximum = SnappyStream::maximum_compressed_len(input.len(), limits)?;
        assert_eq!(maximum, 24);
        let encoded = SnappyStream::compress(&input, limits)?;
        assert!(encoded.len() <= maximum);
        assert_eq!(
            SnappyStream::decompress(&encoded, limits)?.as_bytes(),
            input
        );
        Ok(())
    }

    #[test]
    fn decompressed_length_preflight_matches_decoded_stream() -> crate::Result<()> {
        let input = (0..(SnappyStream::WRITE_CHUNK_SIZE + 17))
            .map(|index| (index as u8).wrapping_mul(13))
            .collect::<Vec<_>>();
        let compressed = SnappyStream::compress(&input, Default::default())?;
        let length = SnappyStream::decompressed_len(&compressed, SnappyLimits::default())?;
        let decoded = SnappyStream::decompress(&compressed, Default::default())?;
        assert_eq!(length, input.len());
        assert_eq!(decoded.as_bytes(), input.as_slice());
        Ok(())
    }
}
