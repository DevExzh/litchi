#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! Compressed RTF support.
//!
//! This module implements the RTF compression algorithm as specified in:
//! <https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp/65dfe2df-1b69-43fc-8ebd-21819a7463fb>
//!
//! Compressed RTF is commonly used in email attachments and other scenarios
//! where file size reduction is important.

#![allow(
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "decoding steps deliberately rebind a working value as it is refined through the parse pipeline"
)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use super::error::{RtfError, RtfResult};
use zerocopy::{FromBytes, IntoBytes};
use zerocopy_derive::{
    FromBytes as DeriveFromBytes, Immutable, IntoBytes as DeriveIntoBytes, KnownLayout,
};

/// Magic signature for compressed RTF
const COMPRESSED_SIGNATURE: &[u8; 4] = b"LZFu";

/// Magic signature for uncompressed RTF (stored with compression header)
const UNCOMPRESSED_SIGNATURE: &[u8; 4] = b"MELA";

/// Header size, including the four-byte `COMPSIZE` field.
const HEADER_SIZE: usize = 16;

/// Header bytes counted by `COMPSIZE` in addition to the content bytes.
const COUNTED_HEADER_SIZE: usize = HEADER_SIZE - size_of::<u32>();

/// Default finite ceiling for a decompressed compressed-RTF payload (256 MiB).
pub const DEFAULT_MAX_DECOMPRESSED_RTF_BYTES: usize = 256 * 1_048_576;

/// Initial dictionary for compression/decompression
const INIT_DICT: &[u8] = b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}\
{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial\
Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\
\\b\\i\\u\\tab\\tx";

/// Size of initial dictionary
const INIT_DICT_SIZE: usize = INIT_DICT.len();

/// Maximum dictionary size
const MAX_DICT_SIZE: usize = 4096;

const POSITION_WORD_BITS: usize = u64::BITS as usize;
const POSITION_WORDS: usize = MAX_DICT_SIZE / POSITION_WORD_BITS;
const POSITION_BUCKETS: usize = u8::MAX as usize + 1;

const _: () = assert!(INIT_DICT_SIZE <= MAX_DICT_SIZE);
const _: () = assert!(MAX_DICT_SIZE.is_multiple_of(POSITION_WORD_BITS));

fn codec_invariant(message: &'static str) -> RtfError {
    RtfError::InvalidStructure(format!("compressed-RTF codec invariant failed: {message}"))
}

fn allocation_failed(resource: &'static str, requested: usize) -> RtfError {
    RtfError::AllocationFailed {
        resource,
        requested,
    }
}

fn reserve_exact_bytes(
    output: &mut Vec<u8>,
    capacity: usize,
    resource: &'static str,
) -> RtfResult<()> {
    output
        .try_reserve_exact(capacity)
        .map_err(|_err| allocation_failed(resource, capacity))
}

fn extend_bytes(output: &mut Vec<u8>, bytes: &[u8], resource: &'static str) -> RtfResult<()> {
    let requested = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| codec_invariant("output size overflow"))?;
    output
        .try_reserve(bytes.len())
        .map_err(|_err| allocation_failed(resource, requested))?;
    output.extend_from_slice(bytes);
    Ok(())
}

/// Resource limits used while expanding a compressed-RTF payload.
///
/// The declared `RAWSIZE` is checked before allocating the output buffer, and
/// the decoder independently prevents every literal or reference from crossing
/// the same boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecompressionLimits {
    max_output_bytes: usize,
}

impl DecompressionLimits {
    /// Create a limit profile with a caller-selected output ceiling.
    #[must_use]
    pub const fn new(max_output_bytes: usize) -> Self {
        Self { max_output_bytes }
    }

    /// Maximum number of bytes the decoder may produce.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
}

impl Default for DecompressionLimits {
    fn default() -> Self {
        Self::new(DEFAULT_MAX_DECOMPRESSED_RTF_BYTES)
    }
}

/// Compressed RTF header (16 bytes)
#[repr(C)]
#[derive(Debug, Clone, Copy, DeriveIntoBytes, DeriveFromBytes, Immutable, KnownLayout)]
struct CompressedRtfHeader {
    /// Content size plus the remaining 12 header bytes (little-endian).
    compressed_size: [u8; 4],
    /// Size of uncompressed data (little-endian)
    raw_size: [u8; 4],
    /// Compression type signature
    compression_type: [u8; 4],
    /// CRC32 checksum (little-endian)
    crc32: [u8; 4],
}

impl CompressedRtfHeader {
    /// Get compressed size as u32.
    #[inline]
    fn get_compressed_size(&self) -> u32 {
        u32::from_le_bytes(self.compressed_size)
    }

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
#[must_use]
pub fn is_compressed_rtf(data: &[u8]) -> bool {
    data.get(8..12).is_some_and(|signature| {
        signature == COMPRESSED_SIGNATURE || signature == UNCOMPRESSED_SIGNATURE
    })
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
/// Returns an error when the frame, token stream, declared size, compression
/// type, checksum, or default output budget is invalid.
pub fn decompress(data: &[u8]) -> RtfResult<Vec<u8>> {
    decompress_with_limits(data, DecompressionLimits::default())
}

/// Decompress compressed RTF with an explicit finite output budget.
///
/// # Errors
/// Returns an error when the input is malformed or a configured limit is exceeded.
pub fn decompress_with_limits(data: &[u8], limits: DecompressionLimits) -> RtfResult<Vec<u8>> {
    let (header, contents) = parse_frame(data)?;
    let raw_size = usize::try_from(header.get_raw_size()).map_err(|_err| {
        RtfError::InvalidStructure("RTF RAWSIZE is not representable on this target".to_string())
    })?;
    if raw_size > limits.max_output_bytes {
        return Err(RtfError::LimitExceeded {
            resource: "decompressed bytes",
            observed: raw_size,
            limit: limits.max_output_bytes,
        });
    }

    if header.is_compressed() {
        decompress_lzfu(contents, &header, raw_size)
    } else if header.is_uncompressed() {
        decompress_uncompressed(contents, &header, raw_size)
    } else {
        Err(RtfError::InvalidStructure(format!(
            "unknown compressed-RTF compression type: {:?}",
            header.compression_type
        )))
    }
}

fn parse_frame(data: &[u8]) -> RtfResult<(CompressedRtfHeader, &[u8])> {
    let bytes = data.get(..HEADER_SIZE).ok_or_else(|| {
        RtfError::InvalidStructure(format!(
            "compressed-RTF header requires {HEADER_SIZE} bytes, found {}",
            data.len()
        ))
    })?;
    let header = *<CompressedRtfHeader as FromBytes>::ref_from_bytes(bytes).map_err(|_err| {
        RtfError::InvalidStructure("failed to parse compressed-RTF header".to_string())
    })?;
    let compressed_size = usize::try_from(header.get_compressed_size()).map_err(|_err| {
        RtfError::InvalidStructure("RTF COMPSIZE is not representable on this target".to_string())
    })?;
    if compressed_size < COUNTED_HEADER_SIZE {
        return Err(RtfError::InvalidStructure(format!(
            "RTF COMPSIZE {compressed_size} is smaller than the {COUNTED_HEADER_SIZE}-byte counted header"
        )));
    }
    let declared_total = compressed_size
        .checked_add(HEADER_SIZE - COUNTED_HEADER_SIZE)
        .ok_or_else(|| {
            RtfError::InvalidStructure("compressed-RTF total size overflow".to_string())
        })?;
    if declared_total != data.len() {
        return Err(RtfError::InvalidStructure(format!(
            "RTF COMPSIZE declares {declared_total} total bytes, found {}",
            data.len()
        )));
    }
    let contents = data.get(HEADER_SIZE..).ok_or_else(|| {
        RtfError::InvalidStructure("compressed-RTF content boundary is invalid".to_string())
    })?;
    Ok((header, contents))
}

/// Decompress LZFu-compressed data.
fn decompress_lzfu(
    data: &[u8],
    header: &CompressedRtfHeader,
    raw_size: usize,
) -> RtfResult<Vec<u8>> {
    let calculated_crc = crc32(data)?;
    if calculated_crc != header.get_crc32() {
        return Err(RtfError::InvalidStructure(format!(
            "compressed-RTF CRC32 mismatch: expected {:#010x}, got {calculated_crc:#010x}",
            header.get_crc32()
        )));
    }

    let mut dictionary = DecodeDictionary::new();
    let mut output = Vec::new();
    reserve_exact_bytes(&mut output, raw_size, "decompressed RTF output")?;
    let mut cursor = 0usize;
    let mut saw_artificial_empty_literal = false;
    while let Some(&control) = data.get(cursor) {
        cursor = cursor
            .checked_add(1)
            .ok_or_else(|| codec_invariant("input cursor overflow"))?;
        for bit in 0..8 {
            if control & (1 << bit) == 0 {
                let literal = *data.get(cursor).ok_or_else(|| {
                    RtfError::InvalidStructure(
                        "compressed RTF ends inside a literal token".to_string(),
                    )
                })?;
                cursor = cursor
                    .checked_add(1)
                    .ok_or_else(|| codec_invariant("input cursor overflow"))?;
                let canonical_empty_literal = raw_size == 0
                    && output.is_empty()
                    && literal == 0
                    && !saw_artificial_empty_literal;
                dictionary.write(literal)?;
                if canonical_empty_literal {
                    saw_artificial_empty_literal = true;
                    continue;
                }
                ensure_output_growth(output.len(), 1, raw_size)?;
                output.push(literal);
                continue;
            }

            let token_end = cursor
                .checked_add(2)
                .ok_or_else(|| codec_invariant("dictionary-reference cursor overflow"))?;
            let token_bytes: [u8; 2] = data
                .get(cursor..token_end)
                .ok_or_else(|| {
                    RtfError::InvalidStructure(
                        "compressed RTF ends inside a dictionary reference".to_string(),
                    )
                })?
                .try_into()
                .map_err(|_err| {
                    RtfError::InvalidStructure(
                        "invalid compressed-RTF dictionary reference".to_string(),
                    )
                })?;
            cursor = token_end;
            let token = u16::from_be_bytes(token_bytes);
            let offset = usize::from(token >> 4);
            if offset == dictionary.write_offset {
                if output.len() == raw_size {
                    return Ok(output);
                }
                return Err(raw_size_mismatch(output.len(), raw_size));
            }
            if !dictionary.contains(offset) {
                return Err(RtfError::InvalidStructure(format!(
                    "compressed RTF references unavailable dictionary offset {offset}"
                )));
            }

            let length = usize::from(token & 0x000f) + 2;
            ensure_output_growth(output.len(), length, raw_size)?;
            let mut read_offset = offset;
            for _ in 0..length {
                let byte = dictionary.read(read_offset)?;
                output.push(byte);
                dictionary.write(byte)?;
                read_offset = read_offset
                    .checked_add(1)
                    .ok_or_else(|| codec_invariant("dictionary read offset overflow"))?
                    % MAX_DICT_SIZE;
            }
        }
    }

    Err(RtfError::InvalidStructure(
        "compressed RTF has no terminating dictionary reference".to_string(),
    ))
}

fn ensure_output_growth(current: usize, additional: usize, declared: usize) -> RtfResult<()> {
    let observed = current
        .checked_add(additional)
        .ok_or_else(|| RtfError::InvalidStructure("decompressed RTF size overflow".to_string()))?;
    if observed > declared {
        return Err(raw_size_mismatch(observed, declared));
    }
    Ok(())
}

fn raw_size_mismatch(observed: usize, declared: usize) -> RtfError {
    RtfError::InvalidStructure(format!(
        "RTF RAWSIZE declares {declared} bytes, decoder produced {observed}"
    ))
}

/// Decompress uncompressed RTF data.
fn decompress_uncompressed(
    data: &[u8],
    header: &CompressedRtfHeader,
    raw_size: usize,
) -> RtfResult<Vec<u8>> {
    if header.get_crc32() != 0 {
        return Err(RtfError::InvalidStructure(
            "CRC32 must be 0x00000000 for uncompressed RTF".to_string(),
        ));
    }
    if data.len() != raw_size {
        return Err(raw_size_mismatch(data.len(), raw_size));
    }
    let mut output = Vec::new();
    reserve_exact_bytes(&mut output, raw_size, "uncompressed RTF output")?;
    output.extend_from_slice(data);
    Ok(output)
}

struct DecodeDictionary {
    bytes: [u8; MAX_DICT_SIZE],
    write_offset: usize,
    len: usize,
}

impl DecodeDictionary {
    fn new() -> Self {
        let mut bytes = [0; MAX_DICT_SIZE];
        for (slot, byte) in bytes.iter_mut().zip(INIT_DICT.iter().copied()) {
            *slot = byte;
        }
        Self {
            bytes,
            write_offset: INIT_DICT_SIZE,
            len: INIT_DICT_SIZE,
        }
    }

    fn contains(&self, offset: usize) -> bool {
        offset < self.len || self.len == MAX_DICT_SIZE
    }

    fn read(&self, offset: usize) -> RtfResult<u8> {
        self.bytes
            .get(offset)
            .copied()
            .ok_or_else(|| codec_invariant("dictionary read offset is out of range"))
    }

    fn write(&mut self, byte: u8) -> RtfResult<()> {
        let slot = self
            .bytes
            .get_mut(self.write_offset)
            .ok_or_else(|| codec_invariant("dictionary write offset is out of range"))?;
        *slot = byte;
        self.write_offset = self
            .write_offset
            .checked_add(1)
            .ok_or_else(|| codec_invariant("dictionary write offset overflow"))?
            % MAX_DICT_SIZE;
        self.len = self.len.saturating_add(1).min(MAX_DICT_SIZE);
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct DictionaryMatch {
    offset: usize,
    len: usize,
}

/// One control byte followed by at most eight two-byte dictionary references.
struct EncodeRun {
    bytes: [u8; 17],
    len: usize,
}

impl EncodeRun {
    fn new() -> Self {
        Self {
            bytes: [0; 17],
            len: 1,
        }
    }

    fn mark_reference(&mut self, bit: usize) -> RtfResult<()> {
        if bit >= u8::BITS as usize {
            return Err(codec_invariant("encode-run control bit is out of range"));
        }
        let control = self
            .bytes
            .first_mut()
            .ok_or_else(|| codec_invariant("encode-run control byte is missing"))?;
        *control |= 1u8 << bit;
        Ok(())
    }

    fn push_literal(&mut self, literal: u8) -> RtfResult<()> {
        let slot = self
            .bytes
            .get_mut(self.len)
            .ok_or_else(|| codec_invariant("encode run exceeds its byte capacity"))?;
        *slot = literal;
        self.len = self
            .len
            .checked_add(1)
            .ok_or_else(|| codec_invariant("encode-run length overflow"))?;
        Ok(())
    }

    fn push_reference(&mut self, bit: usize, reference: [u8; 2]) -> RtfResult<()> {
        self.mark_reference(bit)?;
        let end = self
            .len
            .checked_add(reference.len())
            .ok_or_else(|| codec_invariant("encode-run length overflow"))?;
        let target = self
            .bytes
            .get_mut(self.len..end)
            .ok_or_else(|| codec_invariant("encode run exceeds its byte capacity"))?;
        target.copy_from_slice(&reference);
        self.len = end;
        Ok(())
    }

    fn as_bytes(&self) -> RtfResult<&[u8]> {
        self.bytes
            .get(..self.len)
            .ok_or_else(|| codec_invariant("encode-run length is out of range"))
    }
}

struct EncodeDictionary {
    ring: DecodeDictionary,
    positions: Vec<[u64; POSITION_WORDS]>,
}

impl EncodeDictionary {
    fn new() -> RtfResult<Self> {
        let ring = DecodeDictionary::new();
        let position_bytes = POSITION_BUCKETS
            .checked_mul(size_of::<[u64; POSITION_WORDS]>())
            .ok_or_else(|| codec_invariant("encoder position-table size overflow"))?;
        let mut positions = Vec::new();
        positions
            .try_reserve_exact(POSITION_BUCKETS)
            .map_err(|_err| allocation_failed("compressed RTF position table", position_bytes))?;
        positions.resize(POSITION_BUCKETS, [0; POSITION_WORDS]);
        for (offset, &byte) in INIT_DICT.iter().enumerate() {
            let words = positions
                .get_mut(usize::from(byte))
                .ok_or_else(|| codec_invariant("position bucket is out of range"))?;
            set_position(words, offset)?;
        }
        Ok(Self { ring, positions })
    }

    fn write(&mut self, byte: u8) -> RtfResult<()> {
        let offset = self.ring.write_offset;
        if self.ring.len == MAX_DICT_SIZE {
            let previous = self.ring.read(offset)?;
            let words = self
                .positions
                .get_mut(usize::from(previous))
                .ok_or_else(|| codec_invariant("position bucket is out of range"))?;
            clear_position(words, offset)?;
        }
        self.ring.write(byte)?;
        let words = self
            .positions
            .get_mut(usize::from(byte))
            .ok_or_else(|| codec_invariant("position bucket is out of range"))?;
        set_position(words, offset)
    }

    /// Find the spec-defined longest match while mirroring each newly proven
    /// byte into the ring. That mutation is what lets references overlap the
    /// current write position during decoding.
    fn find_longest_match(&mut self, input: &[u8]) -> RtfResult<Option<DictionaryMatch>> {
        let Some((&first, _)) = input.split_first() else {
            return Ok(None);
        };
        let final_offset = self.ring.write_offset;
        let mut best = None;

        if self.ring.len < MAX_DICT_SIZE {
            self.scan_candidates(first, 0, final_offset, input, &mut best)?;
        } else {
            let next_offset = final_offset
                .checked_add(1)
                .ok_or_else(|| codec_invariant("candidate offset overflow"))?;
            if next_offset < MAX_DICT_SIZE {
                self.scan_candidates(first, next_offset, MAX_DICT_SIZE, input, &mut best)?;
            }
            if best.is_none_or(|matched| matched.len < 17) {
                self.scan_candidates(first, 0, final_offset, input, &mut best)?;
            }
        }

        if best.is_none() {
            self.write(first)?;
        }
        Ok(best.filter(|matched| matched.len >= 2))
    }

    fn scan_candidates(
        &mut self,
        first: u8,
        start: usize,
        end: usize,
        input: &[u8],
        best: &mut Option<DictionaryMatch>,
    ) -> RtfResult<()> {
        let mut next = start;
        loop {
            let offset = {
                let words = self
                    .positions
                    .get(usize::from(first))
                    .ok_or_else(|| codec_invariant("position bucket is out of range"))?;
                next_position(words, next, end)?
            };
            let Some(offset) = offset else {
                break;
            };
            let best_len = best.map_or(0, |matched| matched.len);
            let match_len = self.match_at(offset, input, best_len)?;
            if match_len > best_len {
                *best = Some(DictionaryMatch {
                    offset,
                    len: match_len,
                });
                if match_len == 17 {
                    return Ok(());
                }
            }
            next = offset
                .checked_add(1)
                .ok_or_else(|| codec_invariant("candidate offset overflow"))?;
        }
        Ok(())
    }

    fn match_at(&mut self, offset: usize, input: &[u8], best_len: usize) -> RtfResult<usize> {
        let max_len = input.len().min(17);
        let mut dictionary_offset = offset;
        let mut match_len = 0usize;
        while match_len < max_len {
            let dictionary_byte = self.ring.read(dictionary_offset)?;
            let input_byte = *input
                .get(match_len)
                .ok_or_else(|| codec_invariant("match input offset is out of range"))?;
            if dictionary_byte != input_byte {
                break;
            }
            match_len += 1;
            if match_len > best_len {
                self.write(input_byte)?;
            }
            dictionary_offset = dictionary_offset
                .checked_add(1)
                .ok_or_else(|| codec_invariant("dictionary match offset overflow"))?
                % MAX_DICT_SIZE;
        }
        Ok(match_len)
    }
}

fn set_position(words: &mut [u64; POSITION_WORDS], offset: usize) -> RtfResult<()> {
    let word = words
        .get_mut(offset / POSITION_WORD_BITS)
        .ok_or_else(|| codec_invariant("position-set offset is out of range"))?;
    *word |= 1 << (offset % POSITION_WORD_BITS);
    Ok(())
}

fn clear_position(words: &mut [u64; POSITION_WORDS], offset: usize) -> RtfResult<()> {
    let word = words
        .get_mut(offset / POSITION_WORD_BITS)
        .ok_or_else(|| codec_invariant("position-clear offset is out of range"))?;
    *word &= !(1 << (offset % POSITION_WORD_BITS));
    Ok(())
}

fn next_position(
    words: &[u64; POSITION_WORDS],
    start: usize,
    end: usize,
) -> RtfResult<Option<usize>> {
    if start > end || end > MAX_DICT_SIZE {
        return Err(codec_invariant("position scan range is invalid"));
    }
    let mut position = start;
    while position < end {
        let word_index = position / POSITION_WORD_BITS;
        let bit_index = position % POSITION_WORD_BITS;
        let word_start = word_index
            .checked_mul(POSITION_WORD_BITS)
            .ok_or_else(|| codec_invariant("position word offset overflow"))?;
        let mut word = words
            .get(word_index)
            .copied()
            .ok_or_else(|| codec_invariant("position word is out of range"))?
            & (u64::MAX << bit_index);
        let valid_bits = end
            .checked_sub(word_start)
            .ok_or_else(|| codec_invariant("position word exceeds scan end"))?
            .min(POSITION_WORD_BITS);
        if valid_bits < POSITION_WORD_BITS {
            word &= (1u64 << valid_bits) - 1;
        }
        if word != 0 {
            let bit = usize::try_from(word.trailing_zeros())
                .map_err(|_err| codec_invariant("position bit cannot be represented"))?;
            let found = word_start
                .checked_add(bit)
                .ok_or_else(|| codec_invariant("position result overflow"))?;
            return Ok(Some(found));
        }
        position = word_index
            .checked_add(1)
            .and_then(|index| index.checked_mul(POSITION_WORD_BITS))
            .ok_or_else(|| codec_invariant("position scan overflow"))?;
    }
    Ok(None)
}

/// Compress RTF data
///
/// # Arguments
///
/// * `data` - Uncompressed RTF data
/// * `compress` - If true, use `LZFu` compression; if false, store uncompressed
///
/// # Returns
///
/// Compressed RTF data with header
///
/// # Errors
/// Returns an error when the input is malformed or a configured limit is exceeded.
pub fn compress(data: &[u8], compress: bool) -> RtfResult<Vec<u8>> {
    if compress {
        compress_lzfu(data)
    } else {
        compress_uncompressed(data)
    }
}

/// Compress data using `LZFu` algorithm
fn compress_lzfu(data: &[u8]) -> RtfResult<Vec<u8>> {
    checked_u32_size(data.len(), "uncompressed RTF")?;
    let mut dictionary = EncodeDictionary::new()?;
    let mut output = Vec::new();

    if data.is_empty() {
        // Required by [MS-OXRTFCP] 2.3.3.2: the canonical empty stream
        // carries one artificial NUL literal before its end reference.
        dictionary.write(0)?;
        let mut run = EncodeRun::new();
        run.push_literal(0)?;
        run.push_reference(1, dictionary_reference(dictionary.ring.write_offset, 0)?)?;
        extend_bytes(&mut output, run.as_bytes()?, "compressed RTF output")?;
    } else {
        let mut input_offset = 0usize;
        'runs: loop {
            let mut run = EncodeRun::new();
            for bit in 0..8usize {
                if input_offset == data.len() {
                    run.push_reference(
                        bit,
                        dictionary_reference(dictionary.ring.write_offset, 0)?,
                    )?;
                    extend_bytes(&mut output, run.as_bytes()?, "compressed RTF output")?;
                    break 'runs;
                }

                let remaining = data
                    .get(input_offset..)
                    .ok_or_else(|| codec_invariant("compressor input offset is out of range"))?;
                if let Some(matched) = dictionary.find_longest_match(remaining)? {
                    run.push_reference(bit, dictionary_reference(matched.offset, matched.len)?)?;
                    input_offset = input_offset
                        .checked_add(matched.len)
                        .filter(|offset| *offset <= data.len())
                        .ok_or_else(|| codec_invariant("compressor input offset overflow"))?;
                } else {
                    let literal = *data
                        .get(input_offset)
                        .ok_or_else(|| codec_invariant("compressor literal is out of range"))?;
                    run.push_literal(literal)?;
                    input_offset = input_offset
                        .checked_add(1)
                        .ok_or_else(|| codec_invariant("compressor input offset overflow"))?;
                }
            }
            extend_bytes(&mut output, run.as_bytes()?, "compressed RTF output")?;
        }
    }

    let crc32 = crc32(&output)?;
    build_frame(data.len(), &output, *COMPRESSED_SIGNATURE, crc32)
}

/// Compress data without compression (just add header)
fn compress_uncompressed(data: &[u8]) -> RtfResult<Vec<u8>> {
    build_frame(data.len(), data, *UNCOMPRESSED_SIGNATURE, 0)
}

fn dictionary_reference(offset: usize, length: usize) -> RtfResult<[u8; 2]> {
    if offset >= MAX_DICT_SIZE {
        return Err(RtfError::InvalidStructure(format!(
            "dictionary offset {offset} is not representable in compressed RTF"
        )));
    }
    if length != 0 && !(2..=17).contains(&length) {
        return Err(RtfError::InvalidStructure(format!(
            "dictionary match length {length} is not representable in compressed RTF"
        )));
    }
    let offset = u16::try_from(offset).map_err(|_err| {
        RtfError::InvalidStructure("compressed-RTF dictionary offset overflow".to_string())
    })?;
    let encoded_length = u16::try_from(length.saturating_sub(2)).map_err(|_err| {
        RtfError::InvalidStructure("compressed-RTF match length overflow".to_string())
    })?;
    Ok(((offset << 4) | encoded_length).to_be_bytes())
}

fn checked_u32_size(value: usize, description: &str) -> RtfResult<u32> {
    u32::try_from(value).map_err(|_err| {
        RtfError::InvalidStructure(format!(
            "{description} size {value} exceeds the compressed-RTF 32-bit wire limit"
        ))
    })
}

fn build_frame(
    raw_len: usize,
    contents: &[u8],
    compression_type: [u8; 4],
    crc32: u32,
) -> RtfResult<Vec<u8>> {
    let raw_size = checked_u32_size(raw_len, "uncompressed RTF")?;
    let counted_size = contents
        .len()
        .checked_add(COUNTED_HEADER_SIZE)
        .ok_or_else(|| RtfError::InvalidStructure("RTF COMPSIZE overflow".to_string()))?;
    let compressed_size = checked_u32_size(counted_size, "compressed RTF")?;
    let total_size = HEADER_SIZE
        .checked_add(contents.len())
        .ok_or_else(|| RtfError::InvalidStructure("compressed-RTF size overflow".to_string()))?;
    let header = CompressedRtfHeader::new(compressed_size, raw_size, compression_type, crc32);
    let mut result = Vec::new();
    reserve_exact_bytes(&mut result, total_size, "compressed RTF frame")?;
    result.extend_from_slice(IntoBytes::as_bytes(&header));
    result.extend_from_slice(contents);
    Ok(result)
}

fn crc32(data: &[u8]) -> RtfResult<u32> {
    // [MS-OXRTFCP] uses the reflected ISO polynomial with an all-zero
    // initial state and no final XOR. `crc-fast` supplies the accelerated
    // polynomial implementation; remove its ISO-HDLC final XOR here.
    let mut digest = crc_fast::Digest::new_with_init_state(crc_fast::CrcAlgorithm::Crc32IsoHdlc, 0);
    digest.update(data);
    u32::try_from(digest.finalize() ^ u64::from(u32::MAX))
        .map_err(|_err| RtfError::InvalidStructure("compressed-RTF CRC32 overflow".to_string()))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    #![allow(
        clippy::cast_possible_truncation,
        reason = "test PRNG fixtures intentionally truncate wide state into bytes"
    )]
    use super::*;

    const SPEC_RAW: &[u8] = b"{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n";
    const SPEC_COMPRESSED: &[u8] = &[
        0x2d, 0x00, 0x00, 0x00, 0x2b, 0x00, 0x00, 0x00, 0x4c, 0x5a, 0x46, 0x75, 0xf1, 0xc5, 0xc7,
        0xa7, 0x03, 0x00, 0x0a, 0x00, 0x72, 0x63, 0x70, 0x67, 0x31, 0x32, 0x35, 0x42, 0x32, 0x0a,
        0xf3, 0x20, 0x68, 0x65, 0x6c, 0x09, 0x00, 0x20, 0x62, 0x77, 0x05, 0xb0, 0x6c, 0x64, 0x7d,
        0x0a, 0x80, 0x0f, 0xa0,
    ];

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

    #[test]
    fn specification_example_is_byte_exact() {
        assert_eq!(decompress(SPEC_COMPRESSED).unwrap(), SPEC_RAW);
        assert_eq!(compress(SPEC_RAW, true).unwrap(), SPEC_COMPRESSED);
        assert_eq!(
            crate::RtfDocument::from_bytes(SPEC_COMPRESSED)
                .unwrap()
                .text(),
            "hello world"
        );
    }

    #[test]
    fn compressed_round_trip_crosses_and_wraps_the_dictionary_write_position() {
        let mut original = b"{\\rtf1 WXYZWXYZWXYZWXYZWXYZ}".repeat(512);
        for index in 0..12_000usize {
            original.push(((index * 37 + index / 17 * 11) % 251) as u8);
        }

        let encoded = compress(&original, true).unwrap();
        assert!(encoded.len() < original.len());
        assert_eq!(decompress(&encoded).unwrap(), original);
    }

    #[test]
    fn compressed_round_trip_handles_varied_binary_corpus() {
        const LENGTHS: &[usize] = &[
            0, 1, 2, 3, 16, 17, 18, 255, 256, 257, 4095, 4096, 4097, 8193,
        ];

        for case in 0..3u64 {
            for &len in LENGTHS {
                let mut state = 0x9e37_79b9_7f4a_7c15u64
                    ^ case.wrapping_mul(0xd1b5_4a32_d192_ed03)
                    ^ len as u64;
                let mut original = Vec::with_capacity(len);
                for index in 0..len {
                    state ^= state << 13;
                    state ^= state >> 7;
                    state ^= state << 17;
                    let byte = match case {
                        0 => state as u8,
                        1 => (state & 0x07) as u8,
                        _ => ((index % 251) as u8).wrapping_add((state & 0x03) as u8),
                    };
                    original.push(byte);
                }

                let encoded = compress(&original, true).unwrap();
                assert_eq!(
                    decompress(&encoded).unwrap(),
                    original,
                    "round trip failed for corpus case {case} at length {len}"
                );
            }
        }
    }

    #[test]
    fn canonical_empty_compressed_stream_round_trips() {
        let encoded = compress(&[], true).unwrap();
        assert_eq!(&encoded[HEADER_SIZE..], &[0x02, 0x00, 0x0d, 0x00]);
        assert!(decompress(&encoded).unwrap().is_empty());
    }

    #[test]
    fn declared_frame_and_raw_sizes_are_exact() {
        let mut stored = compress(b"abc", false).unwrap();
        stored[0..4].copy_from_slice(&14u32.to_le_bytes());
        assert!(matches!(
            decompress(&stored),
            Err(RtfError::InvalidStructure(message)) if message.contains("COMPSIZE")
        ));

        let stored = build_frame(4, b"abc", *UNCOMPRESSED_SIGNATURE, 0).unwrap();
        assert!(matches!(
            decompress(&stored),
            Err(RtfError::InvalidStructure(message)) if message.contains("RAWSIZE")
        ));

        let stored = build_frame(2, b"abc", *UNCOMPRESSED_SIGNATURE, 0).unwrap();
        assert!(matches!(
            decompress(&stored),
            Err(RtfError::InvalidStructure(message)) if message.contains("RAWSIZE")
        ));
    }

    #[test]
    fn declared_output_is_budgeted_before_allocation() {
        let mut stored = compress(&[], false).unwrap();
        stored[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
        let limits = DecompressionLimits::new(1024);
        assert_eq!(limits.max_output_bytes(), 1024);
        assert!(matches!(
            decompress_with_limits(&stored, limits),
            Err(RtfError::LimitExceeded {
                resource: "decompressed bytes",
                observed,
                limit: 1024,
            }) if observed == u32::MAX as usize
        ));
    }

    #[test]
    fn failed_capacity_reservations_are_typed() {
        let mut output = Vec::new();
        let error = reserve_exact_bytes(&mut output, usize::MAX, "test output").unwrap_err();

        assert!(matches!(
            error,
            RtfError::AllocationFailed {
                resource: "test output",
                requested: usize::MAX,
            }
        ));
    }

    #[test]
    fn codec_helpers_reject_impossible_indices_and_ranges() {
        let dictionary = DecodeDictionary::new();
        assert!(dictionary.read(MAX_DICT_SIZE).is_err());

        let mut positions = [0; POSITION_WORDS];
        assert!(set_position(&mut positions, MAX_DICT_SIZE).is_err());
        assert!(clear_position(&mut positions, MAX_DICT_SIZE).is_err());
        assert!(next_position(&positions, 2, 1).is_err());
        assert!(next_position(&positions, 0, MAX_DICT_SIZE + 1).is_err());

        let mut run = EncodeRun::new();
        assert!(run.mark_reference(u8::BITS as usize).is_err());
    }

    #[test]
    fn malformed_token_streams_fail_instead_of_returning_partial_output() {
        let missing_end = build_frame(
            1,
            &[0x00, b'A'],
            *COMPRESSED_SIGNATURE,
            crc32(&[0x00, b'A']).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            decompress(&missing_end),
            Err(RtfError::InvalidStructure(message)) if message.contains("ends")
        ));

        let truncated_reference = build_frame(
            0,
            &[0x01, 0x0c],
            *COMPRESSED_SIGNATURE,
            crc32(&[0x01, 0x0c]).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            decompress(&truncated_reference),
            Err(RtfError::InvalidStructure(message)) if message.contains("dictionary reference")
        ));

        let unavailable_reference = build_frame(
            2,
            &[0x01, 0x12, 0xc0],
            *COMPRESSED_SIGNATURE,
            crc32(&[0x01, 0x12, 0xc0]).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            decompress(&unavailable_reference),
            Err(RtfError::InvalidStructure(message)) if message.contains("unavailable")
        ));

        let oversized_literal = build_frame(
            0,
            &[0x02, b'A', 0x0d, 0x00],
            *COMPRESSED_SIGNATURE,
            crc32(&[0x02, b'A', 0x0d, 0x00]).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            decompress(&oversized_literal),
            Err(RtfError::InvalidStructure(message)) if message.contains("RAWSIZE")
        ));

        let duplicate_artificial_literal_contents = [0x04, 0x00, 0x00, 0x0d, 0x10];
        let duplicate_artificial_literal = build_frame(
            0,
            &duplicate_artificial_literal_contents,
            *COMPRESSED_SIGNATURE,
            crc32(&duplicate_artificial_literal_contents).unwrap(),
        )
        .unwrap();
        assert!(matches!(
            decompress(&duplicate_artificial_literal),
            Err(RtfError::InvalidStructure(message)) if message.contains("RAWSIZE")
        ));
    }

    #[test]
    fn padding_is_counted_by_crc_but_not_decoded() {
        let contents = [&SPEC_COMPRESSED[HEADER_SIZE..], &[0xaa, 0x55]].concat();
        let padded = build_frame(
            SPEC_RAW.len(),
            &contents,
            *COMPRESSED_SIGNATURE,
            crc32(&contents).unwrap(),
        )
        .unwrap();
        assert_eq!(decompress(&padded).unwrap(), SPEC_RAW);
    }

    #[test]
    fn corrupt_crc_and_zero_crc_rule_are_rejected() {
        let mut compressed = SPEC_COMPRESSED.to_vec();
        compressed[12] ^= 0x80;
        assert!(matches!(
            decompress(&compressed),
            Err(RtfError::InvalidStructure(message)) if message.contains("CRC32")
        ));

        let stored = build_frame(3, b"abc", *UNCOMPRESSED_SIGNATURE, 1).unwrap();
        assert!(matches!(
            decompress(&stored),
            Err(RtfError::InvalidStructure(message)) if message.contains("CRC32")
        ));
    }

    #[test]
    fn truncation_and_single_byte_mutation_never_panic() {
        for end in 0..SPEC_COMPRESSED.len() {
            assert!(decompress(&SPEC_COMPRESSED[..end]).is_err());
        }
        for index in 0..SPEC_COMPRESSED.len() {
            for replacement in [0x00, 0x01, 0x7f, 0xff] {
                let mut mutated = SPEC_COMPRESSED.to_vec();
                mutated[index] = replacement;
                drop(decompress(&mutated));
            }
        }
    }

    #[test]
    fn wire_size_helpers_reject_unrepresentable_lengths() {
        if usize::BITS > u32::BITS {
            assert!(checked_u32_size(u32::MAX as usize + 1, "test").is_err());
            assert!(build_frame(u32::MAX as usize + 1, &[], *UNCOMPRESSED_SIGNATURE, 0).is_err());
        }
    }
}
