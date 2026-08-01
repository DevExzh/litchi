//! Bounded codecs for the compression algorithms in MS-XLDM section 2.7.
//!
//! Every operation transforms bytes already supplied by the caller.  This module
//! performs no path resolution, file access, or data-source access.

use std::fmt;

const XPRESS_BLOCK_MAX: usize = 65_535;
const XPRESS_LITERAL_BLOCK: usize = 58_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XldmCodecLimits {
    pub max_output_bytes: usize,
    pub max_values: usize,
    pub max_strings: usize,
    pub max_input_bytes: usize,
}

impl Default for XldmCodecLimits {
    fn default() -> Self {
        Self {
            max_output_bytes: 256 * 1024 * 1024,
            max_values: 16 * 1024 * 1024,
            max_strings: 4 * 1024 * 1024,
            max_input_bytes: 256 * 1024 * 1024,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XldmCodecError {
    Invalid(&'static str),
    LimitExceeded(&'static str),
    IntegerOverflow,
}

impl fmt::Display for XldmCodecError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(message) => write!(f, "invalid XLDM compressed data: {message}"),
            Self::LimitExceeded(name) => write!(f, "XLDM codec limit exceeded: {name}"),
            Self::IntegerOverflow => f.write_str("integer overflow while processing XLDM data"),
        }
    }
}

impl std::error::Error for XldmCodecError {}

pub type XldmCodecResult<T> = Result<T, XldmCodecError>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XldmNoSplitWidth {
    W1 = 1,
    W2 = 2,
    W3 = 3,
    W4 = 4,
    W5 = 5,
    W6 = 6,
    W7 = 7,
    W8 = 8,
    W9 = 9,
    W10 = 10,
    W12 = 12,
    W16 = 16,
    W21 = 21,
    W32 = 32,
}

impl XldmNoSplitWidth {
    pub const fn bits(self) -> u8 {
        self as u8
    }
}

impl TryFrom<u8> for XldmNoSplitWidth {
    type Error = XldmCodecError;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            1 => Ok(Self::W1),
            2 => Ok(Self::W2),
            3 => Ok(Self::W3),
            4 => Ok(Self::W4),
            5 => Ok(Self::W5),
            6 => Ok(Self::W6),
            7 => Ok(Self::W7),
            8 => Ok(Self::W8),
            9 => Ok(Self::W9),
            10 => Ok(Self::W10),
            12 => Ok(Self::W12),
            16 => Ok(Self::W16),
            21 => Ok(Self::W21),
            32 => Ok(Self::W32),
            _ => Err(XldmCodecError::Invalid("unsupported XMRENoSplit width")),
        }
    }
}

/// Returns one entry from the 65-by-64 compression mask table in Appendix A.
///
/// `field_bits` is the table row (`0..=64`) and `bit_offset` is the column
/// (`0..63`). Entries whose field would cross the high end of a 64-bit word,
/// plus the zero-width row, are the appendix's zero sentinel. Values outside
/// the table return `None` rather than being truncated.
pub const fn xldm_no_split_compression_mask(field_bits: u8, bit_offset: u8) -> Option<u64> {
    if field_bits > 64 || bit_offset >= 64 {
        return None;
    }
    if field_bits == 0 || field_bits as u16 + bit_offset as u16 > 64 {
        return Some(0);
    }
    if field_bits == 64 {
        return Some(0);
    }
    let field = (1u64 << field_bits) - 1;
    Some(!(field << bit_offset))
}

fn checked_output(
    count: usize,
    element_size: usize,
    limits: XldmCodecLimits,
) -> XldmCodecResult<()> {
    if count > limits.max_values {
        return Err(XldmCodecError::LimitExceeded("max_values"));
    }
    let bytes = count
        .checked_mul(element_size)
        .ok_or(XldmCodecError::IntegerOverflow)?;
    if bytes > limits.max_output_bytes {
        return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
    }
    Ok(())
}

fn no_split_size(count: usize, width: XldmNoSplitWidth) -> XldmCodecResult<usize> {
    let per_word = 64 / usize::from(width.bits());
    let words = count
        .checked_add(per_word - 1)
        .ok_or(XldmCodecError::IntegerOverflow)?
        / per_word;
    words.checked_mul(8).ok_or(XldmCodecError::IntegerOverflow)
}

pub fn decompress_xldm_no_split(
    input: &[u8],
    width: XldmNoSplitWidth,
    min: i32,
    count: usize,
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<i32>> {
    if input.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    checked_output(count, 4, limits)?;
    if input.len() != no_split_size(count, width)? {
        return Err(XldmCodecError::Invalid(
            "XMRENoSplit storage size does not match value count",
        ));
    }
    let bits = usize::from(width.bits());
    let per_word = 64 / bits;
    let mask = (1u64 << bits) - 1;
    let mut result = Vec::with_capacity(count);
    for chunk in input.chunks_exact(8) {
        let word = u64::from_le_bytes(chunk.try_into().expect("exact chunk"));
        for slot in 0..per_word {
            if result.len() == count {
                break;
            }
            let delta = ((word >> (slot * bits)) & mask) as i64;
            let value = i64::from(min)
                .checked_add(delta)
                .ok_or(XldmCodecError::IntegerOverflow)?;
            result.push(i32::try_from(value).map_err(|_| {
                XldmCodecError::Invalid("XMRENoSplit value exceeds signed 32-bit range")
            })?);
        }
    }
    Ok(result)
}

pub fn compress_xldm_no_split(
    values: &[i32],
    width: XldmNoSplitWidth,
    min: i32,
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<u8>> {
    if values.len() > limits.max_values {
        return Err(XldmCodecError::LimitExceeded("max_values"));
    }
    let size = no_split_size(values.len(), width)?;
    if size > limits.max_output_bytes {
        return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
    }
    let bits = usize::from(width.bits());
    let per_word = 64 / bits;
    let max_delta = (1u64 << bits) - 1;
    let mut output = vec![0u8; size];
    for (word_index, group) in values.chunks(per_word).enumerate() {
        let mut word = 0u64;
        for (slot, &value) in group.iter().enumerate() {
            let delta = i64::from(value) - i64::from(min);
            if delta < 0 || (delta as u64) > max_delta {
                return Err(XldmCodecError::Invalid(
                    "value cannot be represented by XMRENoSplit width and Min",
                ));
            }
            let bit_offset = slot * bits;
            let mask = xldm_no_split_compression_mask(width.bits(), bit_offset as u8)
                .expect("supported widths and slots are within the Appendix A table");
            word = (word & mask) | ((delta as u64) << bit_offset);
        }
        output[word_index * 8..word_index * 8 + 8].copy_from_slice(&word.to_le_bytes());
    }
    Ok(output)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmHybridKind {
    NoSplit { width: XldmNoSplitWidth, min: i32 },
    Xm123 { min: i32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmHybridStorage {
    pub primary: Vec<u8>,
    pub subsegment: Vec<u8>,
    pub storage_used_units: u32,
    pub storage_alloc_units: u32,
}

pub fn decompress_xldm_hybrid(
    primary: &[u8],
    subsegment: &[u8],
    kind: XldmHybridKind,
    row_count: usize,
    storage_used_units: u32,
    storage_alloc_units: u32,
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<i32>> {
    checked_output(row_count, 4, limits)?;
    if primary.len() > limits.max_input_bytes || subsegment.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    let used = usize::try_from(storage_used_units)
        .map_err(|_| XldmCodecError::IntegerOverflow)?
        .checked_mul(4)
        .ok_or(XldmCodecError::IntegerOverflow)?;
    let allocated = usize::try_from(storage_alloc_units)
        .map_err(|_| XldmCodecError::IntegerOverflow)?
        .checked_mul(4)
        .ok_or(XldmCodecError::IntegerOverflow)?;
    if used > allocated || allocated != primary.len() || used % 8 != 0 {
        return Err(XldmCodecError::Invalid(
            "hybrid RLE storage sizes are inconsistent",
        ));
    }
    if primary[used..].iter().any(|&byte| byte != 0) {
        return Err(XldmCodecError::Invalid(
            "hybrid RLE allocation padding is not zero",
        ));
    }

    if let XldmHybridKind::Xm123 { min } = kind {
        if used != 8 || subsegment != [0u8; 8] {
            return Err(XldmCodecError::Invalid(
                "XM123 requires one primary entry and an eight-byte zero subsegment",
            ));
        }
        let offset = i32::from_le_bytes(primary[0..4].try_into().expect("length checked"));
        let count = i32::from_le_bytes(primary[4..8].try_into().expect("length checked"));
        if offset != -1 || usize::try_from(count).ok() != Some(row_count) {
            return Err(XldmCodecError::Invalid("XM123 primary entry is invalid"));
        }
        let last = i64::from(min)
            .checked_add(i64::try_from(row_count).map_err(|_| XldmCodecError::IntegerOverflow)?)
            .and_then(|value| value.checked_sub(1))
            .unwrap_or(i64::from(min));
        if row_count != 0 && last > 2_000_000_000 {
            return Err(XldmCodecError::Invalid(
                "XM123 indexed value exceeds 2,000,000,000",
            ));
        }
        return (0..row_count)
            .map(|index| {
                i64::from(min)
                    .checked_add(i64::try_from(index).map_err(|_| XldmCodecError::IntegerOverflow)?)
                    .and_then(|value| i32::try_from(value).ok())
                    .ok_or(XldmCodecError::IntegerOverflow)
            })
            .collect();
    }

    let (width, min) = match kind {
        XldmHybridKind::NoSplit { width, min } => (width, min),
        XldmHybridKind::Xm123 { .. } => unreachable!(),
    };
    enum Entry {
        Run(i32, usize),
        Packed(usize),
    }
    let mut entries = Vec::with_capacity(used / 8);
    let mut packed_count = 0usize;
    let mut logical_count = 0usize;
    for entry in primary[..used].chunks_exact(8) {
        let value_or_offset = i32::from_le_bytes(entry[0..4].try_into().expect("exact chunk"));
        let count = i32::from_le_bytes(entry[4..8].try_into().expect("exact chunk"));
        let count = usize::try_from(count)
            .map_err(|_| XldmCodecError::Invalid("hybrid entry count is not positive"))?;
        if count == 0 {
            return Err(XldmCodecError::Invalid("hybrid entry count is zero"));
        }
        logical_count = logical_count
            .checked_add(count)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if logical_count > row_count {
            return Err(XldmCodecError::Invalid(
                "hybrid entries exceed the declared row count",
            ));
        }
        if value_or_offset >= 0 {
            if count < 64 {
                return Err(XldmCodecError::Invalid(
                    "hybrid RLE run is shorter than 64 values",
                ));
            }
            entries.push(Entry::Run(value_or_offset, count));
        } else {
            let expected = i64::try_from(packed_count)
                .map_err(|_| XldmCodecError::IntegerOverflow)?
                .checked_add(1)
                .and_then(|value| value.checked_neg())
                .ok_or(XldmCodecError::IntegerOverflow)?;
            if i64::from(value_or_offset) != expected {
                return Err(XldmCodecError::Invalid(
                    "hybrid bit-packed offset is not contiguous and one-based",
                ));
            }
            packed_count = packed_count
                .checked_add(count)
                .ok_or(XldmCodecError::IntegerOverflow)?;
            entries.push(Entry::Packed(count));
        }
    }
    if logical_count != row_count {
        return Err(XldmCodecError::Invalid(
            "hybrid entries do not cover the declared row count",
        ));
    }
    let packed = decompress_xldm_no_split(subsegment, width, min, packed_count, limits)?;
    let mut packed_position = 0usize;
    let mut output = Vec::with_capacity(row_count);
    for entry in entries {
        match entry {
            Entry::Run(value, count) => output.extend(std::iter::repeat_n(value, count)),
            Entry::Packed(count) => {
                let end = packed_position
                    .checked_add(count)
                    .ok_or(XldmCodecError::IntegerOverflow)?;
                output.extend_from_slice(&packed[packed_position..end]);
                packed_position = end;
            },
        }
    }
    Ok(output)
}

pub fn compress_xldm_hybrid(
    values: &[i32],
    kind: XldmHybridKind,
    limits: XldmCodecLimits,
) -> XldmCodecResult<XldmHybridStorage> {
    checked_output(values.len(), 4, limits)?;
    if let XldmHybridKind::Xm123 { min } = kind {
        for (index, &value) in values.iter().enumerate() {
            let expected = i64::from(min)
                .checked_add(i64::try_from(index).map_err(|_| XldmCodecError::IntegerOverflow)?)
                .ok_or(XldmCodecError::IntegerOverflow)?;
            if i64::from(value) != expected || expected > 2_000_000_000 {
                return Err(XldmCodecError::Invalid(
                    "values do not form a valid XM123 sequence",
                ));
            }
        }
        let count = i32::try_from(values.len()).map_err(|_| XldmCodecError::IntegerOverflow)?;
        let mut primary = Vec::with_capacity(8);
        primary.extend_from_slice(&(-1i32).to_le_bytes());
        primary.extend_from_slice(&count.to_le_bytes());
        return Ok(XldmHybridStorage {
            primary,
            subsegment: vec![0; 8],
            storage_used_units: 2,
            storage_alloc_units: 2,
        });
    }
    let (width, min) = match kind {
        XldmHybridKind::NoSplit { width, min } => (width, min),
        XldmHybridKind::Xm123 { .. } => unreachable!(),
    };
    let mut primary = Vec::new();
    let mut packed = Vec::new();
    let mut position = 0usize;
    while position < values.len() {
        let mut run_end = position + 1;
        while run_end < values.len() && values[run_end] == values[position] {
            run_end += 1;
        }
        let run_length = run_end - position;
        if run_length >= 64 && values[position] >= 0 {
            primary.extend_from_slice(&values[position].to_le_bytes());
            primary.extend_from_slice(
                &i32::try_from(run_length)
                    .map_err(|_| XldmCodecError::IntegerOverflow)?
                    .to_le_bytes(),
            );
            position = run_end;
            continue;
        }
        let packed_start = position;
        position = run_end;
        while position < values.len() {
            let mut next_end = position + 1;
            while next_end < values.len() && values[next_end] == values[position] {
                next_end += 1;
            }
            if next_end - position >= 64 && values[position] >= 0 {
                break;
            }
            position = next_end;
        }
        let offset = i32::try_from(
            packed
                .len()
                .checked_add(1)
                .ok_or(XldmCodecError::IntegerOverflow)?,
        )
        .map_err(|_| XldmCodecError::IntegerOverflow)?
        .checked_neg()
        .ok_or(XldmCodecError::IntegerOverflow)?;
        let count = position - packed_start;
        primary.extend_from_slice(&offset.to_le_bytes());
        primary.extend_from_slice(
            &i32::try_from(count)
                .map_err(|_| XldmCodecError::IntegerOverflow)?
                .to_le_bytes(),
        );
        packed.extend_from_slice(&values[packed_start..position]);
    }
    if primary.len() > limits.max_output_bytes {
        return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
    }
    let subsegment = compress_xldm_no_split(&packed, width, min, limits)?;
    let units = u32::try_from(primary.len() / 4).map_err(|_| XldmCodecError::IntegerOverflow)?;
    Ok(XldmHybridStorage {
        primary,
        subsegment,
        storage_used_units: units,
        storage_alloc_units: units,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmHuffmanMode {
    MultipleCharacterSets,
    SingleCharacterSet { upper_byte: u8 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmHuffmanEncoded {
    pub total_bits: u32,
    pub offsets: Vec<u32>,
    pub buffer: Vec<u8>,
}

#[derive(Clone, Copy)]
struct HuffmanCode {
    bits: u16,
    len: u8,
}

fn huffman_lengths(encode_array: &[u8]) -> XldmCodecResult<[u8; 256]> {
    if encode_array.len() != 128 {
        return Err(XldmCodecError::Invalid(
            "Huffman encode array is not 128 bytes",
        ));
    }
    let mut lengths = [0u8; 256];
    for (index, &byte) in encode_array.iter().enumerate() {
        lengths[index * 2] = byte & 0x0f;
        lengths[index * 2 + 1] = byte >> 4;
    }
    if lengths.iter().any(|&length| length == 1 || length > 15) {
        return Err(XldmCodecError::Invalid(
            "Huffman code length is outside 2 through 15",
        ));
    }
    Ok(lengths)
}

fn canonical_codes(encode_array: &[u8]) -> XldmCodecResult<[Option<HuffmanCode>; 256]> {
    let lengths = huffman_lengths(encode_array)?;
    let mut counts = [0u32; 16];
    for &length in &lengths {
        if length != 0 {
            counts[usize::from(length)] += 1;
        }
    }
    if counts.iter().sum::<u32>() < 2 {
        return Err(XldmCodecError::Invalid(
            "Huffman alphabet has fewer than two symbols",
        ));
    }
    let mut next = [0u32; 16];
    let mut code = 0u32;
    for bits in 1..=15 {
        code = code
            .checked_add(counts[bits - 1])
            .and_then(|value| value.checked_shl(1))
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if code
            .checked_add(counts[bits])
            .ok_or(XldmCodecError::IntegerOverflow)?
            > (1u32 << bits)
        {
            return Err(XldmCodecError::Invalid(
                "Huffman code lengths are oversubscribed",
            ));
        }
        next[bits] = code;
    }
    let max_len = lengths.iter().copied().max().unwrap_or(0);
    if max_len == 0
        || next[usize::from(max_len)] + counts[usize::from(max_len)] != (1u32 << max_len)
    {
        return Err(XldmCodecError::Invalid(
            "Huffman code lengths do not form a complete prefix tree",
        ));
    }
    let mut result = [None; 256];
    for (symbol, &length) in lengths.iter().enumerate() {
        if length != 0 {
            let slot = &mut next[usize::from(length)];
            result[symbol] = Some(HuffmanCode {
                bits: u16::try_from(*slot).map_err(|_| XldmCodecError::IntegerOverflow)?,
                len: length,
            });
            *slot += 1;
        }
    }
    Ok(result)
}

fn read_bit(buffer: &[u8], position: usize) -> u8 {
    (buffer[position / 8] >> (7 - position % 8)) & 1
}

pub fn decompress_xldm_huffman_strings(
    encode_array: &[u8],
    decode_bits: u32,
    total_bits: u32,
    buffer: &[u8],
    offsets: &[u32],
    mode: XldmHuffmanMode,
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<Vec<u8>>> {
    if decode_bits > 12 {
        return Err(XldmCodecError::Invalid("Huffman DecodeBits exceeds 12"));
    }
    if offsets.len() > limits.max_strings {
        return Err(XldmCodecError::LimitExceeded("max_strings"));
    }
    if buffer.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    let total_bits = usize::try_from(total_bits).map_err(|_| XldmCodecError::IntegerOverflow)?;
    let required_bytes = total_bits
        .checked_add(7)
        .ok_or(XldmCodecError::IntegerOverflow)?
        / 8;
    if buffer.len() != required_bytes {
        return Err(XldmCodecError::Invalid(
            "Huffman buffer size does not match TotalBits",
        ));
    }
    if total_bits % 8 != 0
        && buffer
            .last()
            .is_some_and(|byte| byte & ((1u8 << (8 - total_bits % 8)) - 1) != 0)
    {
        return Err(XldmCodecError::Invalid(
            "Huffman tail padding bits are not zero",
        ));
    }
    let codes = canonical_codes(encode_array)?;
    let mut previous = 0usize;
    for &offset in offsets {
        let offset = usize::try_from(offset).map_err(|_| XldmCodecError::IntegerOverflow)?;
        if offset < previous || offset > total_bits {
            return Err(XldmCodecError::Invalid(
                "Huffman string offsets are not ordered or exceed TotalBits",
            ));
        }
        previous = offset;
    }
    let multiplier = if matches!(mode, XldmHuffmanMode::SingleCharacterSet { .. }) {
        2
    } else {
        1
    };
    let mut total_output = 0usize;
    let mut strings = Vec::with_capacity(offsets.len());
    for (index, &start) in offsets.iter().enumerate() {
        let mut position = usize::try_from(start).map_err(|_| XldmCodecError::IntegerOverflow)?;
        let end = offsets
            .get(index + 1)
            .copied()
            .map(|value| usize::try_from(value).map_err(|_| XldmCodecError::IntegerOverflow))
            .transpose()?
            .unwrap_or(total_bits);
        let mut decoded = Vec::new();
        while position < end {
            let mut code = 0u16;
            let mut matched = None;
            for length in 1..=15u8 {
                if position >= end {
                    return Err(XldmCodecError::Invalid(
                        "Huffman string ends in the middle of a code",
                    ));
                }
                code = (code << 1) | u16::from(read_bit(buffer, position));
                position += 1;
                if let Some((symbol, _)) = codes.iter().enumerate().find(|(_, candidate)| {
                    candidate
                        .is_some_and(|candidate| candidate.len == length && candidate.bits == code)
                }) {
                    matched = Some(symbol as u8);
                    break;
                }
            }
            let symbol = matched.ok_or(XldmCodecError::Invalid(
                "Huffman bit sequence has no symbol",
            ))?;
            decoded.push(symbol);
        }
        let decoded_size = decoded
            .len()
            .checked_mul(multiplier)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        total_output = total_output
            .checked_add(decoded_size)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if total_output > limits.max_output_bytes {
            return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
        }
        if let XldmHuffmanMode::SingleCharacterSet { upper_byte } = mode {
            let mut utf16 = Vec::with_capacity(decoded_size);
            for low_byte in decoded {
                utf16.push(low_byte);
                utf16.push(upper_byte);
            }
            strings.push(utf16);
        } else {
            strings.push(decoded);
        }
    }
    Ok(strings)
}

pub fn compress_xldm_huffman_strings(
    encode_array: &[u8],
    strings: &[&[u8]],
    mode: XldmHuffmanMode,
    limits: XldmCodecLimits,
) -> XldmCodecResult<XldmHuffmanEncoded> {
    if strings.len() > limits.max_strings {
        return Err(XldmCodecError::LimitExceeded("max_strings"));
    }
    let codes = canonical_codes(encode_array)?;
    let mut offsets = Vec::with_capacity(strings.len());
    let mut buffer = Vec::<u8>::new();
    let mut bit_count = 0usize;
    for string in strings {
        offsets.push(u32::try_from(bit_count).map_err(|_| XldmCodecError::IntegerOverflow)?);
        let symbols: Vec<u8> = match mode {
            XldmHuffmanMode::MultipleCharacterSets => string.to_vec(),
            XldmHuffmanMode::SingleCharacterSet { upper_byte } => {
                if string.len() % 2 != 0 || string.chunks_exact(2).any(|pair| pair[1] != upper_byte)
                {
                    return Err(XldmCodecError::Invalid(
                        "single-character-set input is not matching UTF-16LE",
                    ));
                }
                string.chunks_exact(2).map(|pair| pair[0]).collect()
            },
        };
        for symbol in symbols {
            let code = codes[usize::from(symbol)].ok_or(XldmCodecError::Invalid(
                "string uses a symbol absent from the Huffman alphabet",
            ))?;
            for shift in (0..code.len).rev() {
                if bit_count / 8 >= limits.max_output_bytes {
                    return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
                }
                if bit_count.is_multiple_of(8) {
                    buffer.push(0);
                }
                let bit = ((code.bits >> shift) & 1) as u8;
                let byte_index = bit_count / 8;
                buffer[byte_index] |= bit << (7 - bit_count % 8);
                bit_count = bit_count
                    .checked_add(1)
                    .ok_or(XldmCodecError::IntegerOverflow)?;
            }
        }
    }
    Ok(XldmHuffmanEncoded {
        total_bits: u32::try_from(bit_count).map_err(|_| XldmCodecError::IntegerOverflow)?,
        offsets,
        buffer,
    })
}

fn take<'a>(input: &'a [u8], position: &mut usize, count: usize) -> XldmCodecResult<&'a [u8]> {
    let end = position
        .checked_add(count)
        .ok_or(XldmCodecError::IntegerOverflow)?;
    let value = input
        .get(*position..end)
        .ok_or(XldmCodecError::Invalid("Xpress stream is truncated"))?;
    *position = end;
    Ok(value)
}

pub fn decompress_xldm_xpress_block(
    input: &[u8],
    expected_size: usize,
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<u8>> {
    if input.len() > XPRESS_BLOCK_MAX || expected_size > XPRESS_BLOCK_MAX {
        return Err(XldmCodecError::Invalid("Xpress block exceeds 65,535 bytes"));
    }
    checked_output(expected_size, 1, limits)?;
    let mut input_position = 0usize;
    let mut flags = 0u32;
    let mut flag_count = 0u8;
    let mut saved_nibble = None::<u8>;
    let mut output = Vec::with_capacity(expected_size);
    loop {
        if flag_count == 0 {
            flags = u32::from_le_bytes(
                take(input, &mut input_position, 4)?
                    .try_into()
                    .expect("length checked"),
            );
            flag_count = 32;
        }
        flag_count -= 1;
        let is_match = flags & (1u32 << flag_count) != 0;
        if !is_match {
            if output.len() >= expected_size {
                return Err(XldmCodecError::Invalid(
                    "Xpress block expands beyond declared size",
                ));
            }
            output.push(take(input, &mut input_position, 1)?[0]);
            continue;
        }
        if input_position == input.len() {
            if output.len() != expected_size {
                return Err(XldmCodecError::Invalid(
                    "Xpress end marker precedes declared size",
                ));
            }
            return Ok(output);
        }
        let match_bytes = u16::from_le_bytes(
            take(input, &mut input_position, 2)?
                .try_into()
                .expect("length checked"),
        );
        let offset = usize::from(match_bytes >> 3) + 1;
        let mut length = usize::from(match_bytes & 7);
        if length == 7 {
            length = if let Some(nibble) = saved_nibble.take() {
                usize::from(nibble)
            } else {
                let byte = take(input, &mut input_position, 1)?[0];
                saved_nibble = Some(byte >> 4);
                usize::from(byte & 0x0f)
            };
            if length == 15 {
                length = usize::from(take(input, &mut input_position, 1)?[0]);
                if length == 255 {
                    length = usize::from(u16::from_le_bytes(
                        take(input, &mut input_position, 2)?
                            .try_into()
                            .expect("length checked"),
                    ));
                    if length == 0 {
                        length = usize::try_from(u32::from_le_bytes(
                            take(input, &mut input_position, 4)?
                                .try_into()
                                .expect("length checked"),
                        ))
                        .map_err(|_| XldmCodecError::IntegerOverflow)?;
                    }
                    if length < 22 {
                        return Err(XldmCodecError::Invalid(
                            "Xpress extended match length is non-canonical",
                        ));
                    }
                    length -= 22;
                }
                length = length
                    .checked_add(15)
                    .ok_or(XldmCodecError::IntegerOverflow)?;
            }
            length = length
                .checked_add(7)
                .ok_or(XldmCodecError::IntegerOverflow)?;
        }
        length = length
            .checked_add(3)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if offset > output.len() {
            return Err(XldmCodecError::Invalid(
                "Xpress match offset precedes output",
            ));
        }
        let new_size = output
            .len()
            .checked_add(length)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if new_size > expected_size || new_size > limits.max_output_bytes {
            return Err(XldmCodecError::LimitExceeded("Xpress expanded block size"));
        }
        for _ in 0..length {
            let byte = output[output.len() - offset];
            output.push(byte);
        }
    }
}

pub fn compress_xldm_xpress_block_literals(
    input: &[u8],
    limits: XldmCodecLimits,
) -> XldmCodecResult<Vec<u8>> {
    if input.len() > XPRESS_BLOCK_MAX {
        return Err(XldmCodecError::Invalid("Xpress block exceeds 65,535 bytes"));
    }
    if input.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    let groups = input
        .len()
        .checked_add(31)
        .ok_or(XldmCodecError::IntegerOverflow)?
        / 32;
    let terminal_group = usize::from(input.len().is_multiple_of(32));
    let capacity = input
        .len()
        .checked_add(
            groups
                .checked_add(terminal_group)
                .ok_or(XldmCodecError::IntegerOverflow)?
                .checked_mul(4)
                .ok_or(XldmCodecError::IntegerOverflow)?,
        )
        .ok_or(XldmCodecError::IntegerOverflow)?;
    if capacity > XPRESS_BLOCK_MAX || capacity > limits.max_output_bytes {
        return Err(XldmCodecError::LimitExceeded(
            "literal-only Xpress compressed block size",
        ));
    }
    let mut output = Vec::with_capacity(capacity);
    for group in input.chunks(32) {
        let flags = if group.len() == 32 {
            0
        } else {
            (1u32 << (32 - group.len())) - 1
        };
        output.extend_from_slice(&flags.to_le_bytes());
        output.extend_from_slice(group);
    }
    if input.len().is_multiple_of(32) {
        output.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    Ok(output)
}

pub fn decompress_xldm_xpress(input: &[u8], limits: XldmCodecLimits) -> XldmCodecResult<Vec<u8>> {
    if input.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    let mut position = 0usize;
    let mut output = Vec::new();
    while position < input.len() {
        let original = i32::from_le_bytes(
            take(input, &mut position, 4)?
                .try_into()
                .expect("length checked"),
        );
        let compressed = i32::from_le_bytes(
            take(input, &mut position, 4)?
                .try_into()
                .expect("length checked"),
        );
        let original = usize::try_from(original)
            .map_err(|_| XldmCodecError::Invalid("negative Xpress original block size"))?;
        let compressed = usize::try_from(compressed)
            .map_err(|_| XldmCodecError::Invalid("negative Xpress compressed block size"))?;
        if original > XPRESS_BLOCK_MAX || compressed > XPRESS_BLOCK_MAX {
            return Err(XldmCodecError::Invalid(
                "Xpress block header exceeds 65,535 bytes",
            ));
        }
        let new_size = output
            .len()
            .checked_add(original)
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if new_size > limits.max_output_bytes {
            return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
        }
        let block = take(input, &mut position, compressed)?;
        output.extend_from_slice(&decompress_xldm_xpress_block(block, original, limits)?);
    }
    Ok(output)
}

pub fn compress_xldm_xpress(input: &[u8], limits: XldmCodecLimits) -> XldmCodecResult<Vec<u8>> {
    if input.len() > limits.max_input_bytes {
        return Err(XldmCodecError::LimitExceeded("max_input_bytes"));
    }
    let mut output = Vec::new();
    for block in input.chunks(XPRESS_LITERAL_BLOCK) {
        let compressed = compress_xldm_xpress_block_literals(block, limits)?;
        let new_size = output
            .len()
            .checked_add(8)
            .and_then(|value| value.checked_add(compressed.len()))
            .ok_or(XldmCodecError::IntegerOverflow)?;
        if new_size > limits.max_output_bytes {
            return Err(XldmCodecError::LimitExceeded("max_output_bytes"));
        }
        output.extend_from_slice(
            &i32::try_from(block.len())
                .map_err(|_| XldmCodecError::IntegerOverflow)?
                .to_le_bytes(),
        );
        output.extend_from_slice(
            &i32::try_from(compressed.len())
                .map_err(|_| XldmCodecError::IntegerOverflow)?
                .to_le_bytes(),
        );
        output.extend_from_slice(&compressed);
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits() -> XldmCodecLimits {
        XldmCodecLimits {
            max_output_bytes: 1_000_000,
            max_values: 100_000,
            max_strings: 1000,
            max_input_bytes: 1_000_000,
        }
    }

    #[test]
    fn no_split_known_answers_cover_every_width() {
        for raw_width in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 21, 32] {
            let width = XldmNoSplitWidth::try_from(raw_width).unwrap();
            let min = if raw_width == 32 { i32::MIN } else { 7 };
            let max_delta = (1u64 << raw_width) - 1;
            let values = [
                min,
                min + 1,
                i32::try_from(i64::from(min) + i64::try_from(max_delta / 2).unwrap()).unwrap(),
                i32::try_from(i64::from(min) + i64::try_from(max_delta).unwrap()).unwrap(),
            ];
            let encoded = compress_xldm_no_split(&values, width, min, limits()).unwrap();
            assert_eq!(
                decompress_xldm_no_split(&encoded, width, min, values.len(), limits()).unwrap(),
                values
            );
            assert_eq!(encoded.len() % 8, 0);
        }
    }

    #[test]
    fn no_split_rejects_bad_size_and_range() {
        assert!(decompress_xldm_no_split(&[0; 7], XldmNoSplitWidth::W1, 0, 1, limits()).is_err());
        assert!(compress_xldm_no_split(&[-1], XldmNoSplitWidth::W1, 0, limits()).is_err());
    }

    #[test]
    fn appendix_a_compression_mask_known_answers_and_complete_domain() {
        assert_eq!(xldm_no_split_compression_mask(0, 0), Some(0));
        assert_eq!(
            xldm_no_split_compression_mask(1, 0),
            Some(0xffff_ffff_ffff_fffe)
        );
        assert_eq!(
            xldm_no_split_compression_mask(1, 63),
            Some(0x7fff_ffff_ffff_ffff)
        );
        assert_eq!(
            xldm_no_split_compression_mask(2, 1),
            Some(0xffff_ffff_ffff_fff9)
        );
        assert_eq!(
            xldm_no_split_compression_mask(50, 0),
            Some(0xfffc_0000_0000_0000)
        );
        assert_eq!(
            xldm_no_split_compression_mask(50, 1),
            Some(0xfff8_0000_0000_0001)
        );
        assert_eq!(
            xldm_no_split_compression_mask(50, 14),
            Some(0x0000_0000_0000_3fff)
        );
        assert_eq!(xldm_no_split_compression_mask(50, 15), Some(0));
        assert_eq!(xldm_no_split_compression_mask(64, 0), Some(0));
        assert_eq!(xldm_no_split_compression_mask(64, 63), Some(0));
        assert_eq!(xldm_no_split_compression_mask(65, 0), None);
        assert_eq!(xldm_no_split_compression_mask(1, 64), None);

        for field_bits in 0..=64u8 {
            for bit_offset in 0..64u8 {
                let expected = if field_bits == 0
                    || field_bits == 64
                    || u16::from(field_bits) + u16::from(bit_offset) > 64
                {
                    0
                } else {
                    !(((1u64 << field_bits) - 1) << bit_offset)
                };
                assert_eq!(
                    xldm_no_split_compression_mask(field_bits, bit_offset),
                    Some(expected)
                );
            }
        }
    }

    #[test]
    fn hybrid_mixed_round_trip_and_adversarial_offsets() {
        let mut values = vec![3; 64];
        values.extend([1, 2, 4, 5]);
        values.extend(std::iter::repeat_n(9, 70));
        let storage = compress_xldm_hybrid(
            &values,
            XldmHybridKind::NoSplit {
                width: XldmNoSplitWidth::W4,
                min: 0,
            },
            limits(),
        )
        .unwrap();
        let decoded = decompress_xldm_hybrid(
            &storage.primary,
            &storage.subsegment,
            XldmHybridKind::NoSplit {
                width: XldmNoSplitWidth::W4,
                min: 0,
            },
            values.len(),
            storage.storage_used_units,
            storage.storage_alloc_units,
            limits(),
        )
        .unwrap();
        assert_eq!(decoded, values);
        let mut corrupt = storage.primary.clone();
        corrupt[8..12].copy_from_slice(&(-2i32).to_le_bytes());
        assert!(
            decompress_xldm_hybrid(
                &corrupt,
                &storage.subsegment,
                XldmHybridKind::NoSplit {
                    width: XldmNoSplitWidth::W4,
                    min: 0
                },
                values.len(),
                storage.storage_used_units,
                storage.storage_alloc_units,
                limits()
            )
            .is_err()
        );
    }

    #[test]
    fn section_3_2_hybrid_idf_known_answer() {
        let mut primary = vec![0u8; 128];
        for (entry, (value, count)) in [(3i32, 1024i32), (4, 1024), (5, 1024), (6, 1024)]
            .into_iter()
            .enumerate()
        {
            let offset = entry * 8;
            primary[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            primary[offset + 4..offset + 8].copy_from_slice(&count.to_le_bytes());
        }
        primary[32..36].copy_from_slice(&(-1i32).to_le_bytes());
        primary[36..40].copy_from_slice(&8i32.to_le_bytes());
        let subsegment = [0xac, 0xef, 0xfb, 0, 0, 0, 0, 0];
        let decoded = decompress_xldm_hybrid(
            &primary,
            &subsegment,
            XldmHybridKind::NoSplit {
                width: XldmNoSplitWidth::W3,
                min: 3,
            },
            4104,
            10,
            32,
            limits(),
        )
        .unwrap();
        assert_eq!(&decoded[0..1024], vec![3; 1024]);
        assert_eq!(&decoded[1024..2048], vec![4; 1024]);
        assert_eq!(&decoded[2048..3072], vec![5; 1024]);
        assert_eq!(&decoded[3072..4096], vec![6; 1024]);
        assert_eq!(&decoded[4096..], [7, 8, 9, 10, 9, 10, 9, 10]);

        let encoded = compress_xldm_hybrid(
            &decoded,
            XldmHybridKind::NoSplit {
                width: XldmNoSplitWidth::W3,
                min: 3,
            },
            limits(),
        )
        .unwrap();
        assert_eq!(encoded.primary, primary[..40]);
        assert_eq!(encoded.subsegment, subsegment);
        assert_eq!(encoded.storage_used_units, 10);

        primary[127] = 1;
        assert!(
            decompress_xldm_hybrid(
                &primary,
                &subsegment,
                XldmHybridKind::NoSplit {
                    width: XldmNoSplitWidth::W3,
                    min: 3
                },
                4104,
                10,
                32,
                limits(),
            )
            .is_err()
        );
    }

    #[test]
    fn xm123_round_trip_and_bound() {
        let values: Vec<i32> = (100..110).collect();
        let storage =
            compress_xldm_hybrid(&values, XldmHybridKind::Xm123 { min: 100 }, limits()).unwrap();
        assert_eq!(
            decompress_xldm_hybrid(
                &storage.primary,
                &storage.subsegment,
                XldmHybridKind::Xm123 { min: 100 },
                10,
                2,
                2,
                limits()
            )
            .unwrap(),
            values
        );
        assert!(
            decompress_xldm_hybrid(
                &storage.primary,
                &storage.subsegment,
                XldmHybridKind::Xm123 { min: 2_000_000_000 },
                10,
                2,
                2,
                limits()
            )
            .is_err()
        );
    }

    fn example_huffman_table() -> [u8; 128] {
        let mut lengths = [0u8; 256];
        for symbol in *b"FMam" {
            lengths[usize::from(symbol)] = 3;
        }
        for symbol in *b"el" {
            lengths[usize::from(symbol)] = 2;
        }
        let mut packed = [0u8; 128];
        for index in 0..128 {
            packed[index] = lengths[index * 2] | (lengths[index * 2 + 1] << 4);
        }
        packed
    }

    #[test]
    fn huffman_spec_known_answer_female_male() {
        let table = example_huffman_table();
        let decoded = decompress_xldm_huffman_strings(
            &table,
            8,
            25,
            &[0x87, 0xc9, 0x72, 0x00],
            &[0, 15],
            XldmHuffmanMode::MultipleCharacterSets,
            limits(),
        )
        .unwrap();
        assert_eq!(decoded, [b"Female".to_vec(), b"Male".to_vec()]);
        let encoded = compress_xldm_huffman_strings(
            &table,
            &[b"Female", b"Male"],
            XldmHuffmanMode::MultipleCharacterSets,
            limits(),
        )
        .unwrap();
        assert_eq!(encoded.total_bits, 25);
        assert_eq!(encoded.offsets, [0, 15]);
        assert_eq!(encoded.buffer, [0x87, 0xc9, 0x72, 0x00]);
    }

    #[test]
    fn huffman_single_charset_and_invalid_tree() {
        let table = example_huffman_table();
        let encoded = compress_xldm_huffman_strings(
            &table,
            &[&[b'F', 0, b'e', 0]],
            XldmHuffmanMode::SingleCharacterSet { upper_byte: 0 },
            limits(),
        )
        .unwrap();
        assert_eq!(
            decompress_xldm_huffman_strings(
                &table,
                12,
                encoded.total_bits,
                &encoded.buffer,
                &encoded.offsets,
                XldmHuffmanMode::SingleCharacterSet { upper_byte: 0 },
                limits()
            )
            .unwrap()[0],
            [b'F', 0, b'e', 0]
        );
        assert!(
            decompress_xldm_huffman_strings(
                &[0; 128],
                0,
                0,
                &[],
                &[],
                XldmHuffmanMode::MultipleCharacterSets,
                limits()
            )
            .is_err()
        );
    }

    #[test]
    fn xpress_literal_round_trip_and_match_known_answer() {
        let literal = compress_xldm_xpress_block_literals(b"abc", limits()).unwrap();
        assert_eq!(literal, [0xff, 0xff, 0xff, 0x1f, b'a', b'b', b'c']);
        assert_eq!(
            decompress_xldm_xpress_block(&literal, 3, limits()).unwrap(),
            b"abc"
        );
        let match_block = [0xff, 0xff, 0xff, 0x7f, b'a', 0x02, 0x00];
        assert_eq!(
            decompress_xldm_xpress_block(&match_block, 6, limits()).unwrap(),
            b"aaaaaa"
        );
    }

    #[test]
    fn xpress_framing_round_trip_and_bomb_guard() {
        let input = vec![0x5a; 100_000];
        let encoded = compress_xldm_xpress(
            &input,
            XldmCodecLimits {
                max_output_bytes: 200_000,
                ..limits()
            },
        )
        .unwrap();
        assert_eq!(
            decompress_xldm_xpress(
                &encoded,
                XldmCodecLimits {
                    max_output_bytes: 100_000,
                    ..limits()
                }
            )
            .unwrap(),
            input
        );
        let mut bomb = Vec::new();
        bomb.extend_from_slice(&65_535i32.to_le_bytes());
        bomb.extend_from_slice(&4i32.to_le_bytes());
        bomb.extend_from_slice(&u32::MAX.to_le_bytes());
        assert!(
            decompress_xldm_xpress(
                &bomb,
                XldmCodecLimits {
                    max_output_bytes: 16,
                    ..limits()
                }
            )
            .is_err()
        );
    }
}
