//! Explicit Xpress stream framing; no ambiguous format auto-detection.

use super::{
    CodecError, CodecLimits, CodecResult, XPRESS_BLOCK_MAX, checked_output,
    decode_xpress_block_into, take,
};

/// Header and raw-block rules for an Xpress stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XpressFraming {
    /// MS-WUSP 2.1.1: two signed little-endian 32-bit block lengths.
    Wusp32,
    /// Observed native version-150 tabular members: two little-endian 16-bit
    /// lengths; equal lengths indicate a stored, uncompressed block.
    /// This does not enable version-150 outer-storage or metadata interpretation.
    Tabular16,
}

struct Block<'a> {
    bytes: &'a [u8],
    original: usize,
    raw: bool,
}

fn block<'a>(
    input: &'a [u8],
    position: &mut usize,
    framing: XpressFraming,
) -> CodecResult<Block<'a>> {
    let (original, compressed) = match framing {
        XpressFraming::Wusp32 => {
            let header = take(input, position, 8)?;
            let original = i32::from_le_bytes([header[0], header[1], header[2], header[3]]);
            let compressed = i32::from_le_bytes([header[4], header[5], header[6], header[7]]);
            (
                usize::try_from(original)
                    .map_err(|_| CodecError::Invalid("negative Xpress original block size"))?,
                usize::try_from(compressed)
                    .map_err(|_| CodecError::Invalid("negative Xpress compressed block size"))?,
            )
        },
        XpressFraming::Tabular16 => {
            let header = take(input, position, 4)?;
            (
                usize::from(u16::from_le_bytes([header[0], header[1]])),
                usize::from(u16::from_le_bytes([header[2], header[3]])),
            )
        },
    };
    if original > XPRESS_BLOCK_MAX || compressed > XPRESS_BLOCK_MAX {
        return Err(CodecError::Invalid(
            "Xpress block header exceeds 65,535 bytes",
        ));
    }
    let bytes = take(input, position, compressed)?;
    Ok(Block {
        bytes,
        original,
        raw: framing == XpressFraming::Tabular16 && original == compressed,
    })
}

/// Decode explicitly selected stream framing with aggregate size admission.
///
/// The complete frame table is checked before result allocation. Blocks decode
/// directly into one output buffer, and match offsets cannot cross block
/// boundaries. Encoded input is borrowed and never copied into a frame catalog.
pub fn decompress_xpress_framed(
    input: &[u8],
    framing: XpressFraming,
    limits: CodecLimits,
) -> CodecResult<Vec<u8>> {
    let size = stream_size(input, framing, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|_| CodecError::LimitExceeded("Xpress output allocation"))?;
    let mut position = 0;
    while position < input.len() {
        let block = block(input, &mut position, framing)?;
        if block.raw {
            output.extend_from_slice(block.bytes);
        } else {
            decode_xpress_block_into(block.bytes, block.original, limits, &mut output)?;
        }
    }
    Ok(output)
}

pub(super) fn stream_size(
    input: &[u8],
    framing: XpressFraming,
    limits: CodecLimits,
) -> CodecResult<usize> {
    if input.len() > limits.max_input_bytes {
        return Err(CodecError::LimitExceeded("max_input_bytes"));
    }
    let mut position = 0usize;
    let mut size = 0usize;
    while position < input.len() {
        let block = block(input, &mut position, framing)?;
        checked_output(block.original, 1, limits)?;
        size = size
            .checked_add(block.original)
            .ok_or(CodecError::IntegerOverflow)?;
        if size > limits.max_output_bytes {
            return Err(CodecError::LimitExceeded("max_output_bytes"));
        }
    }
    Ok(size)
}
