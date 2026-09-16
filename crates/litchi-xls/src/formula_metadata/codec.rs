//! BIFF8 Formula-record field codec.

use crate::records::FormulaValue;
use crate::utils;
use crate::{Error, Result};

use super::Metadata;
use super::extra::{retain_formula_extra, validate_formula_extra};
use super::validation::{
    FORMULA_FIXED_SIZE, FlagDefect, MAX_FORMULA_PAYLOAD, decode_flags, decode_flags_preserving,
    invalid,
};

/// Parsed cell and formula fields needed by `CellRecord`.
#[derive(Debug)]
pub(crate) struct Parsed {
    pub(crate) row: u16,
    pub(crate) col: u16,
    pub(crate) xf_index: u16,
    pub(crate) value: FormulaValue,
    pub(crate) metadata: Metadata,
    pub(crate) formula: Vec<u8>,
}

/// The fields of a `Formula` payload that are checked before anything is
/// copied, with the token stream and the `RgbExtra` suffix left borrowed.
///
/// Splitting the record here is what lets the materializing parse and the
/// measuring walk share one implementation of the size, value and token-bound
/// checks: `parse_record_with` adds the `.to_vec()` and the retained
/// `Ancillary` on top of this, and [`measure_record`] adds neither.
struct Framed<'a> {
    row: u16,
    col: u16,
    xf_index: u16,
    value: FormulaValue,
    flags: u16,
    calculation_cache: u32,
    tokens: &'a [u8],
    extra: &'a [u8],
}

/// Check a `Formula` payload's fixed header, cached value and token bounds.
///
/// This is every refusal the record can produce before `RgbExtra` is reached,
/// in the order it produced them before the split.
///
/// Inlined deliberately. `Framed` carries a `FormulaValue`, which owns a
/// `String` in one variant, so an out-of-line call returns a droppable
/// temporary through memory; measured at +595,720 instructions on one
/// `15228.xls` text extraction before this attribute was added.
#[inline]
fn frame_record(data: &[u8]) -> Result<Framed<'_>> {
    if data.len() < FORMULA_FIXED_SIZE {
        return Err(Error::InvalidLength {
            expected: FORMULA_FIXED_SIZE,
            found: data.len(),
        });
    }
    if data.len() > MAX_FORMULA_PAYLOAD {
        return Err(invalid(format!(
            "Formula payload exceeds the BIFF8 limit of {MAX_FORMULA_PAYLOAD} bytes"
        )));
    }

    let row = read_u16(data, 0);
    let col = read_u16(data, 2);
    let xf_index = read_u16(data, 4);
    let value = utils::parse_formula_value(&data[6..14])?;
    let flags = read_u16(data, 14);
    let calculation_cache = read_u32(data, 16);
    let token_len = usize::from(read_u16(data, 20));
    let formula_end = FORMULA_FIXED_SIZE
        .checked_add(token_len)
        .ok_or_else(|| invalid("Formula token length overflows"))?;
    if formula_end > data.len() {
        return Err(Error::InvalidLength {
            expected: formula_end,
            found: data.len(),
        });
    }
    Ok(Framed {
        row,
        col,
        xf_index,
        value,
        flags,
        calculation_cache,
        tokens: &data[FORMULA_FIXED_SIZE..formula_end],
        extra: &data[formula_end..],
    })
}

/// The two token-stream checks that follow `RgbExtra` retention.
#[inline]
fn check_token_stream(framed: &Framed<'_>) -> Result<()> {
    // A string-valued Formula deliberately carries no Rgce bytes: its
    // cached result is supplied by the immediately following String record.
    // All other FormulaValue variants require an actual token stream.
    if framed.tokens.is_empty() && !matches!(framed.value, FormulaValue::StringPending) {
        return Err(invalid("Formula token stream cannot be empty"));
    }
    if framed.tokens.is_empty() && !framed.extra.is_empty() {
        return Err(invalid("Formula ancillary bytes require a token stream"));
    }
    Ok(())
}

/// Parse the payload of a BIFF8 `Formula` record.
pub(crate) fn parse_record(data: &[u8]) -> Result<Parsed> {
    parse_record_with(data, false).map(|(parsed, _)| parsed)
}

/// Parse a Formula while preserving the one explicitly modeled producer
/// defect. All framing, size, reserved-bit, and value checks remain strict.
pub(crate) fn parse_record_preserving(data: &[u8]) -> Result<(Parsed, Option<FlagDefect>)> {
    parse_record_with(data, true)
}

/// Run every check [`parse_record`] runs, without retaining anything.
///
/// Change 0576's measure-only pattern, applied to the `Formula` record: a
/// worksheet scan that is looking for one cell frames a `Formula` in order to
/// learn its position and its XF index, and drops the `Vec<u8>` token copy and
/// the `Metadata` it also built. The checks are the shared [`frame_record`] and
/// [`check_token_stream`], plus `validate_formula_extra` where the parse calls
/// `retain_formula_extra` — the same scan of the same `RgbExtra` structures,
/// stopping before the copy that retains them — and the same `decode_flags`,
/// whose `Metadata` is discarded. What comes back is the cell position and the
/// XF index, which is all a scan needs from a record it is not keeping.
///
/// One error is unreachable here and reachable in `parse_record`:
/// `Ancillary::new`'s `Error::Allocation("retaining Formula RgbExtra")`, which
/// needs a reservation of at most `MAX_FORMULA_PAYLOAD` bytes to fail.
pub(crate) fn measure_record(data: &[u8]) -> Result<(u16, u16, u16)> {
    let framed = frame_record(data)?;
    validate_formula_extra(framed.tokens, framed.extra)?;
    check_token_stream(&framed)?;
    decode_flags(framed.flags, framed.tokens)?;
    Ok((framed.row, framed.col, framed.xf_index))
}

fn parse_record_with(data: &[u8], preserve_defect: bool) -> Result<(Parsed, Option<FlagDefect>)> {
    let framed = frame_record(data)?;
    let formula = framed.tokens.to_vec();
    let ancillary = retain_formula_extra(framed.row, framed.col, &formula, framed.extra)?;
    check_token_stream(&framed)?;
    let (mut metadata, defect) = if preserve_defect {
        decode_flags_preserving(framed.flags, &formula)?
    } else {
        (decode_flags(framed.flags, &formula)?, None)
    };
    metadata = metadata.with_calculation_cache(framed.calculation_cache);
    if let Some(ancillary) = ancillary {
        metadata = metadata.with_ancillary(ancillary);
    }

    Ok((
        Parsed {
            row: framed.row,
            col: framed.col,
            xf_index: framed.xf_index,
            value: framed.value,
            metadata,
            formula,
        },
        defect,
    ))
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}
