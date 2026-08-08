//! BIFF8 Formula-record field codec.

use crate::records::FormulaValue;
use crate::utils;
use crate::{Error, Result};

use super::Metadata;
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

/// Parse the payload of a BIFF8 `Formula` record.
pub(crate) fn parse_record(data: &[u8]) -> Result<Parsed> {
    parse_record_with(data, false).map(|(parsed, _)| parsed)
}

/// Parse a Formula while preserving the one explicitly modeled producer
/// defect. All framing, size, reserved-bit, and value checks remain strict.
pub(crate) fn parse_record_preserving(data: &[u8]) -> Result<(Parsed, Option<FlagDefect>)> {
    parse_record_with(data, true)
}

fn parse_record_with(data: &[u8], preserve_defect: bool) -> Result<(Parsed, Option<FlagDefect>)> {
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
    if formula_end != data.len() {
        return Err(Error::InvalidLength {
            expected: formula_end,
            found: data.len(),
        });
    }
    let formula = data[FORMULA_FIXED_SIZE..formula_end].to_vec();
    // A string-valued Formula deliberately carries no Rgce bytes: its
    // cached result is supplied by the immediately following String record.
    // All other FormulaValue variants require an actual token stream.
    if formula.is_empty() && !matches!(value, FormulaValue::StringPending) {
        return Err(invalid("Formula token stream cannot be empty"));
    }
    let (mut metadata, defect) = if preserve_defect {
        decode_flags_preserving(flags, &formula)?
    } else {
        (decode_flags(flags, &formula)?, None)
    };
    metadata = metadata.with_calculation_cache(calculation_cache);

    Ok((
        Parsed {
            row,
            col,
            xf_index,
            value,
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
