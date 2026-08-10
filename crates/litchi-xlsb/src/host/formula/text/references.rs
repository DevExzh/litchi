#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! A1-reference parsing and XLSB reference token encoding.

use super::super::{Error, Result};
use super::ast::A1Reference;

pub(super) fn parse_a1_reference(value: &str) -> Option<A1Reference> {
    let bytes = value.as_bytes();
    let mut offset = 0;
    let col_relative = bytes.get(offset) != Some(&b'$');
    if !col_relative {
        offset += 1;
    }
    let col_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_alphabetic) {
        offset += 1;
    }
    if offset == col_start {
        return None;
    }
    let mut col = 0u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
    }
    if col == 0 || col > 16_384 {
        return None;
    }

    let row_relative = bytes.get(offset) != Some(&b'$');
    if !row_relative {
        offset += 1;
    }
    let row_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_digit) {
        offset += 1;
    }
    if offset == row_start || offset != bytes.len() {
        return None;
    }
    let row = value[row_start..offset].parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 {
        return None;
    }
    Some(A1Reference {
        row: row - 1,
        col: col - 1,
        row_relative,
        col_relative,
    })
}

pub(super) fn reference_column_bits(reference: A1Reference) -> u16 {
    let mut bits = reference.col as u16;
    if reference.col_relative {
        bits |= 0x4000;
    }
    if reference.row_relative {
        bits |= 0x8000;
    }
    bits
}

pub(super) fn emit_reference(output: &mut Vec<u8>, token: u8, reference: A1Reference) {
    output.push(token);
    output.extend_from_slice(&reference.row.to_le_bytes());
    output.extend_from_slice(&reference_column_bits(reference).to_le_bytes());
}

pub(super) fn emit_shared_reference(
    output: &mut Vec<u8>,
    token: u8,
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<()> {
    let (row, col) = encode_shared_reference(reference, base_row, base_col)?;
    output.push(token);
    output.extend_from_slice(&row.to_le_bytes());
    output.extend_from_slice(&col.to_le_bytes());
    Ok(())
}

pub(super) fn encode_shared_reference(
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<(u32, u16)> {
    let row = if reference.row_relative {
        let offset = i64::from(reference.row) - i64::from(base_row);
        i32::try_from(offset)
            .map_err(|_| Error::InvalidFormula("shared row offset overflow".to_string()))?
            as u32
    } else {
        reference.row
    };
    let col_value = if reference.col_relative {
        let offset = i64::from(reference.col) - i64::from(base_col);
        if !(-16_383..=16_383).contains(&offset) {
            return Err(Error::InvalidFormula(format!(
                "shared column offset {offset} is outside the XLSB range"
            )));
        }
        (offset as i32 as u16) & 0x3FFF
    } else {
        reference.col as u16
    };
    let mut col = col_value;
    if reference.col_relative {
        col |= 0x4000;
    }
    if reference.row_relative {
        col |= 0x8000;
    }
    Ok((row, col))
}
