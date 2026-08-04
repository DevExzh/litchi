//! Bounded `BrtName` parsing and defined-name formula construction.

use super::model::{Definition, validate_name};
use super::{Error, Result};
use crate::formula::{CellParsedFormula, ptg_types};
use crate::raw::{Cursor, Limits};

/// Maximum bytes accepted for one `BrtName` payload.
pub const MAX_RECORD_BYTES: usize = 64 * 1024 * 1024;
/// Maximum UTF-16 units accepted by the raw nullable-string reader.
const MAX_WIDE_STRING_UNITS: usize = 32_767;
/// Maximum UTF-16 units in an `XLNameWideString`.
const MAX_NAME_UNITS: usize = 255;
/// Maximum UTF-16 units in the `BrtName` comment field.
const MAX_COMMENT_UNITS: usize = 255;
/// Maximum UTF-16 units in each macro description field.
const MAX_MACRO_DESCRIPTION_UNITS: usize = 32_767;

/// Parse one complete `BrtName` payload.
pub fn parse(data: &[u8]) -> Result<Definition> {
    if data.len() > MAX_RECORD_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_RECORD_BYTES,
            found: data.len(),
        });
    }
    if data.len() < 13 {
        return Err(Error::InvalidLength {
            expected: 13,
            found: data.len(),
        });
    }

    let limits = Limits::new(MAX_RECORD_BYTES, MAX_WIDE_STRING_UNITS);
    let mut cursor = Cursor::with_limits(data, "BrtName", limits);
    let flags = cursor.read_u32()?;
    if flags & 0xFFFC_0000 != 0 {
        return Err(Error::InvalidFormula(format!(
            "BrtName reserved flags are nonzero: 0x{flags:08X}"
        )));
    }
    let hidden = (flags & 0x0001) != 0;
    let f_func = (flags & 0x0002) != 0;
    let f_ob = (flags & 0x0004) != 0;
    let f_proc = (flags & 0x0008) != 0;
    let function = f_proc;
    if (f_func || f_ob) && !f_proc {
        return Err(Error::InvalidFormula(
            "BrtName macro type requires fProc".to_string(),
        ));
    }
    let function_group = (flags >> 6) & 0x01FF;
    if !f_proc && function_group != 0 {
        return Err(Error::InvalidFormula(
            "BrtName non-macro has a function group".to_string(),
        ));
    }
    let ch_key = cursor.read_u8()?;
    if (f_func || !f_proc) && ch_key != 0 {
        return Err(Error::InvalidFormula(format!(
            "BrtName has invalid macro shortcut key 0x{ch_key:02X}"
        )));
    }
    if f_proc && !f_func && ch_key < 0x20 {
        return Err(Error::InvalidFormula(format!(
            "BrtName has invalid macro shortcut key 0x{ch_key:02X}"
        )));
    }

    let sheet_id_raw = cursor.read_u32()? as i32;
    let sheet_id = if sheet_id_raw == -1 {
        None
    } else {
        Some(sheet_id_raw as u32)
    };

    let name = cursor.read_wide_string()?;
    if name.encode_utf16().count() > MAX_NAME_UNITS {
        return Err(Error::InvalidFormula(format!(
            "defined name length {} is outside 1..=255",
            name.encode_utf16().count()
        )));
    }
    validate_name(&name)?;

    let formula_start = cursor.position();
    let (parsed_formula, formula_bytes) =
        CellParsedFormula::parse(&data[formula_start..]).map_err(Error::from)?;
    cursor.skip(formula_bytes)?;

    let comment = cursor.read_nullable_wide_string()?;
    if comment
        .as_ref()
        .is_some_and(|value| value.encode_utf16().count() > MAX_COMMENT_UNITS)
    {
        return Err(Error::InvalidFormula(
            "BrtName comment exceeds 255 characters".to_string(),
        ));
    }

    if f_proc {
        for index in 0..4 {
            let value = cursor.read_nullable_wide_string()?;
            if matches!(index, 0 | 3) && value.is_some() {
                return Err(Error::InvalidFormula(
                    "BrtName macro unused string is not NULL".to_string(),
                ));
            }
            if matches!(index, 1 | 2)
                && value
                    .as_ref()
                    .is_some_and(|text| text.encode_utf16().count() > MAX_MACRO_DESCRIPTION_UNITS)
            {
                return Err(Error::InvalidFormula(
                    "BrtName macro description exceeds 32,767 characters".to_string(),
                ));
            }
        }
    }
    if cursor.remaining() != 0 {
        return Err(Error::InvalidFormula(format!(
            "BrtName has {} trailing bytes",
            cursor.remaining()
        )));
    }

    Ok(Definition {
        name,
        formula: Some(parsed_formula.rgce),
        sheet_id,
        hidden,
        function,
    })
}

impl Definition {
    /// Parse one complete `BrtName` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        parse(data)
    }

    /// Create a 3D area formula token stream for a workbook-local range.
    ///
    /// `sheet_id` is the zero-based workbook sheet index. The workbook's
    /// self extern-sheet table reserves the first two entries for the
    /// workbook and `#REF!`, so sheet references start at index two.
    pub fn area3d_formula(
        sheet_id: u32,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<Vec<u8>> {
        area3d_formula(sheet_id, first_row, last_row, first_col, last_col)
    }
}

/// Create a 3D area formula token stream for a workbook-local sheet range.
pub fn area3d_formula(
    sheet_id: u32,
    first_row: u32,
    last_row: u32,
    first_col: u16,
    last_col: u16,
) -> Result<Vec<u8>> {
    if first_row > last_row || last_row >= 1_048_576 || first_col > last_col || last_col >= 16_384 {
        return Err(Error::Formula(crate::formula::Error::InvalidCellReference(
            format!("named range ({first_row}, {first_col})..=({last_row}, {last_col})"),
        )));
    }
    let sheet_index = u16::try_from(sheet_id)
        .ok()
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| {
            Error::InvalidFormula(format!(
                "sheet index {sheet_id} cannot be represented in the extern-sheet table"
            ))
        })?;
    let mut formula = Vec::with_capacity(15);
    formula.push(ptg_types::PTG_AREA_3D);
    formula.extend_from_slice(&sheet_index.to_le_bytes());
    formula.extend_from_slice(&first_row.to_le_bytes());
    formula.extend_from_slice(&last_row.to_le_bytes());
    formula.extend_from_slice(&first_col.to_le_bytes());
    formula.extend_from_slice(&last_col.to_le_bytes());
    Ok(formula)
}
