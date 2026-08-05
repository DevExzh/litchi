//! Shared formula grammar and bounded-value validation helpers.
use super::super::function_table::BUILTIN_FUNCTIONS;
use super::super::model::TableRowType;
use super::super::{Error, Result};

pub(super) fn parse_table_row_type(value: u8) -> Result<TableRowType> {
    match value {
        0x00 => Ok(TableRowType::Data),
        0x01 => Ok(TableRowType::All),
        0x02 => Ok(TableRowType::Headers),
        0x04 => Ok(TableRowType::DataAlternate),
        0x06 => Ok(TableRowType::DataAndHeaders),
        0x08 => Ok(TableRowType::Totals),
        0x0C => Ok(TableRowType::DataAndTotals),
        0x10 => Ok(TableRowType::Current),
        _ => Err(Error::InvalidFormula(format!(
            "invalid PtgRowType 0x{value:02X}"
        ))),
    }
}

pub(super) fn table_row_type_raw(value: TableRowType) -> u8 {
    match value {
        TableRowType::Data => 0x00,
        TableRowType::All => 0x01,
        TableRowType::Headers => 0x02,
        TableRowType::DataAlternate => 0x04,
        TableRowType::DataAndHeaders => 0x06,
        TableRowType::Totals => 0x08,
        TableRowType::DataAndTotals => 0x0C,
        TableRowType::Current => 0x10,
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct BuiltinFunction {
    pub(super) index: u16,
    pub(super) name: &'static str,
    pub(super) min_args: u8,
    pub(super) max_args: u8,
}

impl BuiltinFunction {
    pub(super) fn accepts_arg_count(self, count: u8) -> bool {
        if count < self.min_args || count > self.max_args {
            return false;
        }
        match self.index {
            // GETPIVOTDATA permits the two mandatory arguments, one optional
            // field, or complete field/item pairs thereafter.
            358 => count <= 3 || count.is_multiple_of(2),
            // COUNTIFS is made solely of range/criteria pairs.
            481 => count.is_multiple_of(2),
            // SUMIFS and AVERAGEIFS have one leading aggregate range followed
            // by range/criteria pairs.
            482 | 484 => !count.is_multiple_of(2),
            _ => true,
        }
    }
}

pub(super) fn builtin_function_by_index(index: u16) -> Option<BuiltinFunction> {
    let position = BUILTIN_FUNCTIONS
        .binary_search_by_key(&index, |entry| entry.0)
        .ok()?;
    let (index, name, min_args, max_args) = BUILTIN_FUNCTIONS[position];
    Some(BuiltinFunction {
        index,
        name,
        min_args,
        max_args,
    })
}

pub(super) fn validate_xnum(value: f64, context: &str) -> Result<()> {
    if !value.is_finite()
        || (value == 0.0 && value.is_sign_negative())
        || (value != 0.0 && !value.is_normal())
    {
        return Err(Error::InvalidFormula(format!(
            "{context} contains a non-finite, denormalized, or negative-zero Xnum"
        )));
    }
    Ok(())
}

const FORMULA_ERRORS: &[(&str, u8)] = &[
    ("#GETTING_DATA", 0x2B),
    ("#DIV/0!", 0x07),
    ("#VALUE!", 0x0F),
    ("#NULL!", 0x00),
    ("#NAME?", 0x1D),
    ("#REF!", 0x17),
    ("#NUM!", 0x24),
    ("#N/A", 0x2A),
];

pub(super) fn is_formula_error_code(value: u8) -> bool {
    FORMULA_ERRORS.iter().any(|(_, code)| *code == value)
}

pub(super) fn add_wrapped_offset(base: u32, offset: i32, modulus: u32) -> u32 {
    (i64::from(base) + i64::from(offset)).rem_euclid(i64::from(modulus)) as u32
}
