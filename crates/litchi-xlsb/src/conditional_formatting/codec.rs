//! Conditional formatting support for XLSB.
//
//! The semantic model and bounded Brt* codecs live at the concrete XLSB
//! boundary. Package/worksheet orchestration remains in the OOXML host.
#![allow(clippy::too_many_arguments)]

use crate::formula::{
    ArrayValue, Compiler, MAX_CELL_FORMULA_BYTES, ParsedFormula, Parser, Resolution,
};
use crate::raw::{Writer, kind};
use std::collections::{HashMap, HashSet};
use std::io::Write;

use thiserror::Error;

/// Result type for conditional-formatting codecs.
pub type Result<T> = std::result::Result<T, Error>;

/// Strict conditional-formatting codec error.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    #[error("invalid formula: {0}")]
    InvalidFormula(String),
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),
    #[error("encoding error: {0}")]
    Encoding(String),
    #[error("unsupported feature: {0}")]
    UnsupportedFeature(String),
    #[error("unrecognized {typ}: {val}")]
    Unrecognized { typ: String, val: String },
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    #[error(transparent)]
    Formula(#[from] crate::formula::Error),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

#[derive(Debug, Clone, Copy, Default)]
struct EmptyFormulaResolution;

impl Resolution for EmptyFormulaResolution {
    fn sheet_prefix(&self, index: u16) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved sheet index {index}"
        )))
    }

    fn defined_name(&self, index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved defined name {index}"
        )))
    }

    fn external_name(&self, sheet_index: u16, name_index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved external name {sheet_index}:{name_index}"
        )))
    }

    fn table_reference(
        &self,
        _reference: &crate::formula::TableReference,
    ) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(
            "formula references unresolved table".to_string(),
        ))
    }

    fn pivot_name(&self, index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved pivot name {index}"
        )))
    }
}

/// Small, owner-local formula emitter used when a conditional-formatting
/// model was authored from formula text instead of retained binary tokens.
///
/// The full workbook formula compiler remains a host concern. Conditional
/// formatting only needs the bounded literal/reference/operator subset here;
/// unsupported constructs are rejected rather than silently rewritten.
struct TextCompiler;

impl TextCompiler {
    fn compile(input: &str) -> Result<ParsedFormula> {
        let input = input.strip_prefix('=').unwrap_or(input).trim();
        if input.is_empty() {
            return Err(Error::InvalidFormula(
                "conditional-format formula is empty".to_string(),
            ));
        }
        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        compile_formula_expression(input, &mut rgce, &mut rgcb)?;
        if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "compiled conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(ParsedFormula { rgce, rgcb })
    }
}

fn compile_formula_expression(input: &str, rgce: &mut Vec<u8>, rgcb: &mut Vec<u8>) -> Result<()> {
    let input = input.trim();
    if input.is_empty() {
        return Err(Error::InvalidFormula(
            "conditional-format formula has an empty operand".to_string(),
        ));
    }
    if let Some(inner) = strip_outer_parentheses(input) {
        compile_formula_expression(inner, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_PAREN);
        return Ok(());
    }

    for operators in [
        ["<>", "<=", ">=", "=", "<", ">"].as_slice(),
        ["&"].as_slice(),
        ["+", "-"].as_slice(),
        ["*", "/"].as_slice(),
        ["^"].as_slice(),
    ] {
        if let Some((offset, operator)) = find_formula_operator(input, operators) {
            let (left, right) = input.split_at(offset);
            let right = &right[operator.len()..];
            compile_formula_expression(left, rgce, rgcb)?;
            compile_formula_expression(right, rgce, rgcb)?;
            rgce.push(match operator {
                "<>" => crate::formula::ptg_types::PTG_NE,
                "<=" => crate::formula::ptg_types::PTG_LE,
                ">=" => crate::formula::ptg_types::PTG_GE,
                "=" => crate::formula::ptg_types::PTG_EQ,
                "<" => crate::formula::ptg_types::PTG_LT,
                ">" => crate::formula::ptg_types::PTG_GT,
                "&" => crate::formula::ptg_types::PTG_CONCAT,
                "+" => crate::formula::ptg_types::PTG_ADD,
                "-" => crate::formula::ptg_types::PTG_SUB,
                "*" => crate::formula::ptg_types::PTG_MUL,
                "/" => crate::formula::ptg_types::PTG_DIV,
                "^" => crate::formula::ptg_types::PTG_POWER,
                _ => unreachable!("operator was selected from the fixed table"),
            });
            return Ok(());
        }
    }

    if let Some(rest) = input.strip_prefix('+') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_UPLUS);
        return Ok(());
    }
    if let Some(rest) = input.strip_prefix('-') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_UMINUS);
        return Ok(());
    }
    if let Some(rest) = input.strip_suffix('%') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_PERCENT);
        return Ok(());
    }

    if input.starts_with('{') && input.ends_with('}') {
        return compile_formula_array(&input[1..input.len() - 1], rgce, rgcb);
    }
    if let Some(value) = parse_formula_string(input)? {
        let units = value.encode_utf16().count();
        let units = u16::try_from(units)
            .map_err(|_| Error::InvalidFormula("formula string is too long".to_string()))?;
        rgce.push(crate::formula::ptg_types::PTG_STR);
        rgce.extend_from_slice(&units.to_le_bytes());
        rgce.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
        return Ok(());
    }
    if input.eq_ignore_ascii_case("TRUE") || input.eq_ignore_ascii_case("FALSE") {
        rgce.extend([
            crate::formula::ptg_types::PTG_BOOL,
            u8::from(input.eq_ignore_ascii_case("TRUE")),
        ]);
        return Ok(());
    }
    if let Some(error) = formula_error_code(input) {
        rgce.extend([crate::formula::ptg_types::PTG_ERR, error]);
        return Ok(());
    }
    if let Ok(number) = input.parse::<f64>() {
        if !number.is_finite() {
            return Err(Error::InvalidFormula(
                "formula number is not finite".to_string(),
            ));
        }
        if number.fract() == 0.0 && (0.0..=f64::from(u16::MAX)).contains(&number) {
            rgce.push(crate::formula::ptg_types::PTG_INT);
            rgce.extend_from_slice(&(number as u16).to_le_bytes());
        } else {
            rgce.push(crate::formula::ptg_types::PTG_NUM);
            rgce.extend_from_slice(&number.to_le_bytes());
        }
        return Ok(());
    }

    if let Some((row, _col, bits)) = parse_formula_reference(input) {
        rgce.push(crate::formula::ptg_types::PTG_REF | 0x20);
        rgce.extend_from_slice(&row.to_le_bytes());
        rgce.extend_from_slice(&bits.to_le_bytes());
        return Ok(());
    }
    if let Some((first, last)) = input.split_once(':') {
        let Some((first_row, first_col, first_bits)) = parse_formula_reference(first.trim()) else {
            return Err(Error::InvalidFormula(format!(
                "invalid conditional-format range {input:?}"
            )));
        };
        let Some((last_row, last_col, last_bits)) = parse_formula_reference(last.trim()) else {
            return Err(Error::InvalidFormula(format!(
                "invalid conditional-format range {input:?}"
            )));
        };
        if first_row > last_row || first_col > last_col {
            return Err(Error::InvalidFormula(
                "conditional-format range is reversed".to_string(),
            ));
        }
        rgce.push(crate::formula::ptg_types::PTG_AREA | 0x20);
        rgce.extend_from_slice(&first_row.to_le_bytes());
        rgce.extend_from_slice(&last_row.to_le_bytes());
        rgce.extend_from_slice(&first_bits.to_le_bytes());
        rgce.extend_from_slice(&last_bits.to_le_bytes());
        return Ok(());
    }

    Err(Error::UnsupportedFeature(format!(
        "conditional-format formula construct {input:?} is not supported by the owner-local emitter"
    )))
}

fn compile_formula_array(input: &str, rgce: &mut Vec<u8>, rgcb: &mut Vec<u8>) -> Result<()> {
    let rows = split_formula_list(input, ';');
    if rows.is_empty() || rows.iter().any(|row| row.trim().is_empty()) {
        return Err(Error::InvalidFormula(
            "conditional-format array has an empty row".to_string(),
        ));
    }
    let columns = rows
        .iter()
        .map(|row| split_formula_list(row, ','))
        .collect::<Vec<_>>();
    let column_count = columns[0].len();
    if column_count == 0 || columns.iter().any(|row| row.len() != column_count) {
        return Err(Error::InvalidFormula(
            "conditional-format array rows have different widths".to_string(),
        ));
    }
    let row_count = u32::try_from(columns.len())
        .map_err(|_| Error::InvalidFormula("array row count overflow".to_string()))?;
    let column_count = u32::try_from(column_count)
        .map_err(|_| Error::InvalidFormula("array column count overflow".to_string()))?;
    if row_count > 1_048_576 || column_count > 16_384 {
        return Err(Error::InvalidFormula(
            "conditional-format array exceeds worksheet bounds".to_string(),
        ));
    }

    let mut values = Vec::new();
    for row in columns {
        for value in row {
            values.push(parse_formula_array_value(value.trim())?);
        }
    }
    // PtgArray uses the VALUE operand class (0x40); 0x20 is the base token
    // value and is not a valid array token on the BIFF12 wire.
    rgce.push(0x40);
    rgce.extend([0; 14]);
    rgcb.extend_from_slice(&row_count.to_le_bytes());
    rgcb.extend_from_slice(&column_count.to_le_bytes());
    for value in values {
        match value {
            ArrayValue::Number(value) => {
                if !value.is_finite() {
                    return Err(Error::InvalidFormula(
                        "conditional-format array number is not finite".to_string(),
                    ));
                }
                rgcb.push(0);
                rgcb.extend_from_slice(&value.to_le_bytes());
            },
            ArrayValue::String(value) => {
                let units = u16::try_from(value.encode_utf16().count()).map_err(|_| {
                    Error::InvalidFormula("conditional-format array string is too long".to_string())
                })?;
                rgcb.push(1);
                rgcb.extend_from_slice(&units.to_le_bytes());
                rgcb.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
            },
            ArrayValue::Bool(value) => rgcb.extend([2, u8::from(value)]),
            ArrayValue::Error(value) => rgcb.extend([4, value, 0, 0, 0]),
        }
    }
    Ok(())
}

fn parse_formula_array_value(input: &str) -> Result<ArrayValue> {
    if let Some(value) = parse_formula_string(input)? {
        return Ok(ArrayValue::String(value));
    }
    if input.eq_ignore_ascii_case("TRUE") || input.eq_ignore_ascii_case("FALSE") {
        return Ok(ArrayValue::Bool(input.eq_ignore_ascii_case("TRUE")));
    }
    if let Some(error) = formula_error_code(input) {
        return Ok(ArrayValue::Error(error));
    }
    let value = input.parse::<f64>().map_err(|_| {
        Error::InvalidFormula(format!("invalid conditional-format array value {input:?}"))
    })?;
    Ok(ArrayValue::Number(value))
}

fn parse_formula_string(input: &str) -> Result<Option<String>> {
    if !input.starts_with('"') {
        return Ok(None);
    }
    if !input.ends_with('"') || input.len() < 2 {
        return Err(Error::InvalidFormula(
            "unterminated conditional-format string".to_string(),
        ));
    }
    let mut value = String::new();
    let mut chars = input[1..input.len() - 1].chars().peekable();
    while let Some(character) = chars.next() {
        if character == '"' && chars.peek() == Some(&'"') {
            chars.next();
        }
        value.push(character);
    }
    Ok(Some(value))
}

fn formula_error_code(input: &str) -> Option<u8> {
    [
        ("#NULL!", 0x00),
        ("#DIV/0!", 0x07),
        ("#VALUE!", 0x0F),
        ("#REF!", 0x17),
        ("#NAME?", 0x1D),
        ("#NUM!", 0x24),
        ("#N/A", 0x2A),
        ("#GETTING_DATA", 0x2B),
    ]
    .into_iter()
    .find_map(|(literal, code)| input.eq_ignore_ascii_case(literal).then_some(code))
}

fn parse_formula_reference(input: &str) -> Option<(u32, u32, u16)> {
    let input = input.trim();
    let bytes = input.as_bytes();
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
    let mut col = 0_u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
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
    let row = input[row_start..offset].parse::<u32>().ok()?;
    if col == 0 || col > 16_384 || row == 0 || row > 1_048_576 {
        return None;
    }
    let mut bits = u16::try_from(col - 1).ok()?;
    if col_relative {
        bits |= 0x4000;
    }
    if row_relative {
        bits |= 0x8000;
    }
    Some((row - 1, col - 1, bits))
}

fn split_formula_list(input: &str, separator: char) -> Vec<&str> {
    let mut result = Vec::new();
    let mut start = 0;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
        } else if character == separator && !in_string {
            result.push(&input[start..offset]);
            start = offset + character.len_utf8();
        }
    }
    result.push(&input[start..]);
    result
}

fn find_formula_operator(input: &str, operators: &[&'static str]) -> Option<(usize, &'static str)> {
    let mut parentheses = 0_usize;
    let mut braces = 0_usize;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
            continue;
        }
        if in_string {
            continue;
        }
        match character {
            '(' => parentheses += 1,
            ')' => parentheses = parentheses.saturating_sub(1),
            '{' => braces += 1,
            '}' => braces = braces.saturating_sub(1),
            _ => {},
        }
        if parentheses != 0 || braces != 0 {
            continue;
        }
        for operator in operators {
            if !input[offset..].starts_with(operator) {
                continue;
            }
            if matches!(*operator, "+" | "-") {
                let previous = input[..offset]
                    .chars()
                    .rev()
                    .find(|character| !character.is_whitespace());
                if previous.is_none_or(|character| "+-*/^&=<>(".contains(character))
                    || (*operator == "+" || *operator == "-")
                        && previous.is_some_and(|character| matches!(character, 'e' | 'E'))
                {
                    continue;
                }
            }
            return Some((offset, operator));
        }
    }
    None
}

fn strip_outer_parentheses(input: &str) -> Option<&str> {
    if !input.starts_with('(') || !input.ends_with(')') {
        return None;
    }
    let mut depth = 0_usize;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
        } else if !in_string {
            match character {
                '(' => depth += 1,
                ')' => {
                    depth = depth.checked_sub(1)?;
                    if depth == 0 && offset + character.len_utf8() != input.len() {
                        return None;
                    }
                },
                _ => {},
            }
        }
    }
    (depth == 0).then_some(&input[1..input.len() - 1])
}

// -----------------------------------------------------------------------------
// Owner-local range and Future Record Type helpers.
//
// These helpers intentionally live at the XLSB boundary. They model the
// BinRangeList and FRTHeader structures used by [MS-XLSB] §§2.2.6.2.1,
// 2.4.23--2.4.24, 2.4.43--2.4.44, 2.4.91--2.4.92, 2.4.332--2.4.335,
// 2.4.380--2.4.381, 2.4.399--2.4.400, 2.4.445--2.4.446, 2.5.19--2.5.20,
// and 2.5.98.7.
// -----------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellRange {
    row_first: u32,
    row_last: u32,
    col_first: u32,
    col_last: u32,
}

impl CellRange {
    fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    fn parse(value: &str) -> Result<Self> {
        let value = value.trim();
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        let (row_first, col_first) = parse_cell_reference(first)?;
        let (row_last, col_last) = parse_cell_reference(last)?;
        Ok(Self::new(row_first, row_last, col_first, col_last))
    }

    fn write<W: Write>(&self, writer: &mut W) -> Result<()> {
        writer.write_all(&self.row_first.to_le_bytes())?;
        writer.write_all(&self.row_last.to_le_bytes())?;
        writer.write_all(&self.col_first.to_le_bytes())?;
        writer.write_all(&self.col_last.to_le_bytes())?;
        Ok(())
    }
}

fn parse_range_list(value: &str) -> Result<Vec<CellRange>> {
    value
        .split([',', ' '])
        .filter(|part| !part.is_empty())
        .map(CellRange::parse)
        .collect()
}

fn write_bin_range_list<W: Write>(ranges: &[CellRange], writer: &mut W) -> Result<()> {
    let count = i32::try_from(ranges.len())
        .map_err(|_| invalid("BinRangeList", "range count overflows i32"))?;
    writer.write_all(&count.to_le_bytes())?;
    for range in ranges {
        range.write(writer)?;
    }
    Ok(())
}

fn column_index_to_name(mut column: u32) -> String {
    if column == 0 {
        return String::new();
    }
    let mut result = String::new();
    while column > 0 {
        column -= 1;
        result.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    result
}

fn cell_reference(row: u32, column: u32) -> String {
    let Some(column) = column.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    let Some(row) = row.checked_add(1) else {
        return format!("R{row}C{column}");
    };
    format!("{}{}", column_index_to_name(column), row)
}

fn parse_cell_reference(value: &str) -> Result<(u32, u32)> {
    let normalized = value.trim().to_ascii_uppercase();
    let mut column = String::new();
    let mut row = String::new();
    let mut digit_seen = false;
    for character in normalized.chars() {
        if character.is_ascii_alphabetic() {
            if digit_seen {
                return Err(Error::InvalidCellReference(normalized));
            }
            column.push(character);
        } else if character.is_ascii_digit() {
            digit_seen = true;
            row.push(character);
        } else {
            return Err(invalid("cell reference", normalized));
        }
    }
    if column.is_empty() || row.is_empty() {
        return Err(Error::InvalidCellReference(normalized));
    }
    let mut column_index = 0_u32;
    for character in column.bytes() {
        if !character.is_ascii_uppercase() {
            return Err(Error::InvalidCellReference(normalized));
        }
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(character - b'A' + 1)))
            .ok_or_else(|| Error::InvalidCellReference(normalized.clone()))?;
    }
    let row_index = row
        .parse::<u32>()
        .map_err(|_| Error::InvalidCellReference(normalized.clone()))?;
    if row_index == 0 || column_index == 0 {
        return Err(Error::InvalidCellReference(normalized));
    }
    Ok((row_index - 1, column_index - 1))
}

type FrtRange = (u32, u32, u32, u32);

fn parse_sqref_header(
    data: &[u8],
    record: &'static str,
    maximum_ranges: usize,
) -> Result<(Vec<FrtRange>, usize)> {
    let mut cursor = FrtCursor::new(data, record);
    if cursor.read_u32()? != 0x02 {
        return Err(invalid(record, "FRTHeader is not sqref-only"));
    }
    if cursor.read_u32()? != 1 {
        return Err(invalid(record, "FRTSqrefs count is not 1"));
    }
    let flags = cursor.read_u32()?;
    if flags & 0x02 == 0 || flags & !0x0001_000f != 0 {
        return Err(invalid(
            record,
            format!("invalid FRTSqref flags 0x{flags:08X}"),
        ));
    }
    let count = usize::try_from(cursor.read_u32()? as i32)
        .map_err(|_| invalid(record, "NULL range collection"))?;
    if count == 0 || count > maximum_ranges || count > cursor.remaining() / 16 {
        return Err(invalid(record, format!("invalid range count {count}")));
    }
    let mut ranges = Vec::with_capacity(count);
    for _ in 0..count {
        let row_first = cursor.read_u32()?;
        let row_last = cursor.read_u32()?;
        let col_first = cursor.read_u32()?;
        let col_last = cursor.read_u32()?;
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(invalid(record, "invalid FRT target range"));
        }
        ranges.push((row_first, row_last, col_first, col_last));
    }
    Ok((ranges, cursor.offset))
}

fn serialize_sqref_header(ranges: &[FrtRange]) -> Result<Vec<u8>> {
    if ranges.is_empty() || ranges.len() > i32::MAX as usize {
        return Err(invalid(
            "FRTHeader",
            format!("invalid range count {}", ranges.len()),
        ));
    }
    let mut data = Vec::with_capacity(16 + ranges.len() * 16);
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend_from_slice(&1u32.to_le_bytes());
    data.extend_from_slice(&2u32.to_le_bytes());
    data.extend_from_slice(&(ranges.len() as u32).to_le_bytes());
    for &(row_first, row_last, col_first, col_last) in ranges {
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(invalid("FRTHeader", "invalid FRT target range"));
        }
        for value in [row_first, row_last, col_first, col_last] {
            data.extend_from_slice(&value.to_le_bytes());
        }
    }
    Ok(data)
}

fn parse_formula_header(
    data: &[u8],
    record: &'static str,
    maximum_formulas: usize,
) -> Result<(Vec<ParsedFormula>, usize)> {
    let mut cursor = FrtCursor::new(data, record);
    let flags = cursor.read_u32()?;
    if flags & !0x04 != 0 {
        return Err(invalid(
            record,
            format!("invalid FRTHeader flags 0x{flags:08X}"),
        ));
    }
    let mut formulas = Vec::new();
    if flags & 0x04 != 0 {
        let count = usize::try_from(cursor.read_u32()?)
            .map_err(|_| invalid(record, "FRT formula count overflow"))?;
        if count == 0 || count > maximum_formulas {
            return Err(invalid(
                record,
                format!("FRT formula count {count} is outside 1..={maximum_formulas}"),
            ));
        }
        formulas
            .try_reserve(count)
            .map_err(|_| Error::Unrecognized {
                typ: record.to_string(),
                val: "formula allocation exceeds bounded capacity".to_string(),
            })?;
        for _ in 0..count {
            formulas.push(cursor.read_formula()?);
        }
    }
    Ok((formulas, cursor.offset))
}

fn serialize_formula_header(
    formulas: &[ParsedFormula],
    maximum_formulas: usize,
) -> Result<Vec<u8>> {
    if formulas.len() > maximum_formulas {
        return Err(invalid(
            "FRTHeader",
            format!(
                "formula count {} exceeds {maximum_formulas}",
                formulas.len()
            ),
        ));
    }
    let mut data = Vec::new();
    data.extend_from_slice(&if formulas.is_empty() { 0u32 } else { 4u32 }.to_le_bytes());
    if formulas.is_empty() {
        return Ok(data);
    }
    data.extend_from_slice(
        &u32::try_from(formulas.len())
            .map_err(|_| invalid("FRTHeader", "formula count overflow"))?
            .to_le_bytes(),
    );
    for formula in formulas {
        if formula.rgce.is_empty() || formula.rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "FRT formula token length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                formula.rgce.len()
            )));
        }
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(formula.rgce.len())
                .map_err(|_| invalid("FRTFormula", "token length overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(
            &u32::try_from(formula.rgcb.len())
                .map_err(|_| invalid("FRTFormula", "ancillary length overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(&formula.rgce);
        data.extend_from_slice(&formula.rgcb);
    }
    Ok(data)
}

struct FrtCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> FrtCursor<'a> {
    fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    fn read_formula(&mut self) -> Result<ParsedFormula> {
        let flags = self.read_u32()?;
        if flags != 2 {
            return Err(invalid(
                self.record,
                format!("invalid FRTFormula flags 0x{flags:08X}"),
            ));
        }
        let cce = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula token length overflow"))?;
        let cb = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid(self.record, "formula ancillary length overflow"))?;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "FRT formula token length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        Ok(ParsedFormula {
            rgce: self.take(cce)?.to_vec(),
            rgcb: self.take(cb)?.to_vec(),
        })
    }
}

// Semantic values are defined in model.rs; this file owns only the
// bounded Brt* framing and formula boundary.
use super::model::*;

pub fn icon_count14(icon_set_type: u8) -> usize {
    match icon_set_type {
        0..=7 | 17 | 18 => 3,
        8..=12 => 4,
        13..=16 | 19 => 5,
        _ => 0,
    }
}

pub fn parse_rule_extension_guid(data: &[u8]) -> Result<[u8; 16]> {
    if data.len() != 20 {
        return Err(Error::InvalidLength {
            expected: 20,
            found: data.len(),
        });
    }
    if data[..4] != [0; 4] {
        return Err(invalid("BrtCFRuleExt", "nonzero FRTBlank"));
    }
    Ok(data[4..].try_into().expect("sixteen-byte GUID"))
}

pub fn serialize_rule_extension_guid(guid: [u8; 16]) -> [u8; 20] {
    let mut data = [0; 20];
    data[4..].copy_from_slice(&guid);
    data
}

impl Value {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 24 {
            return Err(Error::InvalidLength {
                expected: 24,
                found: data.len(),
            });
        }
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFVO");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
            return Err(invalid("BrtCFVO", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO", "non-finite numeric parameter"));
        }
        if matches!(cfvo_type, 4 | 5) && !(0.0..=100.0).contains(&numeric_value) {
            return Err(invalid(
                "BrtCFVO",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        let formula_binary = if declared_formula_size == 0 {
            None
        } else {
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != declared_formula_size {
                return Err(invalid(
                    "BrtCFVO",
                    "declared formula size does not match token stream",
                ));
            }
            Some(formula)
        };
        cursor.finish()?;
        if matches!(cfvo_type, 2 | 3) && formula_binary.is_some() {
            return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
        }
        let value = if let Some(formula) = &formula_binary {
            Some(render_formula(formula, base, context)?)
        } else if matches!(cfvo_type, 1 | 4 | 5) {
            Some(format_number(numeric_value))
        } else {
            None
        };
        Ok(Self {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal,
            greater_than_or_equal,
            formula_binary,
        })
    }

    /// Parse an Office 2013 `BrtCFVO14` record.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtCFVO14", 1)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtCFVO14");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO14", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid("BrtCFVO14", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        cursor.finish()?;
        let formula_binary = formulas.into_iter().next();
        if formula_binary
            .as_ref()
            .map_or(0, |formula| formula.rgce.len())
            != declared_formula_size
        {
            return Err(invalid(
                "BrtCFVO14",
                "FRT formula and declared token size disagree",
            ));
        }
        if matches!(cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        if formula_binary.is_none()
            && matches!(cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
        }
        let value = if let Some(formula) = &formula_binary {
            Some(render_formula(formula, base, context)?)
        } else if matches!(cfvo_type, 1 | 4 | 5) {
            Some(format_number(numeric_value))
        } else {
            None
        };
        Ok(Self {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal,
            greater_than_or_equal,
            formula_binary,
        })
    }

    /// Serialize an Office 2013 `BrtCFVO14` payload using its binary formula.
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
        self.serialize_extension14_with(
            self.formula_binary.as_ref(),
            self.numeric_value,
            self.save_greater_than_or_equal,
        )
    }

    fn serialize_extension14_with(
        &self,
        formula_binary: Option<&ParsedFormula>,
        numeric_value: f64,
        save_greater_than_or_equal: bool,
    ) -> Result<Vec<u8>> {
        if !matches!(self.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid(
                "BrtCFVO14",
                format!("invalid type {}", self.cfvo_type),
            ));
        }
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        if formula_binary.is_none()
            && matches!(self.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {} outside 0..=100", numeric_value),
            ));
        }
        if matches!(self.cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if self.cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        let formulas = formula_binary.map_or(&[][..], std::slice::from_ref);
        let mut data = serialize_formula_header(formulas, 1)?;
        data.extend_from_slice(&u32::from(self.cfvo_type).to_le_bytes());
        data.extend_from_slice(&numeric_value.to_le_bytes());
        data.extend_from_slice(&u32::from(save_greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(&u32::from(self.greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(formula_binary.map_or(0, |formula| formula.rgce.len()))
                .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
                .to_le_bytes(),
        );
        Ok(data)
    }
}

impl Color {
    pub fn theme(index: u8, tint: i16) -> Result<Self> {
        if index > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {index}")));
        }
        let tint_bytes = tint.to_le_bytes();
        Ok(Self {
            color_type: 3,
            index,
            tint,
            argb: None,
            raw: [6, index, tint_bytes[0], tint_bytes[1], 0, 0, 0, 0],
        })
    }

    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let raw: [u8; 8] = data.try_into().map_err(|_| Error::InvalidLength {
            expected: 8,
            found: data.len(),
        })?;
        let color_type = raw[0] >> 1;
        if color_type > 3 {
            return Err(invalid("BrtColor", format!("color type {color_type}")));
        }
        let argb = if color_type == 2 {
            if raw[0] & 1 == 0 {
                return Err(invalid("BrtColor", "direct color is not marked valid"));
            }
            Some(
                (u32::from(raw[7]) << 24)
                    | (u32::from(raw[4]) << 16)
                    | (u32::from(raw[5]) << 8)
                    | u32::from(raw[6]),
            )
        } else {
            None
        };
        if color_type == 3 && raw[1] > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {}", raw[1])));
        }
        Ok(Self {
            color_type,
            index: raw[1],
            tint: i16::from_le_bytes([raw[2], raw[3]]),
            argb,
            raw,
        })
    }

    pub fn to_bytes(self) -> Result<[u8; 8]> {
        if self.color_type > 3 || (self.color_type == 3 && self.index > 0x0b) {
            return Err(invalid("BrtColor", "invalid color type or theme index"));
        }
        if self.color_type == 2 && self.argb.is_none() {
            return Err(invalid("BrtColor", "direct color has no ARGB value"));
        }
        if self.color_type != 2 && self.argb.is_some() {
            return Err(invalid("BrtColor", "non-direct color has an ARGB value"));
        }
        let parsed_raw = Self::parse(&self.raw).ok();
        if parsed_raw.as_ref().is_some_and(|raw| {
            raw.color_type == self.color_type
                && raw.index == self.index
                && raw.tint == self.tint
                && raw.argb == self.argb
        }) {
            return Ok(self.raw);
        }
        let tint = self.tint.to_le_bytes();
        let mut raw = [
            self.color_type << 1,
            self.index,
            tint[0],
            tint[1],
            0,
            0,
            0,
            0,
        ];
        if let Some(argb) = self.argb {
            raw[0] |= 1;
            raw[4] = ((argb >> 16) & 0xff) as u8;
            raw[5] = ((argb >> 8) & 0xff) as u8;
            raw[6] = (argb & 0xff) as u8;
            raw[7] = ((argb >> 24) & 0xff) as u8;
        }
        Ok(raw)
    }

    /// Parse an Office 2013 `BrtColor14` payload.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        if data.len() != 12 {
            return Err(Error::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }
        if data[..4] != [0; 4] {
            return Err(invalid("BrtColor14", "nonzero FRTBlank"));
        }
        Self::parse(&data[4..])
    }

    /// Serialize an Office 2013 `BrtColor14` payload.
    pub fn serialize_extension14(self) -> Result<[u8; 12]> {
        let mut data = [0; 12];
        data[4..].copy_from_slice(&self.to_bytes()?);
        Ok(data)
    }
}

impl Direction14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Context),
            1 => Some(Self::LeftToRight),
            2 => Some(Self::RightToLeft),
            _ => None,
        }
    }
}

impl AxisPosition14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Automatic),
            1 => Some(Self::Midpoint),
            2 => Some(Self::None),
            _ => None,
        }
    }
}

impl Bar14 {
    pub fn parse_header(data: &[u8]) -> Result<BarHeader14> {
        let mut cursor = CfCursor::new(data, "BrtBeginDatabar14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginDatabar14", "nonzero FRTBlank"));
        }
        let min_length = cursor.read_u8()?;
        let max_length = cursor.read_u8()?;
        let show_value = cursor.read_bool8()?;
        let direction = Direction14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid direction"))?;
        let axis_position = AxisPosition14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid axis position"))?;
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if min_length > max_length || max_length > 100 {
            return Err(invalid(
                "BrtBeginDatabar14",
                "invalid minimum/maximum length",
            ));
        }
        Ok(BarHeader14 {
            min_length,
            max_length,
            show_value,
            direction,
            axis_position,
            border: flags & 0x01 != 0,
            gradient: flags & 0x02 != 0,
            custom_negative_fill: flags & 0x04 != 0,
            custom_negative_border: flags & 0x08 != 0,
            unused_flags: flags & 0xfff0,
        })
    }

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
        if self.min_length > self.max_length
            || self.max_length > 100
            || self.unused_flags & 0x0f != 0
        {
            return Err(invalid("BrtBeginDatabar14", "invalid data-bar header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.border);
        flags |= u16::from(self.gradient) << 1;
        flags |= u16::from(self.custom_negative_fill) << 2;
        flags |= u16::from(self.custom_negative_border) << 3;
        let mut data = Vec::with_capacity(11);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&[
            self.min_length,
            self.max_length,
            u8::from(self.show_value),
            self.direction as u8,
            self.axis_position as u8,
        ]);
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

impl Icon {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFIcon");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtCFIcon", "nonzero FRTBlank"));
        }
        let value = Self {
            icon_set: cursor.read_i32()?,
            index: cursor.read_i32()?,
        };
        cursor.finish()?;
        value.validate()?;
        Ok(value)
    }

    pub fn serialize(self) -> Result<[u8; 12]> {
        self.validate()?;
        let mut data = [0; 12];
        data[4..8].copy_from_slice(&self.icon_set.to_le_bytes());
        data[8..].copy_from_slice(&self.index.to_le_bytes());
        Ok(data)
    }

    fn validate(self) -> Result<()> {
        if self.icon_set == -1 {
            if self.index == -1 {
                return Ok(());
            }
        } else if let Ok(icon_set) = u8::try_from(self.icon_set)
            && icon_set <= 19
            && (0..icon_count14(icon_set) as i32).contains(&self.index)
        {
            return Ok(());
        }
        Err(invalid("BrtCFIcon", "invalid icon set or index"))
    }
}

impl IconSet14 {
    pub fn parse_header(data: &[u8]) -> Result<IconHeader14> {
        let mut cursor = CfCursor::new(data, "BrtBeginIconSet14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginIconSet14", "nonzero FRTBlank"));
        }
        let icon_set_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtBeginIconSet14", "icon-set type overflow"))?;
        if icon_set_type > 19 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set type"));
        }
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if flags & 0xff80 != 0 {
            return Err(invalid("BrtBeginIconSet14", "reserved flags are nonzero"));
        }
        Ok(IconHeader14 {
            icon_set_type,
            custom: flags & 0x01 != 0,
            show_value: flags & 0x02 == 0,
            reverse: flags & 0x04 == 0,
            unused_flags: flags & 0x78,
        })
    }

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
        if self.icon_set_type > 19 || self.unused_flags & !0x78 != 0 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.custom_icons.is_some());
        flags |= u16::from(!self.show_value) << 1;
        flags |= u16::from(!self.reverse) << 2;
        let mut data = Vec::with_capacity(10);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::from(self.icon_set_type).to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

impl Rule {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtBeginCFRule");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let dxf_id = (raw_dxf != u32::MAX).then_some(raw_dxf);
        if matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        ) && dxf_id.is_some()
        {
            return Err(invalid(
                "BrtBeginCFRule",
                "visual rule has a differential-format index",
            ));
        }
        let priority = cursor.read_u32()?;
        if priority == 0 || priority > i32::MAX as u32 {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid priority {priority}"),
            ));
        }
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        ) && stop_if_true
        {
            return Err(invalid("BrtBeginCFRule", "visual rule sets stop-if-true"));
        }
        if rule_type != RuleType::TopN && (bottom || percent) {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-filter rule sets bottom/percent flags",
            ));
        }
        validate_parameter_and_flags(
            rule_type,
            template,
            parameter,
            above_average,
            bottom,
            percent,
        )?;
        let declared = [cursor.read_u32()?, cursor.read_u32()?, cursor.read_u32()?];
        let text = cursor.read_nullable_string()?;
        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-text template has a string parameter",
            ));
        }
        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
        for (index, size) in declared.into_iter().enumerate() {
            if size == 0 {
                continue;
            }
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != size as usize {
                return Err(invalid(
                    "BrtBeginCFRule",
                    format!(
                        "formula {} declared {size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        cursor.finish()?;
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == RuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();
        if rule_type == RuleType::CellIs && !matches!(operator, Some(1..=8)) {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid cell comparison operator {parameter}"),
            ));
        }

        Ok(Rule {
            rule_type,
            dxf_id,
            priority,
            stop_if_true,
            formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: None,
            classic_extension_guid: None,
        })
    }

    /// Parse an Office 2013 `BrtBeginCFRule14` payload.
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtBeginCFRule14", 2)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginCFRule14");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_extension14_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let signed_priority = cursor.read_i32()?;
        if signed_priority != -1 && signed_priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {signed_priority}"),
            ));
        }
        if signed_priority == -1 && (rule_type != RuleType::DataBar || raw_dxf != 0) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 requires a data-bar rule and zero DXF index",
            ));
        }
        let visual = matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        );
        if signed_priority > 0 && visual && raw_dxf != u32::MAX {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a differential-format index",
            ));
        }
        let dxf_id = if signed_priority == -1 || raw_dxf == u32::MAX {
            None
        } else {
            Some(raw_dxf)
        };
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule14", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if visual && stop_if_true {
            return Err(invalid("BrtBeginCFRule14", "visual rule sets stop-if-true"));
        }
        validate_parameter_and_flags(
            rule_type,
            template,
            parameter,
            above_average,
            bottom,
            percent,
        )?;
        let declared = [cursor.read_u32()?, cursor.read_u32()?, cursor.read_u32()?];
        let unused = cursor.read_u32()?;
        let guid = cursor.read_array::<16>()?;
        let guid_present = cursor.read_bool32()?;
        let text = cursor.read_nullable_string()?;
        cursor.finish()?;

        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "non-text template has a string parameter",
            ));
        }

        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
        let mut formula_iter = formulas.into_iter();
        for (index, declared_size) in declared.into_iter().enumerate() {
            if declared_size == 0 {
                continue;
            }
            let formula = formula_iter.next().ok_or_else(|| {
                invalid(
                    "BrtBeginCFRule14",
                    "declared formula is absent from FRTHeader",
                )
            })?;
            if formula.rgce.len() != declared_size as usize {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    format!(
                        "formula {} declared {declared_size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        if formula_iter.next().is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "FRTHeader contains an undeclared formula",
            ));
        }
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut binary_formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            binary_formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == RuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();

        Ok(Self {
            rule_type,
            dxf_id,
            priority: u32::try_from(signed_priority).unwrap_or(0),
            stop_if_true,
            formulas: binary_formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: Some(RuleMetadata {
                priority: signed_priority,
                unused,
                guid,
                guid_present,
                linked_classic_priority: None,
            }),
            classic_extension_guid: None,
        })
    }

    /// Serialize an Office 2013 `BrtBeginCFRule14` payload.
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
        let metadata = self.extension14.ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                "rule does not contain Office 2013 metadata",
            )
        })?;
        validate_extension14_template(self.rule_type, self.template)?;
        if metadata.priority != -1 && metadata.priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {}", metadata.priority),
            ));
        }
        if metadata.priority > 0 && self.priority != metadata.priority as u32 {
            return Err(invalid(
                "BrtBeginCFRule14",
                "classic and extension priorities disagree",
            ));
        }
        if metadata.priority == -1 && self.rule_type != RuleType::DataBar {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 is only valid for a data-bar extension",
            ));
        }
        let parameter = effective_rule_parameter(self)?;
        validate_parameter_and_flags(
            self.rule_type,
            self.template,
            parameter,
            self.above_average,
            self.bottom,
            self.percent,
        )?;
        let visual = matches!(
            self.rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        );
        if visual && (self.stop_if_true || (metadata.priority > 0 && self.dxf_id.is_some())) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a DXF or stop-if-true flag",
            ));
        }
        if metadata.priority == -1 && self.dxf_id.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "data-bar extension has a DXF index",
            ));
        }
        validate_rule_text(self.template, self.text.as_deref(), "BrtBeginCFRule14")?;

        let formulas = effective_rule_formulas(self)?;
        validate_formula_count(self.rule_type, self.template, parameter, formulas.len())?;
        let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
        let start = if visual { 2 } else { 0 };
        for (index, formula) in formulas.iter().enumerate() {
            slots[start + index] = Some(formula);
        }
        let owned_slots = slots.each_ref().map(|formula| formula.cloned());
        validate_formula_slots(self.rule_type, self.template, parameter, &owned_slots)?;

        let mut payload = serialize_formula_header(&formulas, 2)?;
        payload.extend_from_slice(&(self.rule_type as u32).to_le_bytes());
        payload.extend_from_slice(&self.template.to_le_bytes());
        let raw_dxf = if metadata.priority == -1 {
            0
        } else {
            self.dxf_id.unwrap_or(u32::MAX)
        };
        payload.extend_from_slice(&raw_dxf.to_le_bytes());
        payload.extend_from_slice(&metadata.priority.to_le_bytes());
        payload.extend_from_slice(&parameter.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let mut flags = 0u16;
        flags |= u16::from(self.stop_if_true) << 1;
        flags |= u16::from(self.above_average) << 2;
        flags |= u16::from(self.bottom) << 3;
        flags |= u16::from(self.percent) << 4;
        payload.extend_from_slice(&flags.to_le_bytes());
        for formula in &slots {
            payload.extend_from_slice(
                &u32::try_from(formula.map_or(0, |formula| formula.rgce.len()))
                    .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
                    .to_le_bytes(),
            );
        }
        payload.extend_from_slice(&metadata.unused.to_le_bytes());
        payload.extend_from_slice(&metadata.guid);
        payload.extend_from_slice(&u32::from(metadata.guid_present).to_le_bytes());
        write_nullable_string(&mut payload, self.text.as_deref())?;
        Ok(payload)
    }
}

impl Formatting {
    /// Parse an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn parse_extension14_header(data: &[u8]) -> Result<(Self, u32)> {
        let (formatting, count, _) = Self::parse_extension14_header_with_base(data)?;
        Ok((formatting, count))
    }

    pub fn parse_extension14_header_with_base(data: &[u8]) -> Result<(Self, u32, (u32, u32))> {
        let (ranges, header_size) =
            parse_sqref_header(data, "BrtBeginConditionalFormatting14", i32::MAX as usize)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginConditionalFormatting14");
        let count = cursor.read_u32()?;
        let pivot_only = cursor.read_bool32()?;
        cursor.finish()?;
        let base = (ranges[0].0, ranges[0].2);
        let ranges = ranges
            .into_iter()
            .map(|(first_row, last_row, first_col, last_col)| {
                let first = cell_reference(first_row, first_col);
                let last = cell_reference(last_row, last_col);
                if first == last {
                    first
                } else {
                    format!("{first}:{last}")
                }
            })
            .collect();
        Ok((
            Self {
                ranges,
                rules: Vec::new(),
                pivot_only,
                record_kind: RecordKind::Extension14,
            },
            count,
            base,
        ))
    }

    /// Serialize an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn serialize_extension14_header(&self) -> Result<Vec<u8>> {
        let mut ranges = Vec::new();
        for range_list in &self.ranges {
            for range in range_list
                .split([',', ' '])
                .filter(|range| !range.is_empty())
            {
                let (first, last) = range.split_once(':').unwrap_or((range, range));
                let (first_row, first_col) = parse_cell_reference(first)?;
                let (last_row, last_col) = parse_cell_reference(last)?;
                ranges.push((first_row, last_row, first_col, last_col));
            }
        }
        let mut data = serialize_sqref_header(&ranges)?;
        data.extend_from_slice(
            &u32::try_from(self.rules.len())
                .map_err(|_| invalid("BrtBeginConditionalFormatting14", "rule count overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(&u32::from(self.pivot_only).to_le_bytes());
        Ok(data)
    }
}

pub fn parse_classic_header(data: &[u8]) -> Result<(Formatting, u32, (u32, u32))> {
    let mut cursor = CfCursor::new(data, "BrtBeginConditionalFormatting");
    let count = cursor.read_u32()?;
    let pivot_only = cursor.read_bool32()?;
    let ranges = cursor.read_ranges(1, 8_192)?;
    cursor.finish()?;
    let base = (ranges[0].0, ranges[0].2);
    let ranges = ranges
        .into_iter()
        .map(|(first_row, last_row, first_col, last_col)| {
            let first = cell_reference(first_row, first_col);
            let last = cell_reference(last_row, last_col);
            if first == last {
                first
            } else {
                format!("{first}:{last}")
            }
        })
        .collect();
    Ok((
        Formatting {
            ranges,
            rules: Vec::new(),
            pivot_only,
            record_kind: RecordKind::Classic,
        },
        count,
        base,
    ))
}

pub fn validate_template(rule_type: RuleType, template: u32) -> Result<()> {
    let valid = match rule_type {
        RuleType::CellIs => template == 0,
        RuleType::Expression => matches!(template, 1 | 7..=12 | 15..=27 | 29 | 30),
        RuleType::ColorScale => template == 2,
        RuleType::DataBar => template == 3,
        RuleType::TopN => template == 5,
        RuleType::IconSet => template == 4,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("template {template} is invalid for {rule_type:?}"),
        ))
    }
}

fn validate_extension14_template(rule_type: RuleType, template: u32) -> Result<()> {
    if rule_type == RuleType::DataBar && template == 0 {
        Ok(())
    } else {
        validate_template(rule_type, template).map_err(|_| {
            invalid(
                "BrtBeginCFRule14",
                format!("template {template} is invalid for {rule_type:?}"),
            )
        })
    }
}

pub fn validate_formula_count(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    count: usize,
) -> Result<()> {
    let expected = if rule_type == RuleType::CellIs {
        if matches!(parameter, 1 | 2) { 2 } else { 1 }
    } else if rule_type == RuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        1
    } else {
        0
    };
    let valid = if matches!(
        rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        count <= 1
    } else {
        count == expected
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula count {count} does not match required {expected}"),
        ))
    }
}

fn validate_formula_slots(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    slots: &[Option<ParsedFormula>; 3],
) -> Result<()> {
    let expected = if rule_type == RuleType::CellIs {
        [true, matches!(parameter, 1 | 2), false]
    } else if rule_type == RuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        [true, false, false]
    } else if matches!(
        rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        [false, false, slots[2].is_some()]
    } else {
        [false, false, false]
    };
    let found = slots.each_ref().map(Option::is_some);
    if found == expected {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula slots {found:?} do not match required {expected:?}"),
        ))
    }
}

fn validate_parameter_and_flags(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    above_average: bool,
    bottom: bool,
    percent: bool,
) -> Result<()> {
    let valid_parameter = match (rule_type, template) {
        (RuleType::CellIs, 0) => (1..=8).contains(&parameter),
        (RuleType::Expression, 8) => parameter <= 3,
        (RuleType::Expression, 15) => parameter == 0,
        (RuleType::Expression, 16) => parameter == 6,
        (RuleType::Expression, 17) => parameter == 1,
        (RuleType::Expression, 18) => parameter == 2,
        (RuleType::Expression, 19) => parameter == 5,
        (RuleType::Expression, 20) => parameter == 8,
        (RuleType::Expression, 21) => parameter == 3,
        (RuleType::Expression, 22) => parameter == 7,
        (RuleType::Expression, 23) => parameter == 4,
        (RuleType::Expression, 24) => parameter == 9,
        (RuleType::Expression, 25 | 26) => parameter < 4,
        (RuleType::TopN, 5) if percent => parameter <= 100,
        (RuleType::TopN, 5) => (1..=1_000).contains(&parameter),
        _ => parameter == 0,
    };
    if !valid_parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid parameter {parameter} for template {template}"),
        ));
    }
    if above_average != matches!(template, 25 | 29) {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid above-average flag for template {template}"),
        ));
    }
    if rule_type != RuleType::TopN && (bottom || percent) {
        return Err(invalid(
            "BrtBeginCFRule",
            "bottom/percent flags are set on a non-filter rule",
        ));
    }
    Ok(())
}

fn render_formula(
    formula: &ParsedFormula,
    base: (u32, u32),
    context: &impl Resolution,
) -> Result<String> {
    let tokens =
        Parser::with_base_cell_and_extra(&formula.rgce, &formula.rgcb, base.0, base.1).parse()?;
    Ok(Compiler::try_tokens_to_string_with_resolution(
        &tokens, context,
    )?)
}

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn effective_rule_parameter(rule: &Rule) -> Result<u32> {
    if rule.rule_type != RuleType::CellIs {
        if rule.operator.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "operator is set on a non-cell-comparison rule",
            ));
        }
        return Ok(rule.parameter);
    }
    let parameter = rule.operator.map_or(rule.parameter, u32::from);
    if rule.parameter != 0 && rule.parameter != parameter {
        return Err(invalid(
            "BrtBeginCFRule14",
            "operator and exact parameter disagree",
        ));
    }
    Ok(parameter)
}

fn effective_rule_formulas(rule: &Rule) -> Result<Vec<ParsedFormula>> {
    if !rule.formulas.is_empty() {
        if !rule.formula_extras.is_empty() && rule.formula_extras.len() != rule.formulas.len() {
            return Err(Error::InvalidFormula(
                "conditional-format ancillary stream count does not match formulas".to_string(),
            ));
        }
        return rule
            .formulas
            .iter()
            .enumerate()
            .map(|(index, rgce)| {
                if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(Error::InvalidFormula(format!(
                        "conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                        rgce.len()
                    )));
                }
                Ok(ParsedFormula {
                    rgce: rgce.clone(),
                    rgcb: rule.formula_extras.get(index).cloned().unwrap_or_default(),
                })
            })
            .collect();
    }
    rule.formula_texts
        .iter()
        .map(|formula| TextCompiler::compile(formula))
        .collect()
}

fn validate_rule_text(template: u32, text: Option<&str>, record: &'static str) -> Result<()> {
    if template == 8 {
        if text.is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255) {
            return Err(invalid(record, "invalid text parameter"));
        }
    } else if text.is_some() {
        return Err(invalid(record, "non-text template has a text parameter"));
    }
    Ok(())
}

fn write_nullable_string(data: &mut Vec<u8>, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().collect::<Vec<_>>();
    data.extend_from_slice(
        &u32::try_from(units.len())
            .map_err(|_| invalid("XLNullableWideString", "string length overflow"))?
            .to_le_bytes(),
    );
    for unit in units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn invalid(typ: impl Into<String>, val: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

struct CfCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> CfCursor<'a> {
    fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, size: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(size)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u16(&mut self) -> Result<u16> {
        let bytes = self.take(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn read_bool8(&mut self) -> Result<bool> {
        match self.read_u8()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_i32(&mut self) -> Result<i32> {
        let bytes = self.take(4)?;
        Ok(i32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_array<const N: usize>(&mut self) -> Result<[u8; N]> {
        self.take(N)?.try_into().map_err(|_| Error::InvalidLength {
            expected: N,
            found: self.data.len().saturating_sub(self.offset),
        })
    }

    fn read_f64(&mut self) -> Result<f64> {
        let bytes = self.take(8)?;
        Ok(f64::from_le_bytes(
            bytes.try_into().expect("eight-byte field"),
        ))
    }

    fn read_bool32(&mut self) -> Result<bool> {
        match self.read_u32()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    fn read_nullable_string(&mut self) -> Result<Option<String>> {
        let count = self.read_u32()?;
        if count == u32::MAX {
            return Ok(None);
        }
        let count = count as usize;
        let bytes = self.take(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid(self.record, "string size overflow"))?,
        )?;
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map(Some)
            .map_err(|error| Error::Encoding(format!("invalid UTF-16: {error}")))
    }

    fn read_formula(&mut self) -> Result<ParsedFormula> {
        let cce = self.read_u32()? as usize;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "conditional-format formula length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = self.read_u32()? as usize;
        let rgcb = self.take(cb)?.to_vec();
        Ok(ParsedFormula { rgce, rgcb })
    }

    fn read_ranges(&mut self, minimum: usize, maximum: usize) -> Result<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()? as i32;
        let count = usize::try_from(raw_count)
            .map_err(|_| invalid(self.record, "NULL range collection"))?;
        if !(minimum..=maximum).contains(&count)
            || count > self.data.len().saturating_sub(self.offset) / 16
        {
            return Err(invalid(self.record, format!("invalid range count {count}")));
        }
        let mut ranges = Vec::with_capacity(count);
        for _ in 0..count {
            let first_row = self.read_u32()?;
            let last_row = self.read_u32()?;
            let first_col = self.read_u32()?;
            let last_col = self.read_u32()?;
            if first_row > last_row
                || first_col > last_col
                || last_row >= 1_048_576
                || last_col >= 16_384
            {
                return Err(invalid(self.record, "invalid target range"));
            }
            ranges.push((first_row, last_row, first_col, last_col));
        }
        Ok(ranges)
    }

    fn finish(self) -> Result<()> {
        if self.offset == self.data.len() {
            Ok(())
        } else {
            Err(Error::InvalidLength {
                expected: self.offset,
                found: self.data.len(),
            })
        }
    }
}

#[cfg(test)]
mod model_tests {
    use super::*;

    fn numeric_cfvo_payload(cfvo_type: u32, value: f64) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&cfvo_type.to_le_bytes());
        data.extend_from_slice(&value.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data
    }

    fn cell_rule_payload(dxf_id: u32, priority: u32, stop: bool, operator: u32) -> Vec<u8> {
        let formula = TextCompiler::compile("1").unwrap();
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&dxf_id.to_le_bytes());
        data.extend_from_slice(&priority.to_le_bytes());
        data.extend_from_slice(&operator.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&(u16::from(stop) << 1).to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        data
    }

    #[test]
    fn test_cf_rule_type_from_u8() {
        assert_eq!(RuleType::from_u8(1), Some(RuleType::CellIs));
        assert_eq!(RuleType::from_u8(2), Some(RuleType::Expression));
        assert_eq!(RuleType::from_u8(3), Some(RuleType::ColorScale));
        assert_eq!(RuleType::from_u8(4), Some(RuleType::DataBar));
        assert_eq!(RuleType::from_u8(5), Some(RuleType::TopN));
        assert_eq!(RuleType::from_u8(6), Some(RuleType::IconSet));
        assert_eq!(RuleType::from_u8(0), None);
        assert_eq!(RuleType::from_u8(7), None);
        assert_eq!(RuleType::from_u8(255), None);
    }

    #[test]
    fn test_cfvo_new() {
        let cfvo = Value::new(1, Some("10".to_string()));
        assert_eq!(cfvo.cfvo_type, 1);
        assert_eq!(cfvo.value, Some("10".to_string()));
    }

    #[test]
    fn test_cfvo_serialize_roundtrip() {
        let parsed = Value::parse(&numeric_cfvo_payload(1, 50.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 1);
        assert_eq!(parsed.value.as_deref(), Some("50"));
        assert_eq!(parsed.numeric_value, 50.0);
    }

    #[test]
    fn test_cfvo_serialize_none_value() {
        let parsed = Value::parse(&numeric_cfvo_payload(2, 0.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 2);
        assert!(parsed.value.is_none());
        assert!(parsed.formula_binary.is_none());
    }

    #[test]
    fn test_cfvo_parse_too_short() {
        let result = Value::parse(&[0x01]);
        assert!(result.is_err());
    }

    #[test]
    fn test_color_scale_new() {
        let min_cfvo = Value::new(2, None); // min
        let max_cfvo = Value::new(3, None); // max
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);

        assert_eq!(cs.min_cfvo.cfvo_type, 2);
        assert_eq!(cs.max_cfvo.cfvo_type, 3);
        assert_eq!(cs.min_color, 0xFFFF0000);
        assert_eq!(cs.max_color, 0xFF00FF00);
        assert!(cs.mid_cfvo.is_none());
        assert!(cs.mid_color.is_none());
    }

    #[test]
    fn test_color_scale_with_middle() {
        let min_cfvo = Value::new(2, None);
        let mid_cfvo = Value::new(1, Some("50".to_string()));
        let max_cfvo = Value::new(3, None);
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00)
            .with_middle(mid_cfvo, 0xFFFFFF00);

        assert!(cs.mid_cfvo.is_some());
        assert!(cs.mid_color.is_some());
        assert_eq!(cs.mid_color.unwrap(), 0xFFFFFF00);
    }

    #[test]
    fn test_data_bar_new() {
        let min_cfvo = Value::new(2, None);
        let max_cfvo = Value::new(3, None);
        let db = Bar::new(min_cfvo, max_cfvo, 0xFF4472C4);

        assert_eq!(db.min_cfvo.cfvo_type, 2);
        assert_eq!(db.max_cfvo.cfvo_type, 3);
        assert_eq!(db.color, 0xFF4472C4);
        assert!(db.show_value);
    }

    #[test]
    fn test_icon_set_new() {
        let cfvos = vec![
            Value::new(1, Some("0".to_string())),
            Value::new(1, Some("33".to_string())),
            Value::new(1, Some("67".to_string())),
        ];
        let icon_set = IconSet::new(0x01, cfvos); // 3Arrows

        assert_eq!(icon_set.icon_set_type, 0x01);
        assert_eq!(icon_set.cfvos.len(), 3);
        assert!(icon_set.show_value);
        assert!(!icon_set.reverse);
    }

    #[test]
    fn test_conditional_formatting_rule_new() {
        let rule = Rule::new(RuleType::CellIs, 1);

        assert_eq!(rule.rule_type, RuleType::CellIs);
        assert_eq!(rule.priority, 1);
        assert!(rule.dxf_id.is_none());
        assert!(!rule.stop_if_true);
        assert!(rule.formulas.is_empty());
        assert!(rule.color_scale.is_none());
        assert!(rule.data_bar.is_none());
        assert!(rule.icon_set.is_none());
        assert!(rule.operator.is_none());
    }

    #[test]
    fn test_conditional_formatting_new() {
        let ranges = vec!["A1:B10".to_string()];
        let cf = Formatting::new(ranges);

        assert_eq!(cf.ranges.len(), 1);
        assert_eq!(cf.ranges[0], "A1:B10");
        assert!(cf.rules.is_empty());
    }

    #[test]
    fn test_conditional_formatting_add_rule() {
        let mut cf = Formatting::new(vec!["A1:A10".to_string()]);
        let rule = Rule::new(RuleType::CellIs, 1);
        cf.add_rule(rule);

        assert_eq!(cf.rules.len(), 1);
        assert_eq!(cf.rules[0].rule_type, RuleType::CellIs);
    }

    #[test]
    fn test_conditional_formatting_rule_parse() {
        let rule = Rule::parse(&cell_rule_payload(u32::MAX, 1, false, 5)).unwrap();
        assert_eq!(rule.rule_type, RuleType::CellIs);
        assert!(rule.dxf_id.is_none());
        assert_eq!(rule.priority, 1);
        assert!(!rule.stop_if_true);
        assert_eq!(rule.operator, Some(5));
    }

    #[test]
    fn test_conditional_formatting_rule_parse_with_dxf() {
        let rule = Rule::parse(&cell_rule_payload(5, 10, true, 3)).unwrap();
        assert_eq!(rule.dxf_id, Some(5));
        assert_eq!(rule.priority, 10);
        assert!(rule.stop_if_true);
    }

    #[test]
    fn test_conditional_formatting_rule_parse_too_short() {
        let data = [0x01, 0x02, 0x03]; // too short
        let result = Rule::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_conditional_formatting_rule_parse_invalid_type() {
        let data = [
            0xFF, // invalid type
            0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x00,
        ];
        let result = Rule::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_read_optional_string_none() {
        let data = u32::MAX.to_le_bytes();
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(cursor.read_nullable_string().unwrap(), None);
        cursor.finish().unwrap();
    }

    #[test]
    fn test_read_optional_string_some() {
        // "Hi" encoded as UTF-16LE with length prefix
        let data = [
            0x02, 0x00, 0x00, 0x00, // length = 2
            0x48, 0x00, // 'H'
            0x69, 0x00, // 'i'
        ];
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(
            cursor.read_nullable_string().unwrap().as_deref(),
            Some("Hi")
        );
        cursor.finish().unwrap();
    }

    #[test]
    fn test_read_optional_string_too_short() {
        let data = [0x01]; // too short
        let mut cursor = CfCursor::new(&data, "test");
        assert!(cursor.read_nullable_string().is_err());
    }

    #[test]
    fn test_write_optional_string_none() {
        let data = u32::MAX.to_le_bytes();
        let mut cursor = CfCursor::new(&data, "test");
        assert!(cursor.read_nullable_string().unwrap().is_none());
    }

    #[test]
    fn test_write_optional_string_some() {
        let data = [0x04, 0x00, 0x00, 0x00, b'T', 0, b'e', 0, b's', 0, b't', 0];
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(
            cursor.read_nullable_string().unwrap().as_deref(),
            Some("Test")
        );
    }

    #[test]
    fn test_cf_rule_type_variants() {
        // Verify all enum variants have correct discriminant values
        assert_eq!(RuleType::CellIs as u8, 1);
        assert_eq!(RuleType::Expression as u8, 2);
        assert_eq!(RuleType::ColorScale as u8, 3);
        assert_eq!(RuleType::DataBar as u8, 4);
        assert_eq!(RuleType::TopN as u8, 5);
        assert_eq!(RuleType::IconSet as u8, 6);
    }

    #[test]
    fn test_conditional_formatting_clone() {
        let mut cf = Formatting::new(vec!["A1:A10".to_string()]);
        let rule = Rule::new(RuleType::CellIs, 1);
        cf.add_rule(rule);

        let cloned = cf.clone();
        assert_eq!(cloned.ranges.len(), cf.ranges.len());
        assert_eq!(cloned.rules.len(), cf.rules.len());
    }

    #[test]
    fn test_color_scale_clone() {
        let min_cfvo = Value::new(2, None);
        let max_cfvo = Value::new(3, None);
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);
        let cloned = cs.clone();

        assert_eq!(cloned.min_color, cs.min_color);
        assert_eq!(cloned.max_color, cs.max_color);
    }
}

#[cfg(test)]
mod writer_tests {
    use super::*;

    fn fixture_cell_is_payload() -> Vec<u8> {
        let formula = TextCompiler::compile("5").unwrap();
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        data
    }

    #[test]
    fn parses_normative_cell_is_rule() {
        let rule = Rule::parse(&fixture_cell_is_payload()).unwrap();
        assert_eq!(rule.rule_type, RuleType::CellIs);
        assert_eq!(rule.template, 0);
        assert_eq!(rule.parameter, 5);
        assert_eq!(rule.operator, Some(5));
        assert_eq!(rule.formula_texts, ["5"]);
        assert_eq!(rule.formulas.len(), 1);
        assert_eq!(rule.formula_extras, [Vec::<u8>::new()]);
    }

    #[test]
    fn rejects_formula_in_wrong_slot_or_with_wrong_declared_size() {
        let mut wrong_slot = fixture_cell_is_payload();
        let size = wrong_slot[30..34].to_vec();
        wrong_slot[30..34].fill(0);
        wrong_slot[34..38].copy_from_slice(&size);
        assert!(Rule::parse(&wrong_slot).is_err());

        let mut wrong_size = fixture_cell_is_payload();
        wrong_size[30..34].copy_from_slice(&4u32.to_le_bytes());
        assert!(Rule::parse(&wrong_size).is_err());
    }

    #[test]
    fn parses_cfvo_with_ancillary_formula_losslessly() {
        let formula = TextCompiler::compile("{1,2}").unwrap();
        assert!(!formula.rgcb.is_empty());
        let mut data = Vec::new();
        data.extend_from_slice(&7u32.to_le_bytes());
        data.extend_from_slice(&0f64.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        let parsed = Value::parse(&data).unwrap();
        assert_eq!(parsed.formula_binary.as_ref().unwrap(), &formula);
    }

    #[test]
    fn extension_cfvo_roundtrips_formula_and_automatic_bounds() {
        let formula = TextCompiler::compile("$A$1").unwrap();
        let formula_value = Value {
            cfvo_type: 7,
            value: Some("$A$1".to_string()),
            numeric_value: 0.0,
            save_greater_than_or_equal: true,
            greater_than_or_equal: false,
            formula_binary: Some(formula.clone()),
        };
        let encoded = formula_value.serialize_extension14().unwrap();
        let parsed = Value::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.cfvo_type, 7);
        assert_eq!(parsed.formula_binary, Some(formula));
        assert!(!parsed.greater_than_or_equal);

        for cfvo_type in [8, 9] {
            let automatic = Value {
                cfvo_type,
                value: None,
                numeric_value: 0.0,
                save_greater_than_or_equal: false,
                greater_than_or_equal: true,
                formula_binary: None,
            };
            let encoded = automatic.serialize_extension14().unwrap();
            assert_eq!(Value::parse_extension14(&encoded).unwrap(), automatic);
        }
    }

    #[test]
    fn extension_cfvo_rejects_inconsistent_formula_metadata() {
        let formula = TextCompiler::compile("1").unwrap();
        let value = Value {
            cfvo_type: 7,
            value: Some("1".to_string()),
            numeric_value: 0.0,
            save_greater_than_or_equal: false,
            greater_than_or_equal: true,
            formula_binary: Some(formula),
        };
        let mut encoded = value.serialize_extension14().unwrap();
        let declared_offset = encoded.len() - 4;
        encoded[declared_offset..].copy_from_slice(&999u32.to_le_bytes());
        assert!(Value::parse_extension14(&encoded).is_err());
    }

    #[test]
    fn parses_direct_and_theme_colors() {
        let direct = Color::parse(&[5, 0, 0, 0, 0x11, 0x22, 0x33, 0xff]).unwrap();
        assert_eq!(direct.argb, Some(0xff11_2233));
        let theme = Color::parse(&[6, 4, 0, 0, 0, 0, 0, 0]).unwrap();
        assert_eq!(theme.color_type, 3);
        assert_eq!(theme.index, 4);
        assert!(Color::parse(&[6, 12, 0, 0, 0, 0, 0, 0]).is_err());

        let theme = Color::theme(5, -1_000).unwrap();
        assert_eq!(Color::parse(&theme.to_bytes().unwrap()).unwrap(), theme);
        let mut indexed = Color::indexed(42, 0);
        indexed.tint = 2_000;
        let reparsed = Color::parse(&indexed.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.index, 42);
        assert_eq!(reparsed.tint, 2_000);
    }

    #[test]
    fn extension_color_and_rule_guid_roundtrip_exactly() {
        let color = Color::theme(4, -2_500).unwrap();
        let encoded = color.serialize_extension14().unwrap();
        assert_eq!(Color::parse_extension14(&encoded).unwrap(), color);
        let mut malformed = encoded;
        malformed[0] = 1;
        assert!(Color::parse_extension14(&malformed).is_err());

        let guid = [0x42; 16];
        let encoded = serialize_rule_extension_guid(guid);
        assert_eq!(parse_rule_extension_guid(&encoded).unwrap(), guid);
        let mut malformed = encoded;
        malformed[3] = 1;
        assert!(parse_rule_extension_guid(&malformed).is_err());
    }

    #[test]
    fn extension_data_bar_header_preserves_flags() {
        let mut bar = Bar14::new(
            Value::new(8, None),
            Value::new(9, None),
            Color::from_argb(0xff44_72c4),
        );
        bar.min_length = 3;
        bar.max_length = 97;
        bar.show_value = false;
        bar.direction = Direction14::RightToLeft;
        bar.axis_position = AxisPosition14::Midpoint;
        bar.border = true;
        bar.custom_negative_fill = true;
        bar.unused_flags = 0xA5F0;
        let encoded = bar.serialize_header().unwrap();
        let parsed = Bar14::parse_header(&encoded).unwrap();
        assert_eq!(parsed.min_length, 3);
        assert_eq!(parsed.max_length, 97);
        assert!(!parsed.show_value);
        assert_eq!(parsed.direction, Direction14::RightToLeft);
        assert_eq!(parsed.axis_position, AxisPosition14::Midpoint);
        assert!(parsed.border);
        assert!(parsed.gradient);
        assert!(parsed.custom_negative_fill);
        assert_eq!(parsed.unused_flags, 0xA5F0);

        let mut malformed = encoded;
        malformed[6] = 2;
        assert!(Bar14::parse_header(&malformed).is_err());
    }

    #[test]
    fn extension_icon_set_and_custom_icons_roundtrip() {
        let mut set = IconSet14::new(19, vec![Value::new(1, Some("0".to_string())); 5]);
        set.show_value = false;
        set.reverse = true;
        set.unused_flags = 0x38;
        set.custom_icons = Some(vec![
            Icon {
                icon_set: -1,
                index: -1,
            };
            5
        ]);
        let encoded = set.serialize_header().unwrap();
        let parsed = IconSet14::parse_header(&encoded).unwrap();
        assert_eq!(parsed.icon_set_type, 19);
        assert!(parsed.custom);
        assert!(!parsed.show_value);
        assert!(parsed.reverse);
        assert_eq!(parsed.unused_flags, 0x38);

        for icon in set.custom_icons.unwrap() {
            let encoded = icon.serialize().unwrap();
            assert_eq!(Icon::parse(&encoded).unwrap(), icon);
        }
        assert!(
            Icon {
                icon_set: 0,
                index: 3,
            }
            .serialize()
            .is_err()
        );
    }

    #[test]
    fn parses_classic_header_with_pivot_and_range() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        let (formatting, count, base) = parse_classic_header(&data).unwrap();
        assert_eq!(count, 1);
        assert!(formatting.pivot_only);
        assert_eq!(formatting.ranges, ["A1:B2"]);
        assert_eq!(base, (0, 0));
    }

    #[test]
    fn extension_header_roundtrips_ranges_and_pivot_flag() {
        let mut formatting = Formatting::new(vec!["A1:B2 C3".to_string()]);
        formatting.pivot_only = true;
        formatting.rules.push(Rule::new(RuleType::Expression, 1));
        let encoded = formatting.serialize_extension14_header().unwrap();
        let (parsed, count) = Formatting::parse_extension14_header(&encoded).unwrap();
        assert_eq!(count, 1);
        assert_eq!(parsed.ranges, ["A1:B2", "C3"]);
        assert!(parsed.pivot_only);
        assert_eq!(parsed.record_kind, RecordKind::Extension14);
    }

    #[test]
    fn extension_rule_roundtrips_two_formulas_and_ancillary_data() {
        let first = TextCompiler::compile("{1,2}").unwrap();
        let second = TextCompiler::compile("10").unwrap();
        assert!(!first.rgcb.is_empty());
        let mut rule = Rule::new(RuleType::CellIs, 7);
        rule.operator = Some(1);
        rule.parameter = 1;
        rule.formulas = vec![first.rgce.clone(), second.rgce.clone()];
        rule.formula_extras = vec![first.rgcb.clone(), second.rgcb.clone()];
        rule.dxf_id = Some(4);
        rule.extension14 = Some(RuleMetadata {
            priority: 7,
            unused: 0xA5A5_5A5A,
            guid: [0x3c; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = Rule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 7);
        assert_eq!(parsed.operator, Some(1));
        assert_eq!(parsed.formulas, [first.rgce, second.rgce]);
        assert_eq!(parsed.formula_extras, [first.rgcb, second.rgcb]);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_preserves_signed_data_bar_linkage() {
        let mut rule = Rule::new(RuleType::DataBar, 0);
        rule.template = 0;
        rule.extension14 = Some(RuleMetadata {
            priority: -1,
            unused: 0xDEAD_BEEF,
            guid: [0x96; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = Rule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 0);
        assert_eq!(parsed.template, 0);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_rejects_malformed_fixed_and_formula_fields() {
        let mut rule = Rule::new(RuleType::Expression, 2);
        rule.formula_texts.push("1".to_string());
        rule.extension14 = Some(RuleMetadata {
            priority: 2,
            unused: 0,
            guid: [0; 16],
            guid_present: false,
            linked_classic_priority: None,
        });
        let encoded = rule.serialize_extension14().unwrap();
        let (_, fixed_offset) = parse_formula_header(&encoded, "test", 2).unwrap();

        let mut reserved = encoded.clone();
        reserved[fixed_offset + 20..fixed_offset + 24].copy_from_slice(&1u32.to_le_bytes());
        assert!(Rule::parse_extension14(&reserved).is_err());

        let mut priority = encoded.clone();
        priority[fixed_offset + 12..fixed_offset + 16].copy_from_slice(&0i32.to_le_bytes());
        assert!(Rule::parse_extension14(&priority).is_err());

        let mut declared = encoded;
        declared[fixed_offset + 30..fixed_offset + 34].copy_from_slice(&999u32.to_le_bytes());
        assert!(Rule::parse_extension14(&declared).is_err());
    }
}

/// Write all classic and Office 2013 conditional-formatting collections for a worksheet.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut Writer<W>,
    cond_fmts: &[Formatting],
) -> Result<()> {
    validate_extension_links(cond_fmts)?;
    let mut priorities = HashSet::new();
    for rule in cond_fmts.iter().flat_map(|formatting| &formatting.rules) {
        let priority = rule
            .extension14
            .map_or(i64::from(rule.priority), |metadata| {
                i64::from(metadata.priority)
            });
        if priority > 0 && !priorities.insert(priority) {
            return Err(invalid(
                "BrtBeginCFRule priority",
                format!("duplicate {priority}"),
            ));
        }
    }
    for formatting in cond_fmts {
        match formatting.record_kind {
            RecordKind::Classic => write_single_cond_formatting(writer, formatting)?,
            RecordKind::Extension14 => write_single_cond_formatting14(writer, formatting)?,
        }
    }
    Ok(())
}

fn validate_extension_links(cond_fmts: &[Formatting]) -> Result<()> {
    let mut classic = HashMap::new();
    for formatting in cond_fmts {
        if formatting.record_kind != RecordKind::Classic {
            continue;
        }
        for rule in &formatting.rules {
            let Some(guid) = rule.classic_extension_guid else {
                continue;
            };
            let bar = rule
                .data_bar
                .as_ref()
                .ok_or_else(|| invalid("BrtCFRuleExt", "is attached to a non-data-bar rule"))?;
            if classic
                .insert(guid, (rule.priority, bar.min_length, bar.max_length))
                .is_some()
            {
                return Err(invalid("BrtCFRuleExt", "duplicate GUID"));
            }
        }
    }
    let mut matched = HashSet::new();
    for formatting in cond_fmts {
        if formatting.record_kind != RecordKind::Extension14 {
            continue;
        }
        for rule in &formatting.rules {
            let Some(metadata) = rule.extension14 else {
                continue;
            };
            if metadata.priority != -1 || !metadata.guid_present {
                continue;
            }
            let Some(&(priority, classic_min, classic_max)) = classic.get(&metadata.guid) else {
                if metadata.linked_classic_priority.is_some() {
                    return Err(invalid(
                        "BrtBeginCFRule14",
                        "resolved classic priority has no matching GUID",
                    ));
                }
                continue;
            };
            if !matched.insert(metadata.guid) {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "multiple data-bar extensions use the same GUID",
                ));
            }
            if metadata
                .linked_classic_priority
                .is_some_and(|linked| linked != priority)
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "resolved classic priority disagrees with its GUID",
                ));
            }
            let bar = rule
                .data_bar14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginDatabar14", "missing linked data bar"))?;
            let expected_lengths = if bar.min_length == 0 && bar.max_length == 100 {
                (10, 90)
            } else {
                (bar.min_length, bar.max_length)
            };
            if (classic_min, classic_max) != expected_lengths {
                return Err(invalid(
                    "BrtBeginDatabar14",
                    "widths do not agree with the linked classic data bar",
                ));
            }
        }
    }
    if classic.keys().any(|guid| !matched.contains(guid)) {
        return Err(invalid(
            "BrtCFRuleExt",
            "GUID has no matching data-bar extension",
        ));
    }
    Ok(())
}

fn write_single_cond_formatting<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING,
        &serialize_cond_formatting_header(formatting)?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE, &serialize_cf_rule(rule)?)?;
        write_rule_visualization(writer, rule)?;
        if let Some(guid) = rule.classic_extension_guid {
            writer.write_record(kind::CF_RULE_EXT, &serialize_rule_extension_guid(guid))?;
        }
        writer.write_record(kind::END_CF_RULE, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING, &[])?;
    Ok(())
}

fn write_single_cond_formatting14<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING14,
        &formatting.serialize_extension14_header()?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE14, &rule.serialize_extension14()?)?;
        write_rule_visualization14(writer, rule)?;
        writer.write_record(kind::END_CF_RULE14, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING14, &[])?;
    Ok(())
}

fn serialize_cond_formatting_header(formatting: &Formatting) -> Result<Vec<u8>> {
    let rule_count = u32::try_from(formatting.rules.len())
        .map_err(|_| invalid("BrtBeginConditionalFormatting", "too many rules"))?;
    let mut ranges = Vec::new();
    for range in &formatting.ranges {
        ranges.extend(parse_range_list(range)?);
    }
    if ranges.is_empty() || ranges.len() > 8_192 {
        return Err(invalid(
            "BrtBeginConditionalFormatting",
            format!("classic range count {} is outside 1..=8192", ranges.len()),
        ));
    }
    let mut payload = Vec::with_capacity(12 + ranges.len() * 16);
    payload.extend_from_slice(&rule_count.to_le_bytes());
    payload.extend_from_slice(&u32::from(formatting.pivot_only).to_le_bytes());
    write_bin_range_list(&ranges, &mut payload)?;
    Ok(payload)
}

fn serialize_cf_rule(rule: &Rule) -> Result<Vec<u8>> {
    validate_rule_metadata(rule)?;
    let parameter = effective_parameter(rule)?;
    let formulas = effective_formulas(rule)?;
    validate_formula_count(rule.rule_type, rule.template, parameter, formulas.len())?;

    let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
    let start = if matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        2
    } else {
        0
    };
    for (index, formula) in formulas.iter().enumerate() {
        slots[start + index] = Some(formula);
    }

    let mut payload = Vec::with_capacity(64);
    payload.extend_from_slice(&(rule.rule_type as u32).to_le_bytes());
    payload.extend_from_slice(&rule.template.to_le_bytes());
    payload.extend_from_slice(&rule.dxf_id.unwrap_or(u32::MAX).to_le_bytes());
    payload.extend_from_slice(&rule.priority.to_le_bytes());
    payload.extend_from_slice(&parameter.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    let mut flags = 0u16;
    if rule.stop_if_true {
        flags |= 0x02;
    }
    if rule.above_average {
        flags |= 0x04;
    }
    if rule.bottom {
        flags |= 0x08;
    }
    if rule.percent {
        flags |= 0x10;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    for formula in &slots {
        let size = formula.map_or(0, |formula| formula.rgce.len());
        let size = u32::try_from(size)
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?;
        payload.extend_from_slice(&size.to_le_bytes());
    }
    write_nullable_wide_string(&mut payload, rule.text.as_deref())?;
    for formula in slots.into_iter().flatten() {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    Ok(payload)
}

fn validate_rule_metadata(rule: &Rule) -> Result<()> {
    if rule.extension14.is_some()
        || rule.color_scale14.is_some()
        || rule.data_bar14.is_some()
        || rule.icon_set14.is_some()
    {
        return Err(invalid(
            "BrtBeginCFRule",
            "Office 2013 fields are set on a classic rule",
        ));
    }
    validate_template(rule.rule_type, rule.template)?;
    if rule.priority == 0 || rule.priority > i32::MAX as u32 {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid priority {}", rule.priority),
        ));
    }
    if rule.dxf_id.is_some_and(|id| id > i32::MAX as u32) {
        return Err(invalid(
            "BrtBeginCFRule",
            "differential-format index overflow",
        ));
    }
    let visual = matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    );
    if visual && (rule.dxf_id.is_some() || rule.stop_if_true) {
        return Err(invalid(
            "BrtBeginCFRule",
            "visual rule has a DXF or stop-if-true flag",
        ));
    }
    let expected_visual = match rule.rule_type {
        RuleType::ColorScale => {
            rule.color_scale.is_some() && rule.data_bar.is_none() && rule.icon_set.is_none()
        },
        RuleType::DataBar => {
            rule.color_scale.is_none() && rule.data_bar.is_some() && rule.icon_set.is_none()
        },
        RuleType::IconSet => {
            rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_some()
        },
        _ => rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_none(),
    };
    if !expected_visual {
        return Err(invalid(
            "BrtBeginCFRule",
            "visualization does not match rule type",
        ));
    }
    if rule.template == 8 {
        let valid = rule
            .text
            .as_deref()
            .is_some_and(|text| !text.is_empty() && text.encode_utf16().count() <= 255);
        if !valid {
            return Err(invalid("BrtBeginCFRule", "invalid text parameter"));
        }
    } else if rule.text.is_some() {
        return Err(invalid(
            "BrtBeginCFRule",
            "non-text template has a text parameter",
        ));
    }
    validate_rule_flags_and_parameter(rule, effective_parameter(rule)?)
}

fn validate_rule_flags_and_parameter(rule: &Rule, parameter: u32) -> Result<()> {
    let valid_parameter = match (rule.rule_type, rule.template) {
        (RuleType::CellIs, 0) => (1..=8).contains(&parameter),
        (RuleType::Expression, 8) => parameter <= 3,
        (RuleType::Expression, 15) => parameter == 0,
        (RuleType::Expression, 16) => parameter == 6,
        (RuleType::Expression, 17) => parameter == 1,
        (RuleType::Expression, 18) => parameter == 2,
        (RuleType::Expression, 19) => parameter == 5,
        (RuleType::Expression, 20) => parameter == 8,
        (RuleType::Expression, 21) => parameter == 3,
        (RuleType::Expression, 22) => parameter == 7,
        (RuleType::Expression, 23) => parameter == 4,
        (RuleType::Expression, 24) => parameter == 9,
        (RuleType::Expression, 25 | 26) => parameter < 4,
        (RuleType::TopN, 5) if rule.percent => parameter <= 100,
        (RuleType::TopN, 5) => (1..=1_000).contains(&parameter),
        _ => parameter == 0,
    };
    if !valid_parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            format!(
                "invalid parameter {parameter} for template {}",
                rule.template
            ),
        ));
    }
    let expected_above = matches!(rule.template, 25 | 29);
    if rule.above_average != expected_above {
        return Err(invalid(
            "BrtBeginCFRule",
            format!(
                "above-average flag is invalid for template {}",
                rule.template
            ),
        ));
    }
    if rule.rule_type != RuleType::TopN && (rule.bottom || rule.percent) {
        return Err(invalid(
            "BrtBeginCFRule",
            "bottom/percent flags are set on a non-filter rule",
        ));
    }
    Ok(())
}

fn effective_parameter(rule: &Rule) -> Result<u32> {
    if rule.rule_type != RuleType::CellIs {
        if rule.operator.is_some() {
            return Err(invalid(
                "BrtBeginCFRule",
                "operator is set on a non-cell-comparison rule",
            ));
        }
        return Ok(rule.parameter);
    }
    let parameter = rule.operator.map_or(rule.parameter, u32::from);
    if rule.parameter != 0 && rule.parameter != parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            "operator and exact parameter disagree",
        ));
    }
    Ok(parameter)
}

fn effective_formulas(rule: &Rule) -> Result<Vec<ParsedFormula>> {
    if !rule.formulas.is_empty() {
        if !rule.formula_extras.is_empty() && rule.formula_extras.len() != rule.formulas.len() {
            return Err(Error::InvalidFormula(
                "conditional-format ancillary stream count does not match formulas".to_string(),
            ));
        }
        return rule
            .formulas
            .iter()
            .enumerate()
            .map(|(index, rgce)| {
                if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(Error::InvalidFormula(format!(
                        "conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                        rgce.len()
                    )));
                }
                Ok(ParsedFormula {
                    rgce: rgce.clone(),
                    rgcb: rule.formula_extras.get(index).cloned().unwrap_or_default(),
                })
            })
            .collect();
    }
    rule.formula_texts
        .iter()
        .map(|formula| TextCompiler::compile(formula))
        .collect()
}

fn write_rule_visualization<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule.color_scale.as_ref().expect("validated color scale");
            validate_scale_thresholds(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE, &[])?;
            write_cfvo(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo(writer, midpoint, false)?;
            }
            write_cfvo(writer, &scale.max_cfvo, false)?;
            write_color(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color(writer, record, argb)?;
            }
            write_color(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(kind::END_COLOR_SCALE, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule.data_bar.as_ref().expect("validated data bar");
            if bar.min_length > bar.max_length || bar.max_length > 100 {
                return Err(invalid("BrtBeginDatabar", "invalid minimum/maximum length"));
            }
            validate_boundary_thresholds(&bar.min_cfvo, &bar.max_cfvo, "BrtBeginDatabar")?;
            writer.write_record(
                kind::BEGIN_DATABAR,
                &[bar.min_length, bar.max_length, u8::from(bar.show_value)],
            )?;
            write_cfvo(writer, &bar.min_cfvo, false)?;
            write_cfvo(writer, &bar.max_cfvo, false)?;
            write_color(writer, bar.color_record, bar.color)?;
            writer.write_record(kind::END_DATABAR, &[])?;
        },
        RuleType::IconSet => {
            let set = rule.icon_set.as_ref().expect("validated icon set");
            let expected = icon_count(set.icon_set_type)?;
            if set.cfvos.len() != expected {
                return Err(invalid(
                    "BrtBeginIconSet",
                    format!("expected {expected} thresholds, found {}", set.cfvos.len()),
                ));
            }
            if set.cfvos.iter().any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3)) {
                return Err(invalid(
                    "BrtBeginIconSet",
                    "min/max threshold is not allowed",
                ));
            }
            let mut flags = 0u16;
            if !set.show_value {
                flags |= 0x02;
            }
            if !set.reverse {
                flags |= 0x04;
            }
            let mut begin = Vec::with_capacity(6);
            begin.extend_from_slice(&u32::from(set.icon_set_type).to_le_bytes());
            begin.extend_from_slice(&flags.to_le_bytes());
            writer.write_record(kind::BEGIN_ICON_SET, &begin)?;
            for cfvo in &set.cfvos {
                write_cfvo(writer, cfvo, true)?;
            }
            writer.write_record(kind::END_ICON_SET, &[])?;
        },
        _ => {},
    }
    Ok(())
}

fn write_rule_visualization14<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    if rule.color_scale.is_some() || rule.data_bar.is_some() || rule.icon_set.is_some() {
        return Err(invalid(
            "BrtBeginCFRule14",
            "classic visualization is set on an Office 2013 rule",
        ));
    }
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule
                .color_scale14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 color scale"))?;
            if rule.data_bar14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_scale_thresholds14(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE14, &[])?;
            write_cfvo14(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo14(writer, midpoint, false)?;
            }
            write_cfvo14(writer, &scale.max_cfvo, false)?;
            write_color14(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color14(writer, record, argb)?;
            }
            write_color14(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(kind::END_COLOR_SCALE14, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule
                .data_bar14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 data bar"))?;
            if rule.color_scale14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            let priority = rule
                .extension14
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing extension metadata"))?
                .priority;
            validate_data_bar14(bar, priority)?;
            writer.write_record(kind::BEGIN_DATABAR14, &bar.serialize_header()?)?;
            write_cfvo14(writer, &bar.min_cfvo, false)?;
            write_cfvo14(writer, &bar.max_cfvo, false)?;
            for color in [
                bar.positive_color,
                bar.border_color,
                bar.negative_color,
                bar.negative_border_color,
                bar.axis_color,
            ]
            .into_iter()
            .flatten()
            {
                writer.write_record(kind::COLOR14, &color.serialize_extension14()?)?;
            }
            writer.write_record(kind::END_DATABAR14, &[])?;
        },
        RuleType::IconSet => {
            let set = rule
                .icon_set14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 icon set"))?;
            if rule.color_scale14.is_some() || rule.data_bar14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_icon_set14(set)?;
            writer.write_record(kind::BEGIN_ICON_SET14, &set.serialize_header()?)?;
            for cfvo in &set.cfvos {
                write_cfvo14(writer, cfvo, true)?;
            }
            if let Some(icons) = &set.custom_icons {
                for icon in icons {
                    writer.write_record(kind::CF_ICON, &icon.serialize()?)?;
                }
            }
            writer.write_record(kind::END_ICON_SET14, &[])?;
        },
        _ => {
            if rule.color_scale14.is_some()
                || rule.data_bar14.is_some()
                || rule.icon_set14.is_some()
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "non-visual rule contains a visualization",
                ));
            }
        },
    }
    Ok(())
}

fn validate_scale_thresholds14(scale: &Scale) -> Result<()> {
    if matches!(scale.min_cfvo.cfvo_type, 3 | 8 | 9)
        || matches!(scale.max_cfvo.cfvo_type, 2 | 8 | 9)
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "minimum/maximum threshold type is reversed",
        ));
    }
    if scale.mid_cfvo.is_some() != scale.mid_color_record.is_some()
        || scale.mid_cfvo.is_some() != scale.mid_color.is_some()
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "middle threshold and color must both be present or absent",
        ));
    }
    if scale
        .mid_cfvo
        .as_ref()
        .is_some_and(|cfvo| matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9))
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "middle threshold cannot be a boundary",
        ));
    }
    Ok(())
}

fn validate_data_bar14(bar: &Bar14, priority: i32) -> Result<()> {
    if matches!(bar.min_cfvo.cfvo_type, 3 | 9) || matches!(bar.max_cfvo.cfvo_type, 2 | 8) {
        return Err(invalid(
            "BrtBeginDatabar14",
            "minimum/maximum threshold type is reversed",
        ));
    }
    let valid_colors = bar.positive_color.is_some() == (priority != -1)
        && bar.border_color.is_some() == bar.border
        && bar.negative_color.is_some() == bar.custom_negative_fill
        && bar.negative_border_color.is_some() == (bar.custom_negative_border && bar.border)
        && bar.axis_color.is_some() == (bar.axis_position != AxisPosition14::None);
    if !valid_colors {
        return Err(invalid(
            "BrtBeginDatabar14",
            "color records do not match data-bar flags",
        ));
    }
    Ok(())
}

fn validate_icon_set14(set: &IconSet14) -> Result<()> {
    let expected = icon_count14(set.icon_set_type);
    if expected == 0 || set.cfvos.len() != expected {
        return Err(invalid(
            "BrtBeginIconSet14",
            format!("expected {expected} thresholds, found {}", set.cfvos.len()),
        ));
    }
    if set
        .cfvos
        .iter()
        .any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9) || !cfvo.save_greater_than_or_equal)
    {
        return Err(invalid(
            "BrtBeginIconSet14",
            "invalid threshold type or fSaveGTE flag",
        ));
    }
    if set
        .custom_icons
        .as_ref()
        .is_some_and(|icons| icons.len() != expected)
    {
        return Err(invalid(
            "BrtBeginIconSet14",
            "custom icon count does not match thresholds",
        ));
    }
    Ok(())
}

fn write_cfvo14<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
    let formula = effective_cfvo_formula(cfvo)?;
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    writer.write_record(
        kind::CFVO14,
        &cfvo.serialize_extension14_with(
            formula.as_ref(),
            numeric_value,
            icon_set || cfvo.save_greater_than_or_equal,
        )?,
    )?;
    Ok(())
}

fn write_color14<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR14, &record.serialize_extension14()?)?;
    Ok(())
}

fn validate_scale_thresholds(scale: &Scale) -> Result<()> {
    validate_boundary_thresholds(&scale.min_cfvo, &scale.max_cfvo, "BrtBeginColorScale")?;
    if scale.mid_cfvo.is_some() != scale.mid_color_record.is_some()
        || scale.mid_cfvo.is_some() != scale.mid_color.is_some()
    {
        return Err(invalid(
            "BrtBeginColorScale",
            "middle threshold and color must both be present or absent",
        ));
    }
    if scale
        .mid_cfvo
        .as_ref()
        .is_some_and(|cfvo| matches!(cfvo.cfvo_type, 2 | 3))
    {
        return Err(invalid(
            "BrtBeginColorScale",
            "middle threshold cannot be min/max",
        ));
    }
    Ok(())
}

fn validate_boundary_thresholds(minimum: &Value, maximum: &Value, record: &str) -> Result<()> {
    if minimum.cfvo_type == 3 || maximum.cfvo_type == 2 {
        return Err(invalid(
            record,
            "minimum/maximum threshold type is reversed",
        ));
    }
    Ok(())
}

fn icon_count(icon_set_type: u8) -> Result<usize> {
    match icon_set_type {
        0..=7 => Ok(3),
        8..=12 => Ok(4),
        13..=16 => Ok(5),
        value => Err(invalid("BrtBeginIconSet", format!("invalid set {value}"))),
    }
}

fn write_cfvo<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
    if !matches!(cfvo.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
        return Err(invalid(
            "BrtCFVO",
            format!("invalid type {}", cfvo.cfvo_type),
        ));
    }
    let formula = effective_cfvo_formula(cfvo)?;
    if matches!(cfvo.cfvo_type, 2 | 3) && formula.is_some() {
        return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
    }
    if cfvo.cfvo_type == 7 && formula.is_none() {
        return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
    }
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    if !numeric_value.is_finite()
        || (formula.is_none()
            && matches!(cfvo.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value))
    {
        return Err(invalid("BrtCFVO", "invalid numeric parameter"));
    }
    let mut payload = Vec::with_capacity(32);
    payload.extend_from_slice(&u32::from(cfvo.cfvo_type).to_le_bytes());
    payload.extend_from_slice(&numeric_value.to_le_bytes());
    payload
        .extend_from_slice(&u32::from(icon_set || cfvo.save_greater_than_or_equal).to_le_bytes());
    payload.extend_from_slice(&u32::from(cfvo.greater_than_or_equal).to_le_bytes());
    let formula_size = formula.as_ref().map_or(0, |formula| formula.rgce.len());
    payload.extend_from_slice(
        &u32::try_from(formula_size)
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
            .to_le_bytes(),
    );
    if let Some(formula) = formula {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    writer.write_record(kind::CFVO, &payload)?;
    Ok(())
}

fn effective_cfvo_formula(cfvo: &Value) -> Result<Option<ParsedFormula>> {
    if let Some(formula) = &cfvo.formula_binary {
        return Ok(Some(formula.clone()));
    }
    let Some(value) = cfvo.value.as_deref().filter(|value| !value.is_empty()) else {
        return Ok(None);
    };
    if matches!(cfvo.cfvo_type, 1 | 4 | 5) && value.parse::<f64>().is_ok() {
        return Ok(None);
    }
    if cfvo.cfvo_type == 7 || matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        return TextCompiler::compile(value).map(Some);
    }
    Ok(None)
}

fn write_color<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR, &record.to_bytes()?)?;
    Ok(())
}

fn write_nullable_wide_string(payload: &mut Vec<u8>, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().count();
    payload.extend_from_slice(
        &u32::try_from(units)
            .map_err(|_| Error::Encoding("conditional-format text is too long".to_string()))?
            .to_le_bytes(),
    );
    payload.reserve(units.saturating_mul(2));
    for unit in value.encode_utf16() {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::{Bar, Bar14, IconSet, RuleMetadata, Scale};
    use crate::raw::Records;

    fn compiled(text: &str) -> ParsedFormula {
        TextCompiler::compile(text).unwrap()
    }

    #[test]
    fn fixture_rule_header_matches_libreoffice_sample() {
        let formula = compiled("5");
        let mut rule = Rule::new(RuleType::CellIs, 1);
        rule.dxf_id = Some(0);
        rule.operator = Some(5);
        rule.parameter = 5;
        rule.formulas.push(formula.rgce);
        let payload = serialize_cf_rule(&rule).unwrap();
        assert_eq!(payload.len(), 57);
        assert_eq!(u32::from_le_bytes(payload[30..34].try_into().unwrap()), 3);
        assert_eq!(&payload[42..46], &u32::MAX.to_le_bytes());
        let parsed = Rule::parse(&payload).unwrap();
        assert_eq!(parsed.operator, Some(5));
        assert_eq!(parsed.formula_texts, ["5"]);
    }

    #[test]
    fn header_preserves_pivot_and_multiple_ranges() {
        let mut formatting = Formatting::new(vec!["A1:B10".into(), "D4".into()]);
        formatting.pivot_only = true;
        let payload = serialize_cond_formatting_header(&formatting).unwrap();
        assert_eq!(u32::from_le_bytes(payload[4..8].try_into().unwrap()), 1);
        assert_eq!(u32::from_le_bytes(payload[8..12].try_into().unwrap()), 2);
    }

    #[test]
    fn writes_color_scale_data_bar_and_icon_set_subrecords() {
        let mut formatting = Formatting::new(vec!["A1:A10".into()]);
        let mut scale = Rule::new(RuleType::ColorScale, 1);
        scale.color_scale = Some(Scale::new(
            Value::new(2, None),
            Value::new(3, None),
            0xffff0000,
            0xff00ff00,
        ));
        formatting.add_rule(scale);

        let mut bar = Rule::new(RuleType::DataBar, 2);
        bar.data_bar = Some(Bar::new(
            Value::new(2, None),
            Value::new(3, None),
            0xff4472c4,
        ));
        formatting.add_rule(bar);

        let mut icons = Rule::new(RuleType::IconSet, 3);
        icons.icon_set = Some(IconSet::new(
            0,
            vec![
                Value::new(1, Some("0".into())),
                Value::new(4, Some("33".into())),
                Value::new(4, Some("67".into())),
            ],
        ));
        formatting.add_rule(icons);

        let mut bytes = Vec::new();
        write_conditional_formattings(&mut Writer::new(&mut bytes), &[formatting]).unwrap();
        let records = Records::new(&bytes);
        let mut found = Vec::new();
        for record in records {
            found.push(record.unwrap().kind());
        }
        for typ in [
            kind::BEGIN_COLOR_SCALE,
            kind::BEGIN_DATABAR,
            kind::BEGIN_ICON_SET,
            kind::CFVO,
            kind::COLOR,
        ] {
            assert!(found.contains(&typ), "record 0x{typ:04x}");
        }
    }

    #[test]
    fn rejects_duplicate_priority_and_wrong_formula_slot_count() {
        let mut first = Formatting::new(vec!["A1".into()]);
        let mut rule = Rule::new(RuleType::CellIs, 1);
        rule.operator = Some(1);
        rule.formulas.push(compiled("1").rgce);
        first.add_rule(rule.clone());
        let mut second = Formatting::new(vec!["B1".into()]);
        second.add_rule(rule);
        assert!(
            write_conditional_formattings(&mut Writer::new(Vec::new()), &[first, second]).is_err()
        );
    }

    #[test]
    fn writes_office_2013_visualization_records() {
        let mut formatting = Formatting::new(vec!["A1:A10".into()]);
        formatting.record_kind = RecordKind::Extension14;
        let mut rule = Rule::new(RuleType::DataBar, 1);
        rule.extension14 = Some(RuleMetadata {
            priority: 1,
            unused: 7,
            guid: [0x24; 16],
            guid_present: true,
            linked_classic_priority: None,
        });
        rule.data_bar14 = Some(Bar14::new(
            Value::new(8, None),
            Value::new(9, None),
            Color::from_argb(0xff44_72c4),
        ));
        formatting.add_rule(rule);

        let mut bytes = Vec::new();
        write_conditional_formattings(&mut Writer::new(&mut bytes), &[formatting]).unwrap();
        let found = Records::new(&bytes)
            .map(|record| record.unwrap().kind())
            .collect::<Vec<_>>();
        assert_eq!(
            found,
            [
                kind::BEGIN_COND_FORMATTING14,
                kind::BEGIN_CF_RULE14,
                kind::BEGIN_DATABAR14,
                kind::CFVO14,
                kind::CFVO14,
                kind::COLOR14,
                kind::COLOR14,
                kind::END_DATABAR14,
                kind::END_CF_RULE14,
                kind::END_COND_FORMATTING14,
            ]
        );
    }

    #[test]
    fn rejects_extension_cross_record_violations() {
        let mut extension = Formatting::new_extension14(vec!["A1".into()]);
        let mut rule = Rule::new(RuleType::ColorScale, 1);
        rule.extension14 = Some(RuleMetadata {
            priority: 1,
            unused: 0,
            guid: [0; 16],
            guid_present: false,
            linked_classic_priority: None,
        });
        rule.color_scale14 = Some(Scale::new(
            Value::new(8, None),
            Value::new(9, None),
            0xffff_0000,
            0xff00_ff00,
        ));
        extension.add_rule(rule);
        assert!(write_conditional_formattings(&mut Writer::new(Vec::new()), &[extension]).is_err());

        let mut classic = Formatting::new(vec!["A1".into()]);
        let mut bar = Rule::new(RuleType::DataBar, 1);
        bar.classic_extension_guid = Some([1; 16]);
        bar.data_bar = Some(Bar::new(
            Value::new(2, None),
            Value::new(3, None),
            0xff44_72c4,
        ));
        classic.add_rule(bar);
        assert!(write_conditional_formattings(&mut Writer::new(Vec::new()), &[classic]).is_err());
    }
}
