use std::collections::{HashMap, HashSet};

use litchi_core::sheet::{CellValue, Result};

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

#[derive(Debug, Clone)]
pub struct Member {
    pub row: u32,
    pub column: u32,
    pub index: u32,
    pub reference: Option<String>,
    pub formula: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellRange {
    first_row: u32,
    first_column: u32,
    last_row: u32,
    last_column: u32,
}

impl CellRange {
    fn parse(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        if last.contains(':') {
            return Err(format!("invalid shared formula range '{value}'").into());
        }
        let (first_row, first_column) = parse_cell_position(first)
            .ok_or_else(|| format!("invalid shared formula range '{value}'"))?;
        let (last_row, last_column) = parse_cell_position(last)
            .ok_or_else(|| format!("invalid shared formula range '{value}'"))?;
        if first_row > last_row || first_column > last_column {
            return Err(format!("reversed shared formula range '{value}'").into());
        }
        Ok(Self {
            first_row,
            first_column,
            last_row,
            last_column,
        })
    }

    fn contains(self, row: u32, column: u32) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_column..=self.last_column).contains(&column)
    }
}

#[derive(Debug)]
struct Master {
    row: u32,
    column: u32,
    range: CellRange,
    formula: String,
}

pub fn resolve(
    cells: &mut HashMap<u32, HashMap<u32, CellValue>>,
    shared_cells: &[Member],
) -> Result<()> {
    if shared_cells.is_empty() {
        return Ok(());
    }

    let mut occupied = HashSet::with_capacity(shared_cells.len());
    for cell in shared_cells {
        if !occupied.insert((cell.row, cell.column)) {
            return Err(format!(
                "duplicate or ambiguous shared formula membership at row {}, column {}",
                cell.row, cell.column
            )
            .into());
        }
    }

    let mut masters = HashMap::<u32, Master>::new();
    for cell in shared_cells {
        if cell.reference.is_none() && cell.formula.is_empty() {
            continue;
        }
        let reference = cell.reference.as_deref().ok_or_else(|| {
            format!(
                "shared formula master at row {}, column {} is missing ref",
                cell.row, cell.column
            )
        })?;
        if cell.formula.is_empty() {
            return Err(format!(
                "shared formula master at row {}, column {} has no formula",
                cell.row, cell.column
            )
            .into());
        }
        let range = CellRange::parse(reference)?;
        if (cell.row, cell.column) != (range.first_row, range.first_column) {
            return Err(format!(
                "shared formula master at row {}, column {} is not the first cell of ref '{reference}'",
                cell.row, cell.column
            )
            .into());
        }
        let master = Master {
            row: cell.row,
            column: cell.column,
            range,
            formula: cell.formula.clone(),
        };
        if masters.insert(cell.index, master).is_some() {
            return Err(format!("duplicate shared formula master for si={}", cell.index).into());
        }
    }

    for cell in shared_cells {
        if !masters.contains_key(&cell.index) {
            return Err(format!(
                "shared formula at row {}, column {} has no master for si={}",
                cell.row, cell.column, cell.index
            )
            .into());
        }
    }

    for cell in shared_cells {
        let master = &masters[&cell.index];
        if !master.range.contains(cell.row, cell.column) {
            return Err(format!(
                "shared formula at row {}, column {} lies outside master range for si={}",
                cell.row, cell.column, cell.index
            )
            .into());
        }
        let is_master = (cell.row, cell.column) == (master.row, master.column);
        if !is_master && (!cell.formula.is_empty() || cell.reference.is_some()) {
            return Err(format!(
                "shared formula follower at row {}, column {} contains master data",
                cell.row, cell.column
            )
            .into());
        }
        if is_master {
            continue;
        }
        let value = cells
            .get_mut(&cell.row)
            .and_then(|row| row.get_mut(&cell.column))
            .ok_or_else(|| {
                format!(
                    "missing materialized shared formula cell at row {}, column {}",
                    cell.row, cell.column
                )
            })?;
        let CellValue::Formula { formula, .. } = value else {
            return Err(format!(
                "shared formula metadata points to a non-formula cell at row {}, column {}",
                cell.row, cell.column
            )
            .into());
        };
        *formula = translate_formula(
            &master.formula,
            master.row,
            master.column,
            cell.row,
            cell.column,
        );
    }

    Ok(())
}

fn parse_cell_position(value: &str) -> Option<(u32, u32)> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start {
        return None;
    }
    let column = decode_column(&value[column_start..index])?;
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return None;
    }
    let row = value[row_start..index].parse::<u32>().ok()?;
    (row != 0 && row <= MAX_ROW && column <= MAX_COLUMN).then_some((row, column))
}

fn decode_column(value: &str) -> Option<u32> {
    let mut column = 0u32;
    for byte in value.bytes() {
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    (column != 0).then_some(column)
}

fn encode_column(mut column: u32, output: &mut String) {
    let start = output.len();
    while column != 0 {
        column -= 1;
        output.push(char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    let suffix = output.split_off(start);
    output.extend(suffix.chars().rev());
}

#[derive(Debug, Clone, Copy)]
struct Axis {
    value: u32,
    absolute: bool,
}

#[derive(Debug, Clone, Copy)]
struct CellReference {
    row: Axis,
    column: Axis,
}

#[derive(Debug, Clone, Copy)]
enum Reference {
    Cell(CellReference),
    Area(CellReference, CellReference),
    Columns(Axis, Axis),
    Rows(Axis, Axis),
}

struct ParsedReference<'a> {
    prefix: &'a str,
    reference: Reference,
    end: usize,
}

pub(crate) fn translate_formula(
    formula: &str,
    origin_row: u32,
    origin_column: u32,
    target_row: u32,
    target_column: u32,
) -> String {
    let row_delta = i64::from(target_row) - i64::from(origin_row);
    let column_delta = i64::from(target_column) - i64::from(origin_column);
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;

    while index < bytes.len() {
        if bytes[index] == b'"' {
            let end = quoted_string_end(bytes, index);
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }
        if let Some(parsed) = parse_reference(formula, index) {
            output.push_str(parsed.prefix);
            render_reference(parsed.reference, row_delta, column_delta, &mut output);
            index = parsed.end;
            continue;
        }
        if bytes[index] == b'[' {
            let end = bracket_end(bytes, index);
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }
        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn quoted_string_end(bytes: &[u8], start: usize) -> usize {
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if bytes.get(index + 1) == Some(&b'"') {
                index += 2;
            } else {
                return index + 1;
            }
        } else {
            index += 1;
        }
    }
    bytes.len()
}

fn bracket_end(bytes: &[u8], start: usize) -> usize {
    let mut depth = 0usize;
    let mut index = start;
    while index < bytes.len() {
        match bytes[index] {
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return index + 1;
                }
            },
            _ => {},
        }
        index += 1;
    }
    bytes.len()
}

fn parse_reference(formula: &str, start: usize) -> Option<ParsedReference<'_>> {
    let bytes = formula.as_bytes();
    if start != 0 && is_identifier_byte(bytes[start - 1]) {
        return None;
    }

    let reference_start = parse_prefix(formula, start).unwrap_or(start);
    let prefix = &formula[start..reference_start];
    let (reference, end) = parse_reference_body(formula, reference_start)?;
    if bytes
        .get(end)
        .is_some_and(|byte| is_identifier_byte(*byte) || *byte == b'(')
    {
        return None;
    }
    Some(ParsedReference {
        prefix,
        reference,
        end,
    })
}

fn parse_prefix(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut index = start + 1;
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index += 2;
                    continue;
                }
                return (bytes.get(index + 1) == Some(&b'!')).then_some(index + 2);
            }
            index += 1;
        }
        return None;
    }

    let mut index = start;
    let mut bracket_depth = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'[' => bracket_depth += 1,
            b']' if bracket_depth != 0 => bracket_depth -= 1,
            b'!' if bracket_depth == 0 && index != start => return Some(index + 1),
            byte if bracket_depth == 0 && !is_prefix_byte(byte) => return None,
            _ => {},
        }
        index += 1;
    }
    None
}

fn is_prefix_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric()
        || matches!(byte, b'_' | b'.' | b':' | b'\\' | b'/' | b'[' | b']' | b'-')
}

fn is_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.')
}

fn parse_reference_body(formula: &str, start: usize) -> Option<(Reference, usize)> {
    if let Some((first, first_end)) = parse_cell_reference(formula, start) {
        if formula.as_bytes().get(first_end) == Some(&b':') {
            let (last, end) = parse_cell_reference(formula, first_end + 1)?;
            return Some((Reference::Area(first, last), end));
        }
        return Some((Reference::Cell(first), first_end));
    }
    if let Some((first, first_end)) = parse_column_reference(formula, start)
        && formula.as_bytes().get(first_end) == Some(&b':')
    {
        let (last, end) = parse_column_reference(formula, first_end + 1)?;
        return Some((Reference::Columns(first, last), end));
    }
    if let Some((first, first_end)) = parse_row_reference(formula, start)
        && formula.as_bytes().get(first_end) == Some(&b':')
    {
        let (last, end) = parse_row_reference(formula, first_end + 1)?;
        return Some((Reference::Rows(first, last), end));
    }
    None
}

fn parse_cell_reference(formula: &str, start: usize) -> Option<(CellReference, usize)> {
    let (column, index) = parse_column_reference(formula, start)?;
    let (row, end) = parse_row_reference(formula, index)?;
    Some((CellReference { row, column }, end))
}

fn parse_column_reference(formula: &str, start: usize) -> Option<(Axis, usize)> {
    let bytes = formula.as_bytes();
    let absolute = bytes.get(start) == Some(&b'$');
    let mut index = start + usize::from(absolute);
    let letters_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) && index - letters_start < 3 {
        index += 1;
    }
    if index == letters_start || bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let value = decode_column(&formula[letters_start..index])?;
    (value <= MAX_COLUMN).then_some((Axis { value, absolute }, index))
}

fn parse_row_reference(formula: &str, start: usize) -> Option<(Axis, usize)> {
    let bytes = formula.as_bytes();
    let absolute = bytes.get(start) == Some(&b'$');
    let mut index = start + usize::from(absolute);
    let digits_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == digits_start {
        return None;
    }
    let value = formula[digits_start..index].parse::<u32>().ok()?;
    (value != 0 && value <= MAX_ROW).then_some((Axis { value, absolute }, index))
}

fn shifted(axis: Axis, delta: i64, maximum: u32) -> Option<u32> {
    if axis.absolute {
        return Some(axis.value);
    }
    let value = i64::from(axis.value).checked_add(delta)?;
    (value >= 1 && value <= i64::from(maximum))
        .then(|| u32::try_from(value).ok())
        .flatten()
}

fn render_reference(reference: Reference, row_delta: i64, column_delta: i64, output: &mut String) {
    match reference {
        Reference::Cell(cell) => render_cell(cell, row_delta, column_delta, output),
        Reference::Area(first, last) => {
            render_cell(first, row_delta, column_delta, output);
            output.push(':');
            render_cell(last, row_delta, column_delta, output);
        },
        Reference::Columns(first, last) => {
            render_column(first, column_delta, output);
            output.push(':');
            render_column(last, column_delta, output);
        },
        Reference::Rows(first, last) => {
            render_row(first, row_delta, output);
            output.push(':');
            render_row(last, row_delta, output);
        },
    }
}

fn render_cell(cell: CellReference, row_delta: i64, column_delta: i64, output: &mut String) {
    let Some(column) = shifted(cell.column, column_delta, MAX_COLUMN) else {
        output.push_str("#REF!");
        return;
    };
    let Some(row) = shifted(cell.row, row_delta, MAX_ROW) else {
        output.push_str("#REF!");
        return;
    };
    if cell.column.absolute {
        output.push('$');
    }
    encode_column(column, output);
    if cell.row.absolute {
        output.push('$');
    }
    output.push_str(&row.to_string());
}

fn render_column(column: Axis, delta: i64, output: &mut String) {
    let Some(value) = shifted(column, delta, MAX_COLUMN) else {
        output.push_str("#REF!");
        return;
    };
    if column.absolute {
        output.push('$');
    }
    encode_column(value, output);
}

fn render_row(row: Axis, delta: i64, output: &mut String) {
    let Some(value) = shifted(row, delta, MAX_ROW) else {
        output.push_str("#REF!");
        return;
    };
    if row.absolute {
        output.push('$');
    }
    output.push_str(&value.to_string());
}
