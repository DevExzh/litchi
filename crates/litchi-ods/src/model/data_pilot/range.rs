//! Spreadsheet-address semantics used by data-pilot declarations.

use litchi_core::Result;

use super::{invalid_message, validation::validate_string};

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ParsedRange {
    pub sheet: String,
    pub start_column: usize,
    pub start_row: usize,
    pub end_column: usize,
    pub end_row: usize,
}

pub(crate) fn parse_data_pilot_range(value: &str) -> Result<ParsedRange> {
    validate_string("data-pilot cell range", value, false)?;
    let mut quoted = false;
    let mut separator = None;
    let mut characters = value.char_indices().peekable();
    while let Some((index, character)) = characters.next() {
        if character == '\'' {
            if quoted && characters.peek().is_some_and(|(_, next)| *next == '\'') {
                characters.next();
                continue;
            }
            quoted = !quoted;
        } else if character == ':' && !quoted && separator.replace(index).is_some() {
            return Err(invalid_message("invalid data-pilot cell range"));
        }
    }
    if quoted {
        return Err(invalid_message(
            "unterminated quoted sheet name in data-pilot range",
        ));
    }
    let (first, second) =
        separator.map_or((value, None), |at| (&value[..at], Some(&value[at + 1..])));
    let (sheet, start_column, start_row) = parse_range_endpoint(first, None)?;
    let (end_sheet, end_column, end_row) = if let Some(second) = second {
        parse_range_endpoint(second, Some(&sheet))?
    } else {
        (sheet.clone(), start_column, start_row)
    };
    if end_sheet != sheet || end_column < start_column || end_row < start_row {
        return Err(invalid_message(
            "data-pilot cell range is reversed or crosses sheets",
        ));
    }
    Ok(ParsedRange {
        sheet,
        start_column,
        start_row,
        end_column,
        end_row,
    })
}

fn parse_range_endpoint(
    value: &str,
    inherited_sheet: Option<&str>,
) -> Result<(String, usize, usize)> {
    let value = value.trim();
    let mut quoted = false;
    let mut dot = None;
    let mut characters = value.char_indices().peekable();
    while let Some((index, character)) = characters.next() {
        if character == '\'' {
            if quoted && characters.peek().is_some_and(|(_, next)| *next == '\'') {
                characters.next();
                continue;
            }
            quoted = !quoted;
        } else if character == '.' && !quoted {
            dot = Some(index);
        }
    }
    let (sheet, coordinate) =
        dot.map_or((None, value), |at| (Some(&value[..at]), &value[at + 1..]));
    let sheet = match sheet {
        Some("") => inherited_sheet.unwrap_or_default().to_string(),
        Some(value) => normalize_sheet_name(value)?,
        None => inherited_sheet.unwrap_or_default().to_string(),
    };
    let coordinate = coordinate.replace('$', "");
    let split = coordinate
        .find(|character: char| character.is_ascii_digit())
        .ok_or_else(|| invalid_message("data-pilot cell address lacks a row"))?;
    let (column, row) = coordinate.split_at(split);
    if column.is_empty()
        || !column.chars().all(|ch| ch.is_ascii_uppercase())
        || row.is_empty()
        || !row.chars().all(|ch| ch.is_ascii_digit())
    {
        return Err(invalid_message("invalid data-pilot cell address"));
    }
    let mut column_index = 0usize;
    for ch in column.bytes() {
        column_index = column_index
            .checked_mul(26)
            .and_then(|value| value.checked_add(usize::from(ch - b'A') + 1))
            .ok_or_else(|| invalid_message("data-pilot column index overflow"))?;
    }
    let row_number = row
        .parse::<usize>()
        .map_err(|_| invalid_message("invalid data-pilot row"))?;
    if row_number == 0 {
        return Err(invalid_message("data-pilot rows are one-based"));
    }
    Ok((sheet, column_index - 1, row_number - 1))
}

fn normalize_sheet_name(value: &str) -> Result<String> {
    let value = value.trim().trim_start_matches('$');
    if value.starts_with('\'') {
        if !value.ends_with('\'') || value.len() < 2 {
            return Err(invalid_message("invalid quoted sheet name"));
        }
        Ok(value[1..value.len() - 1].replace("''", "'"))
    } else {
        if value.contains('\'') {
            return Err(invalid_message("invalid sheet name"));
        }
        Ok(value.to_string())
    }
}

pub(super) fn ranges_overlap(left: &ParsedRange, right: &ParsedRange) -> bool {
    left.sheet == right.sheet
        && left.start_column <= right.end_column
        && right.start_column <= left.end_column
        && left.start_row <= right.end_row
        && right.start_row <= left.end_row
}
