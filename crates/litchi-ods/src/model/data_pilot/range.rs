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

const UNBOUNDED: usize = usize::MAX;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Axis {
    Cell,
    Column,
    Row,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Endpoint {
    sheet: String,
    column: usize,
    row: usize,
    axis: Axis,
}

pub(crate) fn parse_data_pilot_range(value: &str) -> Result<ParsedRange> {
    parse_range(value, false)
}

/// Parse the stricter `table:database-range` address grammar.  Data-pilot
/// declarations historically accept the shorthand `Sheet.A1:B2`; the
/// database-range vocabulary uses the schema's qualified cell-range form,
/// where each endpoint carries the `.` separator and unquoted sheet names
/// cannot contain whitespace.  The typed database-range model represents one
/// sheet, so a syntactically qualified cross-sheet pair is refused during
/// semantic validation rather than being collapsed into one sheet.
pub(crate) fn parse_database_range_address(value: &str) -> Result<ParsedRange> {
    parse_range(value, true)
}

fn parse_range(value: &str, require_endpoint_separator: bool) -> Result<ParsedRange> {
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
    let start = parse_range_endpoint(first, None, require_endpoint_separator)?;
    let end = if let Some(second) = second {
        parse_range_endpoint(second, Some(&start.sheet), require_endpoint_separator)?
    } else {
        if start.axis != Axis::Cell {
            return Err(invalid_message(
                "whole-row and whole-column ranges require two endpoints",
            ));
        }
        start.clone()
    };
    if start.sheet != end.sheet || start.axis != end.axis {
        return Err(invalid_message(
            "data-pilot cell range crosses sheets or mixes endpoint axes",
        ));
    }
    let (start_column, start_row, end_column, end_row) = match start.axis {
        Axis::Cell => {
            if end.column < start.column || end.row < start.row {
                return Err(invalid_message(
                    "data-pilot cell range is reversed or crosses sheets",
                ));
            }
            (start.column, start.row, end.column, end.row)
        },
        Axis::Column => {
            if end.column < start.column {
                return Err(invalid_message("data-pilot whole-column range is reversed"));
            }
            (start.column, 0, end.column, UNBOUNDED)
        },
        Axis::Row => {
            if end.row < start.row {
                return Err(invalid_message("data-pilot whole-row range is reversed"));
            }
            (0, start.row, UNBOUNDED, end.row)
        },
    };
    Ok(ParsedRange {
        sheet: start.sheet,
        start_column,
        start_row,
        end_column,
        end_row,
    })
}

fn parse_range_endpoint(
    value: &str,
    inherited_sheet: Option<&str>,
    require_separator: bool,
) -> Result<Endpoint> {
    if value.is_empty() {
        return Err(invalid_message("data-pilot range endpoint is empty"));
    }
    if require_separator && value != value.trim() {
        return Err(invalid_message(
            "data-pilot range endpoint contains whitespace",
        ));
    }
    let value = if require_separator {
        value
    } else {
        value.trim()
    };
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
            if dot.replace(index).is_some() {
                return Err(invalid_message("invalid data-pilot range endpoint"));
            }
        }
    }
    let (sheet, coordinate) = match dot {
        Some(dot) => (Some(&value[..dot]), &value[dot + 1..]),
        None if require_separator => {
            return Err(invalid_message(
                "data-pilot range endpoint requires a sheet separator",
            ));
        },
        None => (None, value),
    };
    let sheet = match sheet {
        Some("") => inherited_sheet.unwrap_or_default().to_string(),
        Some(value) => normalize_sheet_name(value, require_separator)?,
        None => inherited_sheet.unwrap_or_default().to_string(),
    };
    let coordinate = coordinate.as_bytes();
    let mut index = 0usize;
    if coordinate.get(index) == Some(&b'$') {
        index += 1;
    }
    let column_start = index;
    while coordinate.get(index).is_some_and(u8::is_ascii_uppercase) {
        index += 1;
    }
    let column_end = index;
    if column_end > column_start {
        let mut column_index = 0usize;
        for ch in &coordinate[column_start..column_end] {
            column_index = column_index
                .checked_mul(26)
                .and_then(|value| value.checked_add(usize::from(*ch - b'A') + 1))
                .ok_or_else(|| invalid_message("data-pilot column index overflow"))?;
        }
        let row_marker = coordinate.get(index) == Some(&b'$');
        if row_marker {
            index += 1;
        }
        let row_start = index;
        while coordinate.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == row_start {
            if row_marker || index != coordinate.len() {
                return Err(invalid_message("invalid data-pilot whole-column address"));
            }
            return Ok(Endpoint {
                sheet,
                column: column_index - 1,
                row: 0,
                axis: Axis::Column,
            });
        }
        if index != coordinate.len() {
            return Err(invalid_message("invalid data-pilot cell address"));
        }
        let row_number = parse_row(&coordinate[row_start..index])?;
        return Ok(Endpoint {
            sheet,
            column: column_index - 1,
            row: row_number,
            axis: Axis::Cell,
        });
    }
    let row_start = index;
    while coordinate.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != coordinate.len() {
        return Err(invalid_message("invalid data-pilot cell address"));
    }
    Ok(Endpoint {
        sheet,
        column: 0,
        row: parse_row(&coordinate[row_start..index])?,
        axis: Axis::Row,
    })
}

fn parse_row(value: &[u8]) -> Result<usize> {
    let row_number = std::str::from_utf8(value)
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
        .ok_or_else(|| invalid_message("invalid data-pilot row"))?;
    if row_number == 0 {
        return Err(invalid_message("data-pilot rows are one-based"));
    }
    row_number
        .checked_sub(1)
        .ok_or_else(|| invalid_message("data-pilot row index underflow"))
}

fn normalize_sheet_name(value: &str, strict: bool) -> Result<String> {
    let value = if strict { value } else { value.trim() };
    let value = if value.len() > 1 {
        value.strip_prefix('$').unwrap_or(value)
    } else {
        value
    };
    if value.starts_with('\'') {
        if !value.ends_with('\'') || value.len() < 2 {
            return Err(invalid_message("invalid quoted sheet name"));
        }
        let value = value[1..value.len() - 1].replace("''", "'");
        if strict && value.is_empty() {
            return Err(invalid_message("invalid quoted sheet name"));
        }
        Ok(value)
    } else {
        if value.contains('\'') || (strict && value.chars().any(char::is_whitespace)) {
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

#[cfg(test)]
mod tests {
    use super::{UNBOUNDED, parse_data_pilot_range, parse_database_range_address};

    #[test]
    fn accepts_cell_whole_column_and_whole_row_forms() {
        let cell = parse_database_range_address("Sheet.$A$2:Sheet.$C$4").unwrap();
        assert_eq!(cell.sheet, "Sheet");
        assert_eq!((cell.start_column, cell.start_row), (0, 1));
        assert_eq!((cell.end_column, cell.end_row), (2, 3));

        let columns = parse_database_range_address("'My Sheet'.A:'My Sheet'.C").unwrap();
        assert_eq!(columns.sheet, "My Sheet");
        assert_eq!((columns.start_column, columns.end_column), (0, 2));
        assert_eq!((columns.start_row, columns.end_row), (0, UNBOUNDED));

        let rows = parse_database_range_address("Sheet.$1:Sheet.$3").unwrap();
        assert_eq!(rows.sheet, "Sheet");
        assert_eq!((rows.start_column, rows.end_column), (0, UNBOUNDED));
        assert_eq!((rows.start_row, rows.end_row), (0, 2));

        let dollar_sheet = parse_database_range_address("$.A1:$.A1").unwrap();
        assert_eq!(dollar_sheet.sheet, "$");
    }

    #[test]
    fn rejects_mixed_reversed_and_incomplete_axis_forms() {
        for value in [
            "A1",
            "Sheet A.A1:Sheet A.B2",
            "''.A1",
            "Sheet.A1:B2",
            "Sheet.A1:Sheet.A",
            "Sheet.A:Sheet.1",
            "Sheet.C:Sheet.A",
            "Sheet.3:Sheet.1",
            "Sheet.A$",
            "Sheet.A:",
            "Sheet.1:",
        ] {
            assert!(
                parse_database_range_address(value).is_err(),
                "{value} should be rejected"
            );
        }
    }

    #[test]
    fn accepts_empty_sheet_separator_but_rejects_cross_sheet_ranges() {
        assert!(parse_database_range_address(".A1:.B2").is_ok());
        assert!(parse_database_range_address("Sheet.A1:Other.B2").is_err());
    }

    #[test]
    fn data_pilot_keeps_its_legacy_endpoint_shorthand() {
        assert!(parse_data_pilot_range("Sheet.A1:B2").is_ok());
    }
}
