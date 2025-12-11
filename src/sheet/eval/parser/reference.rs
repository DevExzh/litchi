//! Parsing of cell and range references in A1 notation.

use super::ast::RangeRef;

/// Parse a single-cell reference, optionally qualified with a sheet name.
///
/// Examples:
/// - `A1`
/// - `$B$2`
/// - `Sheet2!C3`
/// - `'My Sheet'!$D$10`
pub fn parse_single_cell_reference(current_sheet: &str, input: &str) -> Option<(String, u32, u32)> {
    let s = input.trim();
    if s.is_empty() {
        return None;
    }

    // Split sheet and cell parts on the last '!'. This is sufficient for
    // typical references like Sheet1!A1 or 'Sheet 1'!A1.
    let (sheet_part, cell_part) = if let Some(excl_pos) = s.rfind('!') {
        let (sheet_s, cell_s) = s.split_at(excl_pos);
        let cell_s = &cell_s[1..]; // skip '!'
        let sheet_name = unescape_sheet_name(sheet_s)?;
        (Some(sheet_name), cell_s)
    } else {
        (None, s)
    };

    let sheet_name = sheet_part.unwrap_or_else(|| current_sheet.to_string());
    let (row, col) = parse_a1_ref(cell_part)?;

    Some((sheet_name, row, col))
}

/// Parse a cell range reference, optionally qualified with a sheet name.
///
/// Examples:
/// - `A1:B3`
/// - `Sheet2!A1:B3`
/// - `'My Sheet'!$A$1:$C$10`
pub fn parse_range_reference(current_sheet: &str, input: &str) -> Option<RangeRef> {
    let s = input.trim();
    if s.is_empty() {
        return None;
    }

    let (sheet_part, cells_part) = if let Some(excl_pos) = s.rfind('!') {
        let (sheet_s, cell_s) = s.split_at(excl_pos);
        let cell_s = &cell_s[1..]; // skip '!'
        let sheet_name = unescape_sheet_name(sheet_s)?;
        (Some(sheet_name), cell_s)
    } else {
        (None, s)
    };

    let sheet = sheet_part.unwrap_or_else(|| current_sheet.to_string());

    let mut parts = cells_part.split(':');
    let start_str = parts.next()?.trim();
    let end_str = parts.next()?.trim();
    if parts.next().is_some() {
        return None;
    }

    let (start_row, start_col) = parse_a1_ref(start_str)?;
    let (end_row, end_col) = parse_a1_ref(end_str)?;

    Some(RangeRef {
        sheet,
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

fn unescape_sheet_name(sheet: &str) -> Option<String> {
    let trimmed = sheet.trim();
    let bytes = trimmed.as_bytes();
    if bytes.len() >= 2 && bytes.first() == Some(&b'\'') && bytes.last() == Some(&b'\'') {
        // Strip outer quotes and unescape doubled quotes.
        let inner = &trimmed[1..trimmed.len() - 1];
        let mut out = String::new();
        let mut chars = inner.chars().peekable();
        while let Some(ch) = chars.next() {
            if ch == '\'' && chars.peek() == Some(&'\'') {
                out.push('\'');
                chars.next();
            } else {
                out.push(ch);
            }
        }
        Some(out)
    } else if !trimmed.is_empty() {
        Some(trimmed.to_string())
    } else {
        None
    }
}

fn parse_a1_ref(cell: &str) -> Option<(u32, u32)> {
    let mut chars = cell.chars().peekable();

    // Optional leading '$' for absolute column – ignored here.
    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    let mut col_letters = String::new();
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_alphabetic() {
            col_letters.push(ch.to_ascii_uppercase());
            chars.next();
        } else {
            break;
        }
    }

    if col_letters.is_empty() {
        return None;
    }

    // Optional '$' before row – ignored here.
    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    let mut row_digits = String::new();
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_digit() {
            row_digits.push(ch);
            chars.next();
        } else {
            break;
        }
    }

    if row_digits.is_empty() {
        return None;
    }

    // No trailing characters allowed.
    if chars.peek().is_some() {
        return None;
    }

    let col = column_letters_to_index(&col_letters)?;
    let row = row_digits.parse::<u32>().ok()?;

    Some((row, col))
}

fn column_letters_to_index(col: &str) -> Option<u32> {
    let mut result: u32 = 0;
    for ch in col.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        let value = (ch as u8).wrapping_sub(b'A') as u32 + 1;
        result = result.checked_mul(26)?;
        result = result.checked_add(value)?;
    }
    Some(result)
}
