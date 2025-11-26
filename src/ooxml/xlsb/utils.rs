//! Utility functions for XLSB parsing

use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};

/// Convert column number to Excel column name (A, B, ..., Z, AA, AB, etc.)
///
/// Input is 1-based (1=A, 2=B, 26=Z, 27=AA, etc.)
pub fn column_index_to_name(mut col: u32) -> String {
    if col == 0 {
        return String::new(); // Invalid input
    }

    let mut name = String::new();

    while col > 0 {
        col -= 1; // Make 0-based for calculation
        let ch = (b'A' + (col % 26) as u8) as char;
        name.insert(0, ch);
        col /= 26;
    }

    name
}

/// Convert Excel column name to column index (A=0, B=1, ..., Z=25, AA=26, etc.)
pub fn column_name_to_index(name: &str) -> Option<u32> {
    let name = name.to_ascii_uppercase();
    let mut result: u32 = 0;

    for ch in name.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        result = result * 26 + (ch as u32 - 'A' as u32) + 1;
    }

    Some(result - 1) // Make 0-based
}

/// Convert row and column to Excel cell reference (e.g., "A1", "B2")
pub fn cell_reference(row: u32, col: u32) -> String {
    format!("{}{}", column_index_to_name(col + 1), row + 1)
}

/// Parse Excel cell reference to row and column indices
///
/// Returns 0-based row and column indices
pub fn parse_cell_reference(ref_str: &str) -> XlsbResult<(u32, u32)> {
    let ref_str = ref_str.to_ascii_uppercase();
    let mut col_str = String::new();
    let mut row_str = String::new();
    let mut found_digit = false;

    for ch in ref_str.chars() {
        if ch.is_ascii_uppercase() {
            // Letters must come before digits
            if found_digit {
                return Err(XlsbError::InvalidCellReference(ref_str.to_string()));
            }
            col_str.push(ch);
        } else if ch.is_ascii_digit() {
            found_digit = true;
            row_str.push(ch);
        } else {
            return Err(XlsbError::InvalidCellReference(ref_str.to_string()));
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return Err(XlsbError::InvalidCellReference(ref_str.to_string()));
    }

    let col = column_name_to_index(&col_str)
        .ok_or_else(|| XlsbError::InvalidCellReference(ref_str.to_string()))?;
    let row: u32 = row_str
        .parse()
        .map_err(|_| XlsbError::InvalidCellReference(ref_str.to_string()))?;

    Ok((row - 1, col)) // Make 0-based
}

/// Convert serial date to datetime
#[allow(dead_code)]
pub fn excel_date_to_datetime(serial: f64, is_1904: bool) -> Option<chrono::NaiveDateTime> {
    use chrono::{Duration, NaiveDate};

    let base_date = if is_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)?
    };

    let days = serial.trunc() as i64;
    let seconds = ((serial.fract() * 86400.0).round() as i64) * 1_000_000; // microseconds

    let date = base_date + Duration::days(days);
    let time = Duration::microseconds(seconds);

    Some(date.and_time(chrono::NaiveTime::from_hms_opt(0, 0, 0)?) + time)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_index_to_name() {
        assert_eq!(column_index_to_name(1), "A");
        assert_eq!(column_index_to_name(2), "B");
        assert_eq!(column_index_to_name(26), "Z");
        assert_eq!(column_index_to_name(27), "AA");
        assert_eq!(column_index_to_name(703), "AAA");
    }

    #[test]
    fn test_column_name_to_index() {
        assert_eq!(column_name_to_index("A"), Some(0));
        assert_eq!(column_name_to_index("B"), Some(1));
        assert_eq!(column_name_to_index("Z"), Some(25));
        assert_eq!(column_name_to_index("AA"), Some(26));
        assert_eq!(column_name_to_index("AAA"), Some(702));
        assert_eq!(column_name_to_index("a"), Some(0)); // case insensitive
        assert_eq!(column_name_to_index("1A"), None); // invalid
    }

    #[test]
    fn test_cell_reference() {
        assert_eq!(cell_reference(0, 0), "A1");
        assert_eq!(cell_reference(1, 1), "B2");
        assert_eq!(cell_reference(0, 26), "AA1");
    }

    #[test]
    fn test_parse_cell_reference() {
        assert!(matches!(parse_cell_reference("A1"), Ok((0, 0))));
        assert!(matches!(parse_cell_reference("B2"), Ok((1, 1))));
        assert!(matches!(parse_cell_reference("AA1"), Ok((0, 26))));
        assert!(matches!(parse_cell_reference("a1"), Ok((0, 0)))); // case insensitive
        assert!(parse_cell_reference("1A").is_err()); // invalid
        assert!(parse_cell_reference("A").is_err()); // no row
        assert!(parse_cell_reference("1").is_err()); // no column
    }
}
