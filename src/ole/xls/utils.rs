//! Utility functions for XLS parsing

use crate::ole::binary;
use crate::ole::xls::error::{XlsError, XlsResult};
use crate::ole::xls::records::{XlsEncoding, FormulaValue};

/// Parse a short string (used in sheet names, etc.)
pub fn parse_short_string(data: &[u8], encoding: &XlsEncoding) -> XlsResult<String> {
    if data.is_empty() {
        return Ok(String::new());
    }

    let len = data[0] as usize;
    if data.len() < 1 + len {
        return Err(XlsError::InvalidLength {
            expected: 1 + len,
            found: data.len(),
        });
    }

    let string_data = &data[1..1 + len];
    encoding.decode(string_data)
}

/// Parse a string record with length prefix
pub fn parse_string_record(data: &[u8], encoding: &XlsEncoding) -> XlsResult<String> {
    if data.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }

    let len = binary::read_u16_le_at(data, 0)? as usize;
    let flags = data[2];

    let high_byte = (flags & 0x01) != 0;
    let offset = if matches!(encoding, XlsEncoding::Utf16Le) { 3 } else { 3 };

    if data.len() < offset + len {
        return Err(XlsError::InvalidLength {
            expected: offset + len,
            found: data.len(),
        });
    }

    let string_data = &data[offset..offset + len];

    // For BIFF8, handle UTF-16
    if matches!(encoding, XlsEncoding::Utf16Le) && high_byte {
        if len % 2 != 0 {
            return Err(XlsError::Encoding("Invalid UTF-16 string length".to_string()));
        }
        let utf16_data: Vec<u16> = string_data.chunks(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        Ok(String::from_utf16(&utf16_data)
            .map_err(|e| XlsError::Encoding(format!("UTF-16 decoding error: {}", e)))?)
    } else {
        // Assume Latin-1 or compatible encoding
        Ok(String::from_utf8(string_data.to_vec())
            .unwrap_or_else(|_| String::from_utf8_lossy(string_data).into_owned()))
    }
}

/// Parse a Unicode string from SST or other records
pub fn parse_unicode_string(data: &[u8], encoding: &XlsEncoding) -> XlsResult<(String, usize)> {
    if data.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }

    let cch = binary::read_u16_le_at(data, 0)? as usize;
    let flags = data[2];

    let mut offset = 3;
    let high_byte = (flags & 0x01) != 0;

    // Handle rich text formatting (skip for now)
    if (flags & 0x08) != 0 {
        if data.len() < offset + 2 {
            return Err(XlsError::InvalidLength {
                expected: offset + 2,
                found: data.len(),
            });
        }
        let c_run = binary::read_u16_le_at(data, offset)?;
        offset += 2 + 4 * c_run as usize; // Skip formatting runs
    }

    // Handle extended text (skip for now)
    if (flags & 0x04) != 0 {
        if data.len() < offset + 4 {
            return Err(XlsError::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            });
        }
        let cb_ext_rst = binary::read_u32_le_at(data, offset)?;
        offset += 4 + cb_ext_rst as usize; // Skip extended data
    }

    if data.len() < offset + cch {
        return Err(XlsError::InvalidLength {
            expected: offset + cch,
            found: data.len(),
        });
    }

    let string_data = &data[offset..offset + cch];
    let consumed = offset + cch;

    let string = if matches!(encoding, XlsEncoding::Utf16Le) && high_byte {
        // UTF-16 LE
        if cch % 2 != 0 {
            return Err(XlsError::Encoding("Invalid UTF-16 string length".to_string()));
        }
        let utf16_data: Vec<u16> = string_data.chunks(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&utf16_data)
            .map_err(|e| XlsError::Encoding(format!("UTF-16 decoding error: {}", e)))?
    } else {
        // Assume compatible encoding
        String::from_utf8(string_data.to_vec())
            .unwrap_or_else(|_| String::from_utf8_lossy(string_data).into_owned())
    };

    Ok((string, consumed))
}

/// Convert RK value to f64
///
/// RK values are compressed numeric values used in Excel.
/// The format is: a 30-bit mantissa, 1 bit for 100x multiplier, 1 bit for int/float
pub fn rk_to_f64(rk: u32) -> f64 {
    let d100 = (rk & 0x02) != 0;
    let is_int = (rk & 0x01) != 0;

    let mut value = if is_int {
        // Integer value
        let int_val = (rk >> 2) as i32;
        if d100 {
            if int_val % 100 != 0 {
                int_val as f64 / 100.0
            } else {
                (int_val / 100) as f64
            }
        } else {
            int_val as f64
        }
    } else {
        // Float value - extract IEEE 754 double from 30 bits
        let mut float_bits = [0u8; 8];
        float_bits[0..4].copy_from_slice(&(rk & 0xFFFFFFFC).to_le_bytes());
        // Set the exponent to proper range
        float_bits[7] = 0x3C; // This is approximate
        f64::from_le_bytes(float_bits)
    };

    if d100 && !is_int {
        value /= 100.0;
    }

    value
}

/// Parse formula value from formula record
pub fn parse_formula_value(data: &[u8]) -> XlsResult<FormulaValue> {
    if data.len() < 8 {
        return Err(XlsError::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }

    match data[6] {
        0x00 => {
            // String (will be in next record)
            Ok(FormulaValue::Empty)
        }
        0x01 => {
            // Boolean
            Ok(FormulaValue::Bool(data[2] != 0))
        }
        0x02 => {
            // Error
            Ok(FormulaValue::Error(data[2]))
        }
        0x03 => {
            // Empty string
            Ok(FormulaValue::String(String::new()))
        }
        _ => {
            // Number
            Ok(FormulaValue::Number(binary::read_f64_le_at(data, 0)?))
        }
    }
}

/// Convert column number to Excel column name (A, B, ..., Z, AA, AB, etc.)
pub fn column_index_to_name(mut col: u32) -> String {
    let mut name = String::new();

    while col > 0 {
        col -= 1; // Make 0-based
        let ch = (b'A' + (col % 26) as u8) as char;
        name.insert(0, ch);
        col /= 26;
    }

    name
}

/// Convert Excel column name to column index (A=0, B=1, ..., Z=25, AA=26, etc.)
pub fn column_name_to_index(name: &str) -> Option<u32> {
    let mut result: u32 = 0;

    for ch in name.chars() {
        let ch = ch.to_ascii_uppercase();
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
pub fn parse_cell_reference(ref_str: &str) -> Option<(u32, u32)> {
    let ref_str = ref_str.to_ascii_uppercase();
    let mut col_str = String::new();
    let mut row_str = String::new();

    for ch in ref_str.chars() {
        if ch.is_ascii_uppercase() {
            col_str.push(ch);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        } else {
            return None;
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }

    let col = column_name_to_index(&col_str)?;
    let row: u32 = row_str.parse().ok()?;

    Some((row - 1, col)) // Make 0-based
}

/// Convert serial date to datetime
pub fn excel_date_to_datetime(serial: f64, is_1904: bool) -> Option<chrono::NaiveDateTime> {
    use chrono::{NaiveDate, Duration};

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
        assert_eq!(column_index_to_name(0), "A");
        assert_eq!(column_index_to_name(1), "B");
        assert_eq!(column_index_to_name(25), "Z");
        assert_eq!(column_index_to_name(26), "AA");
        assert_eq!(column_index_to_name(702), "AAA");
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
        assert_eq!(parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(parse_cell_reference("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_reference("a1"), Some((0, 0))); // case insensitive
        assert_eq!(parse_cell_reference("1A"), None); // invalid
        assert_eq!(parse_cell_reference("A"), None); // no row
        assert_eq!(parse_cell_reference("1"), None); // no column
    }
}
