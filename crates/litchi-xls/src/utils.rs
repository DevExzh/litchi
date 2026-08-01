//! Utility functions for XLS parsing

use crate::error::{XlsError, XlsResult};
use crate::records::{FormulaValue, XlsEncoding};
use litchi_core::binary;
use zerocopy::{FromBytes, LE, U16};

/// Parse a BIFF8 `ShortXLUnicodeString`.
///
/// Layout: `[cch: u8] [flags: u8] [chars...]`
///
/// - `cch` — character count
/// - `flags` bit 0 (`fHighByte`) — 0 = compressed Latin-1 (1 byte/char),
///   1 = uncompressed UTF-16LE (2 bytes/char)
pub fn parse_short_string(data: &[u8], _encoding: &XlsEncoding) -> XlsResult<String> {
    if data.len() < 2 {
        return Ok(String::new());
    }

    let cch = data[0] as usize;
    let flags = data[1];
    let high_byte = (flags & 0x01) != 0;

    let byte_len = if high_byte { cch * 2 } else { cch };
    let offset = 2; // skip cch + flags

    if data.len() < offset + byte_len {
        return Err(XlsError::InvalidLength {
            expected: offset + byte_len,
            found: data.len(),
        });
    }

    let string_data = &data[offset..offset + byte_len];

    if high_byte {
        // UTF-16LE
        let utf16: Vec<u16> = string_data
            .chunks_exact(2)
            .map(|chunk| {
                U16::<LE>::read_from_bytes(chunk)
                    .map(|v| v.get())
                    .unwrap_or(0)
            })
            .collect();
        String::from_utf16(&utf16)
            .map_err(|e| XlsError::Encoding(format!("UTF-16 decoding error: {}", e)))
    } else {
        // Compressed Latin-1 (each byte maps directly to U+00xx)
        Ok(string_data.iter().map(|&b| b as char).collect())
    }
}

/// Parse a BIFF8 `XLUnicodeString` with a 16-bit character count.
pub fn parse_string_record(data: &[u8], _encoding: &XlsEncoding) -> XlsResult<String> {
    if data.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }

    let len = binary::read_u16_le_at(data, 0)? as usize;
    let flags = data[2];

    let high_byte = (flags & 0x01) != 0;
    let offset = 3;
    let byte_len = len
        .checked_mul(if high_byte { 2 } else { 1 })
        .ok_or_else(|| XlsError::InvalidData("XLUnicodeString length overflow".to_string()))?;

    if data.len() < offset + byte_len {
        return Err(XlsError::InvalidLength {
            expected: offset + byte_len,
            found: data.len(),
        });
    }

    let string_data = &data[offset..offset + byte_len];

    if high_byte {
        let utf16_data: Vec<u16> = string_data
            .chunks_exact(2)
            .map(|chunk| {
                U16::<LE>::read_from_bytes(chunk)
                    .map(|v| v.get())
                    .unwrap_or(0)
            })
            .collect();
        String::from_utf16(&utf16_data)
            .map_err(|error| XlsError::Encoding(format!("UTF-16 decoding error: {error}")))
    } else {
        // Compressed Unicode supplies an implicit zero high byte; CODEPAGE
        // does not apply to this BIFF8 structure.
        Ok(string_data.iter().map(|&byte| byte as char).collect())
    }
}

/// `fHighByte` option bit of an `XLUnicodeString` (MS-XLS 2.5.294).
const STRING_HIGH_BYTE: u8 = 0x01;

/// Outcome of decoding a possibly continued BIFF8 `String` record payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StringRecordDecode {
    /// All declared characters were decoded.
    Complete(String),
    /// The declared characters extend past the supplied payloads; another
    /// `Continue` record payload is required.
    NeedContinue,
}

/// Decode the cached result of a `String` record (MS-XLS 2.4.296) whose
/// character data may span `Continue` records (MS-XLS 2.1:
/// `FORMULA = ... [String *Continue]`).
///
/// `first` is the `String` record payload and `continues` the payloads of the
/// `Continue` records that follow it. Every `Continue` payload restarts with
/// an option-flags byte selecting the character width of its chunk, so the
/// encoding may switch between compressed and UTF-16 at record boundaries.
/// Returns [`StringRecordDecode::NeedContinue`] when the declared characters
/// extend past every supplied payload.
pub fn decode_string_record(first: &[u8], continues: &[Vec<u8>]) -> XlsResult<StringRecordDecode> {
    if first.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: first.len(),
        });
    }
    let mut chars_left = usize::from(binary::read_u16_le_at(first, 0)?);
    let mut high_byte = first[2] & STRING_HIGH_BYTE != 0;
    let mut units: Vec<u16> = Vec::with_capacity(chars_left);
    let mut first_segment = true;
    for segment in std::iter::once(&first[3..]).chain(continues.iter().map(Vec::as_slice)) {
        let mut chunk: &[u8] = segment;
        if !first_segment {
            // A Continue payload restarts with an option-flags byte.
            let Some((&flags, rest)) = chunk.split_first() else {
                continue;
            };
            high_byte = flags & STRING_HIGH_BYTE != 0;
            chunk = rest;
        }
        first_segment = false;
        if high_byte {
            let mut consumed = 0usize;
            for pair in chunk.chunks_exact(2) {
                if chars_left == 0 {
                    break;
                }
                units.push(u16::from_le_bytes([pair[0], pair[1]]));
                chars_left -= 1;
                consumed += 2;
            }
            if chars_left > 0 && consumed < chunk.len() {
                return Err(XlsError::InvalidData(
                    "UTF-16 character split across String Continue records".to_string(),
                ));
            }
        } else {
            for &byte in chunk {
                if chars_left == 0 {
                    break;
                }
                units.push(u16::from(byte));
                chars_left -= 1;
            }
        }
        if chars_left == 0 {
            break;
        }
    }
    if chars_left > 0 {
        return Ok(StringRecordDecode::NeedContinue);
    }
    String::from_utf16(&units)
        .map(StringRecordDecode::Complete)
        .map_err(|error| XlsError::Encoding(format!("UTF-16 decoding error: {error}")))
}

/// Convert RK value to f64
///
/// RK values are compressed numeric values used in Excel.
/// Bit 0 requests division by 100; bit 1 selects a signed 30-bit integer.
/// Otherwise the upper 30 bits are the most-significant bits of an IEEE-754 double.
pub fn rk_to_f64(rk: u32) -> f64 {
    let mut value = if rk & 0x02 != 0 {
        f64::from((rk as i32) >> 2)
    } else {
        f64::from_bits(u64::from(rk & 0xFFFF_FFFC) << 32)
    };
    if rk & 0x01 != 0 {
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

    if data[6..8] != [0xFF, 0xFF] {
        return Ok(FormulaValue::Number(binary::read_f64_le_at(data, 0)?));
    }

    match data[0] {
        0x00 => Ok(FormulaValue::StringPending),
        0x01 => Ok(FormulaValue::Bool(data[2] != 0)),
        0x02 => Ok(FormulaValue::Error(data[2])),
        0x03 => Ok(FormulaValue::Empty),
        value_type => Err(XlsError::InvalidData(format!(
            "Invalid formula cached-value type: {value_type}"
        ))),
    }
}

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
    let mut found_digit = false;

    for ch in ref_str.chars() {
        if ch.is_ascii_uppercase() {
            // Letters must come before digits
            if found_digit {
                return None;
            }
            col_str.push(ch);
        } else if ch.is_ascii_digit() {
            found_digit = true;
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
    fn parses_xl_unicode_string_character_counts() {
        let utf16 = [2, 0, 1, 0x22, 0x6F, 0x57, 0x5B];
        assert_eq!(
            parse_string_record(&utf16, &XlsEncoding::Utf16Le).unwrap(),
            "漢字"
        );

        let compressed = [1, 0, 0, 0xC0];
        assert_eq!(
            parse_string_record(&compressed, &XlsEncoding::from_codepage(1251).unwrap()).unwrap(),
            "À"
        );
    }

    #[test]
    fn rejects_truncated_xl_unicode_string() {
        let truncated = [2, 0, 1, 0x22, 0x6F];
        assert!(parse_string_record(&truncated, &XlsEncoding::Utf16Le).is_err());
    }

    #[test]
    fn decodes_all_rk_number_encodings() {
        assert_eq!(rk_to_f64((42u32 << 2) | 0x02), 42.0);
        assert_eq!(rk_to_f64(((-42i32) as u32) << 2 | 0x02), -42.0);
        assert_eq!(rk_to_f64((1234u32 << 2) | 0x03), 12.34);

        let encoded_float = ((1.5f64.to_bits() >> 32) as u32) & 0xFFFF_FFFC;
        assert_eq!(rk_to_f64(encoded_float), 1.5);
        assert_eq!(rk_to_f64(encoded_float | 0x01), 0.015);
    }

    #[test]
    fn decodes_formula_cached_value_marker_and_types() {
        assert!(matches!(
            parse_formula_value(&[0, 0, 0, 0, 0, 0, 0xFF, 0xFF]).unwrap(),
            FormulaValue::StringPending
        ));
        assert!(matches!(
            parse_formula_value(&[1, 0, 1, 0, 0, 0, 0xFF, 0xFF]).unwrap(),
            FormulaValue::Bool(true)
        ));
        assert!(matches!(
            parse_formula_value(&[2, 0, 0x07, 0, 0, 0, 0xFF, 0xFF]).unwrap(),
            FormulaValue::Error(0x07)
        ));
        assert!(matches!(
            parse_formula_value(&[3, 0, 0, 0, 0, 0, 0xFF, 0xFF]).unwrap(),
            FormulaValue::Empty
        ));

        let numeric = 3.5f64.to_le_bytes();
        assert!(matches!(
            parse_formula_value(&numeric).unwrap(),
            FormulaValue::Number(value) if value == 3.5
        ));
    }

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
        assert_eq!(parse_cell_reference("A1"), Some((0, 0)));
        assert_eq!(parse_cell_reference("B2"), Some((1, 1)));
        assert_eq!(parse_cell_reference("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_reference("a1"), Some((0, 0))); // case insensitive
        assert!(parse_cell_reference("1A").is_none()); // invalid - digits before letters
        assert!(parse_cell_reference("A").is_none()); // no row
        assert!(parse_cell_reference("1").is_none()); // no column
    }

    #[test]
    fn decode_string_record_completes_within_one_record() {
        let compressed = [3, 0, 0, b'a', b'b', b'c'];
        assert_eq!(
            decode_string_record(&compressed, &[]).unwrap(),
            StringRecordDecode::Complete("abc".to_string())
        );

        let utf16 = [2, 0, 1, 0x22, 0x6F, 0x57, 0x5B];
        assert_eq!(
            decode_string_record(&utf16, &[]).unwrap(),
            StringRecordDecode::Complete("漢字".to_string())
        );
    }

    #[test]
    fn decode_string_record_reports_missing_continuation() {
        let short = [5, 0, 0, b'a', b'b'];
        assert_eq!(
            decode_string_record(&short, &[]).unwrap(),
            StringRecordDecode::NeedContinue
        );
    }

    #[test]
    fn decode_string_record_switches_encoding_at_continue_boundary() {
        // "ab文": two compressed characters in the String record, then a
        // UTF-16 chunk in the Continue record.
        let first = [3, 0, 0, b'a', b'b'];
        let continues = vec![vec![1, 0x87, 0x65]];
        assert_eq!(
            decode_string_record(&first, &continues).unwrap(),
            StringRecordDecode::Complete("ab文".to_string())
        );
    }

    #[test]
    fn decode_string_record_rejects_split_utf16_character() {
        // The String record ends on a dangling byte of a UTF-16 character.
        let first = [2, 0, 1, 0x22];
        let continues = vec![vec![1, 0x6F, 0x57, 0x5B]];
        assert!(decode_string_record(&first, &continues).is_err());
    }
}
