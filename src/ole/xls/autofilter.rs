//! AutoFilter and Sort parsing for XLS BIFF8 files.
//!
//! Parses:
//! - AUTOFILTERINFO (0x009D): Specifies the number of columns with AutoFilter drop-downs.
//! - AUTOFILTER (0x009E): Specifies an individual column filter condition.
//! - SORT (0x0090): Specifies sort settings for the worksheet.
//! - SORTDATA (0x0895): Extended sort data (BIFF8+).
//!
//! # AutoFilter Record Format (MS-XLS 2.4.6)
//!
//! AUTOFILTERINFO:
//! ```text
//! Offset  Size  Field
//! 0       2     cEntries - Number of AutoFilter drop-down arrows
//! ```
//!
//! AUTOFILTER:
//! ```text
//! Offset  Size  Field
//! 0       2     iEntry - Index of this column within the AutoFilter range (0-based)
//! 2       2     grbit - Option flags
//! 4       10    doper1 - First filter condition (DOPER structure)
//! 14      10    doper2 - Second filter condition (DOPER structure)
//! 24      var   string1 - String for condition 1 (if applicable)
//!         var   string2 - String for condition 2 (if applicable)
//! ```

use crate::common::binary;
use crate::ole::xls::error::{XlsError, XlsResult};

/// AUTOFILTERINFO record type identifier.
pub const AUTOFILTERINFO_TYPE: u16 = 0x009D;

/// AUTOFILTER record type identifier.
pub const AUTOFILTER_TYPE: u16 = 0x009E;

/// SORT record type identifier.
pub const SORT_TYPE: u16 = 0x0090;

/// Comparison operators for AutoFilter conditions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterOperator {
    /// No filter applied.
    NoFilter,
    /// Less than.
    LessThan,
    /// Equal to.
    Equal,
    /// Less than or equal to.
    LessOrEqual,
    /// Greater than.
    GreaterThan,
    /// Not equal to.
    NotEqual,
    /// Greater than or equal to.
    GreaterOrEqual,
}

impl FilterOperator {
    fn from_u8(val: u8) -> Self {
        match val {
            0x01 => Self::LessThan,
            0x02 => Self::Equal,
            0x03 => Self::LessOrEqual,
            0x04 => Self::GreaterThan,
            0x05 => Self::NotEqual,
            0x06 => Self::GreaterOrEqual,
            _ => Self::NoFilter,
        }
    }
}

/// The type of a filter operand value inside a DOPER structure.
#[derive(Debug, Clone, PartialEq)]
pub enum FilterValue {
    /// No value / unused condition.
    None,
    /// RK-encoded number.
    Rk(f64),
    /// IEEE 754 double.
    Number(f64),
    /// A string operand (parsed from trailing bytes).
    String(String),
    /// Boolean value.
    Bool(bool),
    /// Error value.
    Error(u8),
    /// "Match all" / blanks.
    MatchAll,
}

/// A single filter condition extracted from a DOPER structure.
#[derive(Debug, Clone)]
pub struct FilterCondition {
    /// Comparison operator.
    pub operator: FilterOperator,
    /// The value to compare against.
    pub value: FilterValue,
}

/// Join logic between the two conditions on one column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterJoin {
    /// Both conditions must match (AND).
    And,
    /// Either condition must match (OR).
    Or,
}

/// A single AutoFilter column definition.
#[derive(Debug, Clone)]
pub struct AutoFilterColumn {
    /// Column index within the AutoFilter range (0-based relative to the filter area).
    pub column_index: u16,
    /// Join logic between conditions.
    pub join: FilterJoin,
    /// Whether this is a "Top 10" filter.
    pub is_top10: bool,
    /// Whether to show the drop-down arrow.
    pub show_arrow: bool,
    /// Whether a simple "match all" or custom filter.
    pub is_simple_filter: bool,
    /// First filter condition.
    pub condition1: FilterCondition,
    /// Second filter condition (optional; check `operator != NoFilter`).
    pub condition2: FilterCondition,
}

/// Complete AutoFilter state for a worksheet.
#[derive(Debug, Clone)]
pub struct AutoFilterInfo {
    /// Number of columns in the AutoFilter range.
    pub column_count: u16,
    /// Per-column filter definitions (populated from AUTOFILTER records).
    pub columns: Vec<AutoFilterColumn>,
}

/// Sort key direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortOrder {
    Ascending,
    Descending,
}

/// A single sort key.
#[derive(Debug, Clone)]
pub struct SortKey {
    /// Column index being sorted (0-based relative to the sort range).
    pub column_index: u16,
    /// Sort direction.
    pub order: SortOrder,
}

/// Sort configuration for a worksheet.
#[derive(Debug, Clone)]
pub struct SortInfo {
    /// Whether the sort is case-sensitive.
    pub case_sensitive: bool,
    /// Whether sort is by rows (true) or columns (false).
    pub sort_by_rows: bool,
    /// Sort keys (up to 3 in BIFF8).
    pub keys: Vec<SortKey>,
}

/// Parse an AUTOFILTERINFO record.
pub fn parse_autofilterinfo(data: &[u8]) -> XlsResult<u16> {
    if data.len() < 2 {
        return Err(XlsError::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    Ok(binary::read_u16_le_at(data, 0)?)
}

/// Parse a DOPER structure (10 bytes) at the given offset.
///
/// Returns `(FilterCondition, string_byte_count)` where `string_byte_count`
/// is the number of trailing string bytes to read (0 if not a string condition).
fn parse_doper(data: &[u8], offset: usize) -> XlsResult<(FilterCondition, usize)> {
    if offset + 10 > data.len() {
        return Err(XlsError::InvalidLength {
            expected: offset + 10,
            found: data.len(),
        });
    }

    let vt = data[offset]; // value type

    let (operator, value, str_bytes) = match vt {
        0x00 => {
            // Unused / filter not set
            (FilterOperator::NoFilter, FilterValue::None, 0)
        },
        0x02 => {
            // RK value
            let op = FilterOperator::from_u8(data[offset + 1]);
            let rk_bytes = binary::read_u32_le_at(data, offset + 2)?;
            let rk_val = crate::ole::xls::utils::rk_to_f64(rk_bytes);
            (op, FilterValue::Rk(rk_val), 0)
        },
        0x04 => {
            // IEEE double
            let op = FilterOperator::from_u8(data[offset + 1]);
            let val = binary::read_f64_le_at(data, offset + 2)?;
            (op, FilterValue::Number(val), 0)
        },
        0x06 => {
            // String
            let op = FilterOperator::from_u8(data[offset + 1]);
            // data[offset+2] is unused, data[offset+3] is the byte length of string
            let cb = data[offset + 3] as usize;
            // The string bytes follow the two DOPER structures
            (op, FilterValue::None, cb)
        },
        0x08 => {
            // Boolean / error
            let op = FilterOperator::from_u8(data[offset + 1]);
            let bval = data[offset + 2];
            let is_error = data[offset + 3];
            if is_error != 0 {
                (op, FilterValue::Error(bval), 0)
            } else {
                (op, FilterValue::Bool(bval != 0), 0)
            }
        },
        0x0C => {
            // Match all (blanks)
            let op = FilterOperator::from_u8(data[offset + 1]);
            (op, FilterValue::MatchAll, 0)
        },
        0x0E => {
            // Match all (non-blanks)
            let op = FilterOperator::from_u8(data[offset + 1]);
            (op, FilterValue::MatchAll, 0)
        },
        _ => {
            // Unknown; treat as no filter
            (FilterOperator::NoFilter, FilterValue::None, 0)
        },
    };

    Ok((FilterCondition { operator, value }, str_bytes))
}

/// Parse an AUTOFILTER record.
///
/// The record data starts after the standard BIFF header.
pub fn parse_autofilter(data: &[u8]) -> XlsResult<AutoFilterColumn> {
    // Minimum: 2 (iEntry) + 2 (grbit) + 10 (doper1) + 10 (doper2) = 24
    if data.len() < 24 {
        return Err(XlsError::InvalidLength {
            expected: 24,
            found: data.len(),
        });
    }

    let column_index = binary::read_u16_le_at(data, 0)?;
    let grbit = binary::read_u16_le_at(data, 2)?;

    let join = if grbit & 0x0001 != 0 {
        FilterJoin::Or
    } else {
        FilterJoin::And
    };
    let is_simple_filter = grbit & 0x0002 != 0;
    let is_top10 = grbit & 0x0010 != 0;
    let show_arrow = grbit & 0x0020 == 0; // bit set = hide arrow

    let (mut cond1, str_bytes1) = parse_doper(data, 4)?;
    let (mut cond2, str_bytes2) = parse_doper(data, 14)?;

    // Parse trailing string values
    let mut str_offset = 24;

    if str_bytes1 > 0 && str_offset < data.len() {
        let available = data.len() - str_offset;
        if available >= 3 {
            // XLUnicodeStringNoCch: 1-byte flags, then string data
            let flags = data[str_offset];
            str_offset += 1;
            let is_utf16 = flags & 0x01 != 0;
            let byte_count = if is_utf16 { str_bytes1 * 2 } else { str_bytes1 };
            if str_offset + byte_count <= data.len() {
                let s = if is_utf16 {
                    let words: Vec<u16> = data[str_offset..str_offset + byte_count]
                        .chunks_exact(2)
                        .map(|c| u16::from_le_bytes([c[0], c[1]]))
                        .collect();
                    String::from_utf16_lossy(&words)
                } else {
                    data[str_offset..str_offset + byte_count]
                        .iter()
                        .map(|&b| b as char)
                        .collect()
                };
                cond1.value = FilterValue::String(s);
                str_offset += byte_count;
            }
        }
    }

    if str_bytes2 > 0 && str_offset < data.len() {
        let available = data.len() - str_offset;
        if available >= 1 {
            let flags = data[str_offset];
            str_offset += 1;
            let is_utf16 = flags & 0x01 != 0;
            let byte_count = if is_utf16 { str_bytes2 * 2 } else { str_bytes2 };
            if str_offset + byte_count <= data.len() {
                let s = if is_utf16 {
                    let words: Vec<u16> = data[str_offset..str_offset + byte_count]
                        .chunks_exact(2)
                        .map(|c| u16::from_le_bytes([c[0], c[1]]))
                        .collect();
                    String::from_utf16_lossy(&words)
                } else {
                    data[str_offset..str_offset + byte_count]
                        .iter()
                        .map(|&b| b as char)
                        .collect()
                };
                cond2.value = FilterValue::String(s);
            }
        }
    }

    Ok(AutoFilterColumn {
        column_index,
        join,
        is_top10,
        show_arrow,
        is_simple_filter,
        condition1: cond1,
        condition2: cond2,
    })
}

/// Parse a SORT record (0x0090).
pub fn parse_sort(data: &[u8]) -> XlsResult<SortInfo> {
    // Minimum 10 bytes (flags + 3 sort key column indices + 3 reserved)
    if data.len() < 10 {
        return Err(XlsError::InvalidLength {
            expected: 10,
            found: data.len(),
        });
    }

    let flags = binary::read_u16_le_at(data, 0)?;

    let case_sensitive = flags & 0x0001 != 0;
    let sort_by_rows = flags & 0x0004 == 0; // bit clear = rows
    let key1_desc = flags & 0x0010 != 0;
    let key2_desc = flags & 0x0020 != 0;
    let key3_desc = flags & 0x0040 != 0;

    // Number of sort keys (1..3)
    let nsort = ((flags >> 1) & 0x01) + 1; // bit 1 indicates 2+ keys
    // Actually, determine from examining non-zero column indices

    let col1 = binary::read_u16_le_at(data, 2)?;
    let col2 = binary::read_u16_le_at(data, 4)?;
    let col3 = binary::read_u16_le_at(data, 6)?;

    let mut keys = Vec::with_capacity(3);

    // Key 1 always present if SORT record exists
    keys.push(SortKey {
        column_index: col1,
        order: if key1_desc {
            SortOrder::Descending
        } else {
            SortOrder::Ascending
        },
    });

    // Key 2 if column index is non-zero or nsort >= 2
    if col2 != 0 || nsort >= 2 {
        keys.push(SortKey {
            column_index: col2,
            order: if key2_desc {
                SortOrder::Descending
            } else {
                SortOrder::Ascending
            },
        });
    }

    // Key 3
    if col3 != 0 || nsort >= 3 {
        keys.push(SortKey {
            column_index: col3,
            order: if key3_desc {
                SortOrder::Descending
            } else {
                SortOrder::Ascending
            },
        });
    }

    Ok(SortInfo {
        case_sensitive,
        sort_by_rows,
        keys,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_autofilterinfo() {
        let data = 5u16.to_le_bytes();
        assert_eq!(parse_autofilterinfo(&data).unwrap(), 5);
    }

    #[test]
    fn test_parse_autofilter_numeric() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // iEntry = column 2
        data.extend_from_slice(&0u16.to_le_bytes()); // grbit: AND, no top10

        // doper1: IEEE double, operator = GreaterThan (0x04)
        data.push(0x04); // vt = double
        data.push(0x04); // operator = GreaterThan
        data.extend_from_slice(&100.0f64.to_le_bytes()); // value

        // doper2: unused
        data.extend_from_slice(&[0u8; 10]);

        let col = parse_autofilter(&data).unwrap();
        assert_eq!(col.column_index, 2);
        assert_eq!(col.join, FilterJoin::And);
        assert_eq!(col.condition1.operator, FilterOperator::GreaterThan);
        assert!(
            matches!(col.condition1.value, FilterValue::Number(v) if (v - 100.0).abs() < f64::EPSILON)
        );
        assert_eq!(col.condition2.operator, FilterOperator::NoFilter);
    }

    #[test]
    fn test_parse_sort_single_key() {
        let mut data = Vec::new();
        // flags: ascending, row sort, 1 key
        data.extend_from_slice(&0x0000u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes()); // col1
        data.extend_from_slice(&0u16.to_le_bytes()); // col2
        data.extend_from_slice(&0u16.to_le_bytes()); // col3
        data.extend_from_slice(&0u16.to_le_bytes()); // padding

        let sort = parse_sort(&data).unwrap();
        assert!(!sort.case_sensitive);
        assert!(sort.sort_by_rows);
        assert_eq!(sort.keys.len(), 1);
        assert_eq!(sort.keys[0].column_index, 3);
        assert_eq!(sort.keys[0].order, SortOrder::Ascending);
    }
}
