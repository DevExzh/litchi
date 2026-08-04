//! Typed merged-cell records for XLSB.
//!
//! This module owns the semantic representation and fixed-width codec for
//! `BrtMergeCell` from `[MS-XLSB]` section 2.4.713. Package and worksheet
//! orchestration remains in the host crate.

use thiserror::Error;

/// Exact byte length of a `BrtMergeCell` payload.
pub const LEN: usize = 16;

/// Maximum row index in an XLSB worksheet, inclusive.
pub const MAX_MERGED_CELL_ROW: u32 = 1_048_575;
/// Maximum column index in an XLSB worksheet, inclusive.
pub const MAX_MERGED_CELL_COLUMN: u32 = 16_383;
/// Maximum number of `BrtMergeCell` records in one worksheet collection.
pub const MAX_MERGED_CELL_RANGES: usize = 1_048_576;

/// Result type for merged-cell parsing and validation.
pub type Result<T> = std::result::Result<T, Error>;

/// A typed `BrtMergeCell` failure.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// The payload was not the fixed width required by the record.
    #[error("invalid BrtMergeCell payload length: expected {expected} bytes, found {found}")]
    InvalidLength {
        /// Required payload length.
        expected: usize,
        /// Actual payload length.
        found: usize,
    },
    /// The range is reversed or outside the worksheet's indexed bounds.
    #[error("invalid merged range rows {row_first}..={row_last} columns {col_first}..={col_last}")]
    InvalidRange {
        /// First row (zero-based).
        row_first: u32,
        /// Last row (zero-based, inclusive).
        row_last: u32,
        /// First column (zero-based).
        col_first: u32,
        /// Last column (zero-based, inclusive).
        col_last: u32,
    },
}

/// An inclusive merged-cell range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergedCell {
    /// First row (zero-based).
    pub row_first: u32,
    /// Last row (zero-based, inclusive).
    pub row_last: u32,
    /// First column (zero-based).
    pub col_first: u32,
    /// Last column (zero-based, inclusive).
    pub col_last: u32,
}

impl MergedCell {
    /// Create a merged-cell range.
    #[must_use]
    pub const fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    /// Parse one fixed-width `BrtMergeCell` payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != LEN {
            return Err(Error::InvalidLength {
                expected: LEN,
                found: data.len(),
            });
        }

        let range = Self {
            row_first: read_u32(data, 0),
            row_last: read_u32(data, 4),
            col_first: read_u32(data, 8),
            col_last: read_u32(data, 12),
        };
        range.validate()?;
        Ok(range)
    }

    /// Validate worksheet bounds and inclusive endpoint ordering.
    pub fn validate(&self) -> Result<()> {
        if self.row_first > self.row_last
            || self.row_last > MAX_MERGED_CELL_ROW
            || self.col_first > self.col_last
            || self.col_last > MAX_MERGED_CELL_COLUMN
        {
            return Err(Error::InvalidRange {
                row_first: self.row_first,
                row_last: self.row_last,
                col_first: self.col_first,
                col_last: self.col_last,
            });
        }
        Ok(())
    }

    /// Return whether this range shares at least one cell with another range.
    #[must_use]
    pub const fn overlaps(&self, other: &Self) -> bool {
        self.row_first <= other.row_last
            && other.row_first <= self.row_last
            && self.col_first <= other.col_last
            && other.col_first <= self.col_last
    }

    /// Serialize to one fixed-width `BrtMergeCell` payload.
    #[must_use]
    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(LEN);
        data.extend_from_slice(&self.row_first.to_le_bytes());
        data.extend_from_slice(&self.row_last.to_le_bytes());
        data.extend_from_slice(&self.col_first.to_le_bytes());
        data.extend_from_slice(&self.col_last.to_le_bytes());
        data
    }

    /// Render the range using A1 notation, such as `A1:B2`.
    #[must_use]
    pub fn to_range_string(&self) -> String {
        format!(
            "{}:{}",
            cell_reference(self.row_first, self.col_first),
            cell_reference(self.row_last, self.col_last)
        )
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn cell_reference(row: u32, col: u32) -> String {
    let Some(column) = col.checked_add(1) else {
        return format!("R{row}C{col}");
    };
    let Some(row) = row.checked_add(1) else {
        return format!("R{row}C{col}");
    };
    format!("{}{}", column_name(column), row)
}

fn column_name(mut column: u32) -> String {
    let mut name = String::new();
    while column > 0 {
        column -= 1;
        name.push(char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    name.chars().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn range_string_uses_a1_notation() {
        let merged = MergedCell::new(0, 1, 0, 1);
        assert_eq!(merged.to_range_string(), "A1:B2");
    }

    #[test]
    fn serialize_parse_round_trip() {
        let merged = MergedCell::new(0, 1, 0, 1);
        let data = merged.serialize();
        assert_eq!(data.len(), LEN);
        assert_eq!(MergedCell::parse(&data), Ok(merged));
    }

    #[test]
    fn rejects_invalid_payload_and_ranges() {
        assert_eq!(
            MergedCell::parse(&[0; LEN - 1]),
            Err(Error::InvalidLength {
                expected: LEN,
                found: LEN - 1,
            })
        );
        assert!(matches!(
            MergedCell::new(2, 1, 0, 0).validate(),
            Err(Error::InvalidRange { .. })
        ));
        assert!(matches!(
            MergedCell::new(0, MAX_MERGED_CELL_ROW + 1, 0, 0).validate(),
            Err(Error::InvalidRange { .. })
        ));
    }

    #[test]
    fn overlap_is_inclusive() {
        let left = MergedCell::new(0, 1, 0, 1);
        let touching = MergedCell::new(1, 2, 1, 2);
        let separate = MergedCell::new(2, 3, 2, 3);
        assert!(left.overlaps(&touching));
        assert!(!left.overlaps(&separate));
    }
}
