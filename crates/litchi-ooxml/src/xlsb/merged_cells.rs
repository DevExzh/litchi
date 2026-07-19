//! Merged cell range support for XLSB

use crate::xlsb::error::{XlsbError, XlsbResult};
use litchi_core::binary;

/// Maximum row index in an XLSB worksheet, inclusive.
pub const MAX_MERGED_CELL_ROW: u32 = 1_048_575;
/// Maximum column index in an XLSB worksheet, inclusive.
pub const MAX_MERGED_CELL_COLUMN: u32 = 16_383;
/// Maximum number of `BrtMergeCell` records in one worksheet collection.
pub const MAX_MERGED_CELL_RANGES: usize = 1_048_576;

/// Merged cell range
///
/// Represents a range of cells that are merged together.
/// The range is inclusive on all sides.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergedCell {
    /// First row (0-based)
    pub row_first: u32,
    /// Last row (0-based, inclusive)
    pub row_last: u32,
    /// First column (0-based)
    pub col_first: u32,
    /// Last column (0-based, inclusive)
    pub col_last: u32,
}

impl MergedCell {
    /// Create a new merged cell range
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_ooxml::xlsb::merged_cells::MergedCell;
    ///
    /// // Merge cells A1:B2
    /// let merged = MergedCell::new(0, 1, 0, 1);
    /// ```
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        MergedCell {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    /// Parse from XLSB BrtMergeCell record
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() != 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        let range = MergedCell {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        };
        range.validate()?;
        Ok(range)
    }

    /// Validate normalized XLSB worksheet bounds for this range.
    pub fn validate(&self) -> XlsbResult<()> {
        if self.row_first > self.row_last
            || self.row_last > MAX_MERGED_CELL_ROW
            || self.col_first > self.col_last
            || self.col_last > MAX_MERGED_CELL_COLUMN
        {
            return Err(XlsbError::InvalidCellReference(format!(
                "invalid merged range rows {}..={} columns {}..={}",
                self.row_first, self.row_last, self.col_first, self.col_last
            )));
        }
        Ok(())
    }

    /// Return whether this range shares at least one cell with another range.
    pub fn overlaps(&self, other: &Self) -> bool {
        self.row_first <= other.row_last
            && other.row_first <= self.row_last
            && self.col_first <= other.col_last
            && other.col_first <= self.col_last
    }

    /// Serialize to XLSB BrtMergeCell record
    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(16);
        data.extend_from_slice(&self.row_first.to_le_bytes());
        data.extend_from_slice(&self.row_last.to_le_bytes());
        data.extend_from_slice(&self.col_first.to_le_bytes());
        data.extend_from_slice(&self.col_last.to_le_bytes());
        data
    }

    /// Get the cell range as a string (e.g., "A1:B2")
    pub fn to_range_string(&self) -> String {
        format!(
            "{}:{}",
            crate::xlsb::utils::cell_reference(self.row_first, self.col_first),
            crate::xlsb::utils::cell_reference(self.row_last, self.col_last)
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_merged_cell_range_string() {
        let merged = MergedCell::new(0, 1, 0, 1);
        assert_eq!(merged.to_range_string(), "A1:B2");
    }

    #[test]
    fn test_merged_cell_serialize_parse() {
        let merged = MergedCell::new(0, 1, 0, 1);
        let data = merged.serialize();
        let parsed = MergedCell::parse(&data).unwrap();
        assert_eq!(merged, parsed);
    }
}
