//! Merged cell range support for XLSB

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};

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
    /// use litchi::ooxml::xlsb::merged_cells::MergedCell;
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
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        Ok(MergedCell {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        })
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
            crate::ooxml::xlsb::utils::cell_reference(self.row_first, self.col_first),
            crate::ooxml::xlsb::utils::cell_reference(self.row_last, self.col_last)
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
