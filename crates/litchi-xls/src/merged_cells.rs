//! Merged cells parsing for XLS BIFF8 files.
//!
//! Parses MERGECELLS records (0x00E5) which define rectangular regions
//! of cells that are merged into a single display cell.
//!
//! # Record Format (MS-XLS 2.4.168)
//!
//! ```text
//! Offset  Size  Field
//! 0       2     cmcs - Number of merged cell ranges in this record
//! 2       8*n   Array of CellRangeAddress structures (rwFirst, rwLast, colFirst, colLast)
//! ```
//!
//! Multiple MERGECELLS records may appear per worksheet; all are aggregated.
//! Each record holds at most 1027 ranges (enforced by the 8224-byte BIFF limit).

use crate::error::{Error, Result};
use litchi_core::binary;

/// MERGECELLS record type identifier.
pub const RECORD_TYPE: u16 = 0x00E5;

/// A merged cell range: `(first_row, last_row, first_col, last_col)`, all 0-based.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MergedCellRange {
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
}

impl MergedCellRange {
    /// Number of rows spanned by this merge.
    #[inline]
    pub fn row_span(&self) -> u16 {
        self.last_row - self.first_row + 1
    }

    /// Number of columns spanned by this merge.
    #[inline]
    pub fn col_span(&self) -> u16 {
        self.last_col - self.first_col + 1
    }

    /// Whether the given cell `(row, col)` falls within this merged region.
    #[inline]
    pub fn contains(&self, row: u16, col: u16) -> bool {
        row >= self.first_row
            && row <= self.last_row
            && col >= self.first_col
            && col <= self.last_col
    }
}

/// Parse a single MERGECELLS record and append results to `out`.
///
/// # Layout
///
/// - `u16` cmcs: count of CellRangeAddress structs
/// - `cmcs * 8` bytes: array of `(rwFirst: u16, rwLast: u16, colFirst: u16, colLast: u16)`
pub fn parse_mergecells_record(data: &[u8], out: &mut Vec<MergedCellRange>) -> Result<()> {
    if data.len() < 2 {
        return Err(Error::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }

    let count = binary::read_u16_le_at(data, 0)? as usize;
    let expected_len = 2 + count * 8;

    if data.len() < expected_len {
        return Err(Error::InvalidLength {
            expected: expected_len,
            found: data.len(),
        });
    }

    out.reserve(count);
    let mut offset = 2;

    for _ in 0..count {
        let first_row = binary::read_u16_le_at(data, offset)?;
        let last_row = binary::read_u16_le_at(data, offset + 2)?;
        let first_col = binary::read_u16_le_at(data, offset + 4)?;
        let last_col = binary::read_u16_le_at(data, offset + 6)?;

        out.push(MergedCellRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });

        offset += 8;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_single_range() {
        // 1 range: rows 0-1, cols 0-2
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // cmcs = 1
        data.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        data.extend_from_slice(&1u16.to_le_bytes()); // rwLast
        data.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        data.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let mut out = Vec::new();
        parse_mergecells_record(&data, &mut out).unwrap();

        assert_eq!(out.len(), 1);
        assert_eq!(out[0].first_row, 0);
        assert_eq!(out[0].last_row, 1);
        assert_eq!(out[0].first_col, 0);
        assert_eq!(out[0].last_col, 2);
        assert_eq!(out[0].row_span(), 2);
        assert_eq!(out[0].col_span(), 3);
        assert!(out[0].contains(0, 1));
        assert!(!out[0].contains(2, 0));
    }

    #[test]
    fn test_parse_multiple_ranges() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // cmcs = 2
        // Range 1
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        // Range 2
        data.extend_from_slice(&5u16.to_le_bytes());
        data.extend_from_slice(&10u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&4u16.to_le_bytes());

        let mut out = Vec::new();
        parse_mergecells_record(&data, &mut out).unwrap();

        assert_eq!(out.len(), 2);
        assert_eq!(out[1].first_row, 5);
        assert_eq!(out[1].last_row, 10);
    }

    #[test]
    fn test_rejects_truncated_records_without_panicking() {
        let mut out = Vec::new();

        // Too short to hold the cmcs count field.
        assert!(parse_mergecells_record(&[], &mut out).is_err());
        assert!(parse_mergecells_record(&[1], &mut out).is_err());

        // cmcs claims two ranges but the payload holds only one.
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&[0u8; 8]);
        assert!(parse_mergecells_record(&data, &mut out).is_err());

        // Range array truncated mid-entry.
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&[0u8; 5]);
        assert!(parse_mergecells_record(&data, &mut out).is_err());

        // No partial ranges may leak into the output on failure.
        assert!(out.is_empty());
    }

    #[test]
    fn reads_poi_merged_cell_fixtures() {
        use crate::Workbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet")
                .join(name)
        };

        // Single merged region: rows 8-15, columns C-F (0-based 7-14 / 2-5).
        let workbook = Workbook::new(File::open(fixture("53109.xls")).unwrap()).unwrap();
        let merged = workbook.xls_worksheet(0).unwrap().merged_cells();
        assert_eq!(merged.len(), 1);
        assert_eq!(
            merged[0],
            MergedCellRange {
                first_row: 7,
                last_row: 14,
                first_col: 2,
                last_col: 5,
            }
        );
        assert_eq!(merged[0].row_span(), 8);
        assert_eq!(merged[0].col_span(), 4);
        assert!(merged[0].contains(10, 3));
        assert!(!merged[0].contains(6, 3));

        // A workbook holding many merged regions across several sheets.
        let workbook = Workbook::new(File::open(fixture("59858.xls")).unwrap()).unwrap();
        let per_sheet: Vec<usize> = (0..workbook.sheets().len())
            .map(|index| workbook.xls_worksheet(index).unwrap().merged_cells().len())
            .collect();
        assert_eq!(per_sheet, [5, 366, 15, 352, 172, 0, 0]);
    }
}
