//! Binary range serialization for XLSB writer.
//!
//! Provides helpers to serialize cell ranges into the BIFF12 `BinRangeList`
//! format used by both data validation (`BrtDVal`) and conditional formatting
//! (`BrtBeginCondFormatting`) records.
//!
//! # Binary Layout (per LibreOffice `addressconverter.cxx`)
//!
//! - **BinRange** (16 bytes): `row_first(i32) + row_last(i32) + col_first(i32) + col_last(i32)`
//! - **BinRangeList**: `count(i32)` followed by `count` × `BinRange`

use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::utils::parse_cell_reference;
use std::io::Write;

/// A parsed 0-based cell range with inclusive row/column bounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl CellRange {
    /// Create a new cell range from 0-based inclusive coordinates.
    #[inline]
    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Self {
        Self {
            row_first,
            row_last,
            col_first,
            col_last,
        }
    }

    /// Parse a range string like `"A1:B2"` or a single cell like `"A1"`.
    ///
    /// Both parts are expected in standard Excel notation (1-based rows, letter
    /// columns). The returned coordinates are 0-based.
    pub fn parse(range_str: &str) -> XlsbResult<Self> {
        if let Some((left, right)) = range_str.split_once(':') {
            let (r1, c1) = parse_cell_reference(left.trim())?;
            let (r2, c2) = parse_cell_reference(right.trim())?;
            Ok(Self::new(r1, r2, c1, c2))
        } else {
            let (r, c) = parse_cell_reference(range_str.trim())?;
            Ok(Self::new(r, r, c, c))
        }
    }

    /// Serialize as a single BIFF12 `BinRange` (16 bytes, little-endian i32s).
    pub fn write<W: Write>(&self, w: &mut W) -> XlsbResult<()> {
        w.write_all(&(self.row_first as i32).to_le_bytes())?;
        w.write_all(&(self.row_last as i32).to_le_bytes())?;
        w.write_all(&(self.col_first as i32).to_le_bytes())?;
        w.write_all(&(self.col_last as i32).to_le_bytes())?;
        Ok(())
    }
}

/// Parse a comma-separated list of range strings (e.g. `"A1:B2,C3:D4"`) into
/// a vector of [`CellRange`]s.
pub fn parse_range_list(sqref: &str) -> XlsbResult<Vec<CellRange>> {
    sqref
        .split([',', ' '])
        .filter(|s| !s.is_empty())
        .map(CellRange::parse)
        .collect()
}

/// Serialize a slice of [`CellRange`]s as a BIFF12 `BinRangeList`.
///
/// Layout: `count(i32)` + `count` × 16-byte `BinRange`.
pub fn write_bin_range_list<W: Write>(ranges: &[CellRange], w: &mut W) -> XlsbResult<()> {
    w.write_all(&(ranges.len() as i32).to_le_bytes())?;
    for r in ranges {
        r.write(w)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_single_cell() {
        let r = CellRange::parse("A1").unwrap();
        assert_eq!(r, CellRange::new(0, 0, 0, 0));
    }

    #[test]
    fn test_parse_range() {
        let r = CellRange::parse("B2:D5").unwrap();
        assert_eq!(r, CellRange::new(1, 4, 1, 3));
    }

    #[test]
    fn test_parse_range_list() {
        let ranges = parse_range_list("A1:B2,C3:D4").unwrap();
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0], CellRange::new(0, 1, 0, 1));
        assert_eq!(ranges[1], CellRange::new(2, 3, 2, 3));
    }

    #[test]
    fn test_serialize_bin_range_list() {
        let ranges = [CellRange::new(0, 9, 0, 3)];
        let mut buf = Vec::new();
        write_bin_range_list(&ranges, &mut buf).unwrap();
        // count(4) + 1 × BinRange(16) = 20 bytes
        assert_eq!(buf.len(), 20);
        // count = 1
        assert_eq!(&buf[0..4], &1i32.to_le_bytes());
    }
}
