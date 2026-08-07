//! Producer-typed chart-cache dimensions and Excel cache sections.

use super::{Kind, RowCol};

/// Excel `SIIndex` section. The wire grammar requires these three values in
/// this exact order, even when a section contains no cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Index {
    /// Series values (or horizontal scatter/bubble values).
    Values = 1,
    /// Category labels (or vertical scatter/bubble values).
    Categories = 2,
    /// Bubble sizes.
    Bubbles = 3,
}

impl Index {
    /// Canonical SERIESDATA section order.
    pub const ALL: [Self; 3] = [Self::Values, Self::Categories, Self::Bubbles];

    pub(super) const fn from_raw(value: u16) -> Option<Self> {
        match value {
            1 => Some(Self::Values),
            2 => Some(Self::Categories),
            3 => Some(Self::Bubbles),
            _ => None,
        }
    }
}

/// Excel extended-format-table index stored by cached chart cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Xf(u16);

impl Xf {
    /// Preserves one Excel XF index.
    #[must_use]
    pub const fn new(value: u16) -> Self {
        Self(value)
    }

    /// Raw BIFF XF index.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Standalone Graph internal-format-table index stored by cached cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Ifmt(u16);

impl Ifmt {
    /// Preserves one Graph `IFmt` index.
    #[must_use]
    pub const fn new(value: u16) -> Self {
        Self(value)
    }

    /// Raw Graph `IFmt` index.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// BIFF error value stored by an Excel `BoolErr` cache cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Fault {
    /// `#NULL!`
    Null = 0x00,
    /// `#DIV/0!`
    DivZero = 0x07,
    /// `#VALUE!`
    Value = 0x0F,
    /// `#REF!`
    Ref = 0x17,
    /// `#NAME?`
    Name = 0x1D,
    /// `#NUM!`
    Num = 0x24,
    /// `#N/A`
    Na = 0x2A,
    /// `#GETTING_DATA`
    GettingData = 0x2B,
}

impl Fault {
    pub(super) const fn from_raw(value: u8) -> Option<Self> {
        match value {
            0x00 => Some(Self::Null),
            0x07 => Some(Self::DivZero),
            0x0F => Some(Self::Value),
            0x17 => Some(Self::Ref),
            0x1D => Some(Self::Name),
            0x24 => Some(Self::Num),
            0x2A => Some(Self::Na),
            0x2B => Some(Self::GettingData),
            _ => None,
        }
    }
}

/// Excel used-range bounds from `Dimensions`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct ExcelDims {
    first_row: u32,
    row_after: u32,
    first_col: u16,
    col_after: u16,
}

impl ExcelDims {
    /// Creates checked BIFF8 used-range bounds. Zero `row_after`/`col_after`
    /// denotes an empty cache.
    #[must_use]
    pub const fn new(
        first_row: u32,
        row_after: u32,
        first_col: u16,
        col_after: u16,
    ) -> Option<Self> {
        if first_row > 0xFFFF
            || row_after > 0x1_0000
            || first_col > 0x00FF
            || col_after > 0x0100
            || (row_after == 0 && first_row != 0)
            || (row_after != 0 && first_row >= row_after)
            || (col_after == 0 && first_col != 0)
            || (col_after != 0 && first_col >= col_after)
            || ((row_after == 0) != (col_after == 0))
        {
            return None;
        }
        Some(Self {
            first_row,
            row_after,
            first_col,
            col_after,
        })
    }

    /// First used row.
    #[must_use]
    pub const fn first_row(self) -> u32 {
        self.first_row
    }

    /// Row immediately after the used range.
    #[must_use]
    pub const fn row_after(self) -> u32 {
        self.row_after
    }

    /// First used column.
    #[must_use]
    pub const fn first_col(self) -> u16 {
        self.first_col
    }

    /// Column immediately after the used range.
    #[must_use]
    pub const fn col_after(self) -> u16 {
        self.col_after
    }
}

/// Standalone Graph datasheet dimensions.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct GraphDims {
    longest_row: RowCol,
    rows: u8,
}

impl GraphDims {
    /// Creates dimensions from the longest non-empty row width and the number
    /// of non-empty rows.
    #[must_use]
    pub const fn new(longest_row: RowCol, rows: u8) -> Option<Self> {
        if (longest_row.get() == 0) != (rows == 0) {
            return None;
        }
        Some(Self { longest_row, rows })
    }

    /// Number of non-empty cells in the longest row.
    #[must_use]
    pub const fn longest_row(self) -> RowCol {
        self.longest_row
    }

    /// Number of non-empty rows.
    #[must_use]
    pub const fn rows(self) -> u8 {
        self.rows
    }
}

/// Context-specific `Dimensions` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Dims {
    /// Excel BIFF8 used range.
    Excel(ExcelDims),
    /// Standalone Graph datasheet dimensions.
    Graph(GraphDims),
}

impl Dims {
    pub(super) const fn empty(kind: Kind) -> Self {
        match kind {
            Kind::Excel => Self::Excel(ExcelDims {
                first_row: 0,
                row_after: 0,
                first_col: 0,
                col_after: 0,
            }),
            Kind::Graph => Self::Graph(GraphDims {
                longest_row: RowCol::ZERO,
                rows: 0,
            }),
        }
    }

    /// Whether these dimensions belong to the requested producer grammar.
    #[must_use]
    pub const fn matches(self, kind: Kind) -> bool {
        matches!(
            (self, kind),
            (Self::Excel(_), Kind::Excel) | (Self::Graph(_), Kind::Graph)
        )
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic by design"
    )]
    use super::*;

    #[test]
    fn excel_dimensions_reject_reversed_and_out_of_grid_ranges() {
        assert!(ExcelDims::new(0, 0, 0, 0).is_some());
        assert!(ExcelDims::new(1, 2, 3, 4).is_some());
        assert!(ExcelDims::new(2, 2, 0, 1).is_none());
        assert!(ExcelDims::new(0, 0x1_0001, 0, 1).is_none());
        assert!(ExcelDims::new(0, 1, 0, 0x0101).is_none());
    }

    #[test]
    fn only_three_excel_sections_exist() {
        assert_eq!(Index::from_raw(1), Some(Index::Values));
        assert_eq!(Index::from_raw(3), Some(Index::Bubbles));
        assert_eq!(Index::from_raw(0), None);
        assert_eq!(Index::from_raw(4), None);
    }

    #[test]
    fn cache_formats_cannot_cross_producer_grammars() {
        assert_ne!(Xf::new(7).get(), Ifmt::new(8).get());
        assert_eq!(Fault::from_raw(0x07), Some(Fault::DivZero));
        assert_eq!(Fault::from_raw(0x08), None);
    }
}
