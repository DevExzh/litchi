//! Worksheet features of an opened workbook (record 0327): legacy notes,
//! hyperlinks, drawing pictures and shapes, conditional formats, data
//! validations, the auto-filter and manual page breaks.
//!
//! [`Package::sheet_features`] reads them from one worksheet substream
//! without decoding its cells: the records after the cell table, the
//! page-break records before it, the drawing group for pictures, and the
//! globals' `_FilterDatabase` name for the auto-filter range.
//! [`edit`](crate::sheet_features::edit) changes them at the record level: every record it does not edit
//! keeps its bytes.
//!
//! ```
//! use litchi_xls::Package;
//! use litchi_xls::sheet_features::{FeatureSet, edit::PageBreak};
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_sheet_features()?;
//! edit.set_row_breaks(0, &[PageBreak::new(9)])?;
//! let package = edit.commit()?.into_package();
//! let features = package.sheet_features(0, FeatureSet::PAGE_BREAKS)?;
//! assert_eq!(features.row_breaks[0].index, 9);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod drawing;
pub mod edit;
mod encode;
mod read;
mod stream;

use crate::Package;

pub use read::{
    AutoFilter, ColorScale, ConditionalFormat, DataBar, Drawing, DrawingKind, Features, Fill,
    IconSet, Line, Note, Rule, RuleType, Style, Threshold, Validation,
};

/// A zero-based inclusive cell rectangle of the BIFF8 grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Area {
    /// The first row.
    pub first_row: u16,
    /// The last row.
    pub last_row: u16,
    /// The first column.
    pub first_column: u16,
    /// The last column.
    pub last_column: u16,
}

impl Area {
    /// One cell.
    #[must_use]
    pub const fn cell(row: u16, column: u16) -> Self {
        Self {
            first_row: row,
            last_row: row,
            first_column: column,
            last_column: column,
        }
    }

    /// A rectangle; the corners are ordered.
    #[must_use]
    pub fn new(first_row: u16, first_column: u16, last_row: u16, last_column: u16) -> Self {
        Self {
            first_row: first_row.min(last_row),
            last_row: first_row.max(last_row),
            first_column: first_column.min(last_column),
            last_column: first_column.max(last_column),
        }
    }

    /// Whether two rectangles share a cell.
    #[must_use]
    pub const fn overlaps(self, other: Self) -> bool {
        self.first_row <= other.last_row
            && other.first_row <= self.last_row
            && self.first_column <= other.last_column
            && other.first_column <= self.last_column
    }
}

/// A manual page break: the row (or column) it precedes and the column (or
/// row) span it covers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageBreak {
    /// The zero-based row (column) the break precedes.
    pub index: u16,
    /// The first column (row) of the break.
    pub min: u16,
    /// The last column (row) of the break.
    pub max: u16,
}

impl PageBreak {
    /// A break across the whole sheet before `index` (`max` is filled in
    /// for the axis when the break is written).
    #[must_use]
    pub const fn new(index: u16) -> Self {
        Self {
            index,
            min: 0,
            max: 0,
        }
    }
}

/// Which feature kinds a read decodes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FeatureSet(u16);

impl FeatureSet {
    /// Legacy notes.
    pub const NOTES: Self = Self(1);
    /// Hyperlinks.
    pub const HYPERLINKS: Self = Self(1 << 1);
    /// Drawing pictures and shapes.
    pub const DRAWINGS: Self = Self(1 << 2);
    /// Conditional formats.
    pub const FORMATS: Self = Self(1 << 3);
    /// Data validations.
    pub const VALIDATIONS: Self = Self(1 << 4);
    /// The auto-filter.
    pub const AUTO_FILTER: Self = Self(1 << 5);
    /// Manual page breaks.
    pub const PAGE_BREAKS: Self = Self(1 << 6);
    /// Every kind.
    pub const ALL: Self = Self(0x7f);

    /// Whether every kind of `other` is selected.
    #[must_use]
    pub const fn contains(self, other: Self) -> bool {
        self.0 & other.0 == other.0
    }

    /// The union of two selections.
    #[must_use]
    pub const fn union(self, other: Self) -> Self {
        Self(self.0 | other.0)
    }
}

impl std::ops::BitOr for FeatureSet {
    type Output = Self;

    fn bitor(self, other: Self) -> Self {
        self.union(other)
    }
}

impl Package {
    /// The features of the worksheet at catalog position `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is not an ordinary worksheet or
    /// the records the selection reads are malformed.
    pub fn sheet_features(
        &self,
        sheet: usize,
        select: FeatureSet,
    ) -> crate::package::Result<Features> {
        let stream = self.workbook_stream()?;
        let limits = self.raw_record_limits()?;
        read::features(&stream, sheet, select, &limits).map_err(crate::package::Error::Reader)
    }

    /// The workbook palette (`Palette`, or the default one), the colors a
    /// conditional format's palette indexes name.
    ///
    /// # Errors
    ///
    /// Returns an error when the Workbook stream cannot be read.
    pub fn palette(&self) -> crate::package::Result<crate::Palette> {
        let stream = self.workbook_stream()?;
        let globals = stream::Globals::parse(&stream).map_err(crate::package::Error::Reader)?;
        Ok(read::palette(&stream, &globals))
    }

    /// Start a worksheet-feature edit over this package.
    ///
    /// # Errors
    ///
    /// Returns [`edit::Error::Package`] for an encrypted workbook. Workbook
    /// stream errors are returned by the first feature verb or commit after
    /// deferred materialization.
    pub fn edit_sheet_features(&self) -> edit::Result<edit::Edit> {
        edit::Edit::new(self)
    }

    /// Apply a worksheet-feature patch to the exact package it was made
    /// from.
    ///
    /// # Errors
    ///
    /// Returns [`edit::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_sheet_features(&self, patch: &edit::Patch) -> edit::Result<edit::Commit> {
        edit::apply(self, patch)
    }
}
