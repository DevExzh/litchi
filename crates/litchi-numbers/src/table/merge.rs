//! Compact, archive-free merged-cell geometry and topology algebra.
//!
//! A merge is a checked rectangular view over sparse table coordinates. The
//! native IWA adapter owns formula storage, package transactions, and bounds
//! validation; this module owns only the value and the pure coordinate
//! transformations shared by Numbers, Pages, and Keynote.

use super::coordinate::CellPosition;
use std::fmt;
use std::num::NonZeroU32;

/// The table axis affected by a topology operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Axis {
    /// The row coordinate or row span.
    Row,
    /// The column coordinate or column span.
    Column,
}

impl fmt::Display for Axis {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Row => "row",
            Self::Column => "column",
        })
    }
}

/// Failure while constructing or transforming a merged-cell region.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Error {
    /// The row span is zero.
    ZeroRows,
    /// The column span is zero.
    ZeroColumns,
    /// A merge must cover at least two cells.
    SingleCell,
    /// The inclusive end coordinate does not fit in the compact domain.
    CoordinateOverflow { axis: Axis },
    /// A topology insertion cannot enlarge the compact span.
    SpanOverflow { axis: Axis },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ZeroRows => formatter.write_str("merged-cell region has no rows"),
            Self::ZeroColumns => formatter.write_str("merged-cell region has no columns"),
            Self::SingleCell => {
                formatter.write_str("a merged-cell region must contain at least two cells")
            },
            Self::CoordinateOverflow { axis } => {
                write!(formatter, "merged-cell {axis} coordinate overflows u32")
            },
            Self::SpanOverflow { axis } => {
                write!(formatter, "merged-cell {axis} span overflows u32")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for checked merged-cell geometry.
pub type Result<T> = std::result::Result<T, Error>;

/// A compact rectangular merged-cell region containing at least two cells.
///
/// The representation is 16 bytes on supported targets: two `u32` start
/// coordinates and two niche-bearing non-zero `u32` spans. Coordinates are
/// zero-based and the end accessors are inclusive.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Region {
    start: CellPosition,
    row_count: NonZeroU32,
    column_count: NonZeroU32,
}

impl Region {
    /// Creates a checked merged-cell rectangle.
    ///
    /// # Errors
    ///
    /// Rejects zero spans, a one-cell region, and inclusive end coordinates
    /// outside the compact `u32` coordinate domain.
    pub fn new(row: u32, column: u32, row_count: u32, column_count: u32) -> Result<Self> {
        let row_span = NonZeroU32::new(row_count).ok_or(Error::ZeroRows)?;
        let column_span = NonZeroU32::new(column_count).ok_or(Error::ZeroColumns)?;
        if row_span.get() == 1 && column_span.get() == 1 {
            return Err(Error::SingleCell);
        }
        row.checked_add(row_span.get() - 1)
            .ok_or(Error::CoordinateOverflow { axis: Axis::Row })?;
        column
            .checked_add(column_span.get() - 1)
            .ok_or(Error::CoordinateOverflow { axis: Axis::Column })?;
        Ok(Self {
            start: CellPosition::new(row, column),
            row_count: row_span,
            column_count: column_span,
        })
    }

    /// Returns the inclusive start coordinate.
    #[must_use]
    pub const fn start(self) -> CellPosition {
        self.start
    }

    /// Returns the inclusive end coordinate.
    #[must_use]
    pub const fn end(self) -> CellPosition {
        CellPosition::new(self.end_row(), self.end_column())
    }

    /// Returns the zero-based first row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.start.row()
    }

    /// Returns the zero-based first column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.start.column()
    }

    /// Returns the number of rows covered by the region.
    #[must_use]
    pub const fn row_count(self) -> u32 {
        self.row_count.get()
    }

    /// Returns the number of columns covered by the region.
    #[must_use]
    pub const fn column_count(self) -> u32 {
        self.column_count.get()
    }

    /// Returns the inclusive last row.
    #[must_use]
    pub const fn end_row(self) -> u32 {
        self.row() + self.row_count() - 1
    }

    /// Returns the inclusive last column.
    #[must_use]
    pub const fn end_column(self) -> u32 {
        self.column() + self.column_count() - 1
    }

    /// Returns the covered area when it fits in `usize`.
    #[must_use]
    pub const fn area(self) -> Option<usize> {
        (self.row_count() as usize).checked_mul(self.column_count() as usize)
    }

    /// Returns whether a zero-based coordinate is inside the region.
    #[must_use]
    pub const fn contains(self, row: u32, column: u32) -> bool {
        row >= self.row()
            && row <= self.end_row()
            && column >= self.column()
            && column <= self.end_column()
    }

    /// Returns whether a compact coordinate is inside the region.
    #[must_use]
    pub const fn contains_position(self, position: CellPosition) -> bool {
        self.contains(position.row(), position.column())
    }

    /// Returns whether two closed rectangles overlap.
    #[must_use]
    pub const fn overlaps(self, other: Self) -> bool {
        self.row() <= other.end_row()
            && other.row() <= self.end_row()
            && self.column() <= other.end_column()
            && other.column() <= self.end_column()
    }
}

/// The result of removing one row or column from a merged region.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Deletion {
    /// The merge remains with the supplied compact geometry.
    Retain(Region),
    /// The remaining geometry is one cell, so native merge state disappears.
    Remove,
}

/// A merge anchor that must be copied to a surviving cell before an axis is
/// physically deleted.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct AnchorRelocation {
    source: CellPosition,
    destination: CellPosition,
}

impl AnchorRelocation {
    /// Creates an anchor relocation between two compact coordinates.
    #[must_use]
    pub const fn new(source: CellPosition, destination: CellPosition) -> Self {
        Self {
            source,
            destination,
        }
    }

    /// Returns the source cell whose value and references must be retained.
    #[must_use]
    pub const fn source(self) -> CellPosition {
        self.source
    }

    /// Returns the surviving destination cell.
    #[must_use]
    pub const fn destination(self) -> CellPosition {
        self.destination
    }
}

/// Applies native merged-cell behavior for one axis insertion.
///
/// # Errors
///
/// Returns a typed error when the insertion would overflow a compact
/// coordinate or span, or when the resulting rectangle would be invalid.
pub fn after_insertion(region: Region, axis: Axis, insertion: u32) -> Result<Region> {
    let (start, end, count) = match axis {
        Axis::Row => (region.row(), region.end_row(), region.row_count()),
        Axis::Column => (region.column(), region.end_column(), region.column_count()),
    };
    let (new_start, new_count) = if insertion <= start {
        (
            start
                .checked_add(1)
                .ok_or(Error::CoordinateOverflow { axis })?,
            count,
        )
    } else if insertion <= end {
        (
            start,
            count.checked_add(1).ok_or(Error::SpanOverflow { axis })?,
        )
    } else {
        (start, count)
    };

    match axis {
        Axis::Row => Region::new(new_start, region.column(), new_count, region.column_count()),
        Axis::Column => Region::new(region.row(), new_start, region.row_count(), new_count),
    }
}

/// Applies native merged-cell behavior for one axis deletion.
///
/// # Errors
///
/// Returns a typed error when the resulting rectangle cannot be represented in
/// the compact coordinate domain.
pub fn after_deletion(region: Region, axis: Axis, deletion: u32) -> Result<Deletion> {
    let (start, end, count, other_count) = match axis {
        Axis::Row => (
            region.row(),
            region.end_row(),
            region.row_count(),
            region.column_count(),
        ),
        Axis::Column => (
            region.column(),
            region.end_column(),
            region.column_count(),
            region.row_count(),
        ),
    };

    if deletion < start {
        let shifted_start = start - 1;
        return match axis {
            Axis::Row => Region::new(shifted_start, region.column(), count, other_count)
                .map(Deletion::Retain),
            Axis::Column => {
                Region::new(region.row(), shifted_start, other_count, count).map(Deletion::Retain)
            },
        };
    }
    if deletion > end {
        return Ok(Deletion::Retain(region));
    }

    if count == 1 {
        return Ok(Deletion::Remove);
    }
    let contracted_count = count - 1;
    if contracted_count == 1 && other_count == 1 {
        return Ok(Deletion::Remove);
    }
    match axis {
        Axis::Row => Region::new(region.row(), region.column(), contracted_count, other_count)
            .map(Deletion::Retain),
        Axis::Column => Region::new(region.row(), region.column(), other_count, contracted_count)
            .map(Deletion::Retain),
    }
}

/// Returns the anchor move required when deleting a leading merge boundary.
///
/// # Errors
///
/// Returns a typed error when the surviving anchor coordinate overflows the
/// compact domain.
pub fn anchor_relocation_after_deletion(
    region: Region,
    axis: Axis,
    deletion: u32,
) -> Result<Option<AnchorRelocation>> {
    let (start, count) = match axis {
        Axis::Row => (region.row(), region.row_count()),
        Axis::Column => (region.column(), region.column_count()),
    };
    if deletion != start || count == 1 {
        return Ok(None);
    }
    let next = start
        .checked_add(1)
        .ok_or(Error::CoordinateOverflow { axis })?;
    let source = CellPosition::new(region.row(), region.column());
    let destination = match axis {
        Axis::Row => CellPosition::new(next, region.column()),
        Axis::Column => CellPosition::new(region.row(), next),
    };
    Ok(Some(AnchorRelocation::new(source, destination)))
}

#[cfg(test)]
mod tests {
    use super::{
        AnchorRelocation, Axis, Deletion, Error, Region, after_deletion, after_insertion,
        anchor_relocation_after_deletion,
    };
    use crate::CellPosition;
    use std::mem::size_of;

    #[test]
    fn region_is_compact_and_validates_geometry() {
        assert_eq!(size_of::<Region>(), 16);
        assert!(matches!(Region::new(0, 0, 0, 2), Err(Error::ZeroRows)));
        assert!(matches!(Region::new(0, 0, 1, 0), Err(Error::ZeroColumns)));
        assert!(matches!(Region::new(0, 0, 1, 1), Err(Error::SingleCell)));
        assert!(matches!(
            Region::new(u32::MAX, 0, 2, 1),
            Err(Error::CoordinateOverflow { axis: Axis::Row })
        ));

        let region = Region::new(2, 3, 2, 3).unwrap();
        assert_eq!((region.end_row(), region.end_column()), (3, 5));
        assert!(region.contains(3, 5));
        assert!(region.contains_position(CellPosition::new(2, 3)));
        assert!(region.overlaps(Region::new(3, 5, 2, 1).unwrap()));
        assert!(!region.overlaps(Region::new(4, 3, 1, 2).unwrap()));
    }

    #[test]
    fn insertion_moves_or_expands_at_native_boundaries() {
        let region = Region::new(3, 4, 2, 3).unwrap();
        assert_eq!(
            after_insertion(region, Axis::Row, 3).unwrap(),
            Region::new(4, 4, 2, 3).unwrap()
        );
        assert_eq!(
            after_insertion(region, Axis::Row, 4).unwrap(),
            Region::new(3, 4, 3, 3).unwrap()
        );
        assert_eq!(after_insertion(region, Axis::Row, 5).unwrap(), region);
        assert_eq!(
            after_insertion(region, Axis::Column, 5).unwrap(),
            Region::new(3, 4, 2, 4).unwrap()
        );
    }

    #[test]
    fn deletion_shifts_contracts_and_removes_at_native_boundaries() {
        let region = Region::new(3, 4, 2, 3).unwrap();
        assert_eq!(
            after_deletion(region, Axis::Row, 2).unwrap(),
            Deletion::Retain(Region::new(2, 4, 2, 3).unwrap())
        );
        assert_eq!(
            after_deletion(region, Axis::Row, 3).unwrap(),
            Deletion::Retain(Region::new(3, 4, 1, 3).unwrap())
        );
        assert_eq!(
            after_deletion(region, Axis::Column, 5).unwrap(),
            Deletion::Retain(Region::new(3, 4, 2, 2).unwrap())
        );
        assert_eq!(
            after_deletion(region, Axis::Column, 7).unwrap(),
            Deletion::Retain(region)
        );
        assert_eq!(
            after_deletion(Region::new(1, 1, 1, 2).unwrap(), Axis::Column, 1).unwrap(),
            Deletion::Remove
        );
        assert_eq!(
            after_deletion(Region::new(1, 1, 1, 2).unwrap(), Axis::Row, 1).unwrap(),
            Deletion::Remove
        );
    }

    #[test]
    fn deletion_relocates_only_surviving_leading_anchors() {
        let region = Region::new(3, 4, 2, 3).unwrap();
        assert_eq!(
            anchor_relocation_after_deletion(region, Axis::Row, 3).unwrap(),
            Some(AnchorRelocation::new(
                CellPosition::new(3, 4),
                CellPosition::new(4, 4),
            ))
        );
        assert_eq!(
            anchor_relocation_after_deletion(region, Axis::Column, 4).unwrap(),
            Some(AnchorRelocation::new(
                CellPosition::new(3, 4),
                CellPosition::new(3, 5),
            ))
        );
        assert_eq!(
            anchor_relocation_after_deletion(region, Axis::Row, 4).unwrap(),
            None
        );
        assert_eq!(
            anchor_relocation_after_deletion(Region::new(3, 4, 1, 3).unwrap(), Axis::Row, 3)
                .unwrap(),
            None
        );
    }
}
