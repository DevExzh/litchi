//! Checked geometry for formula values.
//!
//! Coordinates are zero-based and every axis is represented by a half-open
//! interval.  This leaf deliberately knows nothing about a workbook, a
//! provider, or the storage used for the value payload.  The value VM maps
//! the `Option` results to its own limits and evaluation errors.

/// A zero-based, nonempty half-open box in sheet, row, and column space.
///
/// The bounds are `[start, end)` on each axis.  A formula array is normally
/// two-dimensional, so its sheet extent is one; keeping the sheet axis here
/// lets references and future matrix results retain their complete location
/// without introducing a second shape type.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub(super) struct Cuboid {
    sheet_start: usize,
    sheet_end: usize,
    row_start: usize,
    row_end: usize,
    column_start: usize,
    column_end: usize,
}

/// A zero-based sheet, row, and column coordinate.
pub(super) type Point = [usize; 3];

impl Cuboid {
    /// Construct a nonempty box from grouped starts and ends.
    ///
    /// The argument order is `(sheet_start, row_start, column_start,
    /// sheet_end, row_end, column_end)`.  Equal or reversed bounds are
    /// rejected, and no coordinate arithmetic is performed by this
    /// constructor.
    #[must_use]
    pub(super) const fn try_new(
        sheet_start: usize,
        row_start: usize,
        column_start: usize,
        sheet_end: usize,
        row_end: usize,
        column_end: usize,
    ) -> Option<Self> {
        if sheet_start >= sheet_end || row_start >= row_end || column_start >= column_end {
            return None;
        }
        Some(Self {
            sheet_start,
            sheet_end,
            row_start,
            row_end,
            column_start,
            column_end,
        })
    }

    /// Construct a nonempty box from `[sheet, row, column]` bounds.
    #[must_use]
    pub(super) const fn from_bounds(starts: Point, ends: Point) -> Option<Self> {
        Self::try_new(starts[0], starts[1], starts[2], ends[0], ends[1], ends[2])
    }

    /// Construct a nonempty box from an origin and checked axis extents.
    #[must_use]
    pub(super) const fn from_origin_extents(origin: Point, extents: Point) -> Option<Self> {
        let sheet_end = match origin[0].checked_add(extents[0]) {
            Some(value) => value,
            None => return None,
        };
        let row_end = match origin[1].checked_add(extents[1]) {
            Some(value) => value,
            None => return None,
        };
        let column_end = match origin[2].checked_add(extents[2]) {
            Some(value) => value,
            None => return None,
        };
        Self::from_bounds(origin, [sheet_end, row_end, column_end])
    }

    /// Return the inclusive starts as `[sheet, row, column]`.
    #[must_use]
    pub(super) const fn starts(self) -> Point {
        [self.sheet_start, self.row_start, self.column_start]
    }

    /// Return the exclusive ends as `[sheet, row, column]`.
    #[must_use]
    pub(super) const fn ends(self) -> Point {
        [self.sheet_end, self.row_end, self.column_end]
    }

    /// Return the checked axis extents as `[sheets, rows, columns]`.
    #[must_use]
    pub(super) const fn extent(self) -> Point {
        [
            self.sheet_end - self.sheet_start,
            self.row_end - self.row_start,
            self.column_end - self.column_start,
        ]
    }

    #[must_use]
    pub(super) const fn sheet_extent(self) -> usize {
        self.sheet_end - self.sheet_start
    }

    #[must_use]
    pub(super) const fn row_extent(self) -> usize {
        self.row_end - self.row_start
    }

    #[must_use]
    pub(super) const fn column_extent(self) -> usize {
        self.column_end - self.column_start
    }

    /// Return the checked number of cells, or `None` if the product overflows.
    #[must_use]
    pub(super) const fn count(self) -> Option<usize> {
        let sheets_times_rows = match self.sheet_extent().checked_mul(self.row_extent()) {
            Some(value) => value,
            None => return None,
        };
        sheets_times_rows.checked_mul(self.column_extent())
    }

    /// Return the smallest cuboid containing both boxes.
    #[must_use]
    pub(super) const fn bounding_cuboid(self, other: Self) -> Self {
        Self {
            sheet_start: min(self.sheet_start, other.sheet_start),
            sheet_end: max(self.sheet_end, other.sheet_end),
            row_start: min(self.row_start, other.row_start),
            row_end: max(self.row_end, other.row_end),
            column_start: min(self.column_start, other.column_start),
            column_end: max(self.column_end, other.column_end),
        }
    }

    /// Return the nonempty intersection, if the boxes overlap on every axis.
    #[must_use]
    pub(super) const fn intersection(self, other: Self) -> Option<Self> {
        Self::try_new(
            max(self.sheet_start, other.sheet_start),
            max(self.row_start, other.row_start),
            max(self.column_start, other.column_start),
            min(self.sheet_end, other.sheet_end),
            min(self.row_end, other.row_end),
            min(self.column_end, other.column_end),
        )
    }

    /// Return whether a point lies inside all three half-open intervals.
    #[must_use]
    pub(super) const fn contains(self, sheet: usize, row: usize, column: usize) -> bool {
        self.sheet_start <= sheet
            && sheet < self.sheet_end
            && self.row_start <= row
            && row < self.row_end
            && self.column_start <= column
            && column < self.column_end
    }

    /// Return whether a `[sheet, row, column]` point is contained.
    #[must_use]
    pub(super) const fn contains_point(self, point: Point) -> bool {
        self.contains(point[0], point[1], point[2])
    }

    /// Project a 2D singleton across every row and column of `output`.
    ///
    /// `self` is the source shape and `point` is an absolute output point.
    /// Both shapes must describe one sheet, and the point must be inside the
    /// output.  The source point is always the singleton's only cell.
    #[must_use]
    pub(super) const fn project_2d_singleton(self, output: Self, point: Point) -> Option<Point> {
        if !self.is_2d_singleton()
            || !self.same_sheet_extent(output)
            || !output.contains_point(point)
        {
            return None;
        }
        Some([self.sheet_start, self.row_start, self.column_start])
    }

    /// Project a one-row 2D value across the rows of `output`.
    ///
    /// The source column extent is projected relative to the output's column
    /// origin. An output point outside `output` or beyond that source extent
    /// returns `None`; its row is mapped to the source's sole row. This
    /// permits a shorter row vector to contribute its prefix while the VM
    /// represents remaining output positions as out-of-shape values.
    #[must_use]
    pub(super) const fn project_2d_row(self, output: Self, point: Point) -> Option<Point> {
        if self.row_extent() != 1
            || !self.same_sheet_extent(output)
            || !output.contains_point(point)
        {
            return None;
        }
        let column_offset = point[2] - output.column_start;
        if column_offset >= self.column_extent() {
            return None;
        }
        Some([
            self.sheet_start,
            self.row_start,
            self.column_start + column_offset,
        ])
    }

    /// Project a one-column 2D value across the columns of `output`.
    ///
    /// The source row extent is projected relative to the output's row
    /// origin. An output point outside `output` or beyond that source extent
    /// returns `None`; its column is mapped to the source's sole column.
    #[must_use]
    pub(super) const fn project_2d_column(self, output: Self, point: Point) -> Option<Point> {
        if self.column_extent() != 1
            || !self.same_sheet_extent(output)
            || !output.contains_point(point)
        {
            return None;
        }
        let row_offset = point[1] - output.row_start;
        if row_offset >= self.row_extent() {
            return None;
        }
        Some([
            self.sheet_start,
            self.row_start + row_offset,
            self.column_start,
        ])
    }

    /// Project a 2D matrix at an output point.
    ///
    /// A matrix does not broadcast along either axis. The output may be
    /// larger than the source (the VM can then represent an out-of-shape
    /// `#N/A`), but only output points whose row and column offsets fit the
    /// source project to a source cell.
    #[must_use]
    pub(super) const fn project_2d_matrix(self, output: Self, point: Point) -> Option<Point> {
        if !self.is_2d() || !self.same_sheet_extent(output) || !output.contains_point(point) {
            return None;
        }
        let row_offset = point[1] - output.row_start;
        let column_offset = point[2] - output.column_start;
        if row_offset >= self.row_extent() || column_offset >= self.column_extent() {
            None
        } else {
            Some([
                self.sheet_start,
                self.row_start + row_offset,
                self.column_start + column_offset,
            ])
        }
    }

    const fn is_2d(self) -> bool {
        self.sheet_extent() == 1
    }

    const fn is_2d_singleton(self) -> bool {
        self.is_2d() && self.row_extent() == 1 && self.column_extent() == 1
    }

    const fn same_sheet_extent(self, other: Self) -> bool {
        self.is_2d()
            && other.is_2d()
            && self.sheet_start == other.sheet_start
            && self.sheet_end == other.sheet_end
    }
}

const fn min(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

const fn max(left: usize, right: usize) -> usize {
    if left > right { left } else { right }
}

#[cfg(test)]
mod tests {
    use super::{Cuboid, Point};

    const fn box_at(origin: Point, extents: Point) -> Cuboid {
        match Cuboid::from_origin_extents(origin, extents) {
            Some(value) => value,
            None => panic!("test box must be nonempty and in range"),
        }
    }

    #[test]
    fn constructor_rejects_empty_and_checked_extent_overflow() {
        assert!(Cuboid::try_new(0, 0, 0, 0, 1, 1).is_none());
        assert!(Cuboid::try_new(0, 1, 1, 1, 1, 1).is_none());
        assert!(Cuboid::from_bounds([usize::MAX, 0, 0], [usize::MAX, 1, 1]).is_none());
        assert!(Cuboid::from_origin_extents([usize::MAX, 0, 0], [1, 1, 1]).is_none());

        let huge = box_at([0, 0, 0], [usize::MAX, usize::MAX, 2]);
        assert_eq!(huge.extent(), [usize::MAX, usize::MAX, 2]);
        assert_eq!(huge.count(), None);
        assert_eq!(box_at([3, 4, 5], [2, 3, 4]).count(), Some(24));
    }

    #[test]
    fn intersection_is_half_open_and_bounding_spans_gaps() {
        let left = box_at([1, 2, 3], [3, 5, 6]);
        let right = box_at([2, 4, 5], [3, 3, 4]);
        assert_eq!(
            left.intersection(right).map(Cuboid::starts),
            Some([2, 4, 5])
        );
        assert_eq!(left.intersection(right).map(Cuboid::ends), Some([4, 7, 9]));
        assert!(left.intersection(box_at([4, 2, 3], [1, 1, 1])).is_none());
        assert!(left.intersection(box_at([1, 7, 3], [1, 1, 1])).is_none());

        let disjoint_gap = box_at([8, 20, 30], [1, 1, 1]);
        let bounding = left.bounding_cuboid(disjoint_gap);
        assert_eq!(bounding.starts(), [1, 2, 3]);
        assert_eq!(bounding.ends(), [9, 21, 31]);
        assert!(left.contains(1, 2, 3));
        assert!(left.contains(3, 6, 8));
        assert!(!left.contains(4, 2, 3));
        assert!(!left.contains(1, 7, 3));
        assert!(!left.contains(1, 2, 9));
    }

    #[test]
    fn two_dimensional_projection_broadcasts_only_its_named_axis() {
        let singleton = box_at([2, 10, 20], [1, 1, 1]);
        let singleton_output = box_at([2, 0, 0], [1, 3, 4]);
        assert_eq!(
            singleton.project_2d_singleton(singleton_output, [2, 2, 3]),
            Some([2, 10, 20])
        );
        assert!(
            singleton
                .project_2d_singleton(singleton_output, [2, 3, 0])
                .is_none()
        );

        let row = box_at([2, 10, 20], [1, 1, 4]);
        let row_output = box_at([2, 10, 20], [1, 3, 4]);
        assert_eq!(
            row.project_2d_row(row_output, [2, 12, 23]),
            Some([2, 10, 23])
        );
        assert_eq!(
            row.project_2d_row(box_at([2, 10, 20], [1, 3, 5]), [2, 12, 23]),
            Some([2, 10, 23])
        );
        assert!(
            row.project_2d_row(box_at([2, 10, 20], [1, 3, 5]), [2, 12, 24])
                .is_none()
        );
        assert!(row.project_2d_row(row_output, [2, 13, 20]).is_none());
        assert_eq!(
            row.project_2d_row(box_at([2, 0, 0], [1, 3, 5]), [2, 2, 2]),
            Some([2, 10, 22])
        );

        let column = box_at([2, 10, 20], [1, 3, 1]);
        let column_output = box_at([2, 10, 20], [1, 3, 4]);
        assert_eq!(
            column.project_2d_column(column_output, [2, 12, 23]),
            Some([2, 12, 20])
        );
        assert_eq!(
            column.project_2d_column(box_at([2, 10, 20], [1, 4, 4]), [2, 12, 23]),
            Some([2, 12, 20])
        );
        assert!(
            column
                .project_2d_column(box_at([2, 10, 20], [1, 4, 4]), [2, 13, 23])
                .is_none()
        );
        assert!(
            column
                .project_2d_column(column_output, [2, 10, 24])
                .is_none()
        );
        assert_eq!(
            column.project_2d_column(box_at([2, 0, 0], [1, 4, 4]), [2, 2, 2]),
            Some([2, 12, 20])
        );

        let matrix = box_at([2, 10, 20], [1, 2, 3]);
        let matrix_output = box_at([2, 10, 20], [1, 4, 4]);
        assert_eq!(
            matrix.project_2d_matrix(matrix_output, [2, 11, 22]),
            Some([2, 11, 22])
        );
        assert!(
            matrix
                .project_2d_matrix(matrix_output, [2, 13, 23])
                .is_none()
        );
        assert!(
            matrix
                .project_2d_matrix(matrix_output, [2, 14, 24])
                .is_none()
        );
        assert_eq!(
            matrix.project_2d_matrix(box_at([2, 0, 0], [1, 4, 4]), [2, 1, 2]),
            Some([2, 11, 22])
        );
    }
}
