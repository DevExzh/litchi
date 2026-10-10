//! Checked, sparse reads from an eager XLSB worksheet.
//!
//! The legacy [`litchi_core::sheet::Worksheet`] implementation remains in
//! [`super`].  This module is the format-facing read path: addresses and
//! ranges are resolved through [`litchi_sheet::At`] and [`litchi_sheet::Area`]
//! before the retained cell map is touched.  A missing stored cell is
//! represented by `Ok(None)`; range traversal yields only stored cells and
//! never expands a sparse worksheet into a dense allocation.

use std::collections::{BTreeMap, btree_map};
use std::iter::FusedIterator;

use litchi_sheet::{Area, At, Rect};

use super::{Cell, Worksheet};
use crate::package::error::Result;

/// Borrowed stored cells selected by a checked half-open rectangle.
///
/// Cells are yielded in row-major order.  The iterator seeks to the first
/// selected column of each occupied row, and skips the suffix after the
/// selected end column, so a narrow range does not scan unrelated columns in
/// every occupied row.
pub struct Cells<'a> {
    cells: &'a BTreeMap<(u32, u32), Cell>,
    range: btree_map::Range<'a, (u32, u32), Cell>,
    first_column: u32,
    end_row: u32,
    end_column: u32,
}

impl std::fmt::Debug for Cells<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Cells")
            .field("first_column", &self.first_column)
            .field("end_row", &self.end_row)
            .field("end_column", &self.end_column)
            .finish_non_exhaustive()
    }
}

impl<'a> Cells<'a> {
    fn new(cells: &'a BTreeMap<(u32, u32), Cell>, rect: Rect) -> Self {
        let start = rect.start();
        let (end_row, end_column) = rect.end();
        let first_row = start.row().get();
        let first_column = start.column().get();
        Self {
            cells,
            range: cells.range((first_row, first_column)..(end_row, 0)),
            first_column,
            end_row,
            end_column,
        }
    }

    /// Seek to the selected columns of `row`, or exhaust the iterator after
    /// the exclusive row bound.
    fn seek_row(&mut self, row: u32) {
        let start = if row < self.end_row {
            (row, self.first_column)
        } else {
            (self.end_row, 0)
        };
        self.range = self.cells.range(start..(self.end_row, 0));
    }
}

impl<'a> Iterator for Cells<'a> {
    type Item = &'a Cell;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let (&(row, column), cell) = self.range.next()?;
            if column < self.first_column {
                self.seek_row(row);
            } else if column >= self.end_column {
                self.seek_row(row.saturating_add(1));
            } else {
                return Some(cell);
            }
        }
    }
}

impl FusedIterator for Cells<'_> {}

impl Worksheet {
    /// Borrow one stored cell by a checked zero-based coordinate or A1
    /// reference.
    ///
    /// The input is resolved before lookup.  An in-grid coordinate with no
    /// stored cell returns `Ok(None)`; it is never replaced with a synthetic
    /// empty cell.
    pub fn cell_at<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<&Cell>> {
        let address = at.into().resolve()?;
        Ok(self.get_cell(address.row().get(), address.column().get()))
    }

    /// Lazily borrow stored cells inside a checked zero-based half-open range.
    ///
    /// A1 references, raw half-open bounds, and reusable [`Rect`] values are
    /// accepted through [`Area`].  The iterator is sparse and row-major; it
    /// does not allocate a dense grid or synthesize missing cells.
    pub fn cells_in<'a>(&self, area: impl Into<Area<'a>>) -> Result<Cells<'_>> {
        let rect = area.into().resolve()?;
        Ok(Cells::new(&self.cells, rect))
    }
}

#[cfg(test)]
mod tests {
    use litchi_core::sheet::{Cell as _, CellValue};
    use litchi_sheet::{COLUMNS, ROWS};

    use crate::package::error::Error;

    use super::*;

    fn worksheet() -> Worksheet {
        let mut worksheet = Worksheet::new("Read test".to_owned());
        worksheet.add_cell(Cell::new(0, 0, CellValue::Int(1)));
        worksheet.add_cell(Cell::new(0, 3, CellValue::Int(4)));
        worksheet.add_cell(Cell::new(1, 0, CellValue::Int(10)));
        worksheet.add_cell(Cell::new(1, 1, CellValue::Int(11)));
        worksheet.add_cell(Cell::new(1, 2, CellValue::Int(12)));
        worksheet.add_cell(Cell::new(1, 5, CellValue::Int(15)));
        worksheet.add_cell(Cell::new(3, 2, CellValue::Int(32)));
        worksheet.add_cell(Cell::new(ROWS - 1, COLUMNS - 1, CellValue::Int(1_048_575)));
        worksheet
    }

    #[test]
    fn cell_at_is_checked_zero_based_and_sparse() {
        let worksheet = worksheet();

        assert_eq!(
            worksheet.cell_at((0, 0)).unwrap().unwrap().value(),
            &CellValue::Int(1)
        );
        assert_eq!(worksheet.cell_at("A1").unwrap().unwrap().row(), 0);
        assert!(worksheet.cell_at((0, 1)).unwrap().is_none());
        assert_eq!(
            worksheet.cell_at("XFD1048576").unwrap().unwrap().column(),
            COLUMNS - 1
        );
    }

    #[test]
    fn cell_at_rejects_invalid_references_without_fallback() {
        let worksheet = worksheet();

        assert!(matches!(
            worksheet.cell_at("XFE1"),
            Err(Error::Coordinate(_))
        ));
        assert!(matches!(
            worksheet.cell_at((ROWS, 0)),
            Err(Error::Coordinate(_))
        ));
    }

    #[test]
    fn cells_in_is_sparse_row_major_and_respects_both_bounds() {
        let worksheet = worksheet();

        let cells = worksheet
            .cells_in("B1:C4")
            .unwrap()
            .map(|cell| (cell.row(), cell.column(), cell.value().clone()))
            .collect::<Vec<_>>();
        assert_eq!(
            cells,
            [
                (1, 1, CellValue::Int(11)),
                (1, 2, CellValue::Int(12)),
                (3, 2, CellValue::Int(32)),
            ]
        );

        let only_last_column = worksheet
            .cells_in((1, 5, 2, 6))
            .unwrap()
            .map(|cell| (cell.row(), cell.column()))
            .collect::<Vec<_>>();
        assert_eq!(only_last_column, [(1, 5)]);
    }

    #[test]
    fn cells_in_supports_single_last_cell_and_full_grid() {
        let worksheet = worksheet();

        assert_eq!(
            worksheet
                .cells_in("XFD1048576")
                .unwrap()
                .map(|cell| (cell.row(), cell.column()))
                .collect::<Vec<_>>(),
            [(ROWS - 1, COLUMNS - 1)]
        );
        assert_eq!(worksheet.cells_in(Rect::ALL).unwrap().count(), 8);
    }

    #[test]
    fn sparse_iterator_debug_omits_authored_values() {
        let mut worksheet = Worksheet::new("private-sheet-name".to_owned());
        worksheet.add_cell(Cell::new(
            0,
            0,
            CellValue::String("private-cell-text".to_owned()),
        ));
        let debug = format!("{worksheet:?} {:?}", worksheet.cells_in(Rect::ALL).unwrap());
        assert!(!debug.contains("private-sheet-name"));
        assert!(!debug.contains("private-cell-text"));
    }

    #[test]
    fn cells_in_rejects_empty_or_invalid_a1_areas() {
        let worksheet = worksheet();

        for area in ["", "B2:A1", "A0", "XFE1"] {
            assert!(matches!(worksheet.cells_in(area), Err(Error::Range(_))));
        }
    }
}
