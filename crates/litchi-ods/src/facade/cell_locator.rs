//! Bounded lazy logical-cell lookup for the immutable spreadsheet facade.

use crate::worksheet::{Cell, CellView, Sheet};
use std::mem::size_of;

pub(super) const BUILD_QUERY_THRESHOLD: usize = 64;
pub(super) const MAX_LOCATOR_BYTES: usize = 4 * 1024 * 1024;

/// A sheet-aligned index over physical row and cell runs.
///
/// Empty endpoint arrays encode the common direct case where every repeat is
/// one. The index therefore retains only one compact row descriptor per
/// physical row for ordinary dense sheets.
pub(super) struct CellLocator {
    sheets: Vec<SheetLocator>,
}

struct SheetLocator {
    row_ends: Vec<u32>,
    rows: Vec<RowLocator>,
    repeated_cell_ends: Vec<u32>,
}

#[derive(Clone, Copy)]
struct RowLocator {
    cell_ends_start: u32,
    cell_ends_len: u32,
    direct_cell_count: u32,
}

impl CellLocator {
    pub(super) fn try_build(sheets: &[Sheet]) -> Option<Self> {
        Self::try_build_with_budget(sheets, MAX_LOCATOR_BYTES)
    }

    pub(super) fn try_build_with_budget(sheets: &[Sheet], byte_budget: usize) -> Option<Self> {
        let required_bytes = requested_bytes(sheets)?;
        if required_bytes > byte_budget {
            return None;
        }

        let mut indexed_sheets = Vec::new();
        indexed_sheets.try_reserve_exact(sheets.len()).ok()?;
        for sheet in sheets {
            indexed_sheets.push(SheetLocator::try_build(sheet)?);
        }
        Some(Self {
            sheets: indexed_sheets,
        })
    }

    pub(super) fn cell_view<'a>(
        &self,
        sheets: &'a [Sheet],
        sheet_index: usize,
        row: usize,
        column: usize,
    ) -> CellView<'a> {
        let Some(locator) = self.sheets.get(sheet_index) else {
            return CellView::Missing;
        };
        locator
            .cell(sheets.get(sheet_index), row, column)
            .map_or(CellView::Missing, CellView::Stored)
    }
}

impl SheetLocator {
    fn try_build(sheet: &Sheet) -> Option<Self> {
        let repeated_rows = sheet.rows.iter().any(|row| row.repeat() != 1);
        let repeated_cell_count = sheet
            .rows
            .iter()
            .filter(|row| row.cells().iter().any(|cell| cell.repeat() != 1))
            .try_fold(0usize, |total, row| total.checked_add(row.cells().len()))?;

        let mut row_ends = Vec::new();
        if repeated_rows {
            row_ends.try_reserve_exact(sheet.rows.len()).ok()?;
        }
        let mut rows = Vec::new();
        rows.try_reserve_exact(sheet.rows.len()).ok()?;
        let mut repeated_cell_ends = Vec::new();
        repeated_cell_ends
            .try_reserve_exact(repeated_cell_count)
            .ok()?;

        let mut logical_row_end = 0usize;
        for row in &sheet.rows {
            logical_row_end = logical_row_end.checked_add(row.repeat())?;
            if repeated_rows {
                row_ends.push(u32::try_from(logical_row_end).ok()?);
            }

            let has_repeated_cells = row.cells().iter().any(|cell| cell.repeat() != 1);
            if has_repeated_cells {
                let start = u32::try_from(repeated_cell_ends.len()).ok()?;
                let mut logical_cell_end = 0usize;
                for cell in row.cells() {
                    logical_cell_end = logical_cell_end.checked_add(cell.repeat())?;
                    repeated_cell_ends.push(u32::try_from(logical_cell_end).ok()?);
                }
                rows.push(RowLocator {
                    cell_ends_start: start,
                    cell_ends_len: u32::try_from(row.cells().len()).ok()?,
                    direct_cell_count: 0,
                });
            } else {
                rows.push(RowLocator {
                    cell_ends_start: 0,
                    cell_ends_len: 0,
                    direct_cell_count: u32::try_from(row.cells().len()).ok()?,
                });
            }
        }

        Some(Self {
            row_ends,
            rows,
            repeated_cell_ends,
        })
    }

    fn cell<'a>(&self, sheet: Option<&'a Sheet>, row: usize, column: usize) -> Option<&'a Cell> {
        let sheet = sheet?;
        let physical_row = if self.row_ends.is_empty() {
            (row < self.rows.len()).then_some(row)?
        } else {
            let row = u32::try_from(row).ok()?;
            let position = self.row_ends.partition_point(|end| *end <= row);
            (position < self.rows.len()).then_some(position)?
        };
        let row_locator = self.rows.get(physical_row)?;
        let physical_cell = row_locator.physical_cell(&self.repeated_cell_ends, column)?;
        sheet.rows.get(physical_row)?.cells().get(physical_cell)
    }
}

impl RowLocator {
    fn physical_cell(&self, repeated_cell_ends: &[u32], column: usize) -> Option<usize> {
        if self.cell_ends_len == 0 {
            return (column < self.direct_cell_count as usize).then_some(column);
        }

        let column = u32::try_from(column).ok()?;
        let start = self.cell_ends_start as usize;
        let end = start.checked_add(self.cell_ends_len as usize)?;
        let ends = repeated_cell_ends.get(start..end)?;
        let position = ends.partition_point(|logical_end| *logical_end <= column);
        (position < ends.len()).then_some(position)
    }
}

fn requested_bytes(sheets: &[Sheet]) -> Option<usize> {
    let mut bytes = sheets.len().checked_mul(size_of::<SheetLocator>())?;
    for sheet in sheets {
        bytes = bytes.checked_add(sheet.rows.len().checked_mul(size_of::<RowLocator>())?)?;
        if sheet.rows.iter().any(|row| row.repeat() != 1) {
            bytes = bytes.checked_add(sheet.rows.len().checked_mul(size_of::<u32>())?)?;
        }
        let repeated_cells = sheet
            .rows
            .iter()
            .filter(|row| row.cells().iter().any(|cell| cell.repeat() != 1))
            .try_fold(0usize, |total, row| total.checked_add(row.cells().len()))?;
        bytes = bytes.checked_add(repeated_cells.checked_mul(size_of::<u32>())?)?;
    }
    Some(bytes)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::worksheet::{CellValue, Row};

    fn stored_text(view: CellView<'_>) -> Option<&str> {
        match view {
            CellView::Missing => None,
            CellView::Stored(cell) => Some(cell.text.as_str()),
        }
    }

    fn repeated_fixture() -> Vec<Sheet> {
        let mut repeated =
            Sheet::new("Repeated").expect("test fixture or operation should succeed");
        let mut first = Row::repeated(2).expect("test fixture or operation should succeed");
        first
            .push_cell(
                Cell::repeated(CellValue::Text("a".to_owned()), "a", 3)
                    .expect("test fixture or operation should succeed"),
            )
            .expect("test fixture or operation should succeed");
        first
            .push_cell(Cell::new(CellValue::Text("b".to_owned()), "b"))
            .expect("test fixture or operation should succeed");
        repeated
            .push_row(first)
            .expect("test fixture or operation should succeed");
        repeated
            .push_row(Row::new())
            .expect("test fixture or operation should succeed");
        let mut last = Row::repeated(3).expect("test fixture or operation should succeed");
        last.push_cell(Cell::new(CellValue::Text("c".to_owned()), "c"))
            .expect("test fixture or operation should succeed");
        repeated
            .push_row(last)
            .expect("test fixture or operation should succeed");

        let mut direct = Sheet::new("Direct").expect("test fixture or operation should succeed");
        for text in ["d", "e"] {
            let mut row = Row::new();
            row.push_cell(Cell::new(CellValue::Text(text.to_owned()), text))
                .expect("test fixture or operation should succeed");
            direct
                .push_row(row)
                .expect("test fixture or operation should succeed");
        }
        vec![repeated, direct]
    }

    #[test]
    fn indexed_lookup_matches_linear_runs_and_cell_identity() {
        let sheets = repeated_fixture();
        let locator = CellLocator::try_build_with_budget(&sheets, MAX_LOCATOR_BYTES)
            .expect("test fixture or operation should succeed");

        for (sheet_index, sheet) in sheets.iter().enumerate() {
            for row in 0..8 {
                for column in 0..7 {
                    let direct = sheet.cell_view(row, column);
                    let indexed = locator.cell_view(&sheets, sheet_index, row, column);
                    assert_eq!(stored_text(indexed), stored_text(direct));
                    if let (CellView::Stored(indexed), CellView::Stored(direct)) = (indexed, direct)
                    {
                        assert!(std::ptr::eq(indexed, direct));
                    }
                }
            }
        }
        assert_eq!(
            locator.cell_view(&sheets, 0, usize::MAX, usize::MAX),
            CellView::Missing
        );
        assert_eq!(
            locator.cell_view(&sheets, sheets.len(), 0, 0),
            CellView::Missing
        );
    }

    #[test]
    fn zero_budget_falls_back_without_logical_expansion() {
        let sheets = repeated_fixture();
        assert!(CellLocator::try_build_with_budget(&sheets, 0).is_none());

        let mut maximum = Sheet::new("Maximum").expect("test fixture or operation should succeed");
        let mut row = Row::repeated(1_048_576).expect("test fixture or operation should succeed");
        row.push_cell(
            Cell::repeated(CellValue::Empty, "", 1_048_576)
                .expect("test fixture or operation should succeed"),
        )
        .expect("test fixture or operation should succeed");
        maximum
            .push_row(row)
            .expect("test fixture or operation should succeed");
        let maximum = vec![maximum];
        let locator = CellLocator::try_build_with_budget(&maximum, MAX_LOCATOR_BYTES)
            .expect("test fixture or operation should succeed");
        assert!(matches!(
            locator.cell_view(&maximum, 0, 1_048_575, 1_048_575),
            CellView::Stored(_)
        ));
        assert_eq!(
            locator.cell_view(&maximum, 0, 1_048_576, 1_048_576),
            CellView::Missing
        );
    }
}
