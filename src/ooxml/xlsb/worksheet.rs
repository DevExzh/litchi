//! Worksheet implementation for XLSB files

use crate::ooxml::xlsb::cell::XlsbCell;
use crate::ooxml::xlsb::comments::Comment;
use crate::ooxml::xlsb::hyperlinks::Hyperlink;
use crate::ooxml::xlsb::merged_cells::MergedCell;
use crate::sheet::{Cell as SheetCell, CellIterator, CellValue, Result, RowIterator, Worksheet};
use std::borrow::Cow;
use std::collections::BTreeMap;

/// XLSB worksheet implementation
#[derive(Debug, Clone)]
pub struct XlsbWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), XlsbCell>,
    max_row: u32,
    max_col: u32,
    merged_cells: Vec<MergedCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
}

impl XlsbWorksheet {
    /// Create a new worksheet
    pub fn new(name: String) -> Self {
        XlsbWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
        }
    }

    /// Add a cell to the worksheet
    pub fn add_cell(&mut self, cell: XlsbCell) {
        let pos = (cell.row(), cell.column());
        self.max_row = self.max_row.max(cell.row());
        self.max_col = self.max_col.max(cell.column());
        self.cells.insert(pos, cell);
    }

    /// Get cell at position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&XlsbCell> {
        self.cells.get(&(row, col))
    }

    /// Add a merged cell range
    pub fn add_merged_cell(&mut self, merged: MergedCell) {
        self.merged_cells.push(merged);
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    /// Add a comment
    pub fn add_comment(&mut self, comment: Comment) {
        self.comments.push(comment);
    }

    /// Get all merged cells
    pub fn merged_cells(&self) -> &[MergedCell] {
        &self.merged_cells
    }

    /// Get all hyperlinks
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// Get all comments
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }
}

impl Worksheet for XlsbWorksheet {
    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        (self.max_row + 1) as usize
    }

    fn column_count(&self) -> usize {
        (self.max_col + 1) as usize
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            None
        } else {
            Some((0, 0, self.max_row, self.max_col))
        }
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Box::new(cell.clone())),
            None => {
                // Return empty cell for missing positions
                let empty_cell = XlsbCell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            },
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        let (row, col) = crate::ooxml::xlsb::utils::parse_cell_reference(coordinate)
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error + Send + Sync>)?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(XlsbCellIterator {
            cells: self.cells.values().collect(),
            index: 0,
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(XlsbRowIterator {
            worksheet: self,
            current_row: 0,
        })
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        let row_idx = row_idx as u32;
        let mut row_data = Vec::new();

        for col in 0..=self.max_col {
            match self.cells.get(&(row_idx, col)) {
                Some(cell) => row_data.push(cell.value().clone()),
                None => row_data.push(CellValue::Empty),
            }
        }

        Ok(Cow::Owned(row_data))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Cow::Borrowed(cell.value())),
            None => Ok(Cow::Borrowed(CellValue::EMPTY)),
        }
    }
}

/// Cell iterator for XLSB worksheets
struct XlsbCellIterator<'a> {
    cells: Vec<&'a XlsbCell>,
    index: usize,
}

impl<'a> CellIterator<'a> for XlsbCellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetCell + 'a>>> {
        if self.index >= self.cells.len() {
            None
        } else {
            let cell = self.cells[self.index];
            self.index += 1;
            Some(Ok(Box::new(cell.clone())))
        }
    }
}

/// Row iterator for XLSB worksheets
struct XlsbRowIterator<'a> {
    worksheet: &'a XlsbWorksheet,
    current_row: usize,
}

impl<'a> RowIterator<'a> for XlsbRowIterator<'a> {
    fn next(&mut self) -> Option<Result<Cow<'a, [CellValue]>>> {
        if self.current_row >= self.worksheet.row_count() {
            None
        } else {
            let result = self.worksheet.row(self.current_row);
            self.current_row += 1;
            Some(result)
        }
    }
}
