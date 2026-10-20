//! Iterator implementations for text-based formats

use super::cell::TextCell;
use super::worksheet::TextWorksheet;
use crate::sheet::{CellValue, Result as SheetResult, WorkbookTrait};
use std::borrow::Cow;

/// Iterator over worksheets in a text workbook
pub struct TextWorksheetIterator<'a> {
    workbook: &'a super::workbook::TextWorkbook,
    yielded: bool,
}

impl<'a> TextWorksheetIterator<'a> {
    /// Create a new worksheet iterator
    pub fn new(workbook: &'a super::workbook::TextWorkbook) -> Self {
        TextWorksheetIterator {
            workbook,
            yielded: false,
        }
    }
}

impl Iterator for TextWorksheetIterator<'_> {
    type Item = SheetResult<TextWorksheet>;

    fn next(&mut self) -> Option<Self::Item> {
        if !self.yielded {
            self.yielded = true;
            Some(self.workbook.active_worksheet())
        } else {
            None
        }
    }
}

/// Iterator over cells in a text worksheet
pub struct TextCellIterator<'a> {
    data: &'a Vec<Vec<CellValue>>,
    current_row: usize,
    current_col: usize,
}

impl<'a> TextCellIterator<'a> {
    /// Create a new cell iterator
    pub fn new(data: &'a Vec<Vec<CellValue>>) -> Self {
        TextCellIterator {
            data,
            current_row: 0,
            current_col: 0,
        }
    }
}

impl Iterator for TextCellIterator<'_> {
    type Item = SheetResult<TextCell>;

    fn next(&mut self) -> Option<Self::Item> {
        // Find the next cell
        while self.current_row < self.data.len() {
            let row = &self.data[self.current_row];
            if self.current_col < row.len() {
                let row_num = (self.current_row + 1) as u32;
                let col_num = (self.current_col + 1) as u32;
                let value = row[self.current_col].clone();

                let cell = TextCell::new(row_num, col_num, value);
                self.current_col += 1;

                return Some(Ok(cell));
            } else {
                // Move to next row
                self.current_row += 1;
                self.current_col = 0;
            }
        }

        None
    }
}

/// Iterator over rows in a text worksheet
pub struct TextRowIterator<'a> {
    data: &'a Vec<Vec<CellValue>>,
    current_row: usize,
}

impl<'a> TextRowIterator<'a> {
    /// Create a new row iterator
    pub fn new(data: &'a Vec<Vec<CellValue>>) -> Self {
        TextRowIterator {
            data,
            current_row: 0,
        }
    }
}

impl<'a> Iterator for TextRowIterator<'a> {
    type Item = SheetResult<Cow<'a, [CellValue]>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_row < self.data.len() {
            let row = Cow::Borrowed(self.data[self.current_row].as_slice()); // Zero-copy!
            self.current_row += 1;
            Some(Ok(row))
        } else {
            None
        }
    }
}
