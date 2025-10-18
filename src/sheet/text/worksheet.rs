//! Worksheet implementation for text-based formats

use std::io::{Read, Seek};
use crate::sheet::{Worksheet, Cell, CellIterator, RowIterator, CellValue, Result as SheetResult};
use super::cell::TextCell;
use super::iterators::{TextCellIterator, TextRowIterator};
use super::parser::TextParser;

/// Worksheet implementation for text-based formats
pub struct TextWorksheet {
    data: Vec<Vec<CellValue>>,
    name: String,
    dimensions: Option<(u32, u32, u32, u32)>, // (min_row, min_col, max_row, max_col)
}

impl TextWorksheet {
    /// Create a new text worksheet by loading all data into memory
    pub fn new<R: Read + Seek>(reader: &mut R, config: super::workbook::TextConfig, name: String) -> SheetResult<Self> {
        let mut parser = TextParser::new(reader, config);
        let mut data = Vec::new();

        while let Some(row_result) = parser.parse_row()? {
            data.push(row_result?);
        }

        Ok(Self::from_data(&data, name))
    }

    /// Create a text worksheet from parsed data
    pub fn from_data(data: &[Vec<CellValue>], name: String) -> Self {
        let dimensions = if data.is_empty() {
            None
        } else {
            let max_cols = data.iter().map(|row| row.len()).max().unwrap_or(0);
            Some((1, 1, data.len() as u32, max_cols as u32))
        };

        TextWorksheet {
            data: data.to_vec(),
            name,
            dimensions,
        }
    }

}

impl Worksheet for TextWorksheet {
    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        self.data.len()
    }

    fn column_count(&self) -> usize {
        self.data.iter().map(|row| row.len()).max().unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn Cell + '_>> {
        if row < 1 || column < 1 {
            return Ok(Box::new(TextCell::new(row, column, CellValue::Empty)));
        }

        let row_idx = (row - 1) as usize;
        let col_idx = (column - 1) as usize;

        if row_idx >= self.data.len() {
            return Ok(Box::new(TextCell::new(row, column, CellValue::Empty)));
        }

        let row_data = &self.data[row_idx];
        if col_idx >= row_data.len() {
            Ok(Box::new(TextCell::new(row, column, CellValue::Empty)))
        } else {
            let value = row_data[col_idx].clone();
            Ok(Box::new(TextCell::new(row, column, value)))
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn Cell + '_>> {
        // Parse coordinate like "A1", "B2", etc.
        let (col_str, row_str) = coordinate.split_at(
            coordinate.chars().position(|c| c.is_ascii_digit()).unwrap_or(coordinate.len())
        );

        if col_str.is_empty() || row_str.is_empty() {
            return Err(format!("Invalid coordinate: {}", coordinate).into());
        }

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        let mut col_num = 0u32;
        for c in col_str.chars() {
            if !c.is_ascii_alphabetic() {
                return Err(format!("Invalid coordinate: {}", coordinate).into());
            }
            col_num = col_num * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }

        let row_num = row_str.parse::<u32>()
            .map_err(|_| format!("Invalid row number in coordinate: {}", coordinate))?;

        self.cell(row_num, col_num)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(TextCellIterator::new(&self.data))
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(TextRowIterator::new(&self.data))
    }

    fn row(&self, row_idx: usize) -> SheetResult<Vec<CellValue>> {
        if row_idx >= self.data.len() {
            return Err(format!("Row {} not found", row_idx + 1).into());
        }
        Ok(self.data[row_idx].clone())
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        if row < 1 || column < 1 {
            return Ok(CellValue::Empty);
        }

        let row_idx = (row - 1) as usize;
        let col_idx = (column - 1) as usize;

        if row_idx >= self.data.len() {
            return Ok(CellValue::Empty);
        }

        let row_data = &self.data[row_idx];
        if col_idx >= row_data.len() {
            Ok(CellValue::Empty)
        } else {
            Ok(row_data[col_idx].clone())
        }
    }
}
