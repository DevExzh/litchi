//! Cell implementation for Excel worksheets.
//!
//! This module provides the concrete implementation of cells
//! for Excel (.xlsx) files.

use crate::sheet::{Cell as CellTrait, CellValue, Result};

/// Concrete implementation of the Cell trait for Excel files.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Row number (1-based)
    pub row: u32,
    /// Column number (1-based)
    pub column: u32,
    /// Cell value
    pub value: CellValue,
}

impl Cell {
    /// Create a new cell.
    pub fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self { row, column, value }
    }

    /// Convert column number to Excel column letters (e.g., 1 -> "A", 26 -> "Z", 27 -> "AA").
    pub fn column_to_letters(col: u32) -> String {
        let mut letters = String::new();
        let mut col = col;

        while col > 0 {
            col -= 1;
            let letter = ((col % 26) as u8 + b'A') as char;
            letters.insert(0, letter);
            col /= 26;
        }

        letters
    }

    /// Convert Excel reference (e.g., "A1") to row and column numbers.
    pub fn reference_to_coords(reference: &str) -> Result<(u32, u32)> {
        let mut chars = reference.chars();
        let mut col_str = String::new();
        let mut row_str = String::new();

        // Extract column part (letters)
        for ch in &mut chars {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch);
            } else {
                row_str.push(ch);
                break;
            }
        }

        // Add remaining characters to row string
        row_str.extend(chars);

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        let mut col_num = 0u32;
        for ch in col_str.chars() {
            col_num = col_num * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }

        let row_num = row_str.parse::<u32>()
            .map_err(|_| format!("Invalid row number in reference: {}", reference))?;

        Ok((col_num, row_num))
    }
}

impl CellTrait for Cell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        format!("{}{}", Self::column_to_letters(self.column), self.row)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }
}

/// Iterator over cells in a worksheet.
pub struct CellIterator<'a> {
    cells: Vec<Cell>,
    index: usize,
    _phantom: std::marker::PhantomData<&'a ()>,
}

impl<'a> CellIterator<'a> {
    /// Create a new cell iterator.
    pub fn new(cells: Vec<Cell>) -> Self {
        Self {
            cells,
            index: 0,
            _phantom: std::marker::PhantomData,
        }
    }
}

impl<'a> crate::sheet::CellIterator<'a> for CellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn CellTrait + 'a>>> {
        if self.index >= self.cells.len() {
            return None;
        }

        let cell = &self.cells[self.index];
        let boxed_cell = Box::new(cell.clone()) as Box<dyn CellTrait + 'a>;
        self.index += 1;
        Some(Ok(boxed_cell))
    }
}

/// Iterator over rows in a worksheet.
pub struct RowIterator {
    rows: Vec<Vec<CellValue>>,
    index: usize,
}

impl RowIterator {
    /// Create a new row iterator.
    pub fn new(rows: Vec<Vec<CellValue>>) -> Self {
        Self { rows, index: 0 }
    }
}

impl crate::sheet::RowIterator<'_> for RowIterator {
    fn next(&mut self) -> Option<Result<Vec<CellValue>>> {
        if self.index >= self.rows.len() {
            return None;
        }

        let row = self.rows[self.index].clone();
        self.index += 1;
        Some(Ok(row))
    }
}
