//! Cell implementation for Excel worksheets.
//!
//! This module provides the concrete implementation of cells
//! for Excel (.xlsx) files.

use litchi_core::sheet::{Cell as CellTrait, CellValue, Result};
use std::borrow::Cow;

const MAX_EXCEL_COLUMN: u32 = 16_384;
const MAX_EXCEL_ROW: u32 = 1_048_576;

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

    /// Convert an Excel A1 reference to one-based column and row numbers.
    pub fn reference_to_coords(reference: &str) -> Result<(u32, u32)> {
        let bytes = reference.as_bytes();
        let column_end = bytes
            .iter()
            .position(u8::is_ascii_digit)
            .ok_or_else(|| format!("Invalid cell reference: {reference}"))?;
        if column_end == 0 || column_end == bytes.len() {
            return Err(format!("Invalid cell reference: {reference}").into());
        }

        let mut col_num = 0u32;
        for byte in &bytes[..column_end] {
            if !byte.is_ascii_alphabetic() {
                return Err(format!("Invalid column in cell reference: {reference}").into());
            }
            let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
            col_num = col_num
                .checked_mul(26)
                .and_then(|value| value.checked_add(digit))
                .ok_or_else(|| format!("Column overflows in cell reference: {reference}"))?;
        }
        if col_num > MAX_EXCEL_COLUMN {
            return Err(
                format!("Column exceeds Excel limits in cell reference: {reference}").into(),
            );
        }

        let row_bytes = &bytes[column_end..];
        if !row_bytes.iter().all(u8::is_ascii_digit) {
            return Err(format!("Invalid row in cell reference: {reference}").into());
        }
        let row_num = atoi_simd::parse::<_, false, false>(row_bytes)
            .map_err(|_| format!("Invalid row number in cell reference: {reference}"))?;
        if row_num == 0 || row_num > MAX_EXCEL_ROW {
            return Err(format!("Row exceeds Excel limits in cell reference: {reference}").into());
        }

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

impl<'a> litchi_core::sheet::CellIterator<'a> for CellIterator<'a> {
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

impl<'a> litchi_core::sheet::RowIterator<'a> for RowIterator {
    fn next(&mut self) -> Option<Result<Cow<'a, [CellValue]>>> {
        if self.index >= self.rows.len() {
            return None;
        }

        // Since we own the data, we must return Cow::Owned
        // We use std::mem::take to move the row out without cloning
        let row = std::mem::take(&mut self.rows[self.index]);
        self.index += 1;
        Some(Ok(Cow::Owned(row)))
    }
}

#[cfg(test)]
mod tests {
    use super::Cell;

    #[test]
    fn parses_valid_a1_references_at_grid_boundaries() {
        assert_eq!(Cell::reference_to_coords("A1").unwrap(), (1, 1));
        assert_eq!(Cell::reference_to_coords("aa42").unwrap(), (27, 42));
        assert_eq!(
            Cell::reference_to_coords("XFD1048576").unwrap(),
            (16_384, 1_048_576)
        );
    }

    #[test]
    fn rejects_malformed_overflowing_and_out_of_grid_references() {
        for reference in [
            "",
            "A",
            "1",
            "A0",
            "A-1",
            "$A$1",
            "A1x",
            "XFE1",
            "A1048577",
            "ZZZZZZZZZZZZZZZZZZZZ1",
        ] {
            assert!(
                Cell::reference_to_coords(reference).is_err(),
                "accepted {reference}"
            );
        }
    }
}
