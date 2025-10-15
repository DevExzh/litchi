//! Cell implementation for text-based formats

use crate::sheet::{Cell, CellValue};

/// Cell implementation for text-based formats
#[derive(Debug, Clone)]
pub struct TextCell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl TextCell {
    /// Create a new text cell
    pub fn new(row: u32, column: u32, value: CellValue) -> Self {
        TextCell { row, column, value }
    }

    /// Create an empty cell
    pub fn empty(row: u32, column: u32) -> Self {
        Self::new(row, column, CellValue::Empty)
    }
}

impl Cell for TextCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        // Convert column number to Excel-style coordinate (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
        let mut col_str = String::new();
        let mut col = self.column;

        while col > 0 {
            col -= 1;
            let c = (b'A' + (col % 26) as u8) as char;
            col_str.insert(0, c);
            col /= 26;
        }

        format!("{}{}", col_str, self.row)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_coordinate_conversion() {
        let cell = TextCell::new(1, 1, CellValue::String("test".to_string()));
        assert_eq!(cell.coordinate(), "A1");

        let cell = TextCell::new(10, 5, CellValue::Empty);
        assert_eq!(cell.coordinate(), "E10");

        let cell = TextCell::new(100, 27, CellValue::Empty); // AA
        assert_eq!(cell.coordinate(), "AA100");
    }
}
