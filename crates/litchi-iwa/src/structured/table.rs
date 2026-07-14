//! Structured table model.
use std::collections::HashMap;

/// Represents a table extracted from a Numbers document
#[derive(Debug, Clone)]
pub struct Table {
    /// Table name
    pub name: String,
    /// Number of rows
    pub row_count: usize,
    /// Number of columns
    pub column_count: usize,
    /// Cell data (row, column) -> value
    pub cells: HashMap<(usize, usize), CellValue>,
}

impl Table {
    /// Create a new empty table
    pub fn new(name: String) -> Self {
        Self {
            name,
            row_count: 0,
            column_count: 0,
            cells: HashMap::new(),
        }
    }

    /// Get a cell value at the specified position
    pub fn get_cell(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    /// Set a cell value at the specified position
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellValue) {
        self.cells.insert((row, col), value);
        self.row_count = self.row_count.max(row + 1);
        self.column_count = self.column_count.max(col + 1);
    }

    /// Convert table to CSV format
    pub fn to_csv(&self) -> String {
        let mut csv = String::new();
        for row in 0..self.row_count {
            for col in 0..self.column_count {
                if col > 0 {
                    csv.push(',');
                }
                if let Some(cell) = self.get_cell(row, col) {
                    csv.push_str(&cell.to_string());
                }
            }
            csv.push('\n');
        }
        csv
    }
}

/// Represents a cell value in a table
#[derive(Debug, Clone)]
pub enum CellValue {
    /// Text/string value
    Text(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Date value (as string)
    Date(String),
    /// Formula (stored as string)
    Formula(String),
    /// Empty cell
    Empty,
}

impl CellValue {
    /// Check if cell is empty
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Text(s) => write!(f, "{}", s),
            CellValue::Number(n) => write!(f, "{}", n),
            CellValue::Boolean(b) => write!(f, "{}", b),
            CellValue::Date(d) => write!(f, "{}", d),
            CellValue::Formula(formula) => write!(f, "{}", formula),
            CellValue::Empty => Ok(()),
        }
    }
}
