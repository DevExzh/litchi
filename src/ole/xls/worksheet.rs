//! Worksheet implementation for XLS files

use std::collections::BTreeMap;
use crate::sheet::{Worksheet, Cell as SheetCell, CellValue, CellIterator, RowIterator};
use crate::ole::xls::cell::XlsCell;
use crate::ole::xls::error::XlsError;

/// XLS worksheet implementation
#[derive(Debug, Clone)]
pub struct XlsWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), XlsCell>,
    max_row: u32,
    max_col: u32,
    shared_strings: Option<Vec<String>>,
}

impl XlsWorksheet {
    /// Create a new worksheet
    pub fn new(name: String) -> Self {
        XlsWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            shared_strings: None,
        }
    }

    /// Create a new worksheet with shared strings
    pub fn with_shared_strings(name: String, shared_strings: Vec<String>) -> Self {
        XlsWorksheet {
            name,
            cells: BTreeMap::new(),
            max_row: 0,
            max_col: 0,
            shared_strings: Some(shared_strings),
        }
    }

    /// Add a cell to the worksheet
    pub fn add_cell(&mut self, cell: XlsCell) {
        let pos = (cell.row(), cell.column());
        self.max_row = self.max_row.max(cell.row());
        self.max_col = self.max_col.max(cell.column());
        self.cells.insert(pos, cell);
    }

    /// Set worksheet dimensions
    pub fn set_dimensions(&mut self, _first_row: u32, last_row: u32, _first_col: u32, last_col: u32) {
        // Adjust max_row and max_col based on dimensions
        self.max_row = self.max_row.max(last_row.saturating_sub(1));
        self.max_col = self.max_col.max(last_col.saturating_sub(1));
    }

    /// Get shared strings reference
    pub fn shared_strings(&self) -> Option<&[String]> {
        self.shared_strings.as_deref()
    }

    /// Get cell at position
    pub fn get_cell(&self, row: u32, col: u32) -> Option<&XlsCell> {
        self.cells.get(&(row, col))
    }
}

impl Worksheet for XlsWorksheet {
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

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>, Box<dyn std::error::Error>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Box::new(cell.clone())),
            None => {
                // Return empty cell for missing positions
                let empty_cell = XlsCell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            }
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>, Box<dyn std::error::Error>> {
        let (row, col) = crate::ole::xls::utils::parse_cell_reference(coordinate)
            .ok_or_else(|| XlsError::InvalidCellReference(coordinate.to_string()))?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(XlsCellIterator {
            cells: self.cells.values().collect(),
            index: 0,
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(XlsRowIterator {
            worksheet: self,
            current_row: 0,
        })
    }

    fn row(&self, row_idx: usize) -> Result<Vec<CellValue>, Box<dyn std::error::Error>> {
        let row_idx = row_idx as u32;
        let mut row_data = Vec::new();

        for col in 0..=self.max_col {
            match self.cells.get(&(row_idx, col)) {
                Some(cell) => row_data.push(cell.value().clone()),
                None => row_data.push(CellValue::Empty),
            }
        }

        Ok(row_data)
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<CellValue, Box<dyn std::error::Error>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(cell.value().clone()),
            None => Ok(CellValue::Empty),
        }
    }
}

/// Cell iterator for XLS worksheets
struct XlsCellIterator<'a> {
    cells: Vec<&'a XlsCell>,
    index: usize,
}

impl<'a> CellIterator<'a> for XlsCellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetCell + 'a>, Box<dyn std::error::Error>>> {
        if self.index >= self.cells.len() {
            None
        } else {
            let cell = self.cells[self.index];
            self.index += 1;
            Some(Ok(Box::new(cell.clone())))
        }
    }
}

/// Row iterator for XLS worksheets
struct XlsRowIterator<'a> {
    worksheet: &'a XlsWorksheet,
    current_row: usize,
}

impl<'a> RowIterator<'a> for XlsRowIterator<'a> {
    fn next(&mut self) -> Option<Result<Vec<CellValue>, Box<dyn std::error::Error>>> {
        if self.current_row >= self.worksheet.row_count() {
            None
        } else {
            let result = self.worksheet.row(self.current_row);
            self.current_row += 1;
            Some(result)
        }
    }
}
