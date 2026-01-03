//! Worksheet implementation for XLS files

use crate::ole::xls::cell::XlsCell;
use crate::ole::xls::error::XlsError;
use crate::sheet::{Cell as SheetCell, CellIterator, CellValue, Result, RowIterator, Worksheet};
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::sync::Arc;

/// XLS worksheet implementation
#[derive(Debug, Clone)]
pub struct XlsWorksheet {
    name: String,
    cells: BTreeMap<(u32, u32), XlsCell>,
    max_row: u32,
    max_col: u32,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    shared_strings: Option<Arc<Vec<String>>>,
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

    /// Create a new worksheet with shared strings (Arc for zero-copy sharing)
    pub fn with_shared_strings(name: String, shared_strings: Arc<Vec<String>>) -> Self {
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
    pub fn set_dimensions(
        &mut self,
        _first_row: u32,
        last_row: u32,
        _first_col: u32,
        last_col: u32,
    ) {
        // Adjust max_row and max_col based on dimensions
        self.max_row = self.max_row.max(last_row.saturating_sub(1));
        self.max_col = self.max_col.max(last_col.saturating_sub(1));
    }

    /// Get shared strings reference
    pub fn shared_strings(&self) -> Option<&[String]> {
        self.shared_strings.as_ref().map(|arc| arc.as_slice())
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

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>> {
        match self.cells.get(&(row, column)) {
            // Return reference instead of clone - zero-copy!
            Some(cell) => Ok(Box::new(cell)),
            None => {
                // Return empty cell for missing positions (owned, unavoidable)
                let empty_cell = XlsCell::new(row, column, CellValue::Empty);
                Ok(Box::new(empty_cell))
            },
        }
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        let (row, col) =
            crate::ole::xls::utils::parse_cell_reference(coordinate).ok_or_else(|| {
                Box::new(XlsError::InvalidCellReference(coordinate.to_string()))
                    as Box<dyn std::error::Error + Send + Sync>
            })?;
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

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        let row_idx = row_idx as u32;
        let mut row_data = Vec::new();

        for col in 0..=self.max_col {
            match self.cells.get(&(row_idx, col)) {
                Some(cell) => row_data.push(cell.value().clone()),
                None => row_data.push(CellValue::Empty),
            }
        }

        // Return owned data wrapped in Cow
        Ok(Cow::Owned(row_data))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        match self.cells.get(&(row, column)) {
            Some(cell) => Ok(Cow::Borrowed(cell.value())),
            None => Ok(Cow::Borrowed(CellValue::EMPTY)),
        }
    }
}

// Implement Worksheet for &XlsWorksheet to allow zero-copy reference returns
impl Worksheet for &XlsWorksheet {
    fn name(&self) -> &str {
        (*self).name()
    }

    fn row_count(&self) -> usize {
        (*self).row_count()
    }

    fn column_count(&self) -> usize {
        (*self).column_count()
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        (*self).dimensions()
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn SheetCell + '_>> {
        (*self).cell(row, column)
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn SheetCell + '_>> {
        (*self).cell_by_coordinate(coordinate)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        (*self).cells()
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        (*self).rows()
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        (*self).row(row_idx)
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        (*self).cell_value(row, column)
    }
}

/// Cell iterator for XLS worksheets
struct XlsCellIterator<'a> {
    cells: Vec<&'a XlsCell>,
    index: usize,
}

impl<'a> CellIterator<'a> for XlsCellIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetCell + 'a>>> {
        if self.index >= self.cells.len() {
            None
        } else {
            let cell = self.cells[self.index];
            self.index += 1;
            // Return reference instead of clone - zero-copy!
            Some(Ok(Box::new(cell)))
        }
    }
}

/// Row iterator for XLS worksheets
struct XlsRowIterator<'a> {
    worksheet: &'a XlsWorksheet,
    current_row: usize,
}

impl<'a> RowIterator<'a> for XlsRowIterator<'a> {
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
