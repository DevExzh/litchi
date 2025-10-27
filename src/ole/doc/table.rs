/// Table, Row, and Cell structures for legacy Word documents.
use super::package::Result;
use super::paragraph::Paragraph;
use super::parts::tap::{CellProperties, TableJustification, TableProperties};
use std::sync::Arc;

/// A table in a Word document.
///
/// Represents a table in the binary DOC format.
///
/// # Example
///
/// ```rust,ignore
/// for table in document.tables()? {
///     println!("Table with {} rows", table.row_count()?);
///     for row in table.rows()? {
///         for cell in row.cells()? {
///             println!("Cell: {}", cell.text()?);
///         }
///     }
/// }
/// ```
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around table data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Table {
    /// Rows in the table (shared via Arc for efficient cloning)
    rows: Arc<Vec<Row>>,
    /// Table-level properties (if available)
    properties: Option<TableProperties>,
}

impl Table {
    /// Create a new Table.
    #[allow(dead_code)]
    pub(crate) fn new(rows: Vec<Row>) -> Self {
        Self {
            rows: Arc::new(rows),
            properties: None,
        }
    }

    /// Create a new Table with properties.
    #[allow(dead_code)]
    pub(crate) fn with_properties(rows: Vec<Row>, properties: TableProperties) -> Self {
        Self {
            rows: Arc::new(rows),
            properties: Some(properties),
        }
    }

    /// Get the number of rows in this table.
    pub fn row_count(&self) -> Result<usize> {
        Ok(self.rows.len())
    }

    /// Get the number of columns in this table.
    ///
    /// Returns the column count from the first row, or 0 if the table is empty.
    pub fn column_count(&self) -> Result<usize> {
        if let Some(first_row) = self.rows.first() {
            first_row.cell_count()
        } else {
            Ok(0)
        }
    }

    /// Get all rows in this table.
    ///
    /// Returns a cloned vector. Due to Arc-based sharing in Row/Cell structures,
    /// cloning is relatively cheap (only increments atomic reference counts).
    pub fn rows(&self) -> Result<Vec<Row>> {
        Ok((*self.rows).clone())
    }

    /// Get a specific cell by row and column index.
    ///
    /// Returns `None` if the indices are out of bounds.
    pub fn cell(&self, row_idx: usize, col_idx: usize) -> Result<Option<Cell>> {
        if let Some(row) = self.rows.get(row_idx) {
            let cells = row.cells()?;
            Ok(cells.get(col_idx).cloned())
        } else {
            Ok(None)
        }
    }

    /// Get the table properties.
    ///
    /// Returns the table-level formatting properties if available.
    pub fn properties(&self) -> Option<&TableProperties> {
        self.properties.as_ref()
    }

    /// Get table justification (alignment).
    pub fn justification(&self) -> Option<TableJustification> {
        self.properties.as_ref().map(|p| p.justification)
    }

    /// Check if the first row is a header row.
    pub fn has_header_row(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.is_header_row)
    }
}

/// A row in a table.
///
/// Represents a table row in the binary DOC format.
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around row data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Row {
    /// Cells in the row (shared via Arc for efficient cloning)
    cells: Arc<Vec<Cell>>,
    /// Row-level properties (if available)
    row_properties: Option<TableProperties>,
}

impl Row {
    /// Create a new Row.
    #[allow(unused)]
    pub(crate) fn new(cells: Vec<Cell>) -> Self {
        Self {
            cells: Arc::new(cells),
            row_properties: None,
        }
    }

    /// Create a new Row with properties.
    #[allow(unused)]
    pub(crate) fn with_properties(cells: Vec<Cell>, properties: TableProperties) -> Self {
        Self {
            cells: Arc::new(cells),
            row_properties: Some(properties),
        }
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        Ok(self.cells.len())
    }

    /// Get all cells in this row.
    ///
    /// Returns a cloned vector. Due to Rc-based sharing in Cell structures,
    /// cloning is relatively cheap (only increments reference counts).
    pub fn cells(&self) -> Result<Vec<Cell>> {
        Ok((*self.cells).clone())
    }

    /// Get the row properties.
    pub fn properties(&self) -> Option<&TableProperties> {
        self.row_properties.as_ref()
    }

    /// Get the row height in twips (1/1440 inch).
    pub fn height(&self) -> Option<i16> {
        self.row_properties.as_ref().and_then(|p| p.row_height)
    }

    /// Check if this is a header row.
    pub fn is_header(&self) -> bool {
        self.row_properties
            .as_ref()
            .is_some_and(|p| p.is_header_row)
    }
}

/// A cell in a table.
///
/// Represents a table cell in the binary DOC format.
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around cell data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Cell content (text) - shared via Arc for efficient cloning
    text: Arc<String>,
    /// Cell content (paragraphs) - shared via Arc for efficient cloning
    paragraphs: Arc<Vec<Paragraph>>,
    /// Cell properties (if available)
    properties: Option<CellProperties>,
}

impl Cell {
    /// Create a new Cell.
    #[allow(unused)]
    pub(crate) fn new(text: String) -> Self {
        let para = Paragraph::new(text.clone());
        Self {
            text: Arc::new(text),
            paragraphs: Arc::new(vec![para]),
            properties: None,
        }
    }

    /// Create a new Cell with paragraphs and properties.
    #[allow(unused)]
    pub(crate) fn with_properties(
        paragraphs: Vec<Paragraph>,
        properties: Option<CellProperties>,
    ) -> Self {
        let text = paragraphs
            .iter()
            .filter_map(|p| p.text().ok())
            .collect::<Vec<&str>>()
            .join("\n");
        Self {
            text: Arc::new(text),
            paragraphs: Arc::new(paragraphs),
            properties,
        }
    }

    /// Get the text content of this cell.
    ///
    /// Concatenates all text from all paragraphs in the cell.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get all paragraphs in this cell.
    ///
    /// Returns a cloned vector. Cloning is relatively cheap due to Rc-based sharing.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        Ok((*self.paragraphs).clone())
    }

    /// Get the cell properties.
    pub fn properties(&self) -> Option<&CellProperties> {
        self.properties.as_ref()
    }

    /// Get the cell's vertical alignment.
    pub fn vertical_alignment(&self) -> Option<super::parts::tap::VerticalAlignment> {
        self.properties.as_ref().map(|p| p.vertical_alignment)
    }

    /// Get the cell's background color as RGB tuple.
    pub fn background_color(&self) -> Option<(u8, u8, u8)> {
        self.properties.as_ref().and_then(|p| p.background_color)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_text() {
        let cell = Cell::new("Cell content".to_string());
        assert_eq!(cell.text().unwrap(), "Cell content");
    }

    #[test]
    fn test_row_cell_count() {
        let cells = vec![
            Cell::new("A".to_string()),
            Cell::new("B".to_string()),
            Cell::new("C".to_string()),
        ];
        let row = Row::new(cells);
        assert_eq!(row.cell_count().unwrap(), 3);
    }

    #[test]
    fn test_table_dimensions() {
        let row1 = Row::new(vec![Cell::new("A".to_string()), Cell::new("B".to_string())]);
        let row2 = Row::new(vec![Cell::new("C".to_string()), Cell::new("D".to_string())]);
        let table = Table::new(vec![row1, row2]);

        assert_eq!(table.row_count().unwrap(), 2);
        assert_eq!(table.column_count().unwrap(), 2);
    }
}
