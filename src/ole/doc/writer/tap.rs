//! TAP (Table Properties) generation for DOC files
//!
//! TAP structures define table layout, borders, and cell properties.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's TableProperties.

use super::sprm::SprmBuilder;

/// Table cell descriptor
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Cell width (in twips)
    pub width: u16,
    /// Merged cell flags
    pub merged: bool,
}

/// Table row properties
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height (in twips)
    pub height: u16,
    /// Header row flag
    pub is_header: bool,
}

/// TAP (Table Properties) builder
#[derive(Debug)]
pub struct TapBuilder {
    rows: Vec<TableRow>,
}

impl TapBuilder {
    /// Create a new TAP builder
    pub fn new() -> Self {
        Self { rows: Vec::new() }
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.rows.push(row);
    }

    /// Generate TAP SPRMs for a specific row
    pub fn generate_row_sprms(&self, row_index: usize) -> Vec<u8> {
        if row_index >= self.rows.len() {
            return Vec::new();
        }

        let row = &self.rows[row_index];
        let mut builder = SprmBuilder::new();

        // Table definition SPRM (sprmTDefTable)
        builder.add_word(0xD608, 0); // Table definition marker

        // Number of cells
        let cell_count = row.cells.len() as u16;
        builder.add_word(0x5400, cell_count);

        // Cell positions (cumulative widths in twips)
        let mut cumulative_width = 0u16;
        for cell in &row.cells {
            cumulative_width = cumulative_width.saturating_add(cell.width);
            builder.add_word(0x5401, cumulative_width);
        }

        // Row height
        if row.height > 0 {
            builder.add_word(0x9407, row.height);
        }

        // Header row flag
        if row.is_header {
            builder.add_bool(0x3403, true);
        }

        builder.build()
    }

    /// Get the number of rows
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }
}

impl Default for TapBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper to create a simple table
pub fn create_simple_table(rows: usize, cols: usize, cell_width: u16) -> TapBuilder {
    let mut builder = TapBuilder::new();

    for _ in 0..rows {
        let cells = vec![
            TableCell {
                width: cell_width,
                merged: false
            };
            cols
        ];
        builder.add_row(TableRow {
            cells,
            height: 0, // Auto height
            is_header: false,
        });
    }

    builder
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tap_builder() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![
                TableCell {
                    width: 1000,
                    merged: false,
                },
                TableCell {
                    width: 1000,
                    merged: false,
                },
            ],
            height: 200,
            is_header: false,
        });

        let sprms = builder.generate_row_sprms(0);
        assert!(!sprms.is_empty());
    }

    #[test]
    fn test_create_simple_table() {
        let table = create_simple_table(3, 4, 1440); // 3 rows, 4 cols, 1 inch cells
        assert_eq!(table.row_count(), 3);
    }
}
