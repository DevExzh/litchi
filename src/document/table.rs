//! Table implementation for Word documents.

#[cfg(any(feature = "ole", feature = "ooxml", feature = "odf"))]
use crate::common::Error;
use crate::common::Result;

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A table in a Word document.
#[derive(Debug, Clone)]
pub enum Table {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Table),
    #[cfg(feature = "ooxml")]
    Docx(Box<ooxml::docx::Table>),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Table<'static>),
    #[cfg(feature = "odf")]
    Odt(crate::odf::Table),
}

impl Table {
    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => Ok(t.row_count()),
            #[cfg(feature = "odf")]
            Table::Odt(t) => t
                .row_count()
                .map_err(|e| Error::ParseError(format!("Failed to get row count: {}", e))),
        }
    }

    /// Get the rows in this table.
    ///
    /// **Performance Note**: This method allocates and clones the entire row collection.
    /// For better performance when iterating, consider using `row_count()` and `row_at(index)`.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(Row::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.into_iter().map(|r| Row::Docx(Box::new(r))).collect())
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows.iter().cloned().map(Row::Rtf).collect())
            },
            #[cfg(feature = "odf")]
            Table::Odt(t) => {
                let rows = t
                    .rows()
                    .map_err(|e| Error::ParseError(format!("Failed to get rows: {}", e)))?;
                Ok(rows.into_iter().map(Row::Odt).collect())
            },
        }
    }

    /// Get a specific row by index without allocating a collection.
    ///
    /// This is more efficient than calling `rows()` and then indexing,
    /// as it avoids cloning the entire row collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn row_at(&self, index: usize) -> Result<Option<Row>> {
        match self {
            #[cfg(feature = "ole")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.get(index).cloned().map(Row::Doc))
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.get(index).cloned().map(|r| Row::Docx(Box::new(r))))
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows.get(index).cloned().map(Row::Rtf))
            },
            #[cfg(feature = "odf")]
            Table::Odt(t) => {
                let rows = t
                    .rows()
                    .map_err(|e| Error::ParseError(format!("Failed to get rows: {}", e)))?;
                Ok(rows.get(index).cloned().map(Row::Odt))
            },
        }
    }
}

/// A table row in a Word document.
#[derive(Debug, Clone)]
pub enum Row {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Row),
    #[cfg(feature = "ooxml")]
    Docx(Box<ooxml::docx::Row>),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Row<'static>),
    #[cfg(feature = "odf")]
    Odt(crate::odf::Row),
}

impl Row {
    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => r.cell_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => r.cell_count().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => Ok(r.cell_count()),
            #[cfg(feature = "odf")]
            Row::Odt(r) => r
                .cell_count()
                .map_err(|e| Error::ParseError(format!("Failed to get cell count: {}", e))),
        }
    }

    /// Get the cells in this row.
    ///
    /// **Performance Note**: This method allocates and clones the entire cell collection.
    /// For better performance when iterating, consider using `cell_count()` and `cell_at(index)`.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Docx).collect())
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.iter().cloned().map(Cell::Rtf).collect())
            },
            #[cfg(feature = "odf")]
            Row::Odt(r) => {
                let cells = r
                    .cells()
                    .map_err(|e| Error::ParseError(format!("Failed to get cells: {}", e)))?;
                Ok(cells.into_iter().map(Cell::Odt).collect())
            },
        }
    }

    /// Get a specific cell by index without allocating a collection.
    ///
    /// This is more efficient than calling `cells()` and then indexing,
    /// as it avoids cloning the entire cell collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn cell_at(&self, index: usize) -> Result<Option<Cell>> {
        match self {
            #[cfg(feature = "ole")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.get(index).cloned().map(Cell::Doc))
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.get(index).cloned().map(Cell::Docx))
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.get(index).cloned().map(Cell::Rtf))
            },
            #[cfg(feature = "odf")]
            Row::Odt(r) => {
                let cells = r
                    .cells()
                    .map_err(|e| Error::ParseError(format!("Failed to get cells: {}", e)))?;
                Ok(cells.get(index).cloned().map(Cell::Odt))
            },
        }
    }
}

/// A table cell in a Word document.
#[derive(Debug, Clone)]
pub enum Cell {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Cell),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Cell),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Cell<'static>),
    #[cfg(feature = "odf")]
    Odt(crate::odf::Cell),
}

impl Cell {
    /// Get the text content of the cell.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Cell::Doc(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "rtf")]
            Cell::Rtf(c) => Ok(c.text().to_string()),
            #[cfg(feature = "odf")]
            Cell::Odt(c) => c
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to get cell text: {}", e))),
        }
    }

    /// Get the grid span (colspan) of this cell.
    ///
    /// Returns the number of columns this cell spans. Default is 1 (no merge).
    ///
    /// **Note**: Currently only implemented for OOXML (.docx) format.
    /// Other formats always return 1.
    pub fn grid_span(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "ole")]
            Cell::Doc(_) => Ok(1), // Not implemented for OLE format
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c.grid_span().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Cell::Rtf(_) => Ok(1), // Not implemented for RTF format
            #[cfg(feature = "odf")]
            Cell::Odt(_) => Ok(1), // Grid span not available in ODF format (always 1)
        }
    }

    /// Get the vertical merge state of this cell.
    ///
    /// Returns the vertical merge state if this cell participates in vertical merging,
    /// or `None` if no vertical merge is present.
    ///
    /// **Note**: Currently only implemented for OOXML (.docx) format.
    /// Other formats always return `None`.
    #[cfg(feature = "ooxml")]
    pub fn v_merge(&self) -> Result<Option<crate::ooxml::docx::VMergeState>> {
        match self {
            #[cfg(feature = "ole")]
            Cell::Doc(_) => Ok(None), // Not implemented for OLE format
            Cell::Docx(c) => c.v_merge().map_err(Error::from),
            #[cfg(feature = "rtf")]
            Cell::Rtf(_) => Ok(None), // Not implemented for RTF format
            #[cfg(feature = "odf")]
            Cell::Odt(_) => Ok(None), // Vertical merge not available in ODF format
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::Document;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("test-data")
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_row_count_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let row_count = table.row_count().expect("Failed to get row count");
            assert!(row_count > 0, "Table should have at least one row");
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_rows_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");
            assert!(!rows.is_empty(), "Expected at least one row");

            for row in &rows {
                let cell_count = row.cell_count().expect("Failed to get cell count");
                assert!(cell_count > 0, "Row should have at least one cell");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_row_at_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let first_row = table.row_at(0).expect("Failed to get row at index 0");
            assert!(first_row.is_some(), "Expected to find first row");

            let nonexistent_row = table
                .row_at(9999)
                .expect("Failed to check row at index 9999");
            assert!(nonexistent_row.is_none(), "Expected no row at index 9999");
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_cells_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");
                assert!(!cells.is_empty(), "Expected at least one cell");

                for cell in &cells {
                    let _text = cell.text().expect("Failed to get cell text");
                }
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_cell_at_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let first_cell = row.cell_at(0).expect("Failed to get cell at index 0");
                assert!(first_cell.is_some(), "Expected to find first cell");

                let nonexistent_cell = row
                    .cell_at(9999)
                    .expect("Failed to check cell at index 9999");
                assert!(nonexistent_cell.is_none(), "Expected no cell at index 9999");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_cell_grid_span_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");

                for cell in &cells {
                    let grid_span = cell.grid_span().expect("Failed to get grid span");
                    assert!(grid_span >= 1, "Grid span should be at least 1");
                }
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_table_document_with_tables() {
        let test_files = [
            "ooxml/docx/table_footnotes.docx",
            "ooxml/docx/table-indent.docx",
            "ooxml/docx/table-alignment.docx",
        ];

        for file in &test_files {
            let path = test_data_path().join(file);
            if path.exists() {
                let doc = Document::open(&path);
                assert!(doc.is_ok(), "Failed to open {}", file);

                if let Ok(d) = doc {
                    let tables = d.tables().expect("Failed to get tables");
                    for table in &tables {
                        let row_count = table.row_count().expect("Failed to get row count");
                        assert!(row_count > 0, "Expected at least one row in {}", file);
                    }
                }
            }
        }
    }

    #[test]
    #[cfg(feature = "rtf")]
    fn test_table_rtf() {
        let path = test_data_path().join("rtf/chtoutline.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let row_count = table.row_count().expect("Failed to get row count");
            assert!(row_count > 0, "RTF table should have at least one row");

            let rows = table.rows().expect("Failed to get rows");
            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");
                assert!(!cells.is_empty(), "RTF row should have cells");

                for cell in &cells {
                    let _text = cell.text().expect("Failed to get cell text");
                }
            }
        }
    }
}
