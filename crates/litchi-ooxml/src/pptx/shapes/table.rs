/// Table shape implementation for PowerPoint presentations.
use crate::error::{OoxmlError, Result};
use crate::pptx::shapes::textframe::{extract_drawingml_text, scan_drawingml_element_ranges};

/// A table in a PowerPoint presentation.
///
/// Tables in PowerPoint are DrawingML tables (a:tbl) contained within
/// graphic frames. They contain rows, which contain cells.
///
/// # Examples
///
/// ```rust,ignore
/// if let Some(table) = graphic_frame.table() {
///     println!("Table: {}x{}", table.row_count()?, table.column_count()?);
///     
///     for (row_idx, row) in table.rows()?.iter().enumerate() {
///         for (col_idx, cell) in row.cells()?.iter().enumerate() {
///             println!("Cell[{},{}]: {}", row_idx, col_idx, cell.text()?);
///         }
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Table {
    /// Raw XML bytes for the table
    xml_bytes: Vec<u8>,
}

impl Table {
    /// Create a new Table from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Extract table XML from graphic frame XML.
    ///
    /// GraphicFrames contain the table within their structure, so we need
    /// to extract just the table portion.
    pub fn from_graphic_frame_xml(xml_bytes: &[u8]) -> Result<Self> {
        let mut table = None;
        scan_drawingml_element_ranges(xml_bytes, b"tbl", |start, length| {
            if table.is_none() {
                let start = start as usize;
                table = Some(Table::new(
                    xml_bytes[start..start + length as usize].to_vec(),
                ));
            }
            Ok(())
        })?;
        table
            .ok_or_else(|| OoxmlError::PartNotFound("Table not found in graphic frame".to_string()))
    }

    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_drawingml_element_ranges(&self.xml_bytes, b"tr", |_, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get the number of columns in the table.
    ///
    /// Returns the number of cells in the first row, or 0 if the table is empty.
    pub fn column_count(&self) -> Result<usize> {
        let rows = self.rows()?;
        if let Some(first_row) = rows.first() {
            first_row.cell_count()
        } else {
            Ok(0)
        }
    }

    /// Get all rows in the table.
    pub fn rows(&self) -> Result<Vec<TableRow>> {
        let mut rows = Vec::new();
        scan_drawingml_element_ranges(&self.xml_bytes, b"tr", |start, length| {
            let start = start as usize;
            rows.push(TableRow::new(
                self.xml_bytes[start..start + length as usize].to_vec(),
            ));
            Ok(())
        })?;
        Ok(rows)
    }

    /// Get a specific cell by row and column index.
    ///
    /// Indices are zero-based. Returns None if the indices are out of bounds.
    pub fn cell(&self, row_idx: usize, col_idx: usize) -> Result<Option<TableCell>> {
        let rows = self.rows()?;
        if let Some(row) = rows.get(row_idx) {
            let cells = row.cells()?;
            Ok(cells.get(col_idx).cloned())
        } else {
            Ok(None)
        }
    }
}

/// A row in a PowerPoint table.
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Raw XML bytes for this row
    xml_bytes: Vec<u8>,
}

impl TableRow {
    /// Create a new TableRow from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_drawingml_element_ranges(&self.xml_bytes, b"tc", |_, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get all cells in this row.
    pub fn cells(&self) -> Result<Vec<TableCell>> {
        let mut cells = Vec::new();
        scan_drawingml_element_ranges(&self.xml_bytes, b"tc", |start, length| {
            let start = start as usize;
            cells.push(TableCell::new(
                self.xml_bytes[start..start + length as usize].to_vec(),
            ));
            Ok(())
        })?;
        Ok(cells)
    }
}

/// A cell in a PowerPoint table.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Raw XML bytes for this cell
    xml_bytes: Vec<u8>,
}

impl TableCell {
    /// Create a new TableCell from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Extract all text from this cell.
    pub fn text(&self) -> Result<String> {
        Ok(extract_drawingml_text(&self.xml_bytes, Some(' '))?
            .trim()
            .to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_cell_text_decodes_runs_and_separates_paragraphs() {
        let cell = TableCell::new(
            br#"<d:tc xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><d:txBody><d:p><d:r><d:t>A &amp; </d:t></d:r><d:r><d:t><![CDATA[B < C]]></d:t></d:r></d:p><d:p><d:r><d:t>D</d:t></d:r></d:p></d:txBody></d:tc>"#
                .to_vec(),
        );
        assert_eq!(cell.text().unwrap(), "A & B < C D");
    }

    #[test]
    fn table_ranges_preserve_aliases_and_ignore_foreign_lookalikes() {
        let xml = br#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml">
            <false:tbl><false:tr><false:tc/></false:tr></false:tbl>
            <d:tbl data-id="kept">
                <false:tr><false:tc/></false:tr>
                <d:tr h="20"><false:tc/><d:tc><d:txBody><d:p><d:r><d:t><![CDATA[A < B]]></d:t></d:r></d:p></d:txBody></d:tc></d:tr>
            </d:tbl>
        </p:graphicFrame>"#;
        let table = Table::from_graphic_frame_xml(xml).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
        let rows = table.rows().unwrap();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].cell_count().unwrap(), 1);
        let cells = rows[0].cells().unwrap();
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].text().unwrap(), "A < B");
    }

    #[test]
    fn table_ranges_reject_unterminated_selected_elements() {
        let xml = br#"<a:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tbl><a:tr/>"#;
        assert!(Table::from_graphic_frame_xml(xml).is_err());
    }
}
