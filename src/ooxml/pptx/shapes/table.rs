/// Table shape implementation for PowerPoint presentations.
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

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
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut table_xml = Vec::new();
        let mut in_table = false;
        let mut depth = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" && !in_table {
                        in_table = true;
                        depth = 1;
                        table_xml.clear();
                        table_xml.extend_from_slice(b"<a:tbl>");
                    } else if in_table {
                        depth += 1;
                        table_xml.push(b'<');
                        table_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            table_xml.push(b' ');
                            table_xml.extend_from_slice(attr.key.as_ref());
                            table_xml.extend_from_slice(b"=\"");
                            table_xml.extend_from_slice(&attr.value);
                            table_xml.push(b'"');
                        }
                        table_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) => {
                    if in_table {
                        table_xml.extend_from_slice(b"</");
                        table_xml.extend_from_slice(e.name().as_ref());
                        table_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tbl" {
                            return Ok(Table::new(table_xml));
                        }
                    }
                },
                Ok(Event::Text(e)) if in_table => {
                    table_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_table => {
                    table_xml.push(b'<');
                    table_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        table_xml.push(b' ');
                        table_xml.extend_from_slice(attr.key.as_ref());
                        table_xml.extend_from_slice(b"=\"");
                        table_xml.extend_from_slice(&attr.value);
                        table_xml.push(b'"');
                    }
                    table_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Err(OoxmlError::PartNotFound(
            "Table not found in graphic frame".to_string(),
        ))
    }

    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    // DrawingML table rows are <a:tr>
                    if e.local_name().as_ref() == b"tr" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut rows = Vec::new();
        let mut current_row_xml = Vec::new();
        let mut in_row = false;
        let mut depth = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tr" && !in_row {
                        in_row = true;
                        depth = 1;
                        current_row_xml.clear();
                        current_row_xml.extend_from_slice(b"<a:tr>");
                    } else if in_row {
                        depth += 1;
                        current_row_xml.push(b'<');
                        current_row_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_row_xml.push(b' ');
                            current_row_xml.extend_from_slice(attr.key.as_ref());
                            current_row_xml.extend_from_slice(b"=\"");
                            current_row_xml.extend_from_slice(&attr.value);
                            current_row_xml.push(b'"');
                        }
                        current_row_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) => {
                    if in_row {
                        current_row_xml.extend_from_slice(b"</");
                        current_row_xml.extend_from_slice(e.name().as_ref());
                        current_row_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tr" {
                            rows.push(TableRow::new(current_row_xml.clone()));
                            in_row = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_row => {
                    current_row_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_row => {
                    current_row_xml.push(b'<');
                    current_row_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_row_xml.push(b' ');
                        current_row_xml.extend_from_slice(attr.key.as_ref());
                        current_row_xml.extend_from_slice(b"=\"");
                        current_row_xml.extend_from_slice(&attr.value);
                        current_row_xml.push(b'"');
                    }
                    current_row_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    // DrawingML table cells are <a:tc>
                    if e.local_name().as_ref() == b"tc" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(count)
    }

    /// Get all cells in this row.
    pub fn cells(&self) -> Result<Vec<TableCell>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut cells = Vec::new();
        let mut current_cell_xml = Vec::new();
        let mut in_cell = false;
        let mut depth = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tc" && !in_cell {
                        in_cell = true;
                        depth = 1;
                        current_cell_xml.clear();
                        current_cell_xml.extend_from_slice(b"<a:tc>");
                    } else if in_cell {
                        depth += 1;
                        current_cell_xml.push(b'<');
                        current_cell_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_cell_xml.push(b' ');
                            current_cell_xml.extend_from_slice(attr.key.as_ref());
                            current_cell_xml.extend_from_slice(b"=\"");
                            current_cell_xml.extend_from_slice(&attr.value);
                            current_cell_xml.push(b'"');
                        }
                        current_cell_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) => {
                    if in_cell {
                        current_cell_xml.extend_from_slice(b"</");
                        current_cell_xml.extend_from_slice(e.name().as_ref());
                        current_cell_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tc" {
                            cells.push(TableCell::new(current_cell_xml.clone()));
                            in_cell = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_cell => {
                    current_cell_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_cell => {
                    current_cell_xml.push(b'<');
                    current_cell_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_cell_xml.push(b' ');
                        current_cell_xml.extend_from_slice(attr.key.as_ref());
                        current_cell_xml.extend_from_slice(b"=\"");
                        current_cell_xml.extend_from_slice(&attr.value);
                        current_cell_xml.push(b'"');
                    }
                    current_cell_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    if !text.is_empty() && !text.ends_with(' ') {
                        text.push(' ');
                    }
                    text.push_str(t);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(text.trim().to_string())
    }
}
