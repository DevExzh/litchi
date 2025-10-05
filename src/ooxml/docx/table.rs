/// Table, Row, and Cell structures for Word documents.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::Reader;
use smallvec::SmallVec;

/// A table in a Word document.
///
/// Represents a `<w:tbl>` element. Tables contain rows, which contain cells,
/// which contain paragraphs.
///
/// # Example
///
/// ```rust,ignore
/// for table in document.tables()? {
///     println!("Table with {} rows", table.row_count()?);
///     for (row_idx, row) in table.rows()?.iter().enumerate() {
///         for (col_idx, cell) in row.cells()?.iter().enumerate() {
///             println!("Cell [{},{}]: {}", row_idx, col_idx, cell.text()?);
///         }
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Table {
    /// The raw XML bytes for this table
    xml_bytes: Vec<u8>,
}

impl Table {
    /// Create a new Table from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the number of rows in this table.
    pub fn row_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tr" {
                        count += 1;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(count)
    }

    /// Get the number of columns in this table.
    ///
    /// Returns the column count from the first row, or 0 if the table is empty.
    pub fn column_count(&self) -> Result<usize> {
        let rows = self.rows()?;
        if let Some(first_row) = rows.first() {
            first_row.cell_count()
        } else {
            Ok(0)
        }
    }

    /// Get all rows in this table.
    ///
    /// # Performance
    ///
    /// Uses SmallVec for efficient storage of typically small row collections.
    pub fn rows(&self) -> Result<SmallVec<[Row; 16]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of row collections
        let mut rows = SmallVec::new();
        let mut current_row_xml = Vec::with_capacity(2048); // Pre-allocate for row XML
        let mut in_row = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tr" && !in_row {
                        in_row = true;
                        depth = 1;
                        current_row_xml.clear();
                        current_row_xml.extend_from_slice(b"<w:tr");
                        for attr in e.attributes().flatten() {
                            current_row_xml.push(b' ');
                            current_row_xml.extend_from_slice(attr.key.as_ref());
                            current_row_xml.extend_from_slice(b"=\"");
                            current_row_xml.extend_from_slice(&attr.value);
                            current_row_xml.push(b'"');
                        }
                        current_row_xml.push(b'>');
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
                }
                Ok(Event::End(e)) => {
                    if in_row {
                        current_row_xml.extend_from_slice(b"</");
                        current_row_xml.extend_from_slice(e.name().as_ref());
                        current_row_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tr" {
                            rows.push(Row::new(current_row_xml.clone()));
                            in_row = false;
                        }
                    }
                }
                Ok(Event::Text(e)) if in_row => {
                    current_row_xml.extend_from_slice(e.as_ref());
                }
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
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(rows)
    }

    /// Get a specific cell by row and column index.
    ///
    /// Returns `None` if the indices are out of bounds.
    pub fn cell(&self, row_idx: usize, col_idx: usize) -> Result<Option<Cell>> {
        let rows = self.rows()?;
        if let Some(row) = rows.get(row_idx) {
            let cells = row.cells()?;
            Ok(cells.get(col_idx).cloned())
        } else {
            Ok(None)
        }
    }
}

/// A row in a table.
///
/// Represents a `<w:tr>` element.
#[derive(Debug, Clone)]
pub struct Row {
    /// The raw XML bytes for this row
    xml_bytes: Vec<u8>,
}

impl Row {
    /// Create a new Row from XML bytes.
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
                    if e.local_name().as_ref() == b"tc" {
                        count += 1;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(count)
    }

    /// Get all cells in this row.
    ///
    /// # Performance
    ///
    /// Uses SmallVec for efficient storage of typically small cell collections.
    pub fn cells(&self) -> Result<SmallVec<[Cell; 16]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of cell collections
        let mut cells = SmallVec::new();
        let mut current_cell_xml = Vec::with_capacity(2048); // Pre-allocate for cell XML
        let mut in_cell = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tc" && !in_cell {
                        in_cell = true;
                        depth = 1;
                        current_cell_xml.clear();
                        current_cell_xml.extend_from_slice(b"<w:tc");
                        for attr in e.attributes().flatten() {
                            current_cell_xml.push(b' ');
                            current_cell_xml.extend_from_slice(attr.key.as_ref());
                            current_cell_xml.extend_from_slice(b"=\"");
                            current_cell_xml.extend_from_slice(&attr.value);
                            current_cell_xml.push(b'"');
                        }
                        current_cell_xml.push(b'>');
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
                }
                Ok(Event::End(e)) => {
                    if in_cell {
                        current_cell_xml.extend_from_slice(b"</");
                        current_cell_xml.extend_from_slice(e.name().as_ref());
                        current_cell_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tc" {
                            cells.push(Cell::new(current_cell_xml.clone()));
                            in_cell = false;
                        }
                    }
                }
                Ok(Event::Text(e)) if in_cell => {
                    current_cell_xml.extend_from_slice(e.as_ref());
                }
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
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(cells)
    }
}

/// A cell in a table.
///
/// Represents a `<w:tc>` element. Cells contain paragraphs.
#[derive(Debug, Clone)]
pub struct Cell {
    /// The raw XML bytes for this cell
    xml_bytes: Vec<u8>,
}

impl Cell {
    /// Create a new Cell from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the text content of this cell.
    ///
    /// Concatenates all text from all paragraphs in the cell.
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut result = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                }
                Ok(Event::Text(e)) if in_text_element => {
                    let text = std::str::from_utf8(e.as_ref()).unwrap_or("");
                    result.push_str(text);
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(result)
    }

    /// Get all paragraphs in this cell.
    ///
    /// # Performance
    ///
    /// Uses SmallVec for efficient storage of typically small paragraph collections.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 8]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of paragraph collections
        let mut paragraphs = SmallVec::new();
        let mut current_para_xml = Vec::with_capacity(1024); // Pre-allocate for paragraph XML
        let mut in_para = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        current_para_xml.extend_from_slice(b"<w:p");
                        for attr in e.attributes().flatten() {
                            current_para_xml.push(b' ');
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.push(b'"');
                        }
                        current_para_xml.push(b'>');
                    } else if in_para {
                        depth += 1;
                        current_para_xml.push(b'<');
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.push(b' ');
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.push(b'"');
                        }
                        current_para_xml.push(b'>');
                    }
                }
                Ok(Event::End(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"p" {
                            paragraphs.push(Paragraph::new(current_para_xml.clone()));
                            in_para = false;
                        }
                    }
                }
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                }
                Ok(Event::Empty(e)) if in_para => {
                    current_para_xml.push(b'<');
                    current_para_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_para_xml.push(b' ');
                        current_para_xml.extend_from_slice(attr.key.as_ref());
                        current_para_xml.extend_from_slice(b"=\"");
                        current_para_xml.extend_from_slice(&attr.value);
                        current_para_xml.push(b'"');
                    }
                    current_para_xml.extend_from_slice(b"/>");
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(paragraphs)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_text() {
        let xml = br#"<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:p><w:r><w:t>Cell text</w:t></w:r></w:p>
        </w:tc>"#;

        let cell = Cell::new(xml.to_vec());
        let text = cell.text().unwrap();
        assert_eq!(text, "Cell text");
    }
}
