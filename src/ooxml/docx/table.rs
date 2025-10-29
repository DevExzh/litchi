/// Table, Row, and Cell structures for Word documents.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::error::{OoxmlError, Result};
use parking_lot::RwLock;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::sync::Arc;

/// Vertical merge state for table cells.
///
/// In OOXML, vertical merging uses the `<w:vMerge>` element:
/// - `restart`: Starts a new vertical merge (first cell in the merge)
/// - `continue`: Continues a vertical merge from the cell above (no `val` attribute or `val="continue"`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VMergeState {
    /// Starts a vertical merge (`<w:vMerge w:val="restart"/>`)
    Restart,
    /// Continues a vertical merge from above (`<w:vMerge/>` or `<w:vMerge w:val="continue"/>`)
    Continue,
}

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
///
/// # Performance
///
/// Uses lazy parsing with caching - XML is parsed once on first access,
/// then cached results are returned on subsequent calls.
/// Uses Arc and parking_lot RwLock for thread-safe caching, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Table {
    /// The raw XML bytes for this table (shared via Arc for efficient cloning)
    xml_bytes: Arc<Vec<u8>>,
    /// Cached parsed rows (lazy initialization with thread-safe parking_lot RwLock)
    cached_rows: Arc<RwLock<Option<SmallVec<[Row; 16]>>>>,
}

impl Table {
    /// Create a new Table from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_bytes: Arc::new(xml_bytes),
            cached_rows: Arc::new(RwLock::new(None)),
        }
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
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
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
    /// Uses lazy parsing with caching - parses XML once on first call,
    /// returns cached results on subsequent calls. Thread-safe via parking_lot RwLock.
    pub fn rows(&self) -> Result<SmallVec<[Row; 16]>> {
        // Check if we have cached rows (read lock)
        {
            let cache = self.cached_rows.read();
            if let Some(ref rows) = *cache {
                return Ok(rows.clone());
            }
        }

        // Parse rows from XML
        let rows = self.parse_rows()?;

        // Cache the result (write lock)
        *self.cached_rows.write() = Some(rows.clone());

        Ok(rows)
    }

    /// Parse rows from XML (internal method).
    fn parse_rows(&self) -> Result<SmallVec<[Row; 16]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of row collections
        let mut rows = SmallVec::new();
        let mut current_row_xml = Vec::with_capacity(4096); // Pre-allocate for row XML (increased from 2048)
        let mut in_row = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(1024); // Reusable buffer (increased from 512)

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
                },
                Ok(Event::End(e)) => {
                    if in_row {
                        current_row_xml.extend_from_slice(b"</");
                        current_row_xml.extend_from_slice(e.name().as_ref());
                        current_row_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tr" {
                            // Move the XML bytes instead of cloning for the last item
                            // For others, we must clone since we reuse the buffer
                            let row_xml = std::mem::take(&mut current_row_xml);
                            rows.push(Row::new(row_xml));
                            current_row_xml = Vec::with_capacity(2048);
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
///
/// # Performance
///
/// Uses lazy parsing with caching - XML is parsed once on first access,
/// then cached results are returned on subsequent calls.
#[derive(Debug, Clone)]
pub struct Row {
    /// The raw XML bytes for this row (shared via Arc for efficient cloning)
    xml_bytes: Arc<Vec<u8>>,
    /// Cached parsed cells (lazy initialization with thread-safe parking_lot RwLock)
    cached_cells: Arc<RwLock<Option<SmallVec<[Cell; 16]>>>>,
}

impl Row {
    /// Create a new Row from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_bytes: Arc::new(xml_bytes),
            cached_cells: Arc::new(RwLock::new(None)),
        }
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
    ///
    /// # Performance
    ///
    /// Uses lazy parsing with caching - parses XML once on first call,
    /// returns cached results on subsequent calls. Thread-safe via parking_lot RwLock.
    pub fn cells(&self) -> Result<SmallVec<[Cell; 16]>> {
        // Check if we have cached cells (read lock)
        {
            let cache = self.cached_cells.read();
            if let Some(ref cells) = *cache {
                return Ok(cells.clone());
            }
        }

        // Parse cells from XML
        let cells = self.parse_cells()?;

        // Cache the result (write lock)
        *self.cached_cells.write() = Some(cells.clone());

        Ok(cells)
    }

    /// Parse cells from XML (internal method).
    fn parse_cells(&self) -> Result<SmallVec<[Cell; 16]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of cell collections
        let mut cells = SmallVec::new();
        let mut current_cell_xml = Vec::with_capacity(4096); // Pre-allocate for cell XML (increased from 2048)
        let mut in_cell = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(1024); // Reusable buffer (increased from 512)

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
                },
                Ok(Event::End(e)) => {
                    if in_cell {
                        current_cell_xml.extend_from_slice(b"</");
                        current_cell_xml.extend_from_slice(e.name().as_ref());
                        current_cell_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tc" {
                            // Move the XML bytes instead of cloning
                            let cell_xml = std::mem::take(&mut current_cell_xml);
                            cells.push(Cell::new(cell_xml));
                            current_cell_xml = Vec::with_capacity(2048);
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

/// A cell in a table.
///
/// Represents a `<w:tc>` element. Cells contain paragraphs.
///
/// # Performance
///
/// Uses lazy parsing with caching - text is extracted once on first access,
/// then cached results are returned on subsequent calls.
#[derive(Debug, Clone)]
pub struct Cell {
    /// The raw XML bytes for this cell (shared via Arc for efficient cloning)
    xml_bytes: Arc<Vec<u8>>,
    /// Cached extracted text (lazy initialization with thread-safe parking_lot RwLock)
    cached_text: Arc<RwLock<Option<String>>>,
}

impl Cell {
    /// Create a new Cell from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_bytes: Arc::new(xml_bytes),
            cached_text: Arc::new(RwLock::new(None)),
        }
    }

    /// Get the grid span (horizontal merge/colspan) of this cell.
    ///
    /// Returns the number of columns this cell spans. A value of 1 (default) means no merge.
    /// This corresponds to the `<w:gridSpan>` element in OOXML.
    ///
    /// # Example
    ///
    /// ```xml
    /// <w:tc>
    ///   <w:tcPr>
    ///     <w:gridSpan w:val="2"/>
    ///   </w:tcPr>
    ///   ...
    /// </w:tc>
    /// ```
    pub fn grid_span(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_tc_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"tcPr" {
                        in_tc_pr = true;
                    } else if in_tc_pr && name.as_ref() == b"gridSpan" {
                        // Extract the w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let val_str = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                let span = val_str.parse::<usize>().unwrap_or(1);
                                return Ok(span);
                            }
                        }
                        // If no val attribute, default to 1
                        return Ok(1);
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"tcPr" {
                        in_tc_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        // Default: no horizontal merge
        Ok(1)
    }

    /// Get the vertical merge (rowspan) state of this cell.
    ///
    /// Returns `Some(VMergeState)` if this cell participates in vertical merging,
    /// or `None` if no vertical merge is present.
    ///
    /// This corresponds to the `<w:vMerge>` element in OOXML.
    ///
    /// # Example
    ///
    /// ```xml
    /// <!-- Start of vertical merge -->
    /// <w:tc>
    ///   <w:tcPr>
    ///     <w:vMerge w:val="restart"/>
    ///   </w:tcPr>
    ///   ...
    /// </w:tc>
    ///
    /// <!-- Continuation of vertical merge -->
    /// <w:tc>
    ///   <w:tcPr>
    ///     <w:vMerge/>
    ///   </w:tcPr>
    ///   ...
    /// </w:tc>
    /// ```
    pub fn v_merge(&self) -> Result<Option<VMergeState>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_tc_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"tcPr" {
                        in_tc_pr = true;
                    } else if in_tc_pr && name.as_ref() == b"vMerge" {
                        // Check for w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                let val_str = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return match val_str {
                                    "restart" => Ok(Some(VMergeState::Restart)),
                                    _ => Ok(Some(VMergeState::Continue)),
                                };
                            }
                        }
                        // No val attribute means continue
                        return Ok(Some(VMergeState::Continue));
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"tcPr" {
                        in_tc_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        // Default: no vertical merge
        Ok(None)
    }

    /// Get the text content of this cell.
    ///
    /// Concatenates all text from all paragraphs in the cell.
    ///
    /// # Performance
    ///
    /// Uses lazy parsing with caching - parses XML once on first call,
    /// returns cached results on subsequent calls. Thread-safe via parking_lot RwLock.
    pub fn text(&self) -> Result<String> {
        // Check if we have cached text (read lock)
        {
            let cache = self.cached_text.read();
            if let Some(ref text) = *cache {
                return Ok(text.clone());
            }
        }

        // Extract text from XML
        let text = self.extract_text()?;

        // Cache the result (write lock)
        *self.cached_text.write() = Some(text.clone());

        Ok(text)
    }

    /// Extract text from XML (internal method).
    ///
    /// Uses proper XML event parsing to correctly extract text nodes.
    fn extract_text(&self) -> Result<String> {
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
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let text = std::str::from_utf8(e.as_ref()).unwrap_or("");
                    result.push_str(text);
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
                },
                Ok(Event::End(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"p" {
                            // Move the XML bytes instead of cloning
                            let para_xml = std::mem::take(&mut current_para_xml);
                            paragraphs.push(Paragraph::new(para_xml));
                            current_para_xml = Vec::with_capacity(1024);
                            in_para = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
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
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
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
