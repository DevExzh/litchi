/// DocumentPart - the main document.xml part of a Word document.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;

/// The main document part of a Word document.
///
/// This corresponds to the `/word/document.xml` part in the package.
/// It contains the main document content including paragraphs, tables,
/// sections, and other block-level elements.
pub struct DocumentPart<'a> {
    /// Reference to the underlying part
    part: &'a dyn Part,
}

impl<'a> DocumentPart<'a> {
    /// Create a DocumentPart from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the document.xml content
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the document.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Extract all paragraph text from the document.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML.
    ///
    /// # Performance
    ///
    /// Uses `quick-xml` for efficient streaming XML parsing with pre-allocated
    /// buffer and unsafe string conversion for optimal performance.
    pub fn extract_text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity to reduce reallocations
        let estimated_capacity = self.xml_bytes().len() / 8; // Rough estimate for text content
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    // Check if this is a w:t element
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Extract text content - use unsafe conversion for better performance
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
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

    /// Count the number of paragraphs in the document.
    ///
    /// Counts `<w:p>` elements in the document body.
    pub fn paragraph_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"p" {
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

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    pub fn table_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" {
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

    /// Get all paragraphs in the document.
    ///
    /// Extracts all `<w:p>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated SmallVec for efficient
    /// storage of typically small paragraph collections.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of paragraph collections
        let mut paragraphs = SmallVec::new();
        let mut current_para_xml = Vec::with_capacity(4096); // Pre-allocate for paragraph XML (increased from 2048)
        let mut in_para = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(2048); // Reusable buffer (increased from 512 to reduce reallocations)

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
                            // PERFORMANCE: Use mem::take to move the buffer instead of cloning
                            // This avoids allocating a new Vec for each paragraph
                            paragraphs.push(Paragraph::new(std::mem::take(&mut current_para_xml)));
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

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses SmallVec for efficient storage of typically small table collections.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of table collections
        let mut tables = SmallVec::new();
        let mut current_table_xml = Vec::with_capacity(8192); // Pre-allocate for table XML (increased from 4096, tables can be large)
        let mut in_table = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(2048); // Reusable buffer (increased from 512 to reduce reallocations)

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" && !in_table {
                        in_table = true;
                        depth = 1;
                        current_table_xml.clear();
                        current_table_xml.extend_from_slice(b"<w:tbl");
                        for attr in e.attributes().flatten() {
                            current_table_xml.push(b' ');
                            current_table_xml.extend_from_slice(attr.key.as_ref());
                            current_table_xml.extend_from_slice(b"=\"");
                            current_table_xml.extend_from_slice(&attr.value);
                            current_table_xml.push(b'"');
                        }
                        current_table_xml.push(b'>');
                    } else if in_table {
                        depth += 1;
                        current_table_xml.push(b'<');
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_table_xml.push(b' ');
                            current_table_xml.extend_from_slice(attr.key.as_ref());
                            current_table_xml.extend_from_slice(b"=\"");
                            current_table_xml.extend_from_slice(&attr.value);
                            current_table_xml.push(b'"');
                        }
                        current_table_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) => {
                    if in_table {
                        current_table_xml.extend_from_slice(b"</");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        current_table_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tbl" {
                            // PERFORMANCE: Use mem::take to move the buffer instead of cloning
                            // This avoids allocating a new Vec for each table
                            tables.push(Table::new(std::mem::take(&mut current_table_xml)));
                            in_table = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_table => {
                    current_table_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_table => {
                    current_table_xml.push(b'<');
                    current_table_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_table_xml.push(b' ');
                        current_table_xml.extend_from_slice(attr.key.as_ref());
                        current_table_xml.extend_from_slice(b"=\"");
                        current_table_xml.extend_from_slice(&attr.value);
                        current_table_xml.push(b'"');
                    }
                    current_table_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(tables)
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method parses the XML once and extracts both paragraphs and tables,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that extracts both `<w:p>` and `<w:tbl>` elements
    /// in document order, which is significantly faster than parsing the XML twice.
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        use crate::document::DocumentElement;

        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut elements = Vec::new();
        let mut current_element_xml = Vec::with_capacity(8192);
        let mut in_paragraph = false;
        let mut in_table = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(2048);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    // Start of a paragraph
                    if e.local_name().as_ref() == b"p" && !in_paragraph && !in_table {
                        in_paragraph = true;
                        depth = 1;
                        current_element_xml.clear();
                        current_element_xml.extend_from_slice(b"<w:p");
                        for attr in e.attributes().flatten() {
                            current_element_xml.push(b' ');
                            current_element_xml.extend_from_slice(attr.key.as_ref());
                            current_element_xml.extend_from_slice(b"=\"");
                            current_element_xml.extend_from_slice(&attr.value);
                            current_element_xml.push(b'"');
                        }
                        current_element_xml.push(b'>');
                    }
                    // Start of a table
                    else if e.local_name().as_ref() == b"tbl" && !in_table && !in_paragraph {
                        in_table = true;
                        depth = 1;
                        current_element_xml.clear();
                        current_element_xml.extend_from_slice(b"<w:tbl");
                        for attr in e.attributes().flatten() {
                            current_element_xml.push(b' ');
                            current_element_xml.extend_from_slice(attr.key.as_ref());
                            current_element_xml.extend_from_slice(b"=\"");
                            current_element_xml.extend_from_slice(&attr.value);
                            current_element_xml.push(b'"');
                        }
                        current_element_xml.push(b'>');
                    }
                    // Nested element inside paragraph or table
                    else if in_paragraph || in_table {
                        depth += 1;
                        current_element_xml.push(b'<');
                        current_element_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_element_xml.push(b' ');
                            current_element_xml.extend_from_slice(attr.key.as_ref());
                            current_element_xml.extend_from_slice(b"=\"");
                            current_element_xml.extend_from_slice(&attr.value);
                            current_element_xml.push(b'"');
                        }
                        current_element_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) => {
                    if in_paragraph || in_table {
                        current_element_xml.extend_from_slice(b"</");
                        current_element_xml.extend_from_slice(e.name().as_ref());
                        current_element_xml.push(b'>');

                        depth -= 1;

                        if depth == 0 && e.local_name().as_ref() == b"p" && in_paragraph {
                            // End of paragraph
                            let para_xml = std::mem::take(&mut current_element_xml);
                            elements.push(DocumentElement::Paragraph(
                                crate::document::Paragraph::Docx(Paragraph::new(para_xml)),
                            ));
                            in_paragraph = false;
                        } else if depth == 0 && e.local_name().as_ref() == b"tbl" && in_table {
                            // End of table
                            let table_xml = std::mem::take(&mut current_element_xml);
                            elements.push(DocumentElement::Table(crate::document::Table::Docx(
                                Table::new(table_xml),
                            )));
                            in_table = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_paragraph || in_table => {
                    current_element_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_paragraph || in_table => {
                    current_element_xml.push(b'<');
                    current_element_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_element_xml.push(b' ');
                        current_element_xml.extend_from_slice(attr.key.as_ref());
                        current_element_xml.extend_from_slice(b"=\"");
                        current_element_xml.extend_from_slice(&attr.value);
                        current_element_xml.push(b'"');
                    }
                    current_element_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(elements)
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as we have a way to construct test XmlParts
}
