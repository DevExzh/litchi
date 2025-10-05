/// DocumentPart - the main document.xml part of a Word document.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::events::Event;
use quick_xml::Reader;

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
    /// Uses `quick-xml` for efficient streaming XML parsing with minimal
    /// allocations.
    pub fn extract_text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut result = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    // Check if this is a w:t element
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                }
                Ok(Event::Text(e)) if in_text_element => {
                    // Extract text content
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
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
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
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
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
    /// Uses streaming XML parsing to efficiently extract paragraph XML.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut paragraphs = Vec::new();
        let mut current_para_xml = Vec::new();
        let mut in_para = false;
        let mut depth = 0;
        let mut buf = Vec::new();

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

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    pub fn tables(&self) -> Result<Vec<Table>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut tables = Vec::new();
        let mut current_table_xml = Vec::new();
        let mut in_table = false;
        let mut depth = 0;
        let mut buf = Vec::new();

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
                }
                Ok(Event::End(e)) => {
                    if in_table {
                        current_table_xml.extend_from_slice(b"</");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        current_table_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tbl" {
                            tables.push(Table::new(current_table_xml.clone()));
                            in_table = false;
                        }
                    }
                }
                Ok(Event::Text(e)) if in_table => {
                    current_table_xml.extend_from_slice(e.as_ref());
                }
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
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(tables)
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as we have a way to construct test XmlParts
}
