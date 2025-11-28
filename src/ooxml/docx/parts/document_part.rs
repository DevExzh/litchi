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
    /// storage of typically small paragraph collections. Optimized to minimize
    /// allocations via pre-sized reserves.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        // Estimate paragraph count
        let estimated = (xml_bytes.len() / 400).max(8);
        let mut paragraphs = SmallVec::with_capacity(estimated);
        let mut current_para_xml = Vec::with_capacity(4096);
        let mut in_para = false;
        let mut depth = 0u32;
        let mut buf = Vec::with_capacity(2048);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        write_start_tag(&mut current_para_xml, b"<w:p", &e);
                    } else if in_para {
                        depth += 1;
                        write_start_tag_dynamic(&mut current_para_xml, e.name().as_ref(), &e);
                    }
                },
                Ok(Event::End(e)) => {
                    if in_para {
                        let name = e.name();
                        let name_ref = name.as_ref();
                        let close_len = 3 + name_ref.len();
                        current_para_xml.reserve(close_len);
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(name_ref);
                        current_para_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"p" {
                            paragraphs.push(Paragraph::new(std::mem::take(&mut current_para_xml)));
                            in_para = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    let text = e.as_ref();
                    current_para_xml.reserve(text.len());
                    current_para_xml.extend_from_slice(text);
                },
                Ok(Event::Empty(e)) if in_para => {
                    write_empty_tag(&mut current_para_xml, e.name().as_ref(), &e);
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
    /// Optimized to minimize allocations via pre-sized reserves.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut tables = SmallVec::new();
        let mut current_table_xml = Vec::with_capacity(8192);
        let mut in_table = false;
        let mut depth = 0u32;
        let mut buf = Vec::with_capacity(2048);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" && !in_table {
                        in_table = true;
                        depth = 1;
                        current_table_xml.clear();
                        write_start_tag(&mut current_table_xml, b"<w:tbl", &e);
                    } else if in_table {
                        depth += 1;
                        write_start_tag_dynamic(&mut current_table_xml, e.name().as_ref(), &e);
                    }
                },
                Ok(Event::End(e)) => {
                    if in_table {
                        let name = e.name();
                        let name_ref = name.as_ref();
                        let close_len = 3 + name_ref.len();
                        current_table_xml.reserve(close_len);
                        current_table_xml.extend_from_slice(b"</");
                        current_table_xml.extend_from_slice(name_ref);
                        current_table_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tbl" {
                            tables.push(Table::new(std::mem::take(&mut current_table_xml)));
                            in_table = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_table => {
                    let text = e.as_ref();
                    current_table_xml.reserve(text.len());
                    current_table_xml.extend_from_slice(text);
                },
                Ok(Event::Empty(e)) if in_table => {
                    write_empty_tag(&mut current_table_xml, e.name().as_ref(), &e);
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
    /// Optimized to minimize allocations by pre-sizing buffers and batching writes.
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        use crate::document::DocumentElement;

        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        // Estimate element count: typical DOCX has ~1 paragraph per 200-500 bytes
        let estimated_elements = (xml_bytes.len() / 300).max(16);
        let mut elements = Vec::with_capacity(estimated_elements);

        // Pre-allocate XML buffer; will be reused via mem::take
        let mut current_element_xml = Vec::with_capacity(8192);
        let mut in_paragraph = false;
        let mut in_table = false;
        let mut depth = 0u32;
        let mut buf = Vec::with_capacity(2048);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    let local_ref = local.as_ref();

                    // Start of a paragraph
                    if local_ref == b"p" && !in_paragraph && !in_table {
                        in_paragraph = true;
                        depth = 1;
                        current_element_xml.clear();
                        write_start_tag(&mut current_element_xml, b"<w:p", &e);
                    }
                    // Start of a table
                    else if local_ref == b"tbl" && !in_table && !in_paragraph {
                        in_table = true;
                        depth = 1;
                        current_element_xml.clear();
                        write_start_tag(&mut current_element_xml, b"<w:tbl", &e);
                    }
                    // Nested element inside paragraph or table
                    else if in_paragraph || in_table {
                        depth += 1;
                        write_start_tag_dynamic(&mut current_element_xml, e.name().as_ref(), &e);
                    }
                },
                Ok(Event::End(e)) => {
                    if in_paragraph || in_table {
                        let name = e.name();
                        let name_ref = name.as_ref();
                        // Write closing tag: "</name>"
                        let close_len = 3 + name_ref.len(); // "</" + name + ">"
                        current_element_xml.reserve(close_len);
                        current_element_xml.extend_from_slice(b"</");
                        current_element_xml.extend_from_slice(name_ref);
                        current_element_xml.push(b'>');

                        depth -= 1;

                        if depth == 0 {
                            let local = e.local_name();
                            let local_ref = local.as_ref();
                            if local_ref == b"p" && in_paragraph {
                                // End of paragraph
                                let para_xml = std::mem::take(&mut current_element_xml);
                                elements.push(DocumentElement::Paragraph(
                                    crate::document::Paragraph::Docx(Paragraph::new(para_xml)),
                                ));
                                in_paragraph = false;
                            } else if local_ref == b"tbl" && in_table {
                                // End of table
                                let table_xml = std::mem::take(&mut current_element_xml);
                                elements.push(DocumentElement::Table(
                                    crate::document::Table::Docx(Table::new(table_xml)),
                                ));
                                in_table = false;
                            }
                        }
                    }
                },
                Ok(Event::Text(e)) if in_paragraph || in_table => {
                    let text = e.as_ref();
                    current_element_xml.reserve(text.len());
                    current_element_xml.extend_from_slice(text);
                },
                Ok(Event::Empty(e)) if in_paragraph || in_table => {
                    write_empty_tag(&mut current_element_xml, e.name().as_ref(), &e);
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

/// Write a start tag with a known prefix (e.g., "<w:p") and attributes, ending with ">".
/// Pre-calculates total size to do a single reserve, avoiding repeated capacity checks.
#[inline]
fn write_start_tag(out: &mut Vec<u8>, prefix: &[u8], e: &quick_xml::events::BytesStart<'_>) {
    // First pass: calculate total size
    let mut attr_len = 0;
    for attr in e.attributes().flatten() {
        // Space + key + ="" + value
        attr_len += 1 + attr.key.as_ref().len() + 2 + attr.value.len() + 1;
    }

    let total = prefix.len() + attr_len + 1; // +1 for '>'
    out.reserve(total);
    out.extend_from_slice(prefix);

    // Second pass: write attributes
    for attr in e.attributes().flatten() {
        out.push(b' ');
        out.extend_from_slice(attr.key.as_ref());
        out.extend_from_slice(b"=\"");
        out.extend_from_slice(&attr.value);
        out.push(b'"');
    }
    out.push(b'>');
}

/// Write a start tag with dynamic tag name and attributes, ending with ">".
#[inline]
fn write_start_tag_dynamic(
    out: &mut Vec<u8>,
    tag_name: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
) {
    // First pass: calculate total size
    let mut attr_len = 0;
    for attr in e.attributes().flatten() {
        attr_len += 1 + attr.key.as_ref().len() + 2 + attr.value.len() + 1;
    }

    // "<" + tag_name + attrs + ">"
    let total = 1 + tag_name.len() + attr_len + 1;
    out.reserve(total);
    out.push(b'<');
    out.extend_from_slice(tag_name);

    // Second pass: write attributes
    for attr in e.attributes().flatten() {
        out.push(b' ');
        out.extend_from_slice(attr.key.as_ref());
        out.extend_from_slice(b"=\"");
        out.extend_from_slice(&attr.value);
        out.push(b'"');
    }
    out.push(b'>');
}

/// Write an empty tag "<name attrs/>".
#[inline]
fn write_empty_tag(out: &mut Vec<u8>, tag_name: &[u8], e: &quick_xml::events::BytesStart<'_>) {
    // First pass: calculate total size
    let mut attr_len = 0;
    for attr in e.attributes().flatten() {
        attr_len += 1 + attr.key.as_ref().len() + 2 + attr.value.len() + 1;
    }

    // "<" + tag_name + attrs + "/>"
    let total = 1 + tag_name.len() + attr_len + 2;
    out.reserve(total);
    out.push(b'<');
    out.extend_from_slice(tag_name);

    // Second pass: write attributes
    for attr in e.attributes().flatten() {
        out.push(b' ');
        out.extend_from_slice(attr.key.as_ref());
        out.extend_from_slice(b"=\"");
        out.extend_from_slice(&attr.value);
        out.push(b'"');
    }
    out.extend_from_slice(b"/>");
}

#[cfg(test)]
mod tests {
    // Tests will be added as we have a way to construct test XmlParts
}
