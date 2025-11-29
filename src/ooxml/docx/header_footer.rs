/// Header and footer support for Word documents.
///
/// This module provides types and methods for accessing headers and footers
/// in Word documents. Headers and footers can be different for first page,
/// even/odd pages, and sections.
use crate::ooxml::docx::enums::WdHeaderFooter;
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A header or footer part in a Word document.
///
/// Headers and footers contain paragraphs and tables just like the main document body.
/// They are stored in separate XML parts (`/word/header*.xml` and `/word/footer*.xml`).
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Get all headers
/// let headers = doc.headers()?;
/// for (hdr_type, header) in headers {
///     println!("{:?} header: {}", hdr_type, header.text()?);
/// }
///
/// // Get all footers
/// let footers = doc.footers()?;
/// for (ftr_type, footer) in footers {
///     println!("{:?} footer: {}", ftr_type, footer.text()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    /// The raw XML bytes for this header/footer
    xml_bytes: Vec<u8>,
    /// The type of header/footer (default, first, even)
    hdr_ftr_type: WdHeaderFooter,
}

impl HeaderFooter {
    /// Create a new HeaderFooter from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the header/footer XML content
    /// * `hdr_ftr_type` - The type of header/footer
    pub fn from_part(part: &dyn Part, hdr_ftr_type: WdHeaderFooter) -> Result<Self> {
        Ok(Self {
            xml_bytes: part.blob().to_vec(),
            hdr_ftr_type,
        })
    }

    /// Create a new HeaderFooter from XML bytes.
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the header/footer element
    /// * `hdr_ftr_type` - The type of header/footer
    #[inline]
    pub fn from_xml_bytes(xml_bytes: Vec<u8>, hdr_ftr_type: WdHeaderFooter) -> Self {
        Self {
            xml_bytes,
            hdr_ftr_type,
        }
    }

    /// Get the type of this header/footer.
    #[inline]
    pub fn header_footer_type(&self) -> WdHeaderFooter {
        self.hdr_ftr_type
    }

    /// Get the XML bytes of this header/footer.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml_bytes
    }

    /// Extract all text content from this header/footer.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML, similar to how Document extracts text.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let headers = doc.headers()?;
    ///
    /// for (_, header) in headers {
    ///     println!("Header text: {}", header.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity to reduce reallocations
        let estimated_capacity = self.xml_bytes.len() / 8;
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Use unsafe conversion for better performance (safe since XML is validated)
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
        }

        result.shrink_to_fit();
        Ok(result)
    }

    /// Get all paragraphs in this header/footer.
    ///
    /// Returns a vector of `Paragraph` objects representing all `<w:p>`
    /// elements in the header/footer.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let headers = doc.headers()?;
    ///
    /// for (_, header) in headers {
    ///     for para in header.paragraphs()? {
    ///         println!("Paragraph: {}", para.text()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut paragraphs = Vec::new();
        let mut current_para_xml = Vec::with_capacity(4096);
        let mut in_para = false;
        let mut depth = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        current_para_xml.extend_from_slice(b"<w:p");
                        // Copy attributes
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b">");
                    } else if in_para {
                        depth += 1;
                        current_para_xml.extend_from_slice(b"<");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.extend_from_slice(b">");

                        if e.local_name().as_ref() == b"p" && depth == 1 {
                            paragraphs.push(Paragraph::new(current_para_xml.clone()));
                            in_para = false;
                        } else {
                            depth -= 1;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"<");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.extend_from_slice(b" ");
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.extend_from_slice(b"\"");
                        }
                        current_para_xml.extend_from_slice(b"/>");
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::CData(e)) if in_para => {
                    current_para_xml.extend_from_slice(b"<![CDATA[");
                    current_para_xml.extend_from_slice(e.as_ref());
                    current_para_xml.extend_from_slice(b"]]>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(paragraphs)
    }

    /// Get all tables in this header/footer.
    ///
    /// Returns a vector of `Table` objects representing all `<w:tbl>`
    /// elements in the header/footer.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let footers = doc.footers()?;
    ///
    /// for (_, footer) in footers {
    ///     for table in footer.tables()? {
    ///         println!("Table with {} rows", table.row_count()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn tables(&self) -> Result<Vec<Table>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut tables = Vec::new();
        let mut current_table_xml = Vec::with_capacity(8192);
        let mut in_table = false;
        let mut depth = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" && !in_table {
                        in_table = true;
                        depth = 1;
                        current_table_xml.clear();
                        current_table_xml.extend_from_slice(b"<w:tbl");
                        for attr in e.attributes().flatten() {
                            current_table_xml.extend_from_slice(b" ");
                            current_table_xml.extend_from_slice(attr.key.as_ref());
                            current_table_xml.extend_from_slice(b"=\"");
                            current_table_xml.extend_from_slice(&attr.value);
                            current_table_xml.extend_from_slice(b"\"");
                        }
                        current_table_xml.extend_from_slice(b">");
                    } else if in_table {
                        depth += 1;
                        current_table_xml.extend_from_slice(b"<");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_table_xml.extend_from_slice(b" ");
                            current_table_xml.extend_from_slice(attr.key.as_ref());
                            current_table_xml.extend_from_slice(b"=\"");
                            current_table_xml.extend_from_slice(&attr.value);
                            current_table_xml.extend_from_slice(b"\"");
                        }
                        current_table_xml.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_table {
                        current_table_xml.extend_from_slice(b"</");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        current_table_xml.extend_from_slice(b">");

                        if e.local_name().as_ref() == b"tbl" && depth == 1 {
                            tables.push(Table::new(current_table_xml.clone()));
                            in_table = false;
                        } else {
                            depth -= 1;
                        }
                    }
                },
                Ok(Event::Empty(e)) => {
                    if in_table {
                        current_table_xml.extend_from_slice(b"<");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_table_xml.extend_from_slice(b" ");
                            current_table_xml.extend_from_slice(attr.key.as_ref());
                            current_table_xml.extend_from_slice(b"=\"");
                            current_table_xml.extend_from_slice(&attr.value);
                            current_table_xml.extend_from_slice(b"\"");
                        }
                        current_table_xml.extend_from_slice(b"/>");
                    }
                },
                Ok(Event::Text(e)) if in_table => {
                    current_table_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::CData(e)) if in_table => {
                    current_table_xml.extend_from_slice(b"<![CDATA[");
                    current_table_xml.extend_from_slice(e.as_ref());
                    current_table_xml.extend_from_slice(b"]]>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(tables)
    }

    /// Get the number of paragraphs in this header/footer.
    pub fn paragraph_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut count = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"p" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(count)
    }

    /// Get the number of tables in this header/footer.
    pub fn table_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut count = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(count)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_text_extraction() {
        let xml = b"<w:hdr><w:p><w:r><w:t>Header Text</w:t></w:r></w:p></w:hdr>";
        let header = HeaderFooter::from_xml_bytes(xml.to_vec(), WdHeaderFooter::Primary);
        assert_eq!(header.text().unwrap(), "Header Text");
    }

    #[test]
    fn test_footer_text_extraction() {
        let xml = b"<w:ftr><w:p><w:r><w:t>Footer Text</w:t></w:r></w:p></w:ftr>";
        let footer = HeaderFooter::from_xml_bytes(xml.to_vec(), WdHeaderFooter::Primary);
        assert_eq!(footer.text().unwrap(), "Footer Text");
    }
}
