//! DocumentPart - the main document.xml part of a Word document.

use crate::docx::namespace::scan_word_element_ranges;
use crate::docx::paragraph::{Paragraph, extract_word_text};
use crate::docx::table::Table;
use crate::error::Result;
use litchi_opc::part::Part;
use smallvec::SmallVec;
use std::sync::Arc;

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

    /// Get the shared Arc of XML bytes (zero-copy from Part).
    #[inline]
    fn get_xml_arc(&self) -> Arc<Vec<u8>> {
        self.part.blob_arc()
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
        extract_word_text(self.xml_bytes())
    }

    /// Count the number of paragraphs in the document.
    ///
    /// Counts `<w:p>` elements in the document body.
    pub fn paragraph_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_word_element_ranges(self.xml_bytes(), &[b"p".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    pub fn table_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_word_element_ranges(self.xml_bytes(), &[b"tbl".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get all paragraphs in the document.
    ///
    /// Extracts all `<w:p>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        let source = self.get_xml_arc();
        let mut paragraphs = SmallVec::new();
        scan_word_element_ranges(source.as_slice(), &[b"p".as_slice()], |_, start, length| {
            paragraphs.push(Paragraph::from_arc_range(
                Arc::clone(&source),
                start,
                length,
            ));
            Ok(())
        })?;
        Ok(paragraphs)
    }

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        let source = self.get_xml_arc();
        let mut tables = SmallVec::new();
        scan_word_element_ranges(
            source.as_slice(),
            &[b"tbl".as_slice()],
            |_, start, length| {
                tables.push(Table::from_arc_range(Arc::clone(&source), start, length));
                Ok(())
            },
        )?;
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
    ///
    /// # Performance
    ///
    /// Uses one-pass, namespace-aware zero-copy parsing.
    pub fn elements(&self) -> Result<Vec<crate::docx::DocxElement>> {
        use crate::docx::DocxElement;

        let source = self.get_xml_arc();
        let mut elements = Vec::new();
        scan_word_element_ranges(
            source.as_slice(),
            &[b"p".as_slice(), b"tbl".as_slice()],
            |target, start, length| {
                let source = Arc::clone(&source);
                elements.push(if target == 0 {
                    DocxElement::Paragraph(Box::new(Paragraph::from_arc_range(
                        source, start, length,
                    )))
                } else {
                    DocxElement::Table(Box::new(Table::from_arc_range(source, start, length)))
                });
                Ok(())
            },
        )?;
        Ok(elements)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::DocxElement;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn document_part(xml: &[u8]) -> BlobPart {
        BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .to_string(),
            xml.to_vec(),
        )
    }

    #[test]
    fn extracts_aliased_word_elements_in_document_order_without_copying_text() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
            <wp:body>
                <false:p><false:r><false:t>ignored</false:t></false:r></false:p>
                <wp:p><wp:r><wp:t><![CDATA[A < B]]></wp:t></wp:r></wp:p>
                <wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl>
                <wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p>
                <wp:p/>
                <false:tbl/>
            </wp:body>
        </wp:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 4);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.tables().unwrap().len(), 1);

        let paragraphs = document.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 4);
        assert_eq!(paragraphs[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[0].runs().unwrap()[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[1].text().unwrap(), "cell");
        assert_eq!(paragraphs[2].text().unwrap(), "tail");
        assert_eq!(paragraphs[3].text().unwrap(), "");

        let elements = document.elements().unwrap();
        assert_eq!(elements.len(), 4);
        assert!(matches!(elements[0], DocxElement::Paragraph(_)));
        assert!(matches!(elements[1], DocxElement::Table(_)));
        assert!(matches!(elements[2], DocxElement::Paragraph(_)));
        assert!(matches!(elements[3], DocxElement::Paragraph(_)));
    }

    #[test]
    fn accepts_strict_wordprocessingml_and_self_closing_blocks() {
        let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:tbl/></s:body></s:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 1);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.elements().unwrap().len(), 2);
    }

    #[test]
    fn rejects_unterminated_selected_elements() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r/>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert!(document.paragraphs().is_err());
        assert!(document.elements().is_err());
    }
}
