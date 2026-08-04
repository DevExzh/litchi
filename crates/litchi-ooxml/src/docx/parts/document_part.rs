//! DocumentPart - the main document.xml part of a Word document.

use crate::docx::namespace::scan_word_element_ranges;
use crate::docx::paragraph::{Paragraph, extract_word_text};
use crate::docx::table::Table;
use crate::error::Result;
use litchi_docx::alt::{Chunk, active, scan};
use litchi_opc::part::Part;
use smallvec::SmallVec;
use std::collections::BTreeSet;
use std::sync::Arc;

/// The main document part of a Word document.
///
/// This corresponds to the `/word/document.xml` part in the package.
/// It contains the main document content including paragraphs, tables,
/// sections, and other block-level elements.
pub struct DocumentPart<'a> {
    /// Reference to the underlying part
    part: &'a dyn Part,
    raw: Arc<Vec<u8>>,
    xml: Arc<Vec<u8>>,
}

/// Select document-level Word blocks in original source order and coordinates.
///
/// Markup-compatibility preprocessing is used only as a visibility oracle; the
/// ranges continue to address the untouched package part.
pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let mut ranges = Vec::new();
    scan_word_element_ranges(
        xml,
        &[b"p".as_slice(), b"tbl".as_slice(), b"altChunk".as_slice()],
        |target, start, length| {
            ranges.push((target, start, length));
            Ok(())
        },
    )?;
    let starts = ranges
        .iter()
        .map(|&(_, start, _)| start)
        .collect::<Vec<_>>();
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)
}

impl<'a> DocumentPart<'a> {
    /// Return the original OPC part; semantic reads use the cached MCE view.
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Create a DocumentPart from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the document.xml content
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let raw = part.blob_arc();
        let xml = match litchi_ooxml_common::mce::process_ooxml(raw.as_slice())? {
            std::borrow::Cow::Borrowed(_) => Arc::clone(&raw),
            std::borrow::Cow::Owned(v) => Arc::new(v),
        };
        Ok(Self { part, raw, xml })
    }

    /// Get the shared Arc of XML bytes (zero-copy from Part).
    #[inline]
    fn get_xml_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    /// Get the original XML backing semantic source ranges.
    #[inline]
    fn get_raw_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.raw)
    }

    /// Get the XML bytes of the document.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Extract all paragraph text from the document.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML.
    ///
    /// # Performance
    ///
    /// Uses `quick-xml` for efficient streaming XML parsing with pre-allocated
    /// buffers and validated text decoding.
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
    pub fn elements(&self) -> Result<Vec<crate::docx::Element>> {
        use crate::docx::Element;

        Ok(self
            .blocks()?
            .into_iter()
            .filter_map(|block| match block {
                crate::docx::Block::Paragraph(paragraph) => Some(Element::Paragraph(paragraph)),
                crate::docx::Block::Table(table) => Some(Element::Table(table)),
                crate::docx::Block::Alt(_) => None,
            })
            .collect())
    }

    /// Get paragraphs, tables, and alternative-format anchors in document order.
    pub fn blocks(&self) -> Result<Vec<crate::docx::Block>> {
        use crate::docx::Block;

        let source = self.get_raw_arc();
        let mut alts = scan(source.as_slice())?;
        let mut elements = Vec::new();
        for (target, start, length) in active_block_ranges(source.as_slice())? {
            let block_source = Arc::clone(&source);
            elements.push(if target == 0 {
                Block::Paragraph(Box::new(Paragraph::from_arc_range(
                    block_source,
                    start,
                    length,
                )))
            } else if target == 1 {
                Block::Table(Box::new(Table::from_arc_range(block_source, start, length)))
            } else {
                let chunk = alts.remove(&start).ok_or_else(|| {
                    crate::error::OoxmlError::InvalidFormat(
                        "ordered altChunk lacks parsed anchor metadata".into(),
                    )
                })?;
                Block::Alt(Box::new(chunk))
            });
        }
        Ok(elements)
    }

    /// Return all alternative-format anchors in XML order.
    pub fn alts(&self) -> Result<Vec<Chunk>> {
        Ok(self
            .blocks()?
            .into_iter()
            .filter_map(|block| match block {
                crate::docx::Block::Alt(chunk) => Some(*chunk),
                crate::docx::Block::Paragraph(_) | crate::docx::Block::Table(_) => None,
            })
            .collect())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::Element;
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
        assert!(matches!(elements[0], Element::Paragraph(_)));
        assert!(matches!(elements[1], Element::Table(_)));
        assert!(matches!(elements[2], Element::Paragraph(_)));
        assert!(matches!(elements[3], Element::Paragraph(_)));
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
