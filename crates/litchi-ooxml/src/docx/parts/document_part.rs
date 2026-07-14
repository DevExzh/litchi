//! DocumentPart - the main document.xml part of a Word document.

use crate::docx::namespace::is_wordprocessing_namespace;
use crate::docx::paragraph::{Paragraph, extract_word_text};
use crate::docx::table::Table;
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
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

#[derive(Clone, Copy)]
enum ElementKind {
    Paragraph,
    Table,
}

#[derive(Clone, Copy)]
struct ElementSelection {
    paragraphs: bool,
    tables: bool,
}

fn selected_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    selection: ElementSelection,
) -> Option<ElementKind> {
    if !is_wordprocessing_namespace(namespace) {
        return None;
    }
    match element.local_name().as_ref() {
        b"p" if selection.paragraphs => Some(ElementKind::Paragraph),
        b"tbl" if selection.tables => Some(ElementKind::Table),
        _ => None,
    }
}

fn scan_element_ranges(
    xml_bytes: &[u8],
    selection: ElementSelection,
    mut emit: impl FnMut(ElementKind, u32, u32) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start(ElementKind),
        NestedStart,
        Empty(ElementKind),
        End,
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    let mut capture: Option<(ElementKind, usize, usize)> = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("Word document offset does not fit usize".to_string())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element) => selected_element(&namespace, &element, selection)
                    .map_or(ScanEvent::Other, ScanEvent::Start),
                Event::Empty(element) if capture.is_none() => {
                    selected_element(&namespace, &element, selection)
                        .map_or(ScanEvent::Other, ScanEvent::Empty)
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("Word document offset does not fit usize".to_string())
        })?;

        match event {
            ScanEvent::Start(kind) => capture = Some((kind, event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word element nesting is too deep".to_string())
                })?;
            },
            ScanEvent::Empty(kind) => emit_element_range(kind, event_start, event_end, &mut emit)?,
            ScanEvent::End => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((kind, start, _)) = capture.take() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    emit_element_range(kind, start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word document element".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }

    Ok(())
}

fn emit_element_range(
    kind: ElementKind,
    start: usize,
    end: usize,
    emit: &mut impl FnMut(ElementKind, u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end
        .checked_sub(start)
        .ok_or_else(|| OoxmlError::InvalidFormat("invalid Word element byte range".to_string()))?;
    let start = u32::try_from(start)
        .map_err(|_| OoxmlError::InvalidFormat("Word element offset exceeds u32".to_string()))?;
    let length = u32::try_from(length)
        .map_err(|_| OoxmlError::InvalidFormat("Word element length exceeds u32".to_string()))?;
    emit(kind, start, length)
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
        scan_element_ranges(
            self.xml_bytes(),
            ElementSelection {
                paragraphs: true,
                tables: false,
            },
            |_, _, _| {
                count += 1;
                Ok(())
            },
        )?;
        Ok(count)
    }

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    pub fn table_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_element_ranges(
            self.xml_bytes(),
            ElementSelection {
                paragraphs: false,
                tables: true,
            },
            |_, _, _| {
                count += 1;
                Ok(())
            },
        )?;
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
        scan_element_ranges(
            source.as_slice(),
            ElementSelection {
                paragraphs: true,
                tables: false,
            },
            |_, start, length| {
                paragraphs.push(Paragraph::from_arc_range(
                    Arc::clone(&source),
                    start,
                    length,
                ));
                Ok(())
            },
        )?;
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
        scan_element_ranges(
            source.as_slice(),
            ElementSelection {
                paragraphs: false,
                tables: true,
            },
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
        scan_element_ranges(
            source.as_slice(),
            ElementSelection {
                paragraphs: true,
                tables: true,
            },
            |kind, start, length| {
                let source = Arc::clone(&source);
                elements.push(match kind {
                    ElementKind::Paragraph => DocxElement::Paragraph(Box::new(
                        Paragraph::from_arc_range(source, start, length),
                    )),
                    ElementKind::Table => {
                        DocxElement::Table(Box::new(Table::from_arc_range(source, start, length)))
                    },
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
