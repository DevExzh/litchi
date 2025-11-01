/// Document writer implementation for DOCX.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::ImageFormat;
// Import from other writer modules
use super::note::Note;
use super::paragraph::{MutableParagraph, ParagraphElement};
use super::section::SectionProperties;
use super::table::MutableTable;

/// A mutable Word document for writing and modification.
///
/// Provides methods to add and modify document content including paragraphs,
/// runs, tables, sections, and other elements.
pub struct MutableDocument {
    /// Document body content (paragraphs, tables, etc.)
    body: DocumentBody,
    /// Header content (optional)
    header: Option<Vec<MutableParagraph>>,
    /// Footer content (optional)
    footer: Option<Vec<MutableParagraph>>,
    /// Footnotes (ID -> Note)
    footnotes: Vec<Note>,
    /// Endnotes (ID -> Note)
    endnotes: Vec<Note>,
    /// Section properties (page setup, margins, orientation)
    section: SectionProperties,
    /// Whether the document has been modified
    modified: bool,
}

impl MutableDocument {
    /// Create a new empty mutable document.
    pub fn new() -> Self {
        Self {
            body: DocumentBody::new(),
            header: None,
            footer: None,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            section: SectionProperties::default(),
            modified: false,
        }
    }

    /// Create a mutable document from existing XML content.
    pub fn from_xml(xml: &str) -> Result<Self> {
        let body = DocumentBody::from_xml(xml)?;
        Ok(Self {
            body,
            header: None,
            footer: None,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            section: SectionProperties::default(),
            modified: false,
        })
    }

    /// Get a mutable reference to the section properties.
    pub fn section_mut(&mut self) -> &mut SectionProperties {
        self.modified = true;
        &mut self.section
    }

    /// Get a reference to the section properties.
    pub fn section(&self) -> &SectionProperties {
        &self.section
    }

    /// Add a new paragraph to the end of the document.
    pub fn add_paragraph(&mut self) -> &mut MutableParagraph {
        self.modified = true;
        self.body.add_paragraph()
    }

    /// Add a paragraph with text.
    pub fn add_paragraph_with_text(&mut self, text: &str) -> &mut MutableParagraph {
        let para = self.add_paragraph();
        para.add_run_with_text(text);
        para
    }

    /// Add a heading paragraph.
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<&mut MutableParagraph> {
        if level > 9 {
            return Err(OoxmlError::InvalidFormat(
                "Heading level must be 0-9".to_string(),
            ));
        }
        let style = if level == 0 {
            "Title".to_string()
        } else {
            format!("Heading {}", level)
        };
        let para = self.add_paragraph();
        para.set_style(&style);
        para.add_run_with_text(text);
        Ok(para)
    }

    /// Add a table with specified rows and columns.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> &mut MutableTable {
        self.modified = true;
        self.body.add_table(rows, cols)
    }

    /// Add a page break.
    pub fn add_page_break(&mut self) -> &mut MutableParagraph {
        let para = self.add_paragraph();
        para.add_run().add_page_break();
        para
    }

    /// Get the number of paragraphs in the document.
    pub fn paragraph_count(&self) -> usize {
        self.body.paragraph_count()
    }

    /// Get the number of tables in the document.
    pub fn table_count(&self) -> usize {
        self.body.table_count()
    }

    /// Check if the document has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    /// Get or create the header.
    pub fn header(&mut self) -> &mut Vec<MutableParagraph> {
        if self.header.is_none() {
            self.header = Some(Vec::new());
            self.modified = true;
        }
        self.header.as_mut().unwrap()
    }

    /// Get or create the footer.
    pub fn footer(&mut self) -> &mut Vec<MutableParagraph> {
        if self.footer.is_none() {
            self.footer = Some(Vec::new());
            self.modified = true;
        }
        self.footer.as_mut().unwrap()
    }

    /// Check if the document has a header.
    pub fn has_header(&self) -> bool {
        self.header.is_some()
    }

    /// Check if the document has a footer.
    pub fn has_footer(&self) -> bool {
        self.footer.is_some()
    }

    /// Add a header to the document.
    pub fn add_header_paragraph(&mut self) -> &mut MutableParagraph {
        if self.header.is_none() {
            self.header = Some(Vec::new());
        }
        let para = MutableParagraph::new();
        self.header.as_mut().unwrap().push(para);
        self.modified = true;
        self.header.as_mut().unwrap().last_mut().unwrap()
    }

    /// Add a footer to the document.
    pub fn add_footer_paragraph(&mut self) -> &mut MutableParagraph {
        if self.footer.is_none() {
            self.footer = Some(Vec::new());
        }
        let para = MutableParagraph::new();
        self.footer.as_mut().unwrap().push(para);
        self.modified = true;
        self.footer.as_mut().unwrap().last_mut().unwrap()
    }

    /// Add a footnote and return its ID and mutable reference.
    pub fn add_footnote(&mut self) -> (u32, &mut Note) {
        let id = self.footnotes.len() as u32 + 1;
        let note = Note::new(id);
        self.footnotes.push(note);
        self.modified = true;
        (id, self.footnotes.last_mut().unwrap())
    }

    /// Add an endnote and return its ID and mutable reference.
    pub fn add_endnote(&mut self) -> (u32, &mut Note) {
        let id = self.endnotes.len() as u32 + 1;
        let note = Note::new(id);
        self.endnotes.push(note);
        self.modified = true;
        (id, self.endnotes.last_mut().unwrap())
    }

    /// Check if the document has footnotes.
    pub fn has_footnotes(&self) -> bool {
        !self.footnotes.is_empty()
    }

    /// Check if the document has endnotes.
    pub fn has_endnotes(&self) -> bool {
        !self.endnotes.is_empty()
    }

    /// Collect all hyperlink URLs from the document in order.
    ///
    /// Note: This collects ALL hyperlinks, not just unique URLs. Each hyperlink
    /// gets its own relationship ID, even if multiple hyperlinks point to the same URL.
    /// This matches the behavior of Microsoft Word and python-docx.
    pub(crate) fn collect_hyperlink_urls(&self) -> Vec<String> {
        let mut urls = Vec::new();

        for element in &self.body.elements {
            if let BodyElement::Paragraph(para) = element {
                for para_element in &para.elements {
                    if let ParagraphElement::Hyperlink(link) = para_element {
                        urls.push(link.url.clone());
                    }
                }
            }
        }

        urls
    }

    /// Collect all images from the document.
    pub(crate) fn collect_images(&self) -> Vec<(&[u8], ImageFormat)> {
        let mut images = Vec::new();

        for element in &self.body.elements {
            if let BodyElement::Paragraph(para) = element {
                for para_element in &para.elements {
                    if let ParagraphElement::InlineImage(image) = para_element {
                        images.push((image.data(), image.format()));
                    }
                }
            }
        }

        images
    }

    /// Generate header XML content.
    #[allow(dead_code)]
    pub(crate) fn generate_header_xml(&self) -> Result<Option<String>> {
        if self.header.is_none() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
        );
        if let Some(ref paragraphs) = self.header {
            if paragraphs.is_empty() {
                xml.push_str(r#"<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>"#);
            } else {
                for para in paragraphs {
                    para.to_xml(&mut xml)?;
                }
            }
        }
        xml.push_str("</w:hdr>");
        Ok(Some(xml))
    }

    /// Generate footer XML content.
    #[allow(dead_code)]
    pub(crate) fn generate_footer_xml(&self) -> Result<Option<String>> {
        if self.footer.is_none() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
        );
        if let Some(ref paragraphs) = self.footer {
            if paragraphs.is_empty() {
                xml.push_str(r#"<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>"#);
            } else {
                for para in paragraphs {
                    para.to_xml(&mut xml)?;
                }
            }
        }
        xml.push_str("</w:ftr>");
        Ok(Some(xml))
    }

    /// Generate footnotes XML content.
    pub(crate) fn generate_footnotes_xml(&self) -> Result<Option<String>> {
        if self.footnotes.is_empty() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);

        xml.push_str(r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"#);
        xml.push_str(r#"<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#);

        for note in &self.footnotes {
            write!(xml, r#"<w:footnote w:id="{}">"#, note.id)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;

            if note.paragraphs.is_empty() {
                xml.push_str("<w:p/>");
            } else {
                for para in &note.paragraphs {
                    para.to_xml(&mut xml)?;
                }
            }

            xml.push_str("</w:footnote>");
        }

        xml.push_str("</w:footnotes>");
        Ok(Some(xml))
    }

    /// Generate endnotes XML content.
    pub(crate) fn generate_endnotes_xml(&self) -> Result<Option<String>> {
        if self.endnotes.is_empty() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);

        xml.push_str(r#"<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>"#);
        xml.push_str(r#"<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>"#);

        for note in &self.endnotes {
            write!(xml, r#"<w:endnote w:id="{}">"#, note.id)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;

            if note.paragraphs.is_empty() {
                xml.push_str("<w:p/>");
            } else {
                for para in &note.paragraphs {
                    para.to_xml(&mut xml)?;
                }
            }

            xml.push_str("</w:endnote>");
        }

        xml.push_str("</w:endnotes>");
        Ok(Some(xml))
    }

    /// Get a reference to a paragraph by index.
    pub fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        self.body.paragraph(index)
    }

    /// Get a reference to a table by index.
    pub fn table(&mut self, index: usize) -> Option<&mut MutableTable> {
        self.body.table(index)
    }

    /// Serialize the document to XML.
    pub fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(4096);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
        self.body.to_xml(&mut xml)?;
        xml.push_str("</w:document>");
        Ok(xml)
    }

    /// Generate XML with actual relationship IDs from the mapper.
    ///
    /// This is the correct method to use when saving documents, as it includes
    /// proper relationship IDs and section properties with header/footer references.
    pub(crate) fn to_xml_with_rels(
        &self,
        rel_mapper: &super::relmap::RelationshipMapper,
    ) -> Result<String> {
        let mut xml = String::with_capacity(4096);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);

        // Generate body with relationship IDs
        self.body.to_xml_with_rels(&mut xml, rel_mapper)?;

        // Add section properties at the end of the body (before </w:body>)
        // The sectPr must be the last element in the body
        self.generate_section_properties(&mut xml, rel_mapper)?;

        xml.push_str("</w:body>");
        xml.push_str("</w:document>");
        Ok(xml)
    }

    /// Generate section properties XML including header/footer references.
    fn generate_section_properties(
        &self,
        xml: &mut String,
        rel_mapper: &super::relmap::RelationshipMapper,
    ) -> Result<()> {
        xml.push_str("<w:sectPr>");

        // Add header reference if present
        if let Some(header_id) = rel_mapper.get_header_id() {
            write!(
                xml,
                r#"<w:headerReference w:type="default" r:id="{}"/>"#,
                header_id
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // Add footer reference if present
        if let Some(footer_id) = rel_mapper.get_footer_id() {
            write!(
                xml,
                r#"<w:footerReference w:type="default" r:id="{}"/>"#,
                footer_id
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // Add page size and margins
        write!(
            xml,
            r#"<w:pgSz w:w="{}" w:h="{}" w:orient="{}"/>"#,
            self.section.page_width,
            self.section.page_height,
            self.section.orientation.as_str()
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        write!(
            xml,
            r#"<w:pgMar w:top="{}" w:right="{}" w:bottom="{}" w:left="{}" w:header="{}" w:footer="{}"/>"#,
            self.section.margin_top,
            self.section.margin_right,
            self.section.margin_bottom,
            self.section.margin_left,
            self.section.header_distance,
            self.section.footer_distance
        ).map_err(|e| OoxmlError::Xml(e.to_string()))?;

        xml.push_str("</w:sectPr>");
        Ok(())
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}

/// The document body containing all content elements.
#[derive(Debug)]
pub(crate) struct DocumentBody {
    /// Content elements (paragraphs, tables, etc.) in document order
    pub(crate) elements: Vec<BodyElement>,
}

impl DocumentBody {
    fn new() -> Self {
        Self {
            elements: Vec::new(),
        }
    }

    fn from_xml(_xml: &str) -> Result<Self> {
        // TODO: Implement XML parsing for reading existing documents
        Ok(Self::new())
    }

    fn add_paragraph(&mut self) -> &mut MutableParagraph {
        self.elements
            .push(BodyElement::Paragraph(MutableParagraph::new()));
        match self.elements.last_mut() {
            Some(BodyElement::Paragraph(p)) => p,
            _ => unreachable!(),
        }
    }

    fn add_table(&mut self, rows: usize, cols: usize) -> &mut MutableTable {
        self.elements
            .push(BodyElement::Table(MutableTable::new(rows, cols)));
        match self.elements.last_mut() {
            Some(BodyElement::Table(t)) => t,
            _ => unreachable!(),
        }
    }

    fn paragraph_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|e| matches!(e, BodyElement::Paragraph(_)))
            .count()
    }

    fn table_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|e| matches!(e, BodyElement::Table(_)))
            .count()
    }

    fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        let mut count = 0;
        for elem in &mut self.elements {
            if let BodyElement::Paragraph(p) = elem {
                if count == index {
                    return Some(p);
                }
                count += 1;
            }
        }
        None
    }

    fn table(&mut self, index: usize) -> Option<&mut MutableTable> {
        let mut count = 0;
        for elem in &mut self.elements {
            if let BodyElement::Table(t) = elem {
                if count == index {
                    return Some(t);
                }
                count += 1;
            }
        }
        None
    }

    fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:body>");

        for element in &self.elements {
            match element {
                BodyElement::Paragraph(p) => p.to_xml(xml)?,
                BodyElement::Table(t) => t.to_xml(xml)?,
            }
        }

        // Add default section properties
        xml.push_str("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>");
        xml.push_str(
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>",
        );
        xml.push_str("</w:sectPr>");

        xml.push_str("</w:body>");

        Ok(())
    }

    /// Generate XML with actual relationship IDs from the mapper.
    /// Note: Does not close </w:body> tag - caller must add sectPr and close it.
    fn to_xml_with_rels(
        &self,
        xml: &mut String,
        rel_mapper: &crate::ooxml::docx::writer::relmap::RelationshipMapper,
    ) -> Result<()> {
        xml.push_str("<w:body>");

        // Global counters for hyperlinks and images across all paragraphs
        let mut hyperlink_counter = 0;
        let mut image_counter = 0;

        for element in &self.elements {
            match element {
                BodyElement::Paragraph(p) => {
                    p.to_xml_with_rels(
                        xml,
                        rel_mapper,
                        &mut hyperlink_counter,
                        &mut image_counter,
                    )?;
                },
                BodyElement::Table(t) => t.to_xml(xml)?, // Tables don't need rel mapping for now
            }
        }

        // Note: Don't close </w:body> here - it will be closed after sectPr is added
        Ok(())
    }
}

/// A body element (paragraph or table).
#[derive(Debug)]
pub(crate) enum BodyElement {
    Paragraph(MutableParagraph),
    Table(MutableTable),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_empty_document() {
        let doc = MutableDocument::new();
        assert_eq!(doc.paragraph_count(), 0);
        assert_eq!(doc.table_count(), 0);
    }

    #[test]
    fn test_add_paragraph() {
        let mut doc = MutableDocument::new();
        doc.add_paragraph_with_text("Hello, World!");
        assert_eq!(doc.paragraph_count(), 1);
    }

    #[test]
    fn test_add_table() {
        let mut doc = MutableDocument::new();
        let table = doc.add_table(2, 3);
        assert_eq!(table.row_count(), 2);
        table.cell(0, 0).unwrap().set_text("Cell 1");
        assert_eq!(doc.table_count(), 1);
    }

    #[test]
    fn test_xml_generation() {
        let mut doc = MutableDocument::new();
        doc.add_paragraph_with_text("Test paragraph");

        let xml = doc.to_xml().unwrap();
        assert!(xml.contains("<w:document"));
        assert!(xml.contains("<w:body>"));
        assert!(xml.contains("<w:p>"));
        assert!(xml.contains("Test paragraph"));
    }

    #[test]
    fn test_run_formatting() {
        let mut doc = MutableDocument::new();
        let para = doc.add_paragraph();
        para.add_run_with_text("Bold text").bold(true);
        para.add_run_with_text("Italic text").italic(true);

        let xml = doc.to_xml().unwrap();
        assert!(xml.contains("<w:b/>"));
        assert!(xml.contains("<w:i/>"));
    }
}
