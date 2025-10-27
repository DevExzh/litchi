/// Document - the main API for working with Word document content.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::parts::DocumentPart;
use crate::ooxml::docx::section::{Section, Sections};
use crate::ooxml::docx::styles::Styles;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::relationship_type;
use quick_xml::Reader;
use quick_xml::events::Event;

/// A Word document.
///
/// This is the main API for reading and manipulating Word document content.
/// It provides access to paragraphs, tables, sections, styles, and other
/// document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Get paragraph count
/// let count = doc.paragraph_count()?;
/// println!("Number of paragraphs: {}", count);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Document<'a> {
    /// The underlying document part
    part: DocumentPart<'a>,
    /// Reference to the OPC package (needed for accessing related parts like styles)
    opc: &'a OpcPackage,
}

impl<'a> Document<'a> {
    /// Create a new Document from a DocumentPart and OpcPackage reference.
    ///
    /// This is typically called internally by `Package::document()`.
    #[inline]
    pub(crate) fn new(part: DocumentPart<'a>, opc: &'a OpcPackage) -> Self {
        Self { part, opc }
    }

    /// Get all text content from the document.
    ///
    /// This extracts all text from all paragraphs in the document,
    /// concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.part.extract_text()
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        self.part.paragraph_count()
    }

    /// Get the number of tables in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.table_count()?;
    /// println!("Tables: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn table_count(&self) -> Result<usize> {
        self.part.table_count()
    }

    /// Get access to the underlying document part.
    ///
    /// This provides lower-level access to the document XML.
    #[inline]
    pub fn part(&self) -> &DocumentPart<'a> {
        &self.part
    }

    /// Get all paragraphs in the document.
    ///
    /// Returns a vector of `Paragraph` objects representing all `<w:p>`
    /// elements in the document body.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    ///
    ///     // Access runs within the paragraph
    ///     for run in para.runs()? {
    ///         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        // Convert SmallVec to Vec for API compatibility
        Ok(self.part.paragraphs()?.into_iter().collect())
    }

    /// Get all tables in the document.
    ///
    /// Returns a vector of `Table` objects representing all `<w:tbl>`
    /// elements in the document body.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for table in doc.tables()? {
    ///     println!("Table with {} rows", table.row_count()?);
    ///
    ///     for (row_idx, row) in table.rows()?.iter().enumerate() {
    ///         for (col_idx, cell) in row.cells()?.iter().enumerate() {
    ///             println!("Cell [{},{}]: {}", row_idx, col_idx, cell.text()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn tables(&self) -> Result<Vec<Table>> {
        // Convert SmallVec to Vec for API compatibility
        Ok(self.part.tables()?.into_iter().collect())
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method extracts both paragraphs and tables in a single pass,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    /// use litchi::DocumentElement;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for element in doc.elements()? {
    ///     match element {
    ///         DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         }
    ///         DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that is significantly faster than
    /// calling `paragraphs()` and `tables()` separately.
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        self.part.elements()
    }

    /// Get all sections in the document.
    ///
    /// Returns a `Sections` collection providing access to each section's
    /// page properties, margins, orientation, etc.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let mut sections = doc.sections()?;
    ///
    /// println!("Document has {} sections", sections.len());
    /// for section in sections.iter_mut() {
    ///     println!("Orientation: {}", section.orientation());
    ///     if let Some(width) = section.page_width() {
    ///         println!("  Page width: {} inches", width.to_inches());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn sections(&self) -> Result<Sections> {
        self.extract_sections()
    }

    /// Get the document styles.
    ///
    /// Returns a `Styles` object providing access to all paragraph, character,
    /// table, and list styles defined in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let mut styles = doc.styles()?;
    ///
    /// // Find a style by name
    /// if let Some(style) = styles.get_by_name("Heading 1")? {
    ///     println!("Found style: {} (id: {})",
    ///         style.name().unwrap_or(""),
    ///         style.style_id());
    /// }
    ///
    /// // Iterate all styles
    /// for style in styles.iter()? {
    ///     println!("Style: {} - Type: {}",
    ///         style.name().unwrap_or("<unnamed>"),
    ///         style.style_type());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn styles(&self) -> Result<Styles<'a>> {
        // Try to find the styles part through the main document part's relationships
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for a relationship to the styles part
        if let Ok(rel) = rels.part_with_reltype(relationship_type::STYLES) {
            let target = rel.target_partname()?;
            let styles_part = self.opc.get_part(&target)?;
            return Ok(Styles::from_part(styles_part));
        }

        // If no styles part is found, return an empty Styles object
        // This can happen in minimal documents
        Err(OoxmlError::PartNotFound(
            "styles part not found".to_string(),
        ))
    }

    /// Extract sections from the document XML.
    ///
    /// Sections are defined by `<w:sectPr>` elements, which can appear
    /// in two places:
    /// 1. Inside `<w:pPr>` (paragraph properties) - defines a section break
    /// 2. At the end of `<w:body>` - defines the last section
    fn extract_sections(&self) -> Result<Sections> {
        let xml_bytes = self.part.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut sections_xml = Vec::new();
        let mut buf = Vec::with_capacity(512);
        let mut depth = 0;
        let mut in_sect_pr = false;
        let mut sect_pr_content = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"sectPr" {
                        in_sect_pr = true;
                        depth = 1;
                        sect_pr_content.clear();
                        // Store the opening tag
                        sect_pr_content.extend_from_slice(b"<w:sectPr>");
                    } else if in_sect_pr {
                        depth += 1;
                        // Store the element
                        sect_pr_content.extend_from_slice(b"<");
                        sect_pr_content.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            sect_pr_content.extend_from_slice(b" ");
                            sect_pr_content.extend_from_slice(attr.key.as_ref());
                            sect_pr_content.extend_from_slice(b"=\"");
                            sect_pr_content.extend_from_slice(&attr.value);
                            sect_pr_content.extend_from_slice(b"\"");
                        }
                        sect_pr_content.extend_from_slice(b">");
                    }
                },
                Ok(Event::End(e)) => {
                    if in_sect_pr {
                        if e.local_name().as_ref() == b"sectPr" && depth == 1 {
                            // End of sectPr element
                            sect_pr_content.extend_from_slice(b"</w:sectPr>");
                            sections_xml.push(Section::from_xml_bytes(sect_pr_content.clone())?);
                            in_sect_pr = false;
                        } else {
                            depth -= 1;
                            sect_pr_content.extend_from_slice(b"</");
                            sect_pr_content.extend_from_slice(e.name().as_ref());
                            sect_pr_content.extend_from_slice(b">");
                        }
                    }
                },
                Ok(Event::Empty(e)) if in_sect_pr => {
                    // Self-closing element inside sectPr
                    sect_pr_content.extend_from_slice(b"<");
                    sect_pr_content.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        sect_pr_content.extend_from_slice(b" ");
                        sect_pr_content.extend_from_slice(attr.key.as_ref());
                        sect_pr_content.extend_from_slice(b"=\"");
                        sect_pr_content.extend_from_slice(&attr.value);
                        sect_pr_content.extend_from_slice(b"\"");
                    }
                    sect_pr_content.extend_from_slice(b"/>");
                },
                Ok(Event::Text(e)) if in_sect_pr => {
                    sect_pr_content.extend_from_slice(e.as_ref());
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        // If no sections were found, create a default section
        if sections_xml.is_empty() {
            sections_xml.push(Section::from_xml_bytes(b"<w:sectPr/>".to_vec())?);
        }

        Ok(Sections::new(sections_xml))
    }

    // TODO: Add more methods:
    // - add_paragraph() -> Paragraph (writing support)
    // - add_table() -> Table (writing support)
    // - save() (writing support)
}

// Note: Paragraph, Run, Table, Row, Cell, Section, Styles are now in separate modules:
// - paragraph.rs: Paragraph and Run
// - table.rs: Table, Row, Cell
// - section.rs: Section and Sections
// - styles.rs: Styles and Style

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
