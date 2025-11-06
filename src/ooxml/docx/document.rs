/// Document - the main API for working with Word document content.
use crate::ooxml::docx::bookmark::Bookmark;
use crate::ooxml::docx::comment::Comment;
use crate::ooxml::docx::content_control::ContentControl;
use crate::ooxml::docx::custom_xml::CustomXmlPart;
use crate::ooxml::docx::enums::WdHeaderFooter;
use crate::ooxml::docx::field::Field;
use crate::ooxml::docx::footnote::Note;
use crate::ooxml::docx::header_footer::HeaderFooter;
use crate::ooxml::docx::hyperlink::Hyperlink;
use crate::ooxml::docx::numbering::Numbering;
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::parts::DocumentPart;
use crate::ooxml::docx::section::{Section, Sections};
use crate::ooxml::docx::settings::DocumentSettings;
use crate::ooxml::docx::statistics::{
    DocumentStatistics, count_characters, count_characters_no_spaces, count_words,
    estimate_line_count, estimate_page_count,
};
use crate::ooxml::docx::styles::Styles;
use crate::ooxml::docx::table::Table;
use crate::ooxml::docx::theme::Theme;
use crate::ooxml::docx::variables::DocumentVariables;
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

    /// Get a specific paragraph by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the paragraph
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get first paragraph
    /// if let Some(para) = doc.paragraph(0)? {
    ///     println!("First paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        let paragraphs = self.paragraphs()?;
        Ok(paragraphs.into_iter().nth(index))
    }

    /// Get a specific table by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the table
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get first table
    /// if let Some(table) = doc.table(0)? {
    ///     println!("Table has {} rows", table.row_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn table(&self, index: usize) -> Result<Option<Table>> {
        let tables = self.tables()?;
        Ok(tables.into_iter().nth(index))
    }

    /// Extract all text from a specific range of paragraphs.
    ///
    /// # Arguments
    /// * `start` - Starting paragraph index (inclusive)
    /// * `end` - Ending paragraph index (exclusive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get text from paragraphs 5-10
    /// let text = doc.text_range(5, 10)?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text_range(&self, start: usize, end: usize) -> Result<String> {
        let paragraphs = self.paragraphs()?;
        let mut result = String::new();

        for (idx, para) in paragraphs.into_iter().enumerate() {
            if idx >= end {
                break;
            }
            if idx >= start {
                if !result.is_empty() {
                    result.push('\n');
                }
                result.push_str(&para.text()?);
            }
        }

        Ok(result)
    }

    /// Check if the document contains any tables.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if doc.has_tables()? {
    ///     println!("Document contains tables");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn has_tables(&self) -> Result<bool> {
        Ok(self.table_count()? > 0)
    }

    /// Get the underlying OPC package reference.
    ///
    /// This provides access to low-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        self.opc
    }

    /// Search for text in the document.
    ///
    /// Returns the indices of paragraphs that contain the search text.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-sensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Find paragraphs containing "important"
    /// let matches = doc.search("important")?;
    /// println!("Found {} matching paragraphs", matches.len());
    ///
    /// for idx in matches {
    ///     if let Some(para) = doc.paragraph(idx)? {
    ///         println!("Match in paragraph {}: {}", idx, para.text()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn search(&self, query: &str) -> Result<Vec<usize>> {
        let paragraphs = self.paragraphs()?;
        let mut matches = Vec::new();

        for (idx, para) in paragraphs.iter().enumerate() {
            if para.text()?.contains(query) {
                matches.push(idx);
            }
        }

        Ok(matches)
    }

    /// Search for text in the document (case-insensitive).
    ///
    /// Returns the indices of paragraphs that contain the search text.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-insensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Find paragraphs containing "important" (case-insensitive)
    /// let matches = doc.search_ignore_case("IMPORTANT")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn search_ignore_case(&self, query: &str) -> Result<Vec<usize>> {
        let paragraphs = self.paragraphs()?;
        let query_lower = query.to_lowercase();
        let mut matches = Vec::new();

        for (idx, para) in paragraphs.iter().enumerate() {
            if para.text()?.to_lowercase().contains(&query_lower) {
                matches.push(idx);
            }
        }

        Ok(matches)
    }

    /// Get all headers in the document.
    ///
    /// Returns a vector of tuples containing the header type and the header itself.
    /// Headers can be of three types: Primary (default), FirstPage, and EvenPage.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for (hdr_type, header) in doc.headers()? {
    ///     println!("{:?} header: {}", hdr_type, header.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn headers(&self) -> Result<Vec<(WdHeaderFooter, HeaderFooter)>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        let mut headers = Vec::new();

        // Iterate through all relationships looking for header parts
        for rel in rels.iter() {
            if rel.reltype() == relationship_type::HEADER {
                // Determine header type from the target name
                let target = rel.target_partname()?;
                let target_str = target.as_str();

                let hdr_type = if target_str.contains("header1.xml")
                    || target_str.contains("Header1.xml")
                {
                    WdHeaderFooter::Primary
                } else if target_str.contains("header2.xml") || target_str.contains("Header2.xml") {
                    WdHeaderFooter::FirstPage
                } else if target_str.contains("header3.xml") || target_str.contains("Header3.xml") {
                    WdHeaderFooter::EvenPage
                } else {
                    // Default to Primary if we can't determine
                    WdHeaderFooter::Primary
                };

                let header_part = self.opc.get_part(&target)?;
                headers.push((hdr_type, HeaderFooter::from_part(header_part, hdr_type)?));
            }
        }

        Ok(headers)
    }

    /// Get all footers in the document.
    ///
    /// Returns a vector of tuples containing the footer type and the footer itself.
    /// Footers can be of three types: Primary (default), FirstPage, and EvenPage.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for (ftr_type, footer) in doc.footers()? {
    ///     println!("{:?} footer: {}", ftr_type, footer.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footers(&self) -> Result<Vec<(WdHeaderFooter, HeaderFooter)>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        let mut footers = Vec::new();

        // Iterate through all relationships looking for footer parts
        for rel in rels.iter() {
            if rel.reltype() == relationship_type::FOOTER {
                // Determine footer type from the target name
                let target = rel.target_partname()?;
                let target_str = target.as_str();

                let ftr_type = if target_str.contains("footer1.xml")
                    || target_str.contains("Footer1.xml")
                {
                    WdHeaderFooter::Primary
                } else if target_str.contains("footer2.xml") || target_str.contains("Footer2.xml") {
                    WdHeaderFooter::FirstPage
                } else if target_str.contains("footer3.xml") || target_str.contains("Footer3.xml") {
                    WdHeaderFooter::EvenPage
                } else {
                    // Default to Primary if we can't determine
                    WdHeaderFooter::Primary
                };

                let footer_part = self.opc.get_part(&target)?;
                footers.push((ftr_type, HeaderFooter::from_part(footer_part, ftr_type)?));
            }
        }

        Ok(footers)
    }

    /// Get a specific header by type.
    ///
    /// # Arguments
    /// * `hdr_type` - The type of header to retrieve
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::{Package, WdHeaderFooter};
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(header) = doc.header(WdHeaderFooter::Primary)? {
    ///     println!("Primary header: {}", header.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn header(&self, hdr_type: WdHeaderFooter) -> Result<Option<HeaderFooter>> {
        let headers = self.headers()?;
        Ok(headers
            .into_iter()
            .find(|(t, _)| *t == hdr_type)
            .map(|(_, h)| h))
    }

    /// Get a specific footer by type.
    ///
    /// # Arguments
    /// * `ftr_type` - The type of footer to retrieve
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::{Package, WdHeaderFooter};
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(footer) = doc.footer(WdHeaderFooter::Primary)? {
    ///     println!("Primary footer: {}", footer.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footer(&self, ftr_type: WdHeaderFooter) -> Result<Option<HeaderFooter>> {
        let footers = self.footers()?;
        Ok(footers
            .into_iter()
            .find(|(t, _)| *t == ftr_type)
            .map(|(_, f)| f))
    }

    /// Get all hyperlinks in the document.
    ///
    /// Returns a vector of `Hyperlink` objects representing all hyperlinks
    /// found in the document. Both external hyperlinks (to URLs) and internal
    /// hyperlinks (to bookmarks) are included.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for link in doc.hyperlinks()? {
    ///     println!("Link text: {}", link.text());
    ///     if let Some(url) = link.url() {
    ///         println!("  URL: {}", url);
    ///     }
    ///     if let Some(anchor) = link.anchor() {
    ///         println!("  Bookmark: {}", anchor);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<Hyperlink>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();
        let xml_bytes = self.part.xml_bytes();

        Hyperlink::extract_from_document(xml_bytes, rels)
    }

    /// Get the number of hyperlinks in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} hyperlinks", doc.hyperlink_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn hyperlink_count(&self) -> Result<usize> {
        Ok(self.hyperlinks()?.len())
    }

    /// Get all footnotes in the document.
    ///
    /// Returns a vector of `Note` objects representing all footnotes
    /// in the document (excluding separators).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for note in doc.footnotes()? {
    ///     println!("Footnote {}: {}", note.id(), note.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footnotes(&self) -> Result<Vec<Note>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for footnotes relationship
        match rels.part_with_reltype(relationship_type::FOOTNOTES) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let footnotes_part = self.opc.get_part(&target)?;
                Note::extract_footnotes_from_part(footnotes_part)
            },
            Err(_) => {
                // No footnotes in document
                Ok(Vec::new())
            },
        }
    }

    /// Get all endnotes in the document.
    ///
    /// Returns a vector of `Note` objects representing all endnotes
    /// in the document (excluding separators).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for note in doc.endnotes()? {
    ///     println!("Endnote {}: {}", note.id(), note.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn endnotes(&self) -> Result<Vec<Note>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for endnotes relationship
        match rels.part_with_reltype(relationship_type::ENDNOTES) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let endnotes_part = self.opc.get_part(&target)?;
                Note::extract_endnotes_from_part(endnotes_part)
            },
            Err(_) => {
                // No endnotes in document
                Ok(Vec::new())
            },
        }
    }

    /// Get the number of footnotes in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} footnotes", doc.footnote_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footnote_count(&self) -> Result<usize> {
        Ok(self.footnotes()?.len())
    }

    /// Get the number of endnotes in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} endnotes", doc.endnote_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn endnote_count(&self) -> Result<usize> {
        Ok(self.endnotes()?.len())
    }

    /// Get all comments in the document.
    ///
    /// Returns a vector of `Comment` objects representing all comments
    /// in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for comment in doc.comments()? {
    ///     println!("{} commented: {}", comment.author(), comment.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for comments relationship
        match rels.part_with_reltype(relationship_type::COMMENTS) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let comments_part = self.opc.get_part(&target)?;
                Comment::extract_from_part(comments_part)
            },
            Err(_) => {
                // No comments in document
                Ok(Vec::new())
            },
        }
    }

    /// Get the number of comments in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} comments", doc.comment_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn comment_count(&self) -> Result<usize> {
        Ok(self.comments()?.len())
    }

    /// Get all bookmarks in the document.
    ///
    /// Returns a vector of `Bookmark` objects representing all bookmarks
    /// in the document (excluding system bookmarks).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for bookmark in doc.bookmarks()? {
    ///     println!("Bookmark: {}", bookmark.name());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn bookmarks(&self) -> Result<Vec<Bookmark>> {
        let xml_bytes = self.part.xml_bytes();
        Bookmark::extract_from_document(xml_bytes)
    }

    /// Get the number of bookmarks in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} bookmarks", doc.bookmark_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn bookmark_count(&self) -> Result<usize> {
        Ok(self.bookmarks()?.len())
    }

    /// Get all fields in the document.
    ///
    /// Returns a vector of `Field` objects representing all fields
    /// in the document (PAGE, DATE, REF, etc.).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for field in doc.fields()? {
    ///     println!("Field {}: {}", field.field_type(), field.instruction());
    ///     if let Some(result) = field.result() {
    ///         println!("  Result: {}", result);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn fields(&self) -> Result<Vec<Field>> {
        let xml_bytes = self.part.xml_bytes();
        Field::extract_from_document(xml_bytes)
    }

    /// Get the number of fields in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} fields", doc.field_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn field_count(&self) -> Result<usize> {
        Ok(self.fields()?.len())
    }

    /// Get the numbering definitions for the document.
    ///
    /// Returns a `Numbering` object providing access to abstract numbering
    /// definitions and numbering instances used for lists.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(numbering) = doc.numbering()? {
    ///     println!("Document has {} numbering definitions", numbering.num_count());
    ///     for num in numbering.nums() {
    ///         println!("Num ID {}: references abstract num {}",
    ///             num.id(), num.abstract_num_id());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn numbering(&self) -> Result<Option<Numbering>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for numbering relationship
        match rels.part_with_reltype(relationship_type::NUMBERING) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let numbering_part = self.opc.get_part(&target)?;
                Ok(Some(Numbering::extract_from_part(numbering_part)?))
            },
            Err(_) => {
                // No numbering in document
                Ok(None)
            },
        }
    }

    /// Get the document settings including protection status.
    ///
    /// Returns a `DocumentSettings` object providing access to document settings
    /// such as protection status, track revisions, and zoom level.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(settings) = doc.settings()? {
    ///     if settings.is_protected() {
    ///         println!("Document is protected");
    ///         if let Some(ptype) = settings.protection_type() {
    ///             println!("Protection type: {:?}", ptype);
    ///         }
    ///     }
    ///     println!("Track revisions: {}", settings.track_revisions());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn settings(&self) -> Result<Option<DocumentSettings>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for settings relationship
        match rels.part_with_reltype(relationship_type::SETTINGS) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let settings_part = self.opc.get_part(&target)?;
                Ok(Some(DocumentSettings::extract_from_part(settings_part)?))
            },
            Err(_) => {
                // No settings in document
                Ok(None)
            },
        }
    }

    /// Check if the document is protected.
    ///
    /// This is a convenience method that checks the settings for protection status.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if doc.is_protected()? {
    ///     println!("This document is protected");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn is_protected(&self) -> Result<bool> {
        Ok(self.settings()?.is_some_and(|s| s.is_protected()))
    }

    /// Get document variables.
    ///
    /// Returns document variables stored in the settings, which can be
    /// referenced by fields and used for mail merge.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(vars) = doc.document_variables()? {
    ///     for (name, value) in vars.iter() {
    ///         println!("{} = {}", name, value);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document_variables(&self) -> Result<Option<DocumentVariables>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        match rels.part_with_reltype(relationship_type::SETTINGS) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let settings_part = self.opc.get_part(&target)?;
                Ok(Some(DocumentVariables::extract_from_settings_part(
                    settings_part,
                )?))
            },
            Err(_) => Ok(None),
        }
    }

    /// Get the document theme.
    ///
    /// Returns the theme containing color scheme, font scheme, and format scheme.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(theme) = doc.theme()? {
    ///     if let Some(name) = theme.name() {
    ///         println!("Theme: {}", name);
    ///     }
    ///     if let Some(major) = theme.major_font() {
    ///         println!("Major font: {}", major);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn theme(&self) -> Result<Option<Theme>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        match rels.part_with_reltype(relationship_type::THEME) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let theme_part = self.opc.get_part(&target)?;
                Ok(Some(Theme::extract_from_part(theme_part)?))
            },
            Err(_) => Ok(None),
        }
    }

    /// Get all content controls in the document.
    ///
    /// Returns a vector of `ContentControl` objects representing structured
    /// content regions in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for control in doc.content_controls()? {
    ///     println!("Control ID {}", control.id());
    ///     if let Some(tag) = control.tag() {
    ///         println!("  Tag: {}", tag);
    ///     }
    ///     if let Some(control_type) = control.control_type() {
    ///         println!("  Type: {}", control_type);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn content_controls(&self) -> Result<Vec<ContentControl>> {
        let xml_bytes = self.part.xml_bytes();
        ContentControl::extract_from_document(xml_bytes)
    }

    /// Get custom XML parts from the document.
    ///
    /// Returns a vector of custom XML parts that store arbitrary XML data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for xml_part in doc.custom_xml_parts()? {
    ///     println!("Custom XML part: {}", xml_part.id());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn custom_xml_parts(&self) -> Result<Vec<CustomXmlPart>> {
        let mut custom_parts = Vec::new();

        // Custom XML parts are stored as relationships from the main document part
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Iterate through all relationships to find custom XML parts
        for rel in rels.iter() {
            if rel.reltype() == relationship_type::CUSTOM_XML {
                let target = rel.target_partname()?;
                let part = self.opc.get_part(&target)?;
                let id = rel.r_id().to_string();
                let custom_xml = CustomXmlPart::from_part(part, id)?;
                custom_parts.push(custom_xml);
            }
        }

        Ok(custom_parts)
    }

    /// Get document statistics.
    ///
    /// Calculates comprehensive statistics about the document including
    /// word count, character count, paragraph count, and other metrics.
    ///
    /// # Performance
    ///
    /// Statistics are calculated on-demand by parsing the entire document.
    /// For large documents, consider caching the result if you need to
    /// access statistics multiple times.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// let stats = doc.statistics()?;
    /// println!("Words: {}", stats.word_count());
    /// println!("Characters: {}", stats.character_count());
    /// println!("Paragraphs: {}", stats.paragraph_count());
    /// println!("Pages (estimate): {}", stats.page_count());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn statistics(&self) -> Result<DocumentStatistics> {
        let mut stats = DocumentStatistics::new();

        // Get all text content
        let text = self.text()?;

        // Calculate text statistics
        stats.set_word_count(count_words(&text));
        stats.set_character_count(count_characters(&text));
        stats.set_character_count_no_spaces(count_characters_no_spaces(&text));

        // Get paragraph and table counts
        stats.set_paragraph_count(self.paragraph_count()?);
        stats.set_table_count(self.table_count()?);

        // Estimate lines and pages (80 chars/line, 45 lines/page)
        let line_count = estimate_line_count(&text, 80);
        stats.set_line_count(line_count);
        stats.set_page_count(estimate_page_count(line_count, 45));

        // Count images and drawings across all paragraphs
        let mut image_count = 0;
        let mut drawing_count = 0;

        for para in self.paragraphs()? {
            image_count += para.images()?.len();
            drawing_count += para.drawing_objects()?.len();
        }

        stats.set_image_count(image_count);
        stats.set_drawing_count(drawing_count);

        Ok(stats)
    }

    // ========================================
    // READING FEATURES - ALL IMPLEMENTED ✅
    // ========================================
    // ✅ Text extraction: text(), paragraph_count(), table_count()
    // ✅ Paragraphs: paragraphs(), paragraph(), text_range()
    // ✅ Tables: tables(), table(), has_tables()
    // ✅ Sections: sections() with page properties, margins, orientation
    // ✅ Styles: styles() with full style information
    // ✅ Headers/Footers: headers(), footers(), header(), footer()
    // ✅ Hyperlinks: hyperlinks(), hyperlink_count()
    // ✅ Footnotes/Endnotes: footnotes(), endnotes(), footnote_count(), endnote_count()
    // ✅ Comments: comments(), comment_count()
    // ✅ Bookmarks: bookmarks(), bookmark_count()
    // ✅ Fields: fields(), field_count()
    // ✅ Numbering: numbering() with abstract numbering and instances
    // ✅ Document Settings: settings(), is_protected()
    // ✅ Document Variables: document_variables()
    // ✅ Statistics: statistics() with word/character/page counts
    // ✅ Theme: theme() with color and font schemes
    // ✅ Content Controls: content_controls()
    // ✅ Custom XML: custom_xml_parts()
    // ✅ Search: search(), search_ignore_case()
    //
    // ========================================
    // WRITE OPERATIONS
    // ========================================
    // Note: Write operations are primarily handled by the MutableDocument API
    // in the writer module. See src/ooxml/docx/writer/doc.rs for full API.
    //
    // TODO: Modification operations
    // - Add/remove paragraphs: add_paragraph(), remove_paragraph()
    // - Add/remove tables: add_table(), remove_table()
    // - Modify runs: set_bold(), set_italic(), set_font(), set_color()
    // - Insert elements: insert_paragraph(), insert_table()
    //
    // ✅ COMPLETED: Track changes reading (November 2024)
    // - See revision.rs module and Paragraph::revisions() method
    // - Full support for insert, delete, move, and format revisions
    // - Includes author, date, and revision ID tracking
    //
    // TODO: Table of contents (MS-DOCX Section 17.16.5)
    // - Requires parsing w:sdt elements with w:docPartGallery="Table of Contents"
    // - insert_toc(), update_toc(), remove_toc()
    //
    // TODO: Mail merge fields (MS-DOCX Section 17.16.5.35)
    // - Requires parsing w:fldSimple and w:fldChar elements with MERGEFIELD
    // - execute_mail_merge(), get_merge_fields()
    //
    // TODO: Page/Section breaks (MS-DOCX Section 17.3.3.3)
    // - Partially supported via Section API, advanced break types pending
    // - insert_page_break(), insert_section_break()
    //
    // TODO: Watermarks (MS-DOCX Section 17.10.2)
    // - Requires parsing VML shapes in headers with watermark styling
    // - add_watermark(), remove_watermark()
    //
    // ✅ COMPLETED: Images reading (November 2024)
    // - See image.rs module and Paragraph::images() method
    // - Full support for inline images with lazy loading
    //
    // ✅ COMPLETED: Drawing objects - shapes, text boxes (November 2024)
    // - See drawing.rs module and Paragraph::drawing_objects() method
    // - Full support for shapes, text boxes, inline/anchored positions
    // - 20+ standard shape types (rectangle, ellipse, arrows, etc.)
    //
    // TODO: Smart tags (MS-DOCX Section 17.5.1)
    // - Requires parsing w:smartTag elements with namespace URIs
    // - get_smart_tags(), add_smart_tag(), remove_smart_tag()
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
