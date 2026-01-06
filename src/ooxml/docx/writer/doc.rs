/// Document writer implementation for DOCX.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::ImageFormat;
// Import from other writer modules
use super::comment::MutableComment;
use super::note::Note;
use super::paragraph::{MutableParagraph, ParagraphElement};
use super::section::SectionProperties;
use super::table::MutableTable;
use super::theme::MutableTheme;
use super::toc::TableOfContents;
use super::watermark::Watermark;
// Import settings types
use super::super::settings::ProtectionType;

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
    /// Comments (ID -> Comment)
    comments: Vec<MutableComment>,
    /// Document protection settings
    protection: Option<DocumentProtection>,
    /// Section properties (page setup, margins, orientation)
    section: SectionProperties,
    /// Theme (optional)
    theme: Option<MutableTheme>,
    /// Watermark (optional)
    pub(crate) watermark: Option<Watermark>,
    /// Table of Contents configuration (optional)
    toc_config: Option<(usize, TableOfContents)>, // (insertion index, config)
    /// Whether the document has been modified
    modified: bool,
}

/// Document protection settings.
#[derive(Debug, Clone)]
pub struct DocumentProtection {
    /// Type of protection
    pub protection_type: ProtectionType,
    /// Password hash (optional, for actual enforcement)
    pub password_hash: Option<String>,
    /// Salt for password hash (optional)
    pub salt: Option<String>,
}

#[cfg(feature = "fonts")]
use crate::fonts::CollectGlyphs;
#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;
#[cfg(feature = "fonts")]
use std::collections::HashMap;

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableDocument {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();

        // Collect from body elements
        for element in &self.body.elements {
            let element_glyphs = match element {
                BodyElement::Paragraph(p) => p.collect_glyphs(),
                BodyElement::Table(t) => t.collect_glyphs(),
            };
            for (font, bitmap) in element_glyphs {
                *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
            }
        }

        // Collect from headers
        if let Some(headers) = &self.header {
            for p in headers {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                }
            }
        }

        // Collect from footers
        if let Some(footers) = &self.footer {
            for p in footers {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                }
            }
        }

        // Collect from footnotes/endnotes
        for note in self.footnotes.iter().chain(self.endnotes.iter()) {
            for p in &note.paragraphs {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                }
            }
        }

        glyphs
    }
}

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableParagraph {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();
        for element in &self.elements {
            let element_glyphs = match element {
                ParagraphElement::Run(r) => r.collect_glyphs(),
                ParagraphElement::Hyperlink(h) => h.collect_glyphs(),
                _ => continue,
            };
            for (font, bitmap) in element_glyphs {
                *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
            }
        }
        glyphs
    }
}

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableTable {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();
        for row in &self.rows {
            for cell in &row.cells {
                for p in &cell.paragraphs {
                    for (font, bitmap) in p.collect_glyphs() {
                        *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                    }
                }
            }
        }
        glyphs
    }
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
            comments: Vec::new(),
            protection: None,
            toc_config: None,
            section: SectionProperties::default(),
            theme: None,
            watermark: None,
            modified: false,
        }
    }

    /// Create a mutable document from existing XML content.
    pub fn from_xml(xml: &str) -> Result<Self> {
        let body = DocumentBody::from_xml(xml)?;
        Ok(Self {
            body,
            toc_config: None,
            header: None,
            footer: None,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            comments: Vec::new(),
            protection: None,
            section: SectionProperties::default(),
            theme: None,
            watermark: None,
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

    /// Add a comment and return its ID and mutable reference.
    ///
    /// # Arguments
    ///
    /// * `author` - Comment author name
    /// * `text` - Comment text content
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let (comment_id, comment) = doc.add_comment("John Doe", "This needs revision");
    /// comment.set_initials(Some("JD".to_string()));
    /// ```
    pub fn add_comment(&mut self, author: &str, text: &str) -> (u32, &mut MutableComment) {
        let id = self.comments.len() as u32 + 1;
        let comment = MutableComment::new(id, author.to_string(), text.to_string());
        self.comments.push(comment);
        self.modified = true;
        (id, self.comments.last_mut().unwrap())
    }

    /// Check if the document has comments.
    pub fn has_comments(&self) -> bool {
        !self.comments.is_empty()
    }

    /// Get the number of comments in the document.
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Set document protection.
    ///
    /// # Arguments
    ///
    /// * `protection_type` - Type of protection to apply
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use litchi::ooxml::docx::settings::ProtectionType;
    ///
    /// // Protect document as read-only
    /// doc.set_protection(ProtectionType::ReadOnly);
    ///
    /// // Allow only comments
    /// doc.set_protection(ProtectionType::Comments);
    /// ```
    pub fn set_protection(&mut self, protection_type: ProtectionType) {
        self.protection = Some(DocumentProtection {
            protection_type,
            password_hash: None,
            salt: None,
        });
        self.modified = true;
    }

    /// Set document protection with password.
    ///
    /// Note: For simplicity, this implementation stores the hash directly.
    /// In a production system, you would use proper password hashing (SHA-256, etc.).
    ///
    /// # Arguments
    ///
    /// * `protection_type` - Type of protection to apply
    /// * `password_hash` - Password hash (from proper hashing algorithm)
    /// * `salt` - Salt used for password hashing
    pub fn set_protection_with_password(
        &mut self,
        protection_type: ProtectionType,
        password_hash: String,
        salt: String,
    ) {
        self.protection = Some(DocumentProtection {
            protection_type,
            password_hash: Some(password_hash),
            salt: Some(salt),
        });
        self.modified = true;
    }

    /// Remove document protection.
    pub fn remove_protection(&mut self) {
        self.protection = None;
        self.modified = true;
    }

    /// Check if the document has protection set.
    pub fn is_protected(&self) -> bool {
        self.protection.is_some()
    }

    /// Get the protection type if set.
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.protection.as_ref().map(|p| p.protection_type)
    }

    /// Set the document theme.
    ///
    /// # Arguments
    ///
    /// * `theme` - Theme to apply to the document
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use litchi::ooxml::docx::writer::MutableTheme;
    ///
    /// let theme = MutableTheme::office_theme();
    /// doc.set_theme(theme);
    /// ```
    pub fn set_theme(&mut self, theme: MutableTheme) {
        self.theme = Some(theme);
        self.modified = true;
    }

    /// Get a reference to the document theme.
    pub fn theme(&self) -> Option<&MutableTheme> {
        self.theme.as_ref()
    }

    /// Get a mutable reference to the document theme.
    pub fn theme_mut(&mut self) -> Option<&mut MutableTheme> {
        self.modified = true;
        self.theme.as_mut()
    }

    /// Set a watermark for the document.
    ///
    /// # Arguments
    ///
    /// * `watermark` - Watermark to apply
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use litchi::ooxml::docx::writer::Watermark;
    ///
    /// let watermark = Watermark::text("CONFIDENTIAL");
    /// doc.set_watermark(watermark);
    /// ```
    pub fn set_watermark(&mut self, watermark: Watermark) {
        self.watermark = Some(watermark);
        self.modified = true;
    }

    /// Remove the watermark from the document.
    pub fn remove_watermark(&mut self) {
        if self.watermark.is_some() {
            self.watermark = None;
            self.modified = true;
        }
    }

    /// Check if the document has a watermark.
    pub fn has_watermark(&self) -> bool {
        self.watermark.is_some()
    }

    /// Add a table of contents at the current position in the document.
    ///
    /// # Arguments
    ///
    /// * `toc` - Table of contents configuration
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use litchi::ooxml::docx::writer::TableOfContents;
    ///
    /// let toc = TableOfContents::new()
    ///     .heading_levels(1, 3)
    ///     .title("Contents");
    /// doc.add_toc(toc);
    /// ```
    pub fn add_toc(&mut self, toc: TableOfContents) -> Result<()> {
        // Add optional title paragraph with TOCHeading style
        if let Some(title) = toc.get_title() {
            let title_para = self.add_paragraph();
            title_para.set_style("TOCHeading");
            let title_run = title_para.add_run();
            title_run.set_text(title);
        }

        // Record the insertion point (after the title if present)
        let insertion_index = self.body.elements.len();

        // Store the TOC configuration for later generation (at save time)
        self.toc_config = Some((insertion_index, toc));

        self.modified = true;
        Ok(())
    }

    /// Generate and insert TOC entries.
    /// This is called automatically before serialization.
    pub(crate) fn generate_toc_if_needed(&mut self) -> Result<()> {
        use super::field::MutableField;
        use std::fmt::Write as FmtWrite;

        // Check if we have a TOC to generate
        let Some((insertion_index, toc)) = self.toc_config.take() else {
            return Ok(());
        };

        // Step 1: Scan document for headings and add bookmarks
        let mut heading_info = Vec::new();
        let mut bookmark_counter = 0u32;
        let start_level = toc.start_level();
        let end_level = toc.end_level();

        // Iterate through all body elements to find headings
        for element in &mut self.body.elements {
            if let BodyElement::Paragraph(para) = element
                && let Some(style) = &para.style
            {
                // Check if this is a heading within our TOC range
                let heading_level = match style.as_str() {
                    "Heading1" => Some(1),
                    "Heading2" => Some(2),
                    "Heading3" => Some(3),
                    "Heading4" => Some(4),
                    "Heading5" => Some(5),
                    "Heading6" => Some(6),
                    "Heading7" => Some(7),
                    "Heading8" => Some(8),
                    "Heading9" => Some(9),
                    _ => None,
                };

                if let Some(level) = heading_level
                    && level >= start_level
                    && level <= end_level
                {
                    // Extract heading text
                    let mut heading_text = String::new();
                    for elem in &para.elements {
                        if let super::paragraph::ParagraphElement::Run(run) = elem {
                            heading_text.push_str(&run.get_text());
                        }
                    }

                    // Generate unique bookmark name
                    let bookmark_name = format!("_Toc{}", 213359267 + bookmark_counter);
                    let bookmark_id = bookmark_counter;
                    bookmark_counter += 1;

                    // Add bookmark to the heading paragraph
                    para.add_bookmark_start(bookmark_id, &bookmark_name);
                    para.add_bookmark_end(bookmark_id);

                    // Store heading info for TOC generation
                    heading_info.push((level, heading_text, bookmark_name));
                }
            }
        }

        // Step 2: Build TOC paragraphs
        let mut toc_paragraphs = Vec::new();

        // First paragraph: TOC field wrapper
        let mut toc_field_para = MutableParagraph::new();
        let instruction = toc.build_field_instruction();
        toc_field_para
            .elements
            .push(super::paragraph::ParagraphElement::Field(
                MutableField::begin(),
            ));
        toc_field_para
            .elements
            .push(super::paragraph::ParagraphElement::Field(
                MutableField::instruction_char(instruction),
            ));
        toc_field_para
            .elements
            .push(super::paragraph::ParagraphElement::Field(
                MutableField::separate(),
            ));

        toc_paragraphs.push(toc_field_para);

        // Generate TOC entry paragraphs
        for (level, heading_text, bookmark_name) in heading_info {
            let mut toc_entry = MutableParagraph::new();

            // Set TOC style
            toc_entry.style = Some(format!("TOC{}", level));

            // Set paragraph properties (tab and indent)
            toc_entry
                .properties
                .tab_stops
                .push(super::paragraph::TabStop {
                    position: 9350,
                    alignment: "right".to_string(),
                    leader: Some("dot".to_string()),
                });

            let indent = match level {
                1 => 0,
                2 => 440,
                3 => 880,
                _ => (level as i32 - 1) * 440,
            };
            toc_entry.properties.indent_left = Some(indent);

            // Add hyperlink with runs and PAGEREF field
            let mut hyperlink =
                super::hyperlink::MutableHyperlink::new_anchor(bookmark_name.clone());

            let mut text_run = super::run::MutableRun::new();
            text_run.set_text(&heading_text);
            text_run.properties.no_proof = true;
            hyperlink.add_run(text_run);

            let mut tab_run = super::run::MutableRun::new();
            tab_run.add_tab();
            tab_run.properties.no_proof = true;
            tab_run.properties.web_hidden = true;
            hyperlink.add_run(tab_run);

            hyperlink
                .elements
                .push(super::hyperlink::HyperlinkElement::Field(
                    MutableField::begin(),
                ));

            let mut pageref_instr = String::new();
            write!(&mut pageref_instr, " PAGEREF {} \\h ", bookmark_name).unwrap();
            hyperlink
                .elements
                .push(super::hyperlink::HyperlinkElement::Field(
                    MutableField::instruction_char(pageref_instr),
                ));

            hyperlink
                .elements
                .push(super::hyperlink::HyperlinkElement::Field(
                    MutableField::separate(),
                ));

            let mut page_run = super::run::MutableRun::new();
            page_run.set_text("1");
            page_run.properties.no_proof = true;
            page_run.properties.web_hidden = true;
            hyperlink.add_run(page_run);

            hyperlink
                .elements
                .push(super::hyperlink::HyperlinkElement::Field(
                    MutableField::end(),
                ));

            toc_entry
                .elements
                .push(super::paragraph::ParagraphElement::Hyperlink(hyperlink));
            toc_paragraphs.push(toc_entry);
        }

        // Add field end to the first TOC paragraph
        if let Some(first_para) = toc_paragraphs.first_mut() {
            first_para
                .elements
                .push(super::paragraph::ParagraphElement::Field(
                    MutableField::end(),
                ));
        }

        // Step 3: Insert TOC paragraphs at the recorded position
        for (i, para) in toc_paragraphs.into_iter().enumerate() {
            self.body
                .elements
                .insert(insertion_index + i, BodyElement::Paragraph(para));
        }

        Ok(())
    }

    /// Generate theme XML for theme1.xml part.
    pub(crate) fn generate_theme_xml(&self) -> Result<Option<String>> {
        if let Some(theme) = &self.theme {
            Ok(Some(theme.to_xml()?))
        } else {
            Ok(None)
        }
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
                    if let ParagraphElement::Hyperlink(link) = para_element
                        && let Some(url) = &link.url
                    {
                        urls.push(url.clone());
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

    /// Generate comments XML content.
    pub(crate) fn generate_comments_xml(&self) -> Result<Option<String>> {
        if self.comments.is_empty() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);

        for comment in &self.comments {
            let comment_xml = comment.to_xml()?;
            xml.push_str(&comment_xml);
        }

        xml.push_str("</w:comments>");
        Ok(Some(xml))
    }

    /// Generate settings XML content with protection if set.
    ///
    /// This generates a complete settings.xml file including document protection
    /// if protection is enabled.
    pub(crate) fn generate_settings_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);

        // Add document protection if set
        if let Some(ref protection) = self.protection {
            xml.push_str(r#"<w:documentProtection w:edit=""#);
            xml.push_str(protection.protection_type.to_xml());
            xml.push_str(r#"" w:enforcement="1""#);

            if let Some(ref hash) = protection.password_hash {
                write!(xml, r#" w:hash="{}""#, hash).map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(ref salt) = protection.salt {
                write!(xml, r#" w:salt="{}""#, salt).map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            xml.push_str("/>");
        }

        // Add default zoom
        xml.push_str(r#"<w:zoom w:percent="100"/>"#);

        xml.push_str("</w:settings>");
        Ok(xml)
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

    /// Generate section properties XML including header/footer/footnote/endnote references.
    fn generate_section_properties(
        &self,
        xml: &mut String,
        rel_mapper: &super::relmap::RelationshipMapper,
    ) -> Result<()> {
        xml.push_str("<w:sectPr>");

        // IMPORTANT: Element order MUST follow OOXML spec (ISO/IEC 29500)
        // Microsoft Word strictly enforces this ordering!

        // 1. Add header reference if present (must come before footnotePr)
        if let Some(header_id) = rel_mapper.get_header_id() {
            write!(
                xml,
                r#"<w:headerReference w:type="default" r:id="{}"/>"#,
                header_id
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // 2. Add footer reference if present (must come before footnotePr)
        if let Some(footer_id) = rel_mapper.get_footer_id() {
            write!(
                xml,
                r#"<w:footerReference w:type="default" r:id="{}"/>"#,
                footer_id
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // 3. Add footnote properties if present
        if rel_mapper.get_footnotes_id().is_some() {
            write!(
                xml,
                r#"<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr>"#
            )
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // 4. Add endnote properties if present
        if rel_mapper.get_endnotes_id().is_some() {
            write!(
                xml,
                r#"<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr>"#
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

    fn from_xml(xml: &str) -> Result<Self> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut body = Self::new();
        let mut current_para_xml = Vec::new();
        let mut current_table_xml = Vec::new();
        let mut in_paragraph = false;
        let mut in_table = false;
        let mut depth = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let tag = e.local_name();
                    if tag.as_ref() == b"p" && !in_paragraph && !in_table {
                        in_paragraph = true;
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
                    } else if tag.as_ref() == b"tbl" && !in_table && !in_paragraph {
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
                    } else if in_paragraph {
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
                    let tag = e.local_name();
                    if in_paragraph {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.push(b'>');
                        depth -= 1;
                        if depth == 0 && tag.as_ref() == b"p" {
                            // Parse paragraph from XML
                            let xml_str = String::from_utf8_lossy(&current_para_xml).into_owned();
                            body.elements
                                .push(BodyElement::Paragraph(MutableParagraph::from_xml(
                                    &xml_str,
                                )?));
                            in_paragraph = false;
                        }
                    } else if in_table {
                        current_table_xml.extend_from_slice(b"</");
                        current_table_xml.extend_from_slice(e.name().as_ref());
                        current_table_xml.push(b'>');
                        depth -= 1;
                        if depth == 0 && tag.as_ref() == b"tbl" {
                            // Parse table from XML
                            let xml_str = String::from_utf8_lossy(&current_table_xml).into_owned();
                            body.elements
                                .push(BodyElement::Table(MutableTable::from_xml(&xml_str)?));
                            in_table = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_paragraph => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Text(e)) if in_table => {
                    current_table_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_paragraph => {
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
        }

        Ok(body)
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
