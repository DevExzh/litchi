/// Document writer implementation for DOCX.
use crate::error::{OoxmlError, Result};
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
    /// Exact document/root/body opening XML retained from an existing document.
    preserved_prefix: Option<String>,
    /// Exact body/document closing XML retained from an existing document.
    preserved_suffix: Option<String>,
    /// Whether section properties must be regenerated instead of preserved verbatim.
    section_dirty: bool,
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
use litchi_fonts::CollectGlyphs;
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
                BodyElement::PreservedParagraph(_)
                | BodyElement::PreservedTable(_)
                | BodyElement::PreservedSectionProperties(_)
                | BodyElement::PreservedOther(_) => continue,
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
            preserved_prefix: None,
            preserved_suffix: None,
            section_dirty: false,
        }
    }

    /// Create a mutable document from existing XML content.
    pub fn from_xml(xml: &str) -> Result<Self> {
        let parsed = DocumentBody::from_xml(xml)?;
        Ok(Self {
            body: parsed.body,
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
            preserved_prefix: Some(parsed.prefix),
            preserved_suffix: Some(parsed.suffix),
            section_dirty: false,
        })
    }

    /// Get a mutable reference to the section properties.
    pub fn section_mut(&mut self) -> &mut SectionProperties {
        self.modified = true;
        self.section_dirty = true;
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
            self.section_dirty = true;
        }
        self.header.as_mut().unwrap()
    }

    /// Get or create the footer.
    pub fn footer(&mut self) -> &mut Vec<MutableParagraph> {
        if self.footer.is_none() {
            self.footer = Some(Vec::new());
            self.modified = true;
            self.section_dirty = true;
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
        self.section_dirty = true;
        (id, self.footnotes.last_mut().unwrap())
    }

    /// Add an endnote and return its ID and mutable reference.
    pub fn add_endnote(&mut self) -> (u32, &mut Note) {
        let id = self.endnotes.len() as u32 + 1;
        let note = Note::new(id);
        self.endnotes.push(note);
        self.modified = true;
        self.section_dirty = true;
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
    /// use litchi_ooxml::docx::settings::ProtectionType;
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
    /// use litchi_ooxml::docx::writer::MutableTheme;
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
    /// use litchi_ooxml::docx::writer::Watermark;
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
    /// use litchi_ooxml::docx::writer::TableOfContents;
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
        let insertion_index = self.body.content_insertion_index();

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
        self.write_document_prefix(&mut xml);
        let preserve_section = !self.section_dirty && self.body.has_preserved_section();
        self.body.write_contents(&mut xml, preserve_section)?;
        if !preserve_section {
            Self::write_default_section_properties(&mut xml);
        }
        self.write_document_suffix(&mut xml);
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
        self.write_document_prefix(&mut xml);

        // Generate body with relationship IDs
        let preserve_section = !self.section_dirty && self.body.has_preserved_section();
        self.body
            .write_contents_with_rels(&mut xml, rel_mapper, preserve_section)?;

        if !preserve_section {
            // The sectPr must be the last element in the body.
            self.generate_section_properties(&mut xml, rel_mapper)?;
        }

        self.write_document_suffix(&mut xml);
        Ok(xml)
    }

    fn write_document_prefix(&self, xml: &mut String) {
        if let Some(prefix) = &self.preserved_prefix {
            xml.push_str(prefix);
        } else {
            xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
            xml.push_str(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body>"#);
        }
    }

    fn write_document_suffix(&self, xml: &mut String) {
        if let Some(suffix) = &self.preserved_suffix {
            xml.push_str(suffix);
        } else {
            xml.push_str("</w:body></w:document>");
        }
    }

    fn write_default_section_properties(xml: &mut String) {
        xml.push_str("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>");
        xml.push_str(
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>",
        );
        xml.push_str("</w:sectPr>");
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

struct ParsedDocumentBody {
    body: DocumentBody,
    prefix: String,
    suffix: String,
}

#[derive(Clone, Copy)]
enum PreservedBodyKind {
    Paragraph,
    Table,
    SectionProperties,
    Other,
}

impl DocumentBody {
    fn new() -> Self {
        Self {
            elements: Vec::new(),
        }
    }

    fn from_xml(xml: &str) -> Result<ParsedDocumentBody> {
        use crate::docx::namespace::is_wordprocessing_namespace;
        use quick_xml::events::Event;
        use quick_xml::reader::NsReader;

        enum ScanEvent {
            StartBody,
            StartChild(PreservedBodyKind),
            NestedStart,
            EmptyChild(PreservedBodyKind),
            EndCaptured,
            EndBody,
            StartOther,
            EndOther,
            Eof,
            Other,
        }

        let bytes = xml.as_bytes();
        let mut reader = NsReader::from_reader(bytes);
        let mut body = Self::new();
        let mut depth = 0usize;
        let mut body_depth = None;
        let mut prefix_end = None;
        let mut suffix_start = None;
        let mut last_content_end = 0usize;
        let mut capture: Option<(PreservedBodyKind, usize, usize)> = None;

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
                    Event::Start(element)
                        if is_wordprocessing_namespace(&namespace)
                            && element.local_name().as_ref() == b"body" =>
                    {
                        ScanEvent::StartBody
                    },
                    Event::Start(element) if body_depth == Some(depth) => ScanEvent::StartChild(
                        preserved_body_kind(&namespace, element.local_name().as_ref()),
                    ),
                    Event::Start(_) => ScanEvent::StartOther,
                    Event::Empty(element) if capture.is_none() && body_depth == Some(depth) => {
                        ScanEvent::EmptyChild(preserved_body_kind(
                            &namespace,
                            element.local_name().as_ref(),
                        ))
                    },
                    Event::End(_) if capture.is_some() => ScanEvent::EndCaptured,
                    Event::End(element)
                        if is_wordprocessing_namespace(&namespace)
                            && element.local_name().as_ref() == b"body"
                            && body_depth == Some(depth) =>
                    {
                        ScanEvent::EndBody
                    },
                    Event::End(_) => ScanEvent::EndOther,
                    Event::Eof => ScanEvent::Eof,
                    _ => ScanEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                OoxmlError::InvalidFormat("Word document offset does not fit usize".to_string())
            })?;

            match event {
                ScanEvent::StartBody => {
                    if body_depth.is_some() || prefix_end.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "document contains multiple Word body elements".to_string(),
                        ));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                    body_depth = Some(depth);
                    prefix_end = Some(event_end);
                    last_content_end = event_end;
                },
                ScanEvent::StartChild(kind) => {
                    capture = Some((kind, event_start, 1));
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::NestedStart => {
                    let Some((_, _, capture_depth)) = capture.as_mut() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing preserved body element".to_string(),
                        ));
                    };
                    *capture_depth = capture_depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::EmptyChild(kind) => {
                    push_preserved_body_range(
                        &mut body,
                        xml,
                        &mut last_content_end,
                        kind,
                        event_start,
                        event_end,
                    )?;
                },
                ScanEvent::EndCaptured => {
                    let Some((_, _, capture_depth)) = capture.as_mut() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing preserved body element".to_string(),
                        ));
                    };
                    *capture_depth = capture_depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                    if *capture_depth == 0 {
                        let Some((kind, start, _)) = capture.take() else {
                            return Err(OoxmlError::InvalidFormat(
                                "missing preserved body element range".to_string(),
                            ));
                        };
                        push_preserved_body_range(
                            &mut body,
                            xml,
                            &mut last_content_end,
                            kind,
                            start,
                            event_end,
                        )?;
                    }
                },
                ScanEvent::EndBody => {
                    if event_start > last_content_end {
                        push_raw_body_xml(
                            &mut body,
                            PreservedBodyKind::Other,
                            xml,
                            last_content_end,
                            event_start,
                        )?;
                    }
                    suffix_start = Some(event_start);
                    body_depth = None;
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                },
                ScanEvent::StartOther => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::EndOther => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                },
                ScanEvent::Eof if depth != 0 || capture.is_some() || body_depth.is_some() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word document XML".to_string(),
                    ));
                },
                ScanEvent::Eof => break,
                ScanEvent::Other => {},
            }
        }

        let prefix_end = prefix_end.ok_or_else(|| {
            OoxmlError::InvalidFormat("Word document has no body element".to_string())
        })?;
        let suffix_start = suffix_start.ok_or_else(|| {
            OoxmlError::InvalidFormat("Word document body is not closed".to_string())
        })?;
        Ok(ParsedDocumentBody {
            body,
            prefix: ensure_writer_namespace_declarations(xml.get(..prefix_end).ok_or_else(
                || OoxmlError::InvalidFormat("invalid Word document prefix range".to_string()),
            )?)?,
            suffix: xml
                .get(suffix_start..)
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word document suffix range".to_string())
                })?
                .to_string(),
        })
    }
    fn add_paragraph(&mut self) -> &mut MutableParagraph {
        let index = self.content_insertion_index();
        self.elements
            .insert(index, BodyElement::Paragraph(MutableParagraph::new()));
        match self.elements.get_mut(index) {
            Some(BodyElement::Paragraph(p)) => p,
            _ => unreachable!(),
        }
    }

    fn add_table(&mut self, rows: usize, cols: usize) -> &mut MutableTable {
        let index = self.content_insertion_index();
        self.elements
            .insert(index, BodyElement::Table(MutableTable::new(rows, cols)));
        match self.elements.get_mut(index) {
            Some(BodyElement::Table(t)) => t,
            _ => unreachable!(),
        }
    }

    fn content_insertion_index(&self) -> usize {
        self.elements
            .iter()
            .position(|element| matches!(element, BodyElement::PreservedSectionProperties(_)))
            .unwrap_or(self.elements.len())
    }

    fn paragraph_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
                )
            })
            .count()
    }

    fn table_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    BodyElement::Table(_) | BodyElement::PreservedTable(_)
                )
            })
            .count()
    }

    fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        let mut count = 0;
        for elem in &mut self.elements {
            match elem {
                BodyElement::Paragraph(paragraph) => {
                    if count == index {
                        return Some(paragraph);
                    }
                    count += 1;
                },
                BodyElement::PreservedParagraph(_) => {
                    if count == index {
                        return None;
                    }
                    count += 1;
                },
                _ => {},
            }
        }
        None
    }

    fn table(&mut self, index: usize) -> Option<&mut MutableTable> {
        let mut count = 0;
        for elem in &mut self.elements {
            match elem {
                BodyElement::Table(table) => {
                    if count == index {
                        return Some(table);
                    }
                    count += 1;
                },
                BodyElement::PreservedTable(_) => {
                    if count == index {
                        return None;
                    }
                    count += 1;
                },
                _ => {},
            }
        }
        None
    }

    fn write_contents(&self, xml: &mut String, preserve_section: bool) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(p) => p.to_xml(xml)?,
                BodyElement::Table(t) => t.to_xml(xml)?,
                BodyElement::PreservedParagraph(raw)
                | BodyElement::PreservedTable(raw)
                | BodyElement::PreservedOther(raw) => xml.push_str(raw),
                BodyElement::PreservedSectionProperties(raw) if preserve_section => {
                    xml.push_str(raw);
                },
                BodyElement::PreservedSectionProperties(_) => {},
            }
        }
        Ok(())
    }

    fn has_preserved_section(&self) -> bool {
        self.elements
            .iter()
            .any(|element| matches!(element, BodyElement::PreservedSectionProperties(_)))
    }

    /// Generate XML with actual relationship IDs from the mapper.
    fn write_contents_with_rels(
        &self,
        xml: &mut String,
        rel_mapper: &crate::docx::writer::relmap::RelationshipMapper,
        preserve_section: bool,
    ) -> Result<()> {
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
                BodyElement::PreservedParagraph(raw)
                | BodyElement::PreservedTable(raw)
                | BodyElement::PreservedOther(raw) => xml.push_str(raw),
                BodyElement::PreservedSectionProperties(raw) if preserve_section => {
                    xml.push_str(raw);
                },
                BodyElement::PreservedSectionProperties(_) => {},
            }
        }
        Ok(())
    }
}

fn preserved_body_kind(
    namespace: &quick_xml::name::ResolveResult<'_>,
    local_name: &[u8],
) -> PreservedBodyKind {
    if crate::docx::namespace::is_wordprocessing_namespace(namespace) {
        return match local_name {
            b"p" => PreservedBodyKind::Paragraph,
            b"tbl" => PreservedBodyKind::Table,
            b"sectPr" => PreservedBodyKind::SectionProperties,
            _ => PreservedBodyKind::Other,
        };
    }
    PreservedBodyKind::Other
}

fn ensure_writer_namespace_declarations(prefix: &str) -> Result<String> {
    const REQUIRED: [(&str, &str); 4] = [
        (
            "w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ),
        (
            "r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        (
            "wp",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    ];

    let declarations = REQUIRED
        .iter()
        .filter(|(namespace_prefix, _)| !has_namespace_declaration(prefix, namespace_prefix))
        .map(|(namespace_prefix, namespace)| format!(r#" xmlns:{namespace_prefix}="{namespace}""#))
        .collect::<String>();
    if declarations.is_empty() {
        return Ok(prefix.to_string());
    }
    let insertion = prefix.rfind('>').ok_or_else(|| {
        OoxmlError::InvalidFormat("Word body opening tag is incomplete".to_string())
    })?;
    let mut augmented = String::with_capacity(prefix.len() + declarations.len());
    augmented.push_str(&prefix[..insertion]);
    augmented.push_str(&declarations);
    augmented.push_str(&prefix[insertion..]);
    Ok(augmented)
}

fn has_namespace_declaration(xml: &str, namespace_prefix: &str) -> bool {
    let needle = format!("xmlns:{namespace_prefix}");
    xml.match_indices(&needle).any(|(start, _)| {
        let before_is_boundary = start == 0
            || xml.as_bytes()[start - 1].is_ascii_whitespace()
            || xml.as_bytes()[start - 1] == b'<';
        let mut after = start + needle.len();
        while xml
            .as_bytes()
            .get(after)
            .is_some_and(u8::is_ascii_whitespace)
        {
            after += 1;
        }
        before_is_boundary && xml.as_bytes().get(after) == Some(&b'=')
    })
}

fn push_preserved_body_range(
    body: &mut DocumentBody,
    xml: &str,
    last_content_end: &mut usize,
    kind: PreservedBodyKind,
    start: usize,
    end: usize,
) -> Result<()> {
    if start > *last_content_end {
        push_raw_body_xml(
            body,
            PreservedBodyKind::Other,
            xml,
            *last_content_end,
            start,
        )?;
    }
    push_raw_body_xml(body, kind, xml, start, end)?;
    *last_content_end = end;
    Ok(())
}

fn push_raw_body_xml(
    body: &mut DocumentBody,
    kind: PreservedBodyKind,
    xml: &str,
    start: usize,
    end: usize,
) -> Result<()> {
    let raw_xml = xml
        .get(start..end)
        .ok_or_else(|| OoxmlError::InvalidFormat("invalid Word body element range".to_string()))?
        .to_string();
    body.elements.push(match kind {
        PreservedBodyKind::Paragraph => BodyElement::PreservedParagraph(raw_xml),
        PreservedBodyKind::Table => BodyElement::PreservedTable(raw_xml),
        PreservedBodyKind::SectionProperties => BodyElement::PreservedSectionProperties(raw_xml),
        PreservedBodyKind::Other => BodyElement::PreservedOther(raw_xml),
    });
    Ok(())
}

/// A body element (paragraph, table, or exact preserved XML).
#[derive(Debug)]
pub(crate) enum BodyElement {
    Paragraph(MutableParagraph),
    Table(MutableTable),
    PreservedParagraph(String),
    PreservedTable(String),
    PreservedSectionProperties(String),
    PreservedOther(String),
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

    #[test]
    fn appending_preserves_existing_body_xml_exactly() {
        let input = r#"<?xml version="1.0"?><q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><q:body>
  <!--keep--><q:p q:rsidR="00AB"><q:r><q:t><![CDATA[A < B]]></q:t></q:r><x:payload data="1 &amp; 2"/></q:p>
  <q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>
  <x:custom><![CDATA[opaque <xml>]]></x:custom>
  <q:sectPr><q:pgSz q:w="20000" q:h="10000"/></q:sectPr>
</q:body></q:document>"#;
        let mut document = MutableDocument::from_xml(input).unwrap();
        assert_eq!(document.paragraph_count(), 1);
        assert_eq!(document.table_count(), 1);

        document.add_paragraph_with_text("appended");
        let output = document.to_xml().unwrap();
        assert!(output.starts_with(
            r#"<?xml version="1.0"?><q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><q:body"#
        ));
        assert!(
            output.contains(
                r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#
            )
        );
        assert!(output.contains(
            r#"<q:p q:rsidR="00AB"><q:r><q:t><![CDATA[A < B]]></q:t></q:r><x:payload data="1 &amp; 2"/></q:p>"#
        ));
        assert!(output.contains(
            r#"<q:tbl><q:tr><q:tc><q:p><q:r><q:t>cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>"#
        ));
        assert!(output.contains(r#"<x:custom><![CDATA[opaque <xml>]]></x:custom>"#));
        assert!(output.contains(r#"<q:sectPr><q:pgSz q:w="20000" q:h="10000"/></q:sectPr>"#));
        assert!(output.contains("appended"));
        assert_eq!(output.matches("sectPr").count(), 2);
        assert!(output.ends_with("</q:body></q:document>"));
    }

    #[test]
    fn existing_document_parser_rejects_missing_or_truncated_body() {
        assert!(MutableDocument::from_xml("<w:document/>").is_err());
        assert!(MutableDocument::from_xml(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>"#
        )
        .is_err());
    }
}
