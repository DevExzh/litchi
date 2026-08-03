/// Document - the main API for working with Word document content.
use crate::docx::bookmark::Bookmark;
use crate::docx::comment::Comment;
use crate::docx::content_control::ContentControl;
use crate::docx::custom_xml::Part;
use crate::docx::enums::WdHeaderFooter;
use crate::docx::field::CompareField;
use crate::docx::field::{
    ActiveContentField, AdvanceField, AutoNumberField, AutoTextField, AutoTextListField,
    BarcodeField, BibliographyField, BidiOutlineField, CitationField, DatabaseField, DdeField,
    DocumentContextField, DocumentInformationField, DocumentPropertyField, DocumentVariableField,
    EmbedField, EquationField, ExternalIncludeField, Field, FormulaField, GoToButtonField,
    HyperlinkField, IfField, IndexEntryField, IndexField, InfoField, LegacyFormField, LinkField,
    ListNumberField, MacroButtonField, MailMergeConditionalControlField, MailMergeCounterField,
    MailMergeDataField, MailMergeNextField, MailMergeRecipientField, MergeField, PrintField,
    PrivateField, PromptField, QuoteField, ReferenceField, ReferencedDocumentField, SequenceField,
    SetField, ShapeField, StyleReferenceField, SymbolField, TableOfAuthoritiesEntryField,
    TableOfAuthoritiesField, TableOfContentsEntryField, TableOfContentsField, UserIdentityField,
};
use crate::docx::footnote::Note;
use crate::docx::header_footer::HeaderFooter;
use crate::docx::hyperlink::Hyperlink;
use crate::docx::mail_merge::{MailMergeRecipients, is_settings_relationship};
use crate::docx::numbering::Numbering;
use crate::docx::paragraph::Paragraph;
use crate::docx::parts::DocumentPart;
use crate::docx::section::{Section, Sections};
use crate::docx::settings::DocumentSettings;
use crate::docx::smart_tag::SmartTag;
use crate::docx::statistics::{
    DocumentStatistics, count_characters, count_characters_no_spaces, count_words,
    estimate_line_count, estimate_page_count,
};
use crate::docx::styles::Styles;
use crate::docx::table::Table;
use crate::docx::theme::Theme;
use crate::docx::variables::DocumentVariables;
use crate::docx::writer::Watermark;
use crate::error::{OoxmlError, Result};
use litchi_docx::alt::{Chunk, Part as AltPart, Target, is_relationship};
use litchi_docx::web;
use litchi_opc::OpcPackage;
use litchi_opc::constants::relationship_type;

/// A Word document.
///
/// This is the main API for reading and manipulating Word document content.
/// It provides access to paragraphs, tables, sections, styles, and other
/// document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
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

/// A picture watermark discovered in a document header, with its media part
/// resolved through the header part's relationships.
///
/// The payload is an inert borrowed byte view into the package; it is never
/// decoded, executed, or displayed.
#[derive(Debug)]
pub struct ImageWatermarkPart<'a> {
    /// Part name of the header carrying the watermark shape.
    pub source_header_name: String,
    /// Relationship ID of the `v:imagedata` reference in the header.
    pub relationship_id: String,
    /// Part name of the media part (e.g. `/word/media/watermarkImage1.png`).
    pub part_name: String,
    /// Declared OPC content type of the media part.
    pub content_type: &'a str,
    /// Original payload bytes held by the package.
    pub bytes: &'a [u8],
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.part.extract_text()
    }

    /// Return the package's glossary/building-block catalog and dialect.
    pub fn glossary(
        &self,
    ) -> Result<
        Option<(
            litchi_docx::glossary::Catalog,
            litchi_docx::glossary::Conformance,
        )>,
    > {
        Ok(litchi_docx::glossary::load(self.opc)?)
    }

    /// Load the typed, inert SmartArt (DrawingML diagram) inventory anchored
    /// in this document.
    ///
    /// Each returned [`crate::docx::smartart::DocxSmartArt`] carries the
    /// parsed data-model node tree, the layout/quick-style/colors part
    /// metadata, and the diagram part names. Both transitional and Strict
    /// namespace dialects are supported.
    pub fn smart_arts(&self) -> Result<Vec<crate::docx::smartart::DocxSmartArt>> {
        crate::docx::smartart::load_smart_arts(self.opc, self.part.part().partname())
    }

    /// Load the typed, inert text-box and WordArt inventory anchored in this
    /// document.
    ///
    /// Each returned [`crate::docx::textbox::DocxTextBox`] carries the shape
    /// identity, the `wps:bodyPr` text-body properties, the story as
    /// paragraphs with runs, and WordArt warp/styling presence flags. Both
    /// DrawingML shapes and legacy VML `w:pict` fallbacks are recognized, in
    /// both the transitional and Strict namespace dialects.
    pub fn text_boxes(&self) -> Result<Vec<crate::docx::textbox::DocxTextBox>> {
        crate::docx::textbox::load_text_boxes(self.part.xml_bytes())
    }

    /// Return numbered paragraphs with resolved, typed list markers.
    ///
    /// This is separate from [`Self::text`], whose behavior remains unchanged.
    pub fn list_items(&self) -> Result<Vec<crate::docx::list::ListItem>> {
        let Some(numbering) = self.numbering()? else {
            return Ok(Vec::new());
        };
        let paragraphs = self.paragraphs()?;
        let mut styles = self.styles()?;
        let mut counters = crate::docx::list::ListCounterState::new();
        let mut items = Vec::new();

        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            let direct = paragraph.numbering()?;
            let style_id = paragraph.style_id()?;
            let inherited = if direct.is_none() {
                match style_id.as_deref() {
                    Some(style_id) => styles.resolved_numbering(style_id)?,
                    None => None,
                }
            } else {
                None
            };
            let associated = if direct.is_none() && inherited.is_none() {
                style_id.as_deref().and_then(|style_id| {
                    let mut found = None;
                    for num in numbering.nums() {
                        if let Some(abstract_num) =
                            numbering.get_abstract_num(num.abstract_num_id())
                        {
                            for level in abstract_num.levels() {
                                if level.paragraph_style.as_deref() == Some(style_id)
                                    && found.is_none()
                                {
                                    found = Some(crate::docx::numbering::ParagraphNumbering {
                                        num_id: num.id(),
                                        level: level.level,
                                    });
                                }
                            }
                        }
                    }
                    found
                })
            } else {
                None
            };
            let Some(mut properties) = direct.or(inherited).or(associated) else {
                continue;
            };
            if properties.num_id == 0 {
                continue;
            }
            let mut linked_num_ids = std::collections::HashSet::new();
            loop {
                if !linked_num_ids.insert(properties.num_id) {
                    return Err(crate::error::OoxmlError::InvalidFormat(format!(
                        "numbering style-link cycle at numId {}",
                        properties.num_id
                    )));
                }
                let num = numbering.get_num(properties.num_id).ok_or_else(|| {
                    crate::error::OoxmlError::InvalidFormat(format!(
                        "paragraph references missing numId {}",
                        properties.num_id
                    ))
                })?;
                let abstract_num = numbering
                    .get_abstract_num(num.abstract_num_id())
                    .ok_or_else(|| {
                        crate::error::OoxmlError::InvalidFormat(format!(
                            "numId {} references a missing abstract numbering definition",
                            properties.num_id
                        ))
                    })?;
                let Some(style_link) = abstract_num.num_style_link() else {
                    break;
                };
                let linked = styles.resolved_numbering(style_link)?.ok_or_else(|| {
                    crate::error::OoxmlError::InvalidFormat(format!(
                        "numbering style link '{style_link}' has no numPr"
                    ))
                })?;
                if linked.num_id == 0 {
                    break;
                }
                properties.num_id = linked.num_id;
            }
            let (marker, suffix) = counters.advance(&numbering, properties)?;
            items.push(crate::docx::list::ListItem {
                paragraph_index,
                numbering: properties,
                marker,
                suffix,
                text: paragraph.text()?,
            });
        }
        Ok(items)
    }

    /// Extract paragraph text with resolvable list labels prepended.
    ///
    /// Picture bullets and unsupported formats remain available as typed markers
    /// through [`Self::list_items`] and are not replaced with fabricated text here.
    pub fn text_with_list_labels(&self) -> Result<String> {
        let paragraphs = self.paragraphs()?;
        let items = self.list_items()?;
        let mut by_paragraph = items.into_iter().peekable();
        let mut output = String::new();
        for (index, paragraph) in paragraphs.iter().enumerate() {
            if index != 0 {
                output.push('\n');
            }
            if by_paragraph
                .peek()
                .is_some_and(|item| item.paragraph_index == index)
            {
                let item = by_paragraph.next().expect("item checked");
                if let crate::docx::list::ListMarker::Text(label) = item.marker {
                    output.push_str(&label);
                    match item.suffix {
                        crate::docx::numbering::NumberingSuffix::Tab => output.push('\t'),
                        crate::docx::numbering::NumberingSuffix::Space => output.push(' '),
                        crate::docx::numbering::NumberingSuffix::Nothing => {},
                    }
                }
            }
            output.push_str(&paragraph.text()?);
        }
        Ok(output)
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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

    /// Return all run-level smart tags in document order.
    pub fn smart_tags(&self) -> Result<Vec<SmartTag>> {
        let mut tags = Vec::new();
        for paragraph in self.part.paragraphs()? {
            tags.extend(paragraph.smart_tags()?);
        }
        Ok(tags)
    }

    /// Return the number of run-level smart tags in the document.
    pub fn smart_tag_count(&self) -> Result<usize> {
        Ok(self.smart_tags()?.len())
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
    /// ```ignore
    /// use litchi_ooxml::docx::Package;
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
    pub fn elements(&self) -> Result<Vec<crate::docx::DocxElement>> {
        self.part.elements()
    }

    /// Return paragraphs, tables, and alternative-format anchors in document order.
    pub fn blocks(&self) -> Result<Vec<crate::docx::DocumentBlock>> {
        self.part.blocks()
    }

    /// Return all alternative-format import anchors in XML order.
    pub fn alts(&self) -> Result<Vec<Chunk>> {
        self.part.alts()
    }

    /// Resolve an alternative-format anchor to its borrowed opaque OPC payload.
    ///
    /// This validates the relationship type and internal target but never parses,
    /// imports, executes, or fetches the foreign content.
    pub fn resolve_alt<'b>(&'b self, chunk: &Chunk) -> Result<AltPart<'b>> {
        let relationship = self
            .part
            .part()
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!(
                    "altChunk relationship '{}' is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if relationship.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "altChunk relationship must have an internal target".into(),
            ));
        }
        if !is_relationship(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "altChunk relationship '{}' has invalid type '{}'",
                chunk.relationship().as_str(),
                relationship.reltype()
            )));
        }
        let target = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidFormat(format!("invalid altChunk target: {error}"))
        })?;
        let part = self.opc.get_part(&target).map_err(|error| {
            OoxmlError::PartNotFound(format!("altChunk target '{}': {error}", target.as_str()))
        })?;
        Ok(AltPart::new(part))
    }

    /// Resolve an alternative-format target without fetching or interpreting it.
    ///
    /// Internal targets are returned as opaque package bytes. External targets
    /// are returned as their relationship URI and are never accessed.
    pub fn alt_target<'b>(&'b self, chunk: &Chunk) -> Result<Target<'b>> {
        let relationship = self
            .part
            .part()
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!(
                    "altChunk relationship {:?} is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if !is_relationship(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "relationship {:?} is not an alternative-format import",
                chunk.relationship().as_str()
            )));
        }
        if relationship.is_external() {
            return Ok(Target::Link(relationship.target_ref()));
        }
        self.resolve_alt(chunk).map(Target::Part)
    }

    /// Get all sections in the document.
    ///
    /// Returns a `Sections` collection providing access to each section's
    /// page properties, margins, orientation, etc.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
        let mut sections_xml = Vec::new();
        crate::docx::namespace::scan_word_element_ranges(
            xml_bytes,
            &[b"sectPr"],
            |_, start, length| {
                let start = usize::try_from(start).map_err(|_| {
                    OoxmlError::InvalidFormat("section offset overflow".to_string())
                })?;
                let length = usize::try_from(length).map_err(|_| {
                    OoxmlError::InvalidFormat("section length overflow".to_string())
                })?;
                let end = start.checked_add(length).ok_or_else(|| {
                    OoxmlError::InvalidFormat("section range overflow".to_string())
                })?;
                let raw = xml_bytes.get(start..end).ok_or_else(|| {
                    OoxmlError::InvalidFormat("section range is outside document XML".to_string())
                })?;
                sections_xml.push(Section::from_xml_bytes(raw.to_vec())?);
                Ok(())
            },
        )?;

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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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

    /// Return distinct standard VML text watermarks from document headers.
    ///
    /// Word commonly repeats the same watermark in default, first-page, and
    /// even-page headers; equivalent copies are returned once in relationship
    /// order.
    pub fn watermarks(&self) -> Result<Vec<Watermark>> {
        let mut watermarks = Vec::new();
        for (_, header) in self.headers()? {
            for watermark in header.watermarks()? {
                if !watermarks.contains(&watermark) {
                    watermarks.push(watermark);
                }
            }
        }
        Ok(watermarks)
    }

    /// Return picture watermarks from document headers with their media
    /// parts resolved.
    ///
    /// Each entry pairs a `v:imagedata` anchor discovered in a header with
    /// the relationship-resolved media part name and payload bytes. The
    /// payload is an inert byte view; it is never decoded or displayed.
    pub fn image_watermarks(&self) -> Result<Vec<ImageWatermarkPart<'_>>> {
        let main_part = self.opc.main_document_part()?;
        let mut parts = Vec::new();
        for rel in main_part.rels().iter() {
            if rel.reltype() != relationship_type::HEADER {
                continue;
            }
            let target = rel.target_partname()?;
            let header_part = self.opc.get_part(&target)?;
            let header = HeaderFooter::from_part(header_part, WdHeaderFooter::Primary)?;
            for anchor in header.image_watermarks()? {
                let image_rel = header_part
                    .rels()
                    .get(anchor.relationship_id())
                    .ok_or_else(|| {
                        OoxmlError::InvalidFormat(format!(
                            "watermark image relationship '{}' is missing from {}",
                            anchor.relationship_id(),
                            target.as_str()
                        ))
                    })?;
                if image_rel.is_external() {
                    return Err(OoxmlError::InvalidFormat(
                        "external watermark image relationship is rejected".to_string(),
                    ));
                }
                let image_target = image_rel.target_partname().map_err(|error| {
                    OoxmlError::InvalidFormat(format!("invalid watermark image target: {error}"))
                })?;
                let image_part = self.opc.get_part(&image_target)?;
                parts.push(ImageWatermarkPart {
                    source_header_name: target.as_str().to_owned(),
                    relationship_id: anchor.relationship_id().to_owned(),
                    part_name: image_target.as_str().to_owned(),
                    content_type: image_part.content_type(),
                    bytes: image_part.blob(),
                });
            }
        }
        Ok(parts)
    }

    /// Get all footers in the document.
    ///
    /// Returns a vector of tuples containing the footer type and the footer itself.
    /// Footers can be of three types: Primary (default), FirstPage, and EvenPage.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::{Package, WdHeaderFooter};
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
    /// use litchi_ooxml::docx::{Package, WdHeaderFooter};
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

    /// Get all `<w:hyperlink>` element hyperlinks in the document.
    ///
    /// Returns external URL and internal bookmark links stored as
    /// `<w:hyperlink>` elements. `HYPERLINK` field instructions are exposed
    /// separately by `Self::hyperlink_fields`.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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

    /// Get the number of `<w:hyperlink>` element hyperlinks in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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

    /// Get typed, inert `HYPERLINK` fields in document order.
    ///
    /// Returned values expose stored targets, bookmarks, display metadata,
    /// switches, cached content, and dirty/lock state only. This method never
    /// opens, resolves, follows, activates, or refreshes a link.
    pub fn hyperlink_fields(&self) -> Result<Vec<HyperlinkField>> {
        self.fields()?
            .iter()
            .map(Field::hyperlink_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `HYPERLINK` fields in the main document.
    pub fn hyperlink_field_count(&self) -> Result<usize> {
        Ok(self.hyperlink_fields()?.len())
    }

    /// Get typed, inert bibliography citation (`CITATION`) fields in document order.
    ///
    /// Returned values expose stored source tags, switches, cached content, and
    /// dirty/lock state only. This method never looks up bibliography sources,
    /// formats citations, or refreshes fields.
    pub fn citations(&self) -> Result<Vec<CitationField>> {
        self.fields()?
            .iter()
            .map(Field::citation)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of bibliography citation fields in the main document.
    pub fn citation_count(&self) -> Result<usize> {
        Ok(self.citations()?.len())
    }

    /// Get typed, inert `BIBLIOGRAPHY` fields in document order.
    ///
    /// Returned values expose stored switches and cached content only. This
    /// method never loads source XML, sorts sources, or regenerates a
    /// bibliography.
    pub fn bibliographies(&self) -> Result<Vec<BibliographyField>> {
        self.fields()?
            .iter()
            .map(Field::bibliography)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of bibliography fields in the main document.
    pub fn bibliography_count(&self) -> Result<usize> {
        Ok(self.bibliographies()?.len())
    }

    /// Get typed, inert `DOCVARIABLE` fields in document order.
    ///
    /// Returned values expose stored names, switches, cached content, and
    /// dirty/lock state only. This method never reads the settings part,
    /// resolves document-variable values, or refreshes fields.
    pub fn document_variable_fields(&self) -> Result<Vec<DocumentVariableField>> {
        self.fields()?
            .iter()
            .map(Field::document_variable)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of document-variable fields in the main document.
    pub fn document_variable_field_count(&self) -> Result<usize> {
        Ok(self.document_variable_fields()?.len())
    }

    /// Get typed, inert `DOCPROPERTY` fields in document order.
    ///
    /// Returned values expose stored property names, switches, cached content,
    /// and dirty/lock state only. This method never reads core, extended, or
    /// custom package properties, resolves a value, or refreshes fields.
    pub fn document_property_fields(&self) -> Result<Vec<DocumentPropertyField>> {
        self.fields()?
            .iter()
            .map(Field::document_property)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DOCPROPERTY` fields in the main document.
    pub fn document_property_field_count(&self) -> Result<usize> {
        Ok(self.document_property_fields()?.len())
    }

    /// Get typed, inert explicit legacy `INFO` fields in document order.
    ///
    /// Returned values expose stored property selectors, optional replacement
    /// values, switches, cached content, and dirty/lock state only. This method
    /// never reads, resolves, modifies, or writes document or template
    /// properties, or refreshes a field.
    pub fn info_fields(&self) -> Result<Vec<InfoField>> {
        self.fields()?
            .iter()
            .map(Field::info_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert explicit legacy `INFO` fields.
    pub fn info_field_count(&self) -> Result<usize> {
        Ok(self.info_fields()?.len())
    }

    /// Get typed, inert built-in document-information fields in document order.
    ///
    /// Returned values expose only stored kinds, switches, cached content, and
    /// dirty/lock state. This method never reads package metadata or host
    /// identity data, calculates dates, revisions, or statistics, resolves a
    /// value, or refreshes fields.
    pub fn document_information_fields(&self) -> Result<Vec<DocumentInformationField>> {
        self.fields()?
            .iter()
            .map(Field::document_information)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert built-in document-information fields.
    pub fn document_information_field_count(&self) -> Result<usize> {
        Ok(self.document_information_fields()?.len())
    }

    /// Get typed, inert built-in document-context and runtime fields in document order.
    ///
    /// Returned values expose only stored kinds, switches, cached content, and
    /// dirty/lock state. This method never reads a document path, attached
    /// template, host filesystem state or file size, current clock, or page and
    /// section layout; resolves a value; or refreshes fields.
    pub fn document_context_fields(&self) -> Result<Vec<DocumentContextField>> {
        self.fields()?
            .iter()
            .map(Field::document_context)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert built-in document-context and runtime fields.
    pub fn document_context_field_count(&self) -> Result<usize> {
        Ok(self.document_context_fields()?.len())
    }

    /// Get typed, inert `MACROBUTTON` fields in document order.
    ///
    /// Returned values expose only stored macro or command names, button text,
    /// cached results, and dirty/lock state. This method never resolves, loads,
    /// invokes, or otherwise executes a macro or command.
    pub fn macro_button_fields(&self) -> Result<Vec<MacroButtonField>> {
        self.fields()?
            .iter()
            .map(Field::macro_button)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `MACROBUTTON` fields in the main document.
    pub fn macro_button_field_count(&self) -> Result<usize> {
        Ok(self.macro_button_fields()?.len())
    }

    /// Get typed, inert `ADDIN`, `CONTROL`, and `HTMLCONTROL` fields in document
    /// order.
    ///
    /// Returned values expose stored kinds, instructions, cached content, and
    /// dirty/lock state only. This method never loads an add-in, instantiates
    /// an OCX or HTML control, invokes code, executes script, renders content,
    /// accesses an external resource, or refreshes a field.
    pub fn active_content_fields(&self) -> Result<Vec<ActiveContentField>> {
        self.fields()?
            .iter()
            .map(Field::active_content_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert active-content fields in the main document.
    pub fn active_content_field_count(&self) -> Result<usize> {
        Ok(self.active_content_fields()?.len())
    }

    /// Get typed, inert `GLOSSARY` and `AUTOTEXT` fields in document order.
    ///
    /// Returned values expose stored kinds, entry names, switches, cached
    /// content, and dirty/lock state only. This method never looks up a
    /// building block, reads a template, inserts content, changes bookmarks,
    /// accesses an external resource, or refreshes a field.
    pub fn auto_text_fields(&self) -> Result<Vec<AutoTextField>> {
        self.fields()?
            .iter()
            .map(Field::auto_text_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert building-block fields in the main document.
    pub fn auto_text_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_fields()?.len())
    }

    /// Get typed, inert `AUTOTEXTLIST` fields in document order.
    ///
    /// Returned values expose stored display text, style/tip options, unknown
    /// switches, cached content, and dirty/lock state only. This method never
    /// shows a selection UI, looks up a building block, reads a template,
    /// inserts content, accesses an external resource, or refreshes a field.
    pub fn auto_text_list_fields(&self) -> Result<Vec<AutoTextListField>> {
        self.fields()?
            .iter()
            .map(Field::auto_text_list_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `AUTOTEXTLIST` fields in the main document.
    pub fn auto_text_list_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_list_fields()?.len())
    }

    /// Get typed, inert `GOTOBUTTON` fields in document order.
    ///
    /// Returned values expose only stored destinations, button text, cached
    /// results, and dirty/lock state. This method never resolves a destination,
    /// changes the insertion point, activates a jump, or refreshes a field.
    pub fn go_to_button_fields(&self) -> Result<Vec<GoToButtonField>> {
        self.fields()?
            .iter()
            .map(Field::go_to_button)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `GOTOBUTTON` fields in the main document.
    pub fn go_to_button_field_count(&self) -> Result<usize> {
        Ok(self.go_to_button_fields()?.len())
    }

    /// Get typed, inert `PRINT` fields in document order.
    ///
    /// Returned values expose only stored printer-instruction text, cached
    /// results, and dirty/lock state. This method never interprets control
    /// codes, opens a printer, sends output, changes print settings, or
    /// refreshes a field.
    pub fn print_fields(&self) -> Result<Vec<PrintField>> {
        self.fields()?
            .iter()
            .map(Field::print_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `PRINT` fields in the main document.
    pub fn print_field_count(&self) -> Result<usize> {
        Ok(self.print_fields()?.len())
    }

    /// Get typed, inert `EMBED` fields in document order.
    ///
    /// Returned values expose only stored opaque object instructions, cached
    /// content, and dirty/lock state. This method never loads, inspects,
    /// deserializes, activates, renders, or executes an embedded object,
    /// accesses an external resource, or refreshes a field.
    pub fn embed_fields(&self) -> Result<Vec<EmbedField>> {
        self.fields()?
            .iter()
            .map(Field::embed_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `EMBED` fields in the main document.
    pub fn embed_field_count(&self) -> Result<usize> {
        Ok(self.embed_fields()?.len())
    }

    /// Get typed, inert `BARCODE` fields in document order.
    ///
    /// Returned values expose only stored opaque barcode instructions, cached
    /// content, and dirty/lock state. This method never parses or validates
    /// barcode data or symbology, generates or renders a barcode, accesses an
    /// external resource, or refreshes a field.
    pub fn barcode_fields(&self) -> Result<Vec<BarcodeField>> {
        self.fields()?
            .iter()
            .map(Field::barcode_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `BARCODE` fields in the main document.
    pub fn barcode_field_count(&self) -> Result<usize> {
        Ok(self.barcode_fields()?.len())
    }

    /// Get typed, inert `BIDIOUTLINE` fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never reads right-to-left language,
    /// paragraph outline, or layout state; chooses a numbering system;
    /// calculates a result; or refreshes a field.
    pub fn bidi_outline_fields(&self) -> Result<Vec<BidiOutlineField>> {
        self.fields()?
            .iter()
            .map(Field::bidi_outline_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `BIDIOUTLINE` fields in the main document.
    pub fn bidi_outline_field_count(&self) -> Result<usize> {
        Ok(self.bidi_outline_fields()?.len())
    }

    /// Get typed, inert `SHAPE` drawing-canvas anchor fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never locates, links, loads, positions,
    /// lays out, or renders a drawing or canvas, or refreshes a field.
    pub fn shape_fields(&self) -> Result<Vec<ShapeField>> {
        self.fields()?
            .iter()
            .map(Field::shape_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SHAPE` drawing-canvas anchor fields.
    pub fn shape_field_count(&self) -> Result<usize> {
        Ok(self.shape_fields()?.len())
    }

    /// Get typed, inert legacy form-code fields in document order.
    ///
    /// Returned values expose only stored text/checkbox/drop-down kind, opaque
    /// instructions, cached content, and dirty/lock state. This method never
    /// reads associated form-property XML, fills a form, changes a selection or
    /// checkbox state, invokes entry or exit macros, or refreshes a field.
    pub fn legacy_form_fields(&self) -> Result<Vec<LegacyFormField>> {
        self.fields()?
            .iter()
            .map(Field::legacy_form_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert legacy form-code fields.
    pub fn legacy_form_field_count(&self) -> Result<usize> {
        Ok(self.legacy_form_fields()?.len())
    }

    /// Get typed, inert `PRIVATE` conversion-data fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never converts a document, interprets
    /// field data, changes hidden-text visibility or layout, or refreshes a
    /// field. `PRIVATE` is not treated as a confidentiality mechanism.
    pub fn private_fields(&self) -> Result<Vec<PrivateField>> {
        self.fields()?
            .iter()
            .map(Field::private_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `PRIVATE` conversion-data fields.
    pub fn private_field_count(&self) -> Result<usize> {
        Ok(self.private_fields()?.len())
    }

    /// Get typed, inert `DATABASE` query fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never opens a data source or database,
    /// uses connection information, executes SQL, generates or inserts a table,
    /// changes layout, or refreshes a field.
    pub fn database_fields(&self) -> Result<Vec<DatabaseField>> {
        self.fields()?
            .iter()
            .map(Field::database_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DATABASE` query fields.
    pub fn database_field_count(&self) -> Result<usize> {
        Ok(self.database_fields()?.len())
    }

    /// Get typed, inert user-identity fields in document order.
    ///
    /// Returned values expose only stored kind, override, formatting, cached
    /// content, and dirty/lock state. This method never reads or modifies a host
    /// user's identity, applies formatting, or refreshes a field.
    pub fn user_identity_fields(&self) -> Result<Vec<UserIdentityField>> {
        self.fields()?
            .iter()
            .map(Field::user_identity_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert user-identity fields in the main document.
    pub fn user_identity_field_count(&self) -> Result<usize> {
        Ok(self.user_identity_fields()?.len())
    }

    /// Get typed, inert `ADVANCE` fields in document order.
    ///
    /// Returned values expose stored point adjustments, cached content, and
    /// dirty/lock state only. This method never moves text, changes layout,
    /// reflows content, or refreshes a field.
    pub fn advance_fields(&self) -> Result<Vec<AdvanceField>> {
        self.fields()?
            .iter()
            .map(Field::advance_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ADVANCE` fields in the main document.
    pub fn advance_field_count(&self) -> Result<usize> {
        Ok(self.advance_fields()?.len())
    }

    /// Get typed, inert `DDE` and `DDEAUTO` fields in document order.
    ///
    /// Returned fields expose stored application, source, item, representation,
    /// storage, cached content, and dirty/lock metadata only. This method never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// requests data, refreshes, converts, evaluates, or executes anything.
    pub fn dde_links(&self) -> Result<Vec<DdeField>> {
        self.fields()?
            .iter()
            .map(Field::dde_link)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DDE` and `DDEAUTO` fields in the main document.
    pub fn dde_link_count(&self) -> Result<usize> {
        Ok(self.dde_links()?.len())
    }

    /// Get typed, inert `INCLUDETEXT`/`INCLUDEPICTURE` fields and historical
    /// `INCLUDE`/`IMPORT` aliases in document order.
    ///
    /// Returned fields expose stored source, bookmark, converter, XML, cached,
    /// and dirty/lock metadata only. This method never opens, resolves,
    /// imports, fetches, refreshes, converts, transforms, evaluates, or
    /// executes anything.
    pub fn external_includes(&self) -> Result<Vec<ExternalIncludeField>> {
        self.fields()?
            .iter()
            .map(Field::external_include)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert external-include fields in the main document.
    pub fn external_include_count(&self) -> Result<usize> {
        Ok(self.external_includes()?.len())
    }

    /// Get typed, inert RD referenced-document fields in document order.
    ///
    /// Returned fields expose stored paths, relative-path requests, switches,
    /// cached content, and dirty/lock metadata only. This method never opens,
    /// resolves, reads, imports, refreshes, evaluates, or executes a referenced
    /// document.
    pub fn referenced_documents(&self) -> Result<Vec<ReferencedDocumentField>> {
        self.fields()?
            .iter()
            .map(Field::referenced_document)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert referenced-document fields in the main document.
    pub fn referenced_document_count(&self) -> Result<usize> {
        Ok(self.referenced_documents()?.len())
    }

    /// Get typed, inert `LINK` fields in document order.
    ///
    /// Returned fields expose stored application, source, item, result,
    /// formatting, cached content, and dirty/lock metadata only. This method
    /// never activates an OLE server, launches an application, opens a source,
    /// requests data, refreshes, converts, evaluates, or executes anything.
    pub fn link_fields(&self) -> Result<Vec<LinkField>> {
        self.fields()?
            .iter()
            .map(Field::link)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `LINK` fields in the main document.
    pub fn link_field_count(&self) -> Result<usize> {
        Ok(self.link_fields()?.len())
    }

    /// Get typed, inert table-of-contents fields in document order.
    ///
    /// Both simple (`w:fldSimple`) and complex (`w:fldChar`) TOC fields are
    /// discovered. Returned values expose the stored instruction, switches,
    /// cached result, and dirty/lock state; this method never paginates,
    /// regenerates a table of contents, follows its links, or executes fields.
    pub fn table_of_contents(&self) -> Result<Vec<TableOfContentsField>> {
        self.fields()?
            .iter()
            .map(Field::table_of_contents)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-contents fields in the main document.
    pub fn table_of_contents_count(&self) -> Result<usize> {
        Ok(self.table_of_contents()?.len())
    }

    /// Get typed, inert table-of-contents entry (`TC`) fields in document order.
    ///
    /// Returned fields expose only stored entry text, list identifiers, levels,
    /// page-number omission requests, switches, cached content, and dirty/lock
    /// state. This method never changes hidden text, calculates page numbers,
    /// generates a table of contents, or refreshes fields.
    pub fn table_of_contents_entries(&self) -> Result<Vec<TableOfContentsEntryField>> {
        self.fields()?
            .iter()
            .map(Field::table_of_contents_entry)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert table-of-contents entry fields.
    pub fn table_of_contents_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_contents_entries()?.len())
    }

    /// Get typed, inert table-of-authorities fields in document order.
    ///
    /// Returned fields expose stored switches and cached content only. This
    /// method never locates citation text, paginates the document, generates a
    /// table of authorities, or refreshes fields.
    pub fn tables_of_authorities(&self) -> Result<Vec<TableOfAuthoritiesField>> {
        self.fields()?
            .iter()
            .map(Field::table_of_authorities)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-authorities fields in the main document.
    pub fn table_of_authorities_count(&self) -> Result<usize> {
        Ok(self.tables_of_authorities()?.len())
    }

    /// Get typed, inert table-of-authorities entry (`TA`) fields in document order.
    ///
    /// These are stored citation markers. This method does not search for
    /// matching visible text, change hidden-text state, or generate a `TOA`.
    pub fn table_of_authorities_entries(&self) -> Result<Vec<TableOfAuthoritiesEntryField>> {
        self.fields()?
            .iter()
            .map(Field::table_of_authorities_entry)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-authorities entry fields in the main document.
    pub fn table_of_authorities_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_authorities_entries()?.len())
    }

    /// Get typed, inert generated-index (`INDEX`) fields in document order.
    ///
    /// Returned fields expose stored switches and cached content only. This
    /// method never searches for index markers, sorts entries, calculates page
    /// references, generates an index, or refreshes fields.
    pub fn indexes(&self) -> Result<Vec<IndexField>> {
        self.fields()?
            .iter()
            .map(Field::index)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of generated-index fields in the main document.
    pub fn index_count(&self) -> Result<usize> {
        Ok(self.indexes()?.len())
    }

    /// Get typed, inert index-entry (`XE`) fields in document order.
    ///
    /// These are stored index markers. This method does not change hidden text,
    /// resolve page-range bookmarks, sort entries, or generate an `INDEX`.
    pub fn index_entries(&self) -> Result<Vec<IndexEntryField>> {
        self.fields()?
            .iter()
            .map(Field::index_entry)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of index-entry fields in the main document.
    pub fn index_entry_count(&self) -> Result<usize> {
        Ok(self.index_entries()?.len())
    }

    /// Get typed, inert `MERGEFIELD` fields in document order.
    ///
    /// Returned values expose stored data-column names, switches, cached
    /// content, and dirty/lock state only. This method never opens a data
    /// source, resolves records, performs a merge, or refreshes field results.
    ///
    /// For backward-compatible access to the raw fields, use
    /// [`Self::merge_fields`].
    pub fn typed_merge_fields(&self) -> Result<Vec<MergeField>> {
        self.fields()?
            .iter()
            .map(Field::merge_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `MERGEFIELD` fields in the main document.
    pub fn typed_merge_field_count(&self) -> Result<usize> {
        Ok(self.typed_merge_fields()?.len())
    }

    /// Get typed, inert `DATA` mail-merge source fields in document order.
    ///
    /// Returned values expose only stored data-source and header-source
    /// identifiers, switches, cached content, and dirty/lock state. This method
    /// never opens, reads, connects to, resolves, or modifies either source; it
    /// never selects a record, performs a merge, or refreshes a field result.
    pub fn mail_merge_data_fields(&self) -> Result<Vec<MailMergeDataField>> {
        self.fields()?
            .iter()
            .map(Field::mail_merge_data)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DATA` mail-merge source fields.
    pub fn mail_merge_data_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_data_fields()?.len())
    }

    /// Get typed, inert `MERGEREC` and `MERGESEQ` fields in document order.
    ///
    /// Returned values expose stored kind, cached content, and dirty/lock state
    /// only. This method never selects or counts records, opens a data source,
    /// performs a merge, or refreshes field results.
    pub fn mail_merge_counters(&self) -> Result<Vec<MailMergeCounterField>> {
        self.fields()?
            .iter()
            .map(Field::mail_merge_counter)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert mail-merge counter fields in the main document.
    pub fn mail_merge_counter_count(&self) -> Result<usize> {
        Ok(self.mail_merge_counters()?.len())
    }

    /// Get typed, inert `NEXT` mail-merge control fields in document order.
    ///
    /// Returned values expose stored cached content and dirty/lock state only.
    /// This method never advances a record, opens a data source, performs a
    /// merge, or refreshes field results.
    pub fn mail_merge_next_fields(&self) -> Result<Vec<MailMergeNextField>> {
        self.fields()?
            .iter()
            .map(Field::mail_merge_next)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `NEXT` mail-merge control fields.
    pub fn mail_merge_next_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_next_fields()?.len())
    }

    /// Get typed, inert `NEXTIF` and `SKIPIF` fields in document order.
    ///
    /// Returned values expose stored comparison text, cached content, and
    /// dirty/lock state only. This method never evaluates a comparison, changes
    /// record selection, opens a data source, performs a merge, or refreshes
    /// field results.
    pub fn mail_merge_conditional_controls(&self) -> Result<Vec<MailMergeConditionalControlField>> {
        self.fields()?
            .iter()
            .map(Field::mail_merge_conditional_control)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert conditional mail-merge control fields.
    pub fn mail_merge_conditional_control_count(&self) -> Result<usize> {
        Ok(self.mail_merge_conditional_controls()?.len())
    }

    /// Get typed, inert `IF` fields in document order.
    ///
    /// Returned values expose stored expression text, cached content, and
    /// dirty/lock state only. This method never parses or evaluates an
    /// expression, resolves field values, or refreshes a field result.
    pub fn if_fields(&self) -> Result<Vec<IfField>> {
        self.fields()?
            .iter()
            .map(Field::if_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `IF` fields.
    pub fn if_field_count(&self) -> Result<usize> {
        Ok(self.if_fields()?.len())
    }

    /// Get typed, inert `COMPARE` fields in document order.
    ///
    /// Returned values expose stored comparisons, cached content, and
    /// dirty/lock state only. This method never parses or evaluates a
    /// comparison, resolves nested field values, or refreshes a field.
    pub fn compare_fields(&self) -> Result<Vec<CompareField>> {
        self.fields()?
            .iter()
            .map(Field::compare_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `COMPARE` fields in the main document.
    pub fn compare_field_count(&self) -> Result<usize> {
        Ok(self.compare_fields()?.len())
    }

    /// Get typed, inert bookmark-reference fields in document order.
    ///
    /// Returned values expose stored kinds, targets, options, unknown switches,
    /// cached content, and dirty/lock state only. This method never looks up a
    /// bookmark, reads a referenced range or note, resolves a page number,
    /// creates a link, calculates a relative position, or refreshes a field.
    pub fn reference_fields(&self) -> Result<Vec<ReferenceField>> {
        self.fields()?
            .iter()
            .map(Field::reference_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert bookmark-reference fields in the main document.
    pub fn reference_field_count(&self) -> Result<usize> {
        Ok(self.reference_fields()?.len())
    }

    /// Get typed, inert `SET` fields in document order.
    ///
    /// Returned values expose stored target names, opaque expressions, cached
    /// content, and dirty/lock state only. This method never evaluates an
    /// expression, looks up or changes a bookmark, changes document state, or
    /// refreshes a field.
    pub fn set_fields(&self) -> Result<Vec<SetField>> {
        self.fields()?
            .iter()
            .map(Field::set_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SET` fields in the main document.
    pub fn set_field_count(&self) -> Result<usize> {
        Ok(self.set_fields()?.len())
    }

    /// Get typed, inert `=` formula fields in document order.
    ///
    /// Returned values expose stored formulas, cached content, and dirty/lock
    /// state only. This method never parses or evaluates a formula, reads table
    /// cells or bookmarks, resolves field values, or refreshes a field.
    pub fn formula_fields(&self) -> Result<Vec<FormulaField>> {
        self.fields()?
            .iter()
            .map(Field::formula_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert formula fields in the main document.
    pub fn formula_field_count(&self) -> Result<usize> {
        Ok(self.formula_fields()?.len())
    }

    /// Get typed, inert `EQ` equation fields in document order.
    ///
    /// Returned values expose stored expressions, cached content, and dirty/lock
    /// state only. This method never parses, calculates, formats, renders, or
    /// refreshes an equation.
    pub fn equations(&self) -> Result<Vec<EquationField>> {
        self.fields()?
            .iter()
            .map(Field::equation)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `EQ` fields in the main document.
    pub fn equation_count(&self) -> Result<usize> {
        Ok(self.equations()?.len())
    }

    /// Get typed, inert `SEQ` fields in document order.
    ///
    /// Returned values expose stored identifiers, optional bookmarks, opaque
    /// tails, cached content, and dirty/lock state only. This method never
    /// looks up a bookmark, increments or resets a sequence, calculates a
    /// number, or refreshes a field.
    pub fn sequence_fields(&self) -> Result<Vec<SequenceField>> {
        self.fields()?
            .iter()
            .map(Field::sequence_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SEQ` fields in the main document.
    pub fn sequence_field_count(&self) -> Result<usize> {
        Ok(self.sequence_fields()?.len())
    }

    /// Get typed, inert `STYLEREF` fields in document order.
    ///
    /// Returned values expose stored style names, options, switches, cached
    /// content, and dirty/lock state only. This method never looks up styled
    /// text, searches document stories, calculates paragraph numbers or
    /// relative positions, resolves page layout, or refreshes a field.
    pub fn style_reference_fields(&self) -> Result<Vec<StyleReferenceField>> {
        self.fields()?
            .iter()
            .map(Field::style_reference_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `STYLEREF` fields in the main document.
    pub fn style_reference_field_count(&self) -> Result<usize> {
        Ok(self.style_reference_fields()?.len())
    }

    /// Get typed, inert `QUOTE` fields in document order.
    ///
    /// Returned values expose stored text arguments, switches, cached content,
    /// and dirty/lock state only. This method never interprets character codes,
    /// expands nested fields, inserts text, or refreshes a field result.
    pub fn quote_fields(&self) -> Result<Vec<QuoteField>> {
        self.fields()?
            .iter()
            .map(Field::quote_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `QUOTE` fields in the main document.
    pub fn quote_field_count(&self) -> Result<usize> {
        Ok(self.quote_fields()?.len())
    }

    /// Get typed, inert `SYMBOL` fields in document order.
    ///
    /// Returned values expose stored character arguments, switches, cached
    /// content, and dirty/lock state only. This method never maps a character
    /// code, looks up a font, inserts a glyph, changes formatting or layout, or
    /// refreshes a field result.
    pub fn symbol_fields(&self) -> Result<Vec<SymbolField>> {
        self.fields()?
            .iter()
            .map(Field::symbol_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SYMBOL` fields in the main document.
    pub fn symbol_field_count(&self) -> Result<usize> {
        Ok(self.symbol_fields()?.len())
    }

    /// Get typed, inert legacy automatic-numbering fields in document order.
    ///
    /// Returned values expose stored kinds, switches, cached content, and
    /// dirty/lock state only. This method never calculates paragraph numbers,
    /// reads heading or style state, changes paragraphs or layout, or refreshes
    /// a field result.
    pub fn auto_number_fields(&self) -> Result<Vec<AutoNumberField>> {
        self.fields()?
            .iter()
            .map(Field::auto_number_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert legacy automatic-numbering fields.
    pub fn auto_number_field_count(&self) -> Result<usize> {
        Ok(self.auto_number_fields()?.len())
    }

    /// Get typed, inert `LISTNUM` fields in document order.
    ///
    /// Returned values expose stored optional list names, switches, cached
    /// content, and dirty/lock state only. This method never looks up a list,
    /// determines a level or start value, calculates a number, changes layout,
    /// or refreshes a field result.
    pub fn list_number_fields(&self) -> Result<Vec<ListNumberField>> {
        self.fields()?
            .iter()
            .map(Field::list_number_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `LISTNUM` fields in the main document.
    pub fn list_number_field_count(&self) -> Result<usize> {
        Ok(self.list_number_fields()?.len())
    }

    /// Get typed, inert `ASK` and `FILLIN` fields in document order.
    ///
    /// Returned values expose stored prompt, bookmark, default-response, cached
    /// content, and dirty/lock state only. This method never displays a prompt,
    /// captures a response, creates or updates a bookmark, performs a merge, or
    /// refreshes a field result.
    pub fn prompt_fields(&self) -> Result<Vec<PromptField>> {
        self.fields()?
            .iter()
            .map(Field::prompt_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ASK` and `FILLIN` fields.
    pub fn prompt_field_count(&self) -> Result<usize> {
        Ok(self.prompt_fields()?.len())
    }

    /// Get typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields in document
    /// order.
    ///
    /// Returned values expose stored recipient layout, locale, country, fallback,
    /// cached-content, and dirty/lock state only. This method never opens a data
    /// source, selects a record, performs a merge, expands placeholders, generates
    /// text, or refreshes a field result.
    pub fn mail_merge_recipient_fields(&self) -> Result<Vec<MailMergeRecipientField>> {
        self.fields()?
            .iter()
            .map(Field::mail_merge_recipient_field)
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields.
    pub fn mail_merge_recipient_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_recipient_fields()?.len())
    }

    /// Get all mail-merge fields in document order.
    ///
    /// This recognizes `MERGEFIELD` instructions represented by either
    /// `<w:fldSimple>` or complex `<w:fldChar>` field sequences.
    pub fn merge_fields(&self) -> Result<Vec<Field>> {
        Ok(self
            .fields()?
            .into_iter()
            .filter(Field::is_merge_field)
            .collect())
    }

    /// Get the data-source column names referenced by mail-merge fields.
    pub fn merge_field_names(&self) -> Result<Vec<String>> {
        Ok(self
            .merge_fields()?
            .into_iter()
            .filter_map(|field| field.merge_field_name().map(str::to_owned))
            .collect())
    }

    /// Get the numbering definitions for the document.
    ///
    /// Returns a `Numbering` object providing access to abstract numbering
    /// definitions and numbering instances used for lists.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
        let mut matches = main_part
            .rels()
            .iter()
            .filter(|rel| is_settings_relationship(rel.reltype()));
        let Some(rel) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(OoxmlError::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        if rel.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = rel.target_partname()?;
        let settings_part = self.opc.get_part(&target)?;
        Ok(Some(DocumentSettings::extract_from_part(settings_part)?))
    }

    /// Load the ISO mail-merge recipient-data part referenced by `settings.xml`.
    pub fn mail_merge_recipients(&self) -> Result<Option<MailMergeRecipients>> {
        let Some(settings) = self.settings()? else {
            return Ok(None);
        };
        let Some(relationship_id) = settings
            .mail_merge()
            .and_then(|merge| merge.odso())
            .and_then(|odso| odso.recipient_data_relationship_id())
        else {
            return Ok(None);
        };
        let main_part = self.opc.main_document_part()?;
        let settings_relationship = main_part
            .rels()
            .iter()
            .find(|rel| is_settings_relationship(rel.reltype()))
            .ok_or_else(|| OoxmlError::InvalidFormat("settings relationship is missing".into()))?;
        let settings_part = self
            .opc
            .get_part(&settings_relationship.target_partname()?)?;
        let relationship = settings_part.rels().get(relationship_id).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "recipient-data relationship '{relationship_id}' is missing"
            ))
        })?;
        let recipient_part = self.opc.get_part(&relationship.target_partname()?)?;
        Ok(Some(MailMergeRecipients::extract_from_part(
            recipient_part,
        )?))
    }

    /// Read the document's typed web-output settings and conformance family.
    pub fn web(&self) -> Result<Option<(web::Settings, web::Conformance)>> {
        Ok(web::load(self.opc)?)
    }

    /// Check if the document is protected.
    ///
    /// This is a convenience method that checks the settings for protection status.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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

    /// Load the bounded, inert classic-chart graph owned by this document.
    ///
    /// Returns every classic DrawingML chart anchored in the main document
    /// body together with its style, color-style, and embedded-workbook
    /// companion parts. See [`crate::docx::chart::load_chart_graph`].
    pub fn chart_graph(&self) -> Result<crate::docx::chart::DocxChartGraph> {
        let main = self.opc.main_document_part()?.partname().clone();
        crate::docx::chart::load_chart_graph(self.opc, &main)
    }

    /// Get document variables.
    ///
    /// Returns document variables stored in the settings, which can be
    /// referenced by fields and used for mail merge.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let mut matches = main_part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::SETTINGS | STRICT_SETTINGS_RELATIONSHIP
            )
        });
        let Some(relationship) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(OoxmlError::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let settings_part = self.opc.get_part(&target)?;
        Ok(Some(DocumentVariables::extract_from_settings_part(
            settings_part,
        )?))
    }

    /// Get the document theme.
    ///
    /// Returns the theme containing color scheme, font scheme, and format scheme.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for xml_part in doc.custom_xml()? {
    ///     println!("Custom XML part: {}", xml_part.id());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn custom_xml(&self) -> Result<Vec<Part>> {
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
                let custom_xml = Part::from_part(part, id)?;
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
    /// use litchi_ooxml::docx::Package;
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
    // ✅ Custom XML: custom_xml()
    // ✅ Search: search(), search_ignore_case()
    //
    // ========================================
    // WRITE OPERATIONS
    // ========================================
    // Note: Write operations are primarily handled by the MutableDocument API
    // in the writer module. See src/ooxml/docx/writer/doc.rs for full API.
    //
    // ✅ COMPLETED: Modification operations (ECMA-376 Section 17.2.2)
    // - Add/insert/remove paragraphs: add_paragraph(), insert_paragraph(),
    //   remove_paragraph() on MutableDocument
    // - Add/insert/remove tables: add_table(), insert_table(), remove_table()
    // - Modify runs: MutableRun::bold(), italic(), font_name(), color()
    //   reached through the paragraph()/table() accessors
    // ✅ COMPLETED: Note and comment removal
    // - remove_footnote()/remove_endnote() strip typed w:footnoteReference /
    //   w:endnoteReference runs (ECMA-376 Sections 17.11.10, 17.11.2)
    // - remove_comment() drops authored w:comment entries (ECMA-376
    //   Section 17.13.4.2); note/comment IDs stay unique after removals
    //
    // ✅ COMPLETED: Track changes reading (November 2024)
    // - See revision.rs module and Paragraph::revisions() method
    // - Full support for insert, delete, move, and format revisions
    // - Includes author, date, and revision ID tracking
    //
    // ✅ Table of contents: table_of_contents(), table_of_contents_count()
    // - Typed inert field discovery and cached results; no pagination,
    //   automatic refresh, or structured-document-tag mutation
    //
    // ✅ Tables of authorities: tables_of_authorities(), table_of_authorities_entries()
    // - Typed inert TOA/TA discovery; no citation search, pagination, or refresh
    //
    // ✅ Generated indexes: indexes(), index_entries()
    // - Typed inert INDEX/XE discovery; no marker search, sorting, pagination, or refresh
    //
    // ✅ Mail merge field discovery: merge_fields(), merge_field_names(), typed_merge_fields(),
    //    mail_merge_data_fields(), mail_merge_counters()
    // ✅ COMPLETED: Mail merge settings mutation (MS-DOCX Section 17.16.5.35)
    // - Typed inert `w:mailMerge` model: MailMergeSettings, MailMergeDataSourceObject,
    //   MailMergeFieldMap, MailMergeRecipients (see mail_merge.rs)
    // - Package-level create/replace/remove: Package::set_mail_merge(),
    //   update_mail_merge(), update_mail_merge_recipients(), clear_mail_merge()
    // - Connection strings, queries, and data sources stay inert typed data;
    //   executing the merge against a data source is out of scope by design
    //
    // ✅ COMPLETED: Typed run breaks (MS-DOCX Section 17.3.3.1)
    // - Run::breaks() preserves text-wrapping, page, and column breaks plus clear behavior
    // - Rendered pagination hints remain distinguishable from authored breaks
    // ✅ COMPLETED: Section-break insertion and mutation
    // - MutableDocument::insert_section_break() authors paragraph-level w:pPr/w:sectPr;
    //   section_mut() edits the body-final section
    // - section_break(), update_section_break(), remove_section_break(),
    //   move_section_break() cover per-section page setup mutation
    //
    // ✅ Watermarks: typed VML header discovery plus mutable add/remove support
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
    // ✅ Smart tags: namespace-aware reading and mutable nested-tag writing
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
