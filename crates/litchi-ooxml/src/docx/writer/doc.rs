use crate::docx::OfficeMath;
/// Document writer implementation for DOCX.
use crate::docx::alt_chunk::{AltChunk, AltChunkNamespace, scan_alt_chunks};
use crate::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::ImageFormat;
// Import from other writer modules
use super::comment::MutableComment;
use super::note::Note;
use super::ole_object::MutableOleObject;
use super::paragraph::{MutableParagraph, ParagraphElement};
use super::section::SectionProperties;
use super::smartart::{MAX_SMART_ARTS, MutableSmartArt};
use super::table::MutableTable;
use super::theme::MutableTheme;
use super::toc::TableOfContents;
use super::vml_shape::MutableVmlShape;
use super::watermark::{ImageWatermark, Watermark};
use std::collections::HashSet;
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
    /// Whether document protection was explicitly changed.
    protection_dirty: bool,
    /// Section properties (page setup, margins, orientation)
    section: SectionProperties,
    /// Theme (optional)
    theme: Option<MutableTheme>,
    /// Watermark (optional)
    pub(crate) watermark: Option<Watermark>,
    /// Image watermark (optional)
    pub(crate) image_watermark: Option<ImageWatermark>,
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
    /// VML shape IDs already assigned to embedded OLE objects and VML shapes
    /// in this document.
    assigned_shape_ids: HashSet<String>,
    /// Next VML shape number tried when allocating OLE object identities.
    next_ole_shape_number: u32,
    /// Next VML shape number tried when allocating VML shape identities.
    next_vml_shape_number: u32,
    /// Next SmartArt anchor number used when allocating anchor keys.
    next_smartart_anchor: u32,
}

/// First VML shape number used when allocating OLE object identities
/// (`_x0000_i1025`, matching Word's numbering convention).
const FIRST_OLE_SHAPE_NUMBER: u32 = 1025;
/// First VML shape number used when allocating VML shape identities
/// (`_x0000_s1025`, matching Word's numbering convention).
const FIRST_VML_SHAPE_NUMBER: u32 = 1025;

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
use super::smart_tag::MutableSmartTag;
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
                | BodyElement::PreservedAltChunk(_, _)
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
                ParagraphElement::SmartTag(tag) => tag.collect_glyphs(),
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
impl CollectGlyphs for MutableSmartTag {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();
        for element in &self.elements {
            let element_glyphs = match element {
                ParagraphElement::Run(run) => run.collect_glyphs(),
                ParagraphElement::Hyperlink(hyperlink) => hyperlink.collect_glyphs(),
                ParagraphElement::SmartTag(tag) => tag.collect_glyphs(),
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
            protection_dirty: false,
            toc_config: None,
            section: SectionProperties::default(),
            theme: None,
            watermark: None,
            image_watermark: None,
            modified: false,
            preserved_prefix: None,
            preserved_suffix: None,
            section_dirty: false,
            assigned_shape_ids: HashSet::new(),
            next_ole_shape_number: FIRST_OLE_SHAPE_NUMBER,
            next_vml_shape_number: FIRST_VML_SHAPE_NUMBER,
            next_smartart_anchor: 1,
        }
    }

    /// Create a mutable document from existing XML content.
    pub fn from_xml(xml: &str) -> Result<Self> {
        let parsed = DocumentBody::from_xml(xml)?;
        parsed.body.validate_section_placement()?;
        let section = parsed.body.final_section_properties()?.unwrap_or_default();
        Ok(Self {
            body: parsed.body,
            toc_config: None,
            header: None,
            footer: None,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            comments: Vec::new(),
            protection: None,
            protection_dirty: false,
            section,
            theme: None,
            watermark: None,
            image_watermark: None,
            modified: false,
            preserved_prefix: Some(parsed.prefix),
            preserved_suffix: Some(parsed.suffix),
            section_dirty: false,
            assigned_shape_ids: HashSet::new(),
            next_ole_shape_number: FIRST_OLE_SHAPE_NUMBER,
            next_vml_shape_number: FIRST_VML_SHAPE_NUMBER,
            next_smartart_anchor: 1,
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

    /// Number of paragraph-level section breaks, excluding the body-final section.
    pub fn section_break_count(&self) -> Result<usize> {
        self.body.section_break_count()
    }

    /// Insert a section break at the end of the selected paragraph.
    pub fn insert_section_break(
        &mut self,
        paragraph_index: usize,
        properties: SectionProperties,
    ) -> Result<()> {
        properties.validate()?;
        self.body
            .insert_section_break(paragraph_index, properties)?;
        self.modified = true;
        Ok(())
    }

    /// Return an owned snapshot of a paragraph-level section break.
    pub fn section_break(&self, index: usize) -> Result<SectionProperties> {
        self.body.section_break(index)
    }

    /// Mutate a paragraph-level section break without rewriting unrelated paragraph XML.
    pub fn update_section_break(
        &mut self,
        index: usize,
        update: impl FnOnce(&mut SectionProperties),
    ) -> Result<()> {
        self.body.update_section_break(index, update)?;
        self.modified = true;
        Ok(())
    }

    /// Remove and return a paragraph-level section break.
    pub fn remove_section_break(&mut self, index: usize) -> Result<SectionProperties> {
        let properties = self.body.remove_section_break(index)?;
        self.modified = true;
        Ok(properties)
    }

    /// Move a section break to the end of another paragraph.
    pub fn move_section_break(&mut self, index: usize, after_paragraph: usize) -> Result<()> {
        let properties = self.remove_section_break(index)?;
        self.insert_section_break(after_paragraph, properties)
    }

    pub(crate) fn collect_section_header_footer_parts(
        &self,
    ) -> Result<Vec<(bool, super::section::SectionHeaderFooterPart)>> {
        let mut parts = Vec::new();
        collect_section_parts(&self.section, &mut parts)?;
        self.body.collect_section_parts(&mut parts)?;
        let mut unique: Vec<(bool, super::section::SectionHeaderFooterPart)> = Vec::new();
        for (header, part) in parts {
            if let Some((existing_header, existing)) =
                unique.iter().find(|(_, existing)| existing.key == part.key)
            {
                if *existing_header != header || existing.xml != part.xml {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "section header/footer key {:?} has conflicting definitions",
                        part.key
                    )));
                }
            } else {
                unique.push((header, part));
            }
        }
        Ok(unique)
    }

    pub(crate) fn collect_explicit_section_header_footer_relationships(
        &self,
    ) -> Result<Vec<(String, bool)>> {
        let mut relationships = Vec::new();
        collect_explicit_section_relationships(&self.section, &mut relationships);
        self.body
            .collect_explicit_section_relationships(&mut relationships)?;
        relationships.sort();
        relationships.dedup();
        Ok(relationships)
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

    /// Add a paragraph containing one display Office Math equation.
    pub fn add_display_office_math(&mut self, equation: OfficeMath) -> &mut MutableParagraph {
        let paragraph = self.add_paragraph();
        paragraph.add_display_office_math(equation);
        paragraph
    }

    /// Parse and add a paragraph containing one display Office Math equation.
    pub fn add_display_office_math_xml(
        &mut self,
        xml: impl Into<String>,
    ) -> Result<&mut MutableParagraph> {
        let equation = OfficeMath::from_xml(xml)?;
        Ok(self.add_display_office_math(equation))
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

    /// Add an inline text box in a new paragraph at the end of the document.
    ///
    /// The text box is serialized as a DrawingML wordprocessing shape
    /// (`wps:wsp`) and reappears in the
    /// [`crate::docx::Document::text_boxes`] inventory after save and reopen.
    pub fn add_text_box(
        &mut self,
        text_box: super::textbox::MutableTextBox,
    ) -> &mut super::textbox::MutableTextBox {
        self.add_paragraph().add_text_box(text_box)
    }

    /// Embed an OLE/package object in a new paragraph at the end of the
    /// document.
    ///
    /// Assigns the object's VML shape identity when unset, rejecting explicit
    /// IDs that collide with shapes already present in the document. The
    /// payload is stored verbatim as an inert `/word/embeddings/oleObjectN.bin`
    /// part and is discoverable through
    /// [`crate::docx::Package::embedded`] after save and reopen.
    pub fn add_ole_object(
        &mut self,
        mut object: MutableOleObject,
    ) -> Result<&mut MutableOleObject> {
        if object.shape_id.is_empty() {
            let mut number = self.next_ole_shape_number;
            let shape_id = loop {
                let candidate = format!("_x0000_i{number}");
                if !self.shape_id_in_use(&candidate) {
                    break candidate;
                }
                number = number.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("OLE shape ID space exhausted".to_string())
                })?;
            };
            self.next_ole_shape_number = number.saturating_add(1);
            object.shape_id = shape_id;
            object.object_id = number;
        } else {
            if self.shape_id_in_use(&object.shape_id) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "OLE shape ID '{}' collides with an existing shape",
                    object.shape_id
                )));
            }
            if object.object_id == 0 {
                object.object_id = self.next_ole_shape_number;
                self.next_ole_shape_number = self.next_ole_shape_number.saturating_add(1);
            }
        }
        self.assigned_shape_ids.insert(object.shape_id.clone());
        self.modified = true;
        Ok(self.add_paragraph().add_ole_object(object))
    }

    /// Check whether a VML shape ID is already used by an authored OLE object
    /// or VML shape, or appears in preserved document XML.
    fn shape_id_in_use(&self, shape_id: &str) -> bool {
        if self.assigned_shape_ids.contains(shape_id) {
            return true;
        }
        let preserved_hit = |raw: &str| raw.contains(shape_id);
        if self.preserved_prefix.as_deref().is_some_and(preserved_hit)
            || self.preserved_suffix.as_deref().is_some_and(preserved_hit)
        {
            return true;
        }
        self.body.elements.iter().any(|element| match element {
            BodyElement::PreservedParagraph(raw)
            | BodyElement::PreservedTable(raw)
            | BodyElement::PreservedSectionProperties(raw)
            | BodyElement::PreservedOther(raw) => raw.contains(shape_id),
            BodyElement::PreservedAltChunk(raw, _) => raw.contains(shape_id),
            _ => false,
        })
    }

    /// Add a legacy VML shape in a new paragraph at the end of the document.
    ///
    /// Assigns the shape's VML identity (`_x0000_s1025`, …, matching Word's
    /// numbering convention) when unset, skipping IDs already used by
    /// authored shapes or present in preserved document XML. A shape with a
    /// `v:textbox` story is discoverable through
    /// [`crate::docx::Document::text_boxes`] after save and reopen.
    pub fn add_vml_shape(&mut self, mut shape: MutableVmlShape) -> Result<&mut MutableVmlShape> {
        if shape.id.is_empty() {
            let mut number = self.next_vml_shape_number;
            let id = loop {
                let candidate = format!("_x0000_s{number}");
                if !self.shape_id_in_use(&candidate) {
                    break candidate;
                }
                number = number.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("VML shape ID space exhausted".to_string())
                })?;
            };
            self.next_vml_shape_number = number.saturating_add(1);
            shape.id = id;
        } else if self.shape_id_in_use(&shape.id) {
            return Err(OoxmlError::InvalidFormat(format!(
                "VML shape ID '{}' collides with an existing shape",
                shape.id
            )));
        }
        self.assigned_shape_ids.insert(shape.id.clone());
        self.modified = true;
        Ok(self.add_paragraph().add_vml_shape(shape))
    }

    /// Add a SmartArt (DrawingML diagram) graphic in a new paragraph at the
    /// end of the document.
    ///
    /// Assigns the anchor key binding the four diagram relationship IDs at
    /// save time. The data/layout/quick-style/colors parts are generated
    /// under `/word/diagrams/`, and the diagram is discoverable through
    /// [`crate::docx::Document::smart_arts`] after save and reopen. The
    /// optional pre-rendered drawing part is not generated; Word and
    /// LibreOffice re-render from the layout and data parts.
    pub fn add_smart_art(&mut self, mut smartart: MutableSmartArt) -> Result<&mut MutableSmartArt> {
        if self.collect_smart_arts().len() >= MAX_SMART_ARTS {
            return Err(OoxmlError::InvalidFormat(format!(
                "SmartArt count exceeds {MAX_SMART_ARTS}"
            )));
        }
        let number = self.next_smartart_anchor;
        self.next_smartart_anchor = number.saturating_add(1);
        smartart.anchor_key = format!("smartart{number}");
        self.modified = true;
        Ok(self.add_paragraph().add_smart_art(smartart))
    }

    /// Add a page break.
    pub fn add_page_break(&mut self) -> &mut MutableParagraph {
        let para = self.add_paragraph();
        para.add_run().add_page_break();
        para
    }

    /// Insert a new empty paragraph before the paragraph at `index`
    /// (`w:p`, ECMA-376 §17.3.1.22).
    ///
    /// Indices follow paragraph order across the whole body: typed and
    /// preserved paragraphs share one sequence, matching [`Self::paragraph`].
    /// Passing `index == paragraph_count()` appends at the end of the body
    /// content, before the body-final `w:sectPr` (ECMA-376 §17.2.2).
    pub fn insert_paragraph(&mut self, index: usize) -> Result<&mut MutableParagraph> {
        let (position, paragraph) = self.body.insert_paragraph(index)?;
        shift_toc_index_on_insert(&mut self.toc_config, position);
        self.modified = true;
        Ok(paragraph)
    }

    /// Insert a new empty table before the table at `index`
    /// (`w:tbl`, ECMA-376 §17.4.38).
    ///
    /// Indices follow table order across the whole body, matching
    /// [`Self::table`]; `index == table_count()` appends at the end of the
    /// body content, before the body-final `w:sectPr`.
    pub fn insert_table(
        &mut self,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> Result<&mut MutableTable> {
        let (position, table) = self.body.insert_table(index, rows, cols)?;
        shift_toc_index_on_insert(&mut self.toc_config, position);
        self.modified = true;
        Ok(table)
    }

    /// Remove the paragraph at `index` (paragraph order, matching
    /// [`Self::paragraph`]).
    ///
    /// Removing a paragraph whose `w:pPr` holds a `w:sectPr` removes that
    /// section break as well, merging the section with the following one —
    /// the same behavior as deleting the paragraph mark in Word.
    pub fn remove_paragraph(&mut self, index: usize) -> Result<()> {
        let position = self.body.remove_paragraph(index)?;
        shift_toc_index_on_remove(&mut self.toc_config, position);
        self.modified = true;
        Ok(())
    }

    /// Remove the table at `index` (table order, matching [`Self::table`]).
    pub fn remove_table(&mut self, index: usize) -> Result<()> {
        let position = self.body.remove_table(index)?;
        shift_toc_index_on_remove(&mut self.toc_config, position);
        self.modified = true;
        Ok(())
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
        let id = Self::next_note_id(self.footnotes.iter().map(|note| note.id));
        let note = Note::new(id);
        self.footnotes.push(note);
        self.modified = true;
        self.section_dirty = true;
        (id, self.footnotes.last_mut().unwrap())
    }

    /// Remove a footnote by ID and return it (`w:footnote`,
    /// ECMA-376 §17.11.10).
    ///
    /// Runs referencing the removed footnote (`w:footnoteReference`,
    /// ECMA-376 §17.11.14) are stripped from typed body, table, header,
    /// and footer paragraphs so the saved document never dangles a
    /// reference into the footnotes part.
    pub fn remove_footnote(&mut self, id: u32) -> Result<Note> {
        let index = self
            .footnotes
            .iter()
            .position(|note| note.id == id)
            .ok_or_else(|| OoxmlError::InvalidFormat(format!("footnote ID {id} does not exist")))?;
        let removed = self.footnotes.remove(index);
        self.strip_note_references(true, id);
        self.modified = true;
        self.section_dirty = true;
        Ok(removed)
    }

    /// Add an endnote and return its ID and mutable reference.
    pub fn add_endnote(&mut self) -> (u32, &mut Note) {
        let id = Self::next_note_id(self.endnotes.iter().map(|note| note.id));
        let note = Note::new(id);
        self.endnotes.push(note);
        self.modified = true;
        self.section_dirty = true;
        (id, self.endnotes.last_mut().unwrap())
    }

    /// Remove an endnote by ID and return it (`w:endnote`,
    /// ECMA-376 §17.11.2).
    ///
    /// Runs referencing the removed endnote (`w:endnoteReference`,
    /// ECMA-376 §17.11.7) are stripped like in [`Self::remove_footnote`].
    pub fn remove_endnote(&mut self, id: u32) -> Result<Note> {
        let index = self
            .endnotes
            .iter()
            .position(|note| note.id == id)
            .ok_or_else(|| OoxmlError::InvalidFormat(format!("endnote ID {id} does not exist")))?;
        let removed = self.endnotes.remove(index);
        self.strip_note_references(false, id);
        self.modified = true;
        self.section_dirty = true;
        Ok(removed)
    }

    /// Next note ID: one above the current maximum, so IDs stay unique
    /// even after removals.
    fn next_note_id(ids: impl Iterator<Item = u32>) -> u32 {
        ids.max().map_or(1, |id| id.saturating_add(1))
    }

    /// Strip every typed run referencing the note `id` from body, table,
    /// header, and footer paragraphs.
    fn strip_note_references(&mut self, footnote: bool, id: u32) {
        fn strip(elements: &mut Vec<ParagraphElement>, footnote: bool, id: u32) {
            elements.retain(|element| {
                !matches!(element, ParagraphElement::Run(run) if run.is_note_reference(footnote, id))
            });
        }
        for element in &mut self.body.elements {
            match element {
                BodyElement::Paragraph(paragraph) => strip(&mut paragraph.elements, footnote, id),
                BodyElement::Table(table) => {
                    for row in &mut table.rows {
                        for cell in &mut row.cells {
                            for paragraph in &mut cell.paragraphs {
                                strip(&mut paragraph.elements, footnote, id);
                            }
                        }
                    }
                },
                _ => {},
            }
        }
        for paragraphs in [&mut self.header, &mut self.footer].into_iter().flatten() {
            for paragraph in paragraphs {
                strip(&mut paragraph.elements, footnote, id);
            }
        }
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
        let id = Self::next_note_id(self.comments.iter().map(|comment| comment.id()));
        let comment = MutableComment::new(id, author.to_string(), text.to_string());
        self.comments.push(comment);
        self.modified = true;
        (id, self.comments.last_mut().unwrap())
    }

    /// Remove a comment by ID and return it (`w:comment`,
    /// ECMA-376 §17.13.4.2).
    ///
    /// Authored comments carry no range markers or reference runs in this
    /// writer model, so removal only affects the comments part emitted on
    /// save.
    pub fn remove_comment(&mut self, id: u32) -> Result<MutableComment> {
        let index = self
            .comments
            .iter()
            .position(|comment| comment.id() == id)
            .ok_or_else(|| OoxmlError::InvalidFormat(format!("comment ID {id} does not exist")))?;
        self.modified = true;
        Ok(self.comments.remove(index))
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
        self.protection_dirty = true;
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
        self.protection_dirty = true;
        self.modified = true;
    }

    /// Remove document protection.
    pub fn remove_protection(&mut self) {
        self.protection = None;
        self.protection_dirty = true;
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

    pub(crate) fn protection_is_dirty(&self) -> bool {
        self.protection_dirty
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

    /// Set an image watermark for the document.
    ///
    /// The image is stored verbatim as a media part and referenced from VML
    /// watermark shapes in the headers with centered default geometry; it is
    /// discoverable through [`crate::docx::Document::image_watermarks`] after
    /// save and reopen.
    pub fn set_image_watermark(&mut self, watermark: ImageWatermark) {
        self.image_watermark = Some(watermark);
        self.modified = true;
    }

    /// Remove the image watermark from the document.
    pub fn remove_image_watermark(&mut self) {
        if self.image_watermark.is_some() {
            self.image_watermark = None;
            self.modified = true;
        }
    }

    /// Check if the document has an image watermark.
    pub fn has_image_watermark(&self) -> bool {
        self.image_watermark.is_some()
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
                        elem.append_run_text(&mut heading_text);
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
                    para_element.collect_hyperlink_urls(&mut urls);
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
                    para_element.collect_images(&mut images);
                }
            }
        }

        images
    }

    /// Collect all embedded OLE objects from the document in document order.
    pub(crate) fn collect_ole_objects(&self) -> Vec<&MutableOleObject> {
        let mut objects = Vec::new();

        for element in &self.body.elements {
            if let BodyElement::Paragraph(para) = element {
                for para_element in &para.elements {
                    if let super::paragraph::ParagraphElement::OleObject(object) = para_element {
                        objects.push(object);
                    }
                }
            }
        }

        objects
    }

    /// Collect all SmartArt diagrams from the document in document order.
    pub(crate) fn collect_smart_arts(&self) -> Vec<&MutableSmartArt> {
        let mut smartarts = Vec::new();

        for element in &self.body.elements {
            if let BodyElement::Paragraph(para) = element {
                for para_element in &para.elements {
                    if let super::paragraph::ParagraphElement::SmartArt(smartart) = para_element {
                        smartarts.push(smartart);
                    }
                }
            }
        }

        smartarts
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

    /// Patch document protection while preserving every unrelated setting byte-for-byte.
    pub(crate) fn generate_settings_xml(&self, existing: Option<&[u8]>) -> Result<Vec<u8>> {
        match existing {
            Some(existing) => patch_document_protection(existing, self.protection.as_ref()),
            None => {
                let mut xml = String::with_capacity(512);
                xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
                xml.push_str(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);
                if let Some(protection) = &self.protection {
                    write_document_protection(&mut xml, protection, "w", None)?;
                }
                xml.push_str("</w:settings>");
                Ok(xml.into_bytes())
            },
        }
    }

    /// Get a reference to a paragraph by index.
    pub fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        self.body.paragraph(index)
    }

    /// Get a reference to a table by index.
    pub fn table(&mut self, index: usize) -> Option<&mut MutableTable> {
        self.body.table(index)
    }

    /// Return the direct alternative-format anchors in body order.
    pub fn alt_chunks(&self) -> Vec<AltChunk> {
        self.body.alt_chunks()
    }

    /// Insert an alternative-format anchor at an anchor-relative index.
    pub fn insert_alt_chunk(
        &mut self,
        index: usize,
        chunk: AltChunk,
        namespace: AltChunkNamespace,
    ) -> Result<()> {
        self.body.insert_alt_chunk(index, chunk, namespace)?;
        self.modified = true;
        Ok(())
    }

    /// Replace an alternative-format anchor without disturbing adjacent XML.
    pub fn replace_alt_chunk(
        &mut self,
        index: usize,
        chunk: AltChunk,
        namespace: AltChunkNamespace,
    ) -> Result<AltChunk> {
        let old = self.body.replace_alt_chunk(index, chunk, namespace)?;
        self.modified = true;
        Ok(old)
    }

    /// Remove an alternative-format anchor.
    pub fn remove_alt_chunk(&mut self, index: usize) -> Result<AltChunk> {
        let old = self.body.remove_alt_chunk(index)?;
        self.modified = true;
        Ok(old)
    }

    /// Move an alternative-format anchor to another anchor-relative index.
    pub fn move_alt_chunk(&mut self, from: usize, to: usize) -> Result<()> {
        self.body.move_alt_chunk(from, to)?;
        self.modified = true;
        Ok(())
    }

    /// Serialize the document to XML.
    pub fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(4096);
        self.write_document_prefix(&mut xml);
        let preserve_section = !self.section_dirty && self.body.has_preserved_section();
        self.body.write_contents(&mut xml, preserve_section)?;
        if !preserve_section {
            self.section.write_xml(&mut xml, None)?;
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
            xml.push_str(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>"#);
        }
    }

    fn write_document_suffix(&self, xml: &mut String) {
        if let Some(suffix) = &self.preserved_suffix {
            xml.push_str(suffix);
        } else {
            xml.push_str("</w:body></w:document>");
        }
    }

    /// Generate section properties XML including header/footer/footnote/endnote references.
    fn generate_section_properties(
        &self,
        xml: &mut String,
        rel_mapper: &super::relmap::RelationshipMapper,
    ) -> Result<()> {
        self.section.write_xml(xml, Some(rel_mapper))
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}

fn write_document_protection(
    xml: &mut String,
    protection: &DocumentProtection,
    prefix: &str,
    local_namespace: Option<&str>,
) -> Result<()> {
    write!(xml, "<{prefix}:documentProtection")
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(namespace) = local_namespace {
        write!(
            xml,
            " xmlns:{prefix}=\"{}\"",
            litchi_core::xml::escape_xml(namespace)
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    write!(
        xml,
        " {prefix}:edit=\"{}\" {prefix}:enforcement=\"1\"",
        protection.protection_type.to_xml()
    )
    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(hash) = &protection.password_hash {
        write!(
            xml,
            " {prefix}:hash=\"{}\"",
            litchi_core::xml::escape_xml(hash)
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(salt) = &protection.salt {
        write!(
            xml,
            " {prefix}:salt=\"{}\"",
            litchi_core::xml::escape_xml(salt)
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn patch_document_protection(
    existing: &[u8],
    protection: Option<&DocumentProtection>,
) -> Result<Vec<u8>> {
    use crate::docx::namespace::scan_word_element_ranges;

    let existing = std::str::from_utf8(existing).map_err(|_| {
        OoxmlError::InvalidFormat("settings.xml must be UTF-8 to modify document protection".into())
    })?;
    let mut ranges = Vec::new();
    scan_word_element_ranges(
        existing.as_bytes(),
        &[b"documentProtection"],
        |_, start, len| {
            let start = usize::try_from(start).map_err(|_| {
                OoxmlError::InvalidFormat("settings protection offset does not fit usize".into())
            })?;
            let len = usize::try_from(len).map_err(|_| {
                OoxmlError::InvalidFormat("settings protection length does not fit usize".into())
            })?;
            let end = start.checked_add(len).ok_or_else(|| {
                OoxmlError::InvalidFormat("settings protection range overflows usize".into())
            })?;
            ranges.push((start, end));
            Ok(())
        },
    )?;
    if ranges.len() > 1 {
        return Err(OoxmlError::InvalidFormat(
            "settings.xml contains duplicate documentProtection elements".into(),
        ));
    }

    let root = locate_settings_root(existing.as_bytes())?;
    let (root_name, root_namespace) = root.name_and_namespace();
    let (prefix, local_namespace) = match root_name.split_once(':') {
        Some((prefix, _)) => (prefix, None),
        None => ("w", Some(root_namespace)),
    };
    let mut replacement = String::new();
    if let Some(protection) = protection {
        write_document_protection(&mut replacement, protection, prefix, local_namespace)?;
    }

    if let Some((start, end)) = ranges.first().copied() {
        let mut output = String::with_capacity(existing.len() - (end - start) + replacement.len());
        output.push_str(&existing[..start]);
        output.push_str(&replacement);
        output.push_str(&existing[end..]);
        return Ok(output.into_bytes());
    }
    if replacement.is_empty() {
        return Ok(existing.as_bytes().to_vec());
    }

    match root {
        SettingsRoot::Paired { close_offset, .. } => {
            let mut output = String::with_capacity(existing.len() + replacement.len());
            output.push_str(&existing[..close_offset]);
            output.push_str(&replacement);
            output.push_str(&existing[close_offset..]);
            Ok(output.into_bytes())
        },
        SettingsRoot::Empty { end, name, .. } => {
            let empty_close = end.checked_sub(2).ok_or_else(|| {
                OoxmlError::InvalidFormat("invalid empty settings root range".into())
            })?;
            if existing.as_bytes().get(empty_close..end) != Some(b"/>") {
                return Err(OoxmlError::InvalidFormat(
                    "invalid empty settings root syntax".into(),
                ));
            }
            let mut output =
                String::with_capacity(existing.len() + replacement.len() + name.len() + 4);
            output.push_str(&existing[..empty_close]);
            output.push('>');
            output.push_str(&replacement);
            output.push_str("</");
            output.push_str(&name);
            output.push('>');
            output.push_str(&existing[end..]);
            Ok(output.into_bytes())
        },
    }
}

enum SettingsRoot {
    Paired {
        close_offset: usize,
        name: String,
        namespace: String,
    },
    Empty {
        end: usize,
        name: String,
        namespace: String,
    },
}

impl SettingsRoot {
    fn name_and_namespace(&self) -> (&str, &str) {
        match self {
            Self::Paired {
                name, namespace, ..
            }
            | Self::Empty {
                name, namespace, ..
            } => (name, namespace),
        }
    }
}

fn locate_settings_root(xml: &[u8]) -> Result<SettingsRoot> {
    use crate::docx::namespace::is_wordprocessing_namespace;
    use quick_xml::events::Event;
    use quick_xml::reader::NsReader;

    enum RootEvent {
        Start(Option<(String, String)>),
        Empty(Option<(String, String)>),
        End(bool),
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_info = None;
    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("settings root offset does not fit usize".into())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => RootEvent::Start(settings_root_info(
                    &namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )),
                Event::Empty(element) => RootEvent::Empty(settings_root_info(
                    &namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )),
                Event::End(element) => RootEvent::End(
                    is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"settings",
                ),
                Event::Eof => RootEvent::Eof,
                _ => RootEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("settings root offset does not fit usize".into())
        })?;

        match event {
            RootEvent::Start(info) if depth == 0 => {
                if saw_root || info.is_none() {
                    return Err(OoxmlError::InvalidFormat(
                        "settings.xml has an invalid or trailing root".into(),
                    ));
                }
                saw_root = true;
                root_info = info;
                depth = 1;
            },
            RootEvent::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("settings XML nesting is too deep".into())
                })?;
            },
            RootEvent::Empty(info) if depth == 0 => {
                if saw_root || info.is_none() {
                    return Err(OoxmlError::InvalidFormat(
                        "settings.xml has an invalid or trailing root".into(),
                    ));
                }
                let (name, namespace) = info.ok_or_else(|| {
                    OoxmlError::InvalidFormat("empty settings root has no name".into())
                })?;
                return Ok(SettingsRoot::Empty {
                    end: event_end,
                    name,
                    namespace,
                });
            },
            RootEvent::End(is_root) => {
                if depth == 1 {
                    if !is_root {
                        return Err(OoxmlError::InvalidFormat(
                            "settings.xml has an invalid root closing element".into(),
                        ));
                    }
                    let (name, namespace) = root_info.take().ok_or_else(|| {
                        OoxmlError::InvalidFormat("settings root metadata is missing".into())
                    })?;
                    return Ok(SettingsRoot::Paired {
                        close_offset: event_start,
                        name,
                        namespace,
                    });
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid settings XML nesting".into())
                })?;
            },
            RootEvent::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "settings.xml has no complete settings root".into(),
                ));
            },
            _ => {},
        }
    }
}

fn settings_root_info(
    namespace: &quick_xml::name::ResolveResult<'_>,
    qualified_name: &[u8],
    local_name: &[u8],
) -> Option<(String, String)> {
    use quick_xml::name::{Namespace, ResolveResult};

    if local_name != b"settings" || !crate::docx::namespace::is_wordprocessing_namespace(namespace)
    {
        return None;
    }
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return None;
    };
    Some((
        String::from_utf8_lossy(qualified_name).into_owned(),
        String::from_utf8_lossy(namespace).into_owned(),
    ))
}

/// The document body containing all content elements.
#[derive(Debug)]
pub(crate) struct DocumentBody {
    /// Content elements (paragraphs, tables, etc.) in document order
    pub(crate) elements: Vec<BodyElement>,
}

/// Keep a pending TOC insertion point anchored after an inserted element.
fn shift_toc_index_on_insert(toc_config: &mut Option<(usize, TableOfContents)>, position: usize) {
    if let Some((index, _)) = toc_config
        && position <= *index
    {
        *index += 1;
    }
}

/// Keep a pending TOC insertion point anchored after a removed element.
fn shift_toc_index_on_remove(toc_config: &mut Option<(usize, TableOfContents)>, position: usize) {
    if let Some((index, _)) = toc_config
        && position < *index
    {
        *index -= 1;
    }
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
    AltChunk,
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

    /// Element positions of all paragraphs, typed and preserved, in body order.
    fn paragraph_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(
                    element,
                    BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
                )
                .then_some(index)
            })
            .collect()
    }

    /// Element positions of all tables, typed and preserved, in body order.
    fn table_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(
                    element,
                    BodyElement::Table(_) | BodyElement::PreservedTable(_)
                )
                .then_some(index)
            })
            .collect()
    }

    /// Insert an empty paragraph before the paragraph at paragraph-relative
    /// `index`; returns the element position and the new paragraph.
    fn insert_paragraph(&mut self, index: usize) -> Result<(usize, &mut MutableParagraph)> {
        let positions = self.paragraph_positions();
        if index > positions.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "paragraph insertion index {index} is out of range"
            )));
        }
        let position = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        self.elements
            .insert(position, BodyElement::Paragraph(MutableParagraph::new()));
        match self.elements.get_mut(position) {
            Some(BodyElement::Paragraph(paragraph)) => Ok((position, paragraph)),
            _ => unreachable!(),
        }
    }

    /// Insert an empty table before the table at table-relative `index`;
    /// returns the element position and the new table.
    fn insert_table(
        &mut self,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> Result<(usize, &mut MutableTable)> {
        let positions = self.table_positions();
        if index > positions.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "table insertion index {index} is out of range"
            )));
        }
        let position = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        self.elements
            .insert(position, BodyElement::Table(MutableTable::new(rows, cols)));
        match self.elements.get_mut(position) {
            Some(BodyElement::Table(table)) => Ok((position, table)),
            _ => unreachable!(),
        }
    }

    /// Remove the paragraph at paragraph-relative `index`; returns the
    /// vacated element position.
    fn remove_paragraph(&mut self, index: usize) -> Result<usize> {
        let position = self
            .paragraph_positions()
            .get(index)
            .copied()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("paragraph index {index} is out of range"))
            })?;
        self.elements.remove(position);
        Ok(position)
    }

    /// Remove the table at table-relative `index`; returns the vacated
    /// element position.
    fn remove_table(&mut self, index: usize) -> Result<usize> {
        let position = self.table_positions().get(index).copied().ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("table index {index} is out of range"))
        })?;
        self.elements.remove(position);
        Ok(position)
    }

    fn alt_chunk_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(element, BodyElement::PreservedAltChunk(_, _)).then_some(index)
            })
            .collect()
    }

    fn alt_chunks(&self) -> Vec<AltChunk> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                BodyElement::PreservedAltChunk(_, chunk) => Some(chunk.clone()),
                _ => None,
            })
            .collect()
    }

    fn insert_alt_chunk(
        &mut self,
        index: usize,
        chunk: AltChunk,
        namespace: AltChunkNamespace,
    ) -> Result<()> {
        let positions = self.alt_chunk_positions();
        if index > positions.len() {
            return Err(OoxmlError::InvalidFormat(format!(
                "altChunk index {index} is out of range"
            )));
        }
        let position = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        let xml = chunk.to_xml(namespace);
        self.elements
            .insert(position, BodyElement::PreservedAltChunk(xml, chunk));
        Ok(())
    }

    fn replace_alt_chunk(
        &mut self,
        index: usize,
        chunk: AltChunk,
        namespace: AltChunkNamespace,
    ) -> Result<AltChunk> {
        let position = self
            .alt_chunk_positions()
            .get(index)
            .copied()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let xml = chunk.to_xml(namespace);
        match std::mem::replace(
            &mut self.elements[position],
            BodyElement::PreservedAltChunk(xml, chunk),
        ) {
            BodyElement::PreservedAltChunk(_, old) => Ok(old),
            _ => unreachable!(),
        }
    }

    fn remove_alt_chunk(&mut self, index: usize) -> Result<AltChunk> {
        let position = self
            .alt_chunk_positions()
            .get(index)
            .copied()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        match self.elements.remove(position) {
            BodyElement::PreservedAltChunk(_, chunk) => Ok(chunk),
            _ => unreachable!(),
        }
    }

    fn move_alt_chunk(&mut self, from: usize, to: usize) -> Result<()> {
        let count = self.alt_chunk_positions().len();
        if from >= count || to >= count {
            return Err(OoxmlError::InvalidFormat(format!(
                "altChunk move {from} -> {to} is out of range"
            )));
        }
        if from == to {
            return Ok(());
        }
        let source = self.alt_chunk_positions()[from];
        let element = self.elements.remove(source);
        let remaining = self.alt_chunk_positions();
        let destination = remaining
            .get(to)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        self.elements.insert(destination, element);
        Ok(())
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
                | BodyElement::PreservedOther(raw)
                | BodyElement::PreservedAltChunk(raw, _) => xml.push_str(raw),
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

    fn final_section_properties(&self) -> Result<Option<SectionProperties>> {
        self.elements
            .iter()
            .find_map(|element| match element {
                BodyElement::PreservedSectionProperties(raw) => {
                    Some(SectionProperties::from_xml(raw))
                },
                _ => None,
            })
            .transpose()
    }

    fn validate_section_placement(&self) -> Result<()> {
        let mut final_section = None;
        for (index, element) in self.elements.iter().enumerate() {
            match element {
                BodyElement::PreservedSectionProperties(raw) => {
                    if final_section.replace(index).is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "document body contains multiple final section properties".to_string(),
                        ));
                    }
                    SectionProperties::from_xml(raw)?;
                },
                BodyElement::PreservedParagraph(raw) => {
                    paragraph_section_range(raw)?;
                },
                _ => {},
            }
        }
        if let Some(index) = final_section
            && self.elements[index + 1..].iter().any(|element| {
                !matches!(element, BodyElement::PreservedOther(raw) if raw.trim().is_empty())
            })
        {
            return Err(OoxmlError::InvalidFormat(
                "body-final section properties are not the final body child".to_string(),
            ));
        }
        Ok(())
    }

    fn section_break_count(&self) -> Result<usize> {
        let mut count = 0usize;
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) if paragraph.properties.section.is_some() => {
                    count += 1;
                },
                BodyElement::PreservedParagraph(raw) if paragraph_section_range(raw)?.is_some() => {
                    count += 1;
                },
                _ => {},
            }
        }
        Ok(count)
    }

    fn insert_section_break(
        &mut self,
        paragraph_index: usize,
        properties: SectionProperties,
    ) -> Result<()> {
        let element = self.paragraph_element_mut(paragraph_index).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("paragraph index {paragraph_index} is out of range"))
        })?;
        match element {
            BodyElement::Paragraph(paragraph) => {
                if paragraph.properties.section.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "paragraph already ends a section".to_string(),
                    ));
                }
                paragraph.set_section_break(properties)
            },
            BodyElement::PreservedParagraph(raw) => {
                if paragraph_section_range(raw)?.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "paragraph already ends a section".to_string(),
                    ));
                }
                let mut section_xml = String::new();
                properties.write_xml(&mut section_xml, None)?;
                *raw = insert_paragraph_property(raw, &section_xml)?;
                Ok(())
            },
            _ => unreachable!(),
        }
    }

    fn section_break(&self, index: usize) -> Result<SectionProperties> {
        let mut current = 0usize;
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        if current == index {
                            return Ok(section.clone());
                        }
                        current += 1;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            return SectionProperties::from_xml(&raw[start..end]);
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(OoxmlError::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    fn update_section_break(
        &mut self,
        index: usize,
        update: impl FnOnce(&mut SectionProperties),
    ) -> Result<()> {
        let mut current = 0usize;
        let mut update = Some(update);
        for element in &mut self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = paragraph.properties.section.as_mut() {
                        if current == index {
                            update.take().expect("called once")(section);
                            return section.validate();
                        }
                        current += 1;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            let mut section = SectionProperties::from_xml(&raw[start..end])?;
                            update.take().expect("called once")(&mut section);
                            section.validate()?;
                            let mut replacement = String::new();
                            section.write_xml(&mut replacement, None)?;
                            raw.replace_range(start..end, &replacement);
                            return Ok(());
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(OoxmlError::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    fn remove_section_break(&mut self, index: usize) -> Result<SectionProperties> {
        let mut current = 0usize;
        for element in &mut self.elements {
            match element {
                BodyElement::Paragraph(paragraph) if paragraph.properties.section.is_some() => {
                    if current == index {
                        return paragraph.remove_section_break().ok_or_else(|| {
                            OoxmlError::InvalidFormat("section break disappeared".to_string())
                        });
                    }
                    current += 1;
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            let section = SectionProperties::from_xml(&raw[start..end])?;
                            raw.replace_range(start..end, "");
                            return Ok(section);
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(OoxmlError::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    fn paragraph_element_mut(&mut self, index: usize) -> Option<&mut BodyElement> {
        let mut current = 0usize;
        for element in &mut self.elements {
            if matches!(
                element,
                BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
            ) {
                if current == index {
                    return Some(element);
                }
                current += 1;
            }
        }
        None
    }

    fn collect_section_parts(
        &self,
        parts: &mut Vec<(bool, super::section::SectionHeaderFooterPart)>,
    ) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        collect_section_parts(section, parts)?;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        collect_section_parts(
                            &SectionProperties::from_xml(&raw[start..end])?,
                            parts,
                        )?;
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }

    fn collect_explicit_section_relationships(
        &self,
        relationships: &mut Vec<(String, bool)>,
    ) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        collect_explicit_section_relationships(section, relationships);
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        collect_explicit_section_relationships(
                            &SectionProperties::from_xml(&raw[start..end])?,
                            relationships,
                        );
                    }
                },
                _ => {},
            }
        }
        Ok(())
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
                | BodyElement::PreservedOther(raw)
                | BodyElement::PreservedAltChunk(raw, _) => xml.push_str(raw),
                BodyElement::PreservedSectionProperties(raw) if preserve_section => {
                    xml.push_str(raw);
                },
                BodyElement::PreservedSectionProperties(_) => {},
            }
        }
        Ok(())
    }
}

fn collect_section_parts(
    section: &SectionProperties,
    parts: &mut Vec<(bool, super::section::SectionHeaderFooterPart)>,
) -> Result<()> {
    section.validate()?;
    for reference in &section.headers {
        if let Some(part) = &reference.part {
            parts.push((true, part.clone()));
        }
    }
    for reference in &section.footers {
        if let Some(part) = &reference.part {
            parts.push((false, part.clone()));
        }
    }
    Ok(())
}

fn collect_explicit_section_relationships(
    section: &SectionProperties,
    relationships: &mut Vec<(String, bool)>,
) {
    for reference in &section.headers {
        if let Some(id) = &reference.relationship_id {
            relationships.push((id.clone(), true));
        }
    }
    for reference in &section.footers {
        if let Some(id) = &reference.relationship_id {
            relationships.push((id.clone(), false));
        }
    }
}

fn word_ranges(xml: &str, target: &[u8]) -> Result<Vec<(usize, usize)>> {
    let mut ranges = Vec::new();
    crate::docx::namespace::scan_word_element_ranges(
        xml.as_bytes(),
        &[target],
        |_, start, length| {
            let start = usize::try_from(start)
                .map_err(|_| OoxmlError::InvalidFormat("Word range overflow".to_string()))?;
            let length = usize::try_from(length)
                .map_err(|_| OoxmlError::InvalidFormat("Word range overflow".to_string()))?;
            ranges.push((start, start + length));
            Ok(())
        },
    )?;
    Ok(ranges)
}

fn paragraph_section_range(xml: &str) -> Result<Option<(usize, usize)>> {
    let sections = word_ranges(xml, b"sectPr")?;
    if sections.len() > 1 {
        return Err(OoxmlError::InvalidFormat(
            "paragraph contains multiple section properties".to_string(),
        ));
    }
    let Some(section) = sections.first().copied() else {
        return Ok(None);
    };
    let properties = word_ranges(xml, b"pPr")?;
    if properties.len() != 1 || section.0 < properties[0].0 || section.1 > properties[0].1 {
        return Err(OoxmlError::InvalidFormat(
            "paragraph section properties must be inside one pPr".to_string(),
        ));
    }
    let close = xml[..properties[0].1]
        .rfind("</")
        .unwrap_or(properties[0].1);
    if !xml[section.1..close].trim().is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "paragraph section properties must be the final pPr child".to_string(),
        ));
    }
    SectionProperties::from_xml(&xml[section.0..section.1])?;
    Ok(Some(section))
}

fn insert_paragraph_property(xml: &str, property: &str) -> Result<String> {
    let properties = word_ranges(xml, b"pPr")?;
    if properties.len() > 1 {
        return Err(OoxmlError::InvalidFormat(
            "paragraph contains multiple pPr elements".to_string(),
        ));
    }
    if let Some((start, end)) = properties.first().copied() {
        if xml[start..end].trim_end().ends_with("/>") {
            let empty_end = xml[..end].rfind("/>").ok_or_else(|| {
                OoxmlError::InvalidFormat("invalid empty paragraph properties".to_string())
            })?;
            let name_end = xml[start + 1..]
                .find(|ch: char| ch.is_whitespace() || ch == '/' || ch == '>')
                .map(|offset| start + 1 + offset)
                .ok_or_else(|| OoxmlError::InvalidFormat("invalid pPr name".to_string()))?;
            let name = &xml[start + 1..name_end];
            return Ok(format!(
                "{}>{property}</{name}>{}",
                &xml[..empty_end],
                &xml[end..]
            ));
        }
        let close = xml[..end].rfind("</").ok_or_else(|| {
            OoxmlError::InvalidFormat("paragraph properties are not closed".to_string())
        })?;
        return Ok(format!("{}{property}{}", &xml[..close], &xml[close..]));
    }

    let open_end = xml.find('>').ok_or_else(|| {
        OoxmlError::InvalidFormat("paragraph opening element is missing".to_string())
    })?;
    if xml[..=open_end].trim_end().ends_with("/>") {
        let empty_end = xml[..=open_end].rfind("/>").expect("checked");
        let name_end = xml[1..]
            .find(|ch: char| ch.is_whitespace() || ch == '/' || ch == '>')
            .map(|offset| 1 + offset)
            .ok_or_else(|| OoxmlError::InvalidFormat("invalid paragraph name".to_string()))?;
        let name = &xml[1..name_end];
        return Ok(format!(
            "{}><w:pPr>{property}</w:pPr></{name}>{}",
            &xml[..empty_end],
            &xml[open_end + 1..]
        ));
    }
    Ok(format!(
        "{}<w:pPr>{property}</w:pPr>{}",
        &xml[..=open_end],
        &xml[open_end + 1..]
    ))
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
            b"altChunk" => PreservedBodyKind::AltChunk,
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
        PreservedBodyKind::AltChunk => {
            let namespace_pairs = [
                (
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ),
                (
                    "http://purl.oclc.org/ooxml/wordprocessingml/main",
                    "http://purl.oclc.org/ooxml/officeDocument/relationships",
                ),
            ];
            let parsed = namespace_pairs
                .into_iter()
                .find_map(|(word, relationship)| {
                    let wrapped = format!(
                        r#"<root xmlns:w="{word}" xmlns:r="{relationship}">{raw_xml}</root>"#
                    );
                    let mut chunks = scan_alt_chunks(wrapped.as_bytes()).ok()?;
                    (chunks.len() == 1).then(|| chunks.pop_first().expect("length checked").1)
                });
            BodyElement::PreservedAltChunk(
                raw_xml,
                parsed.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "direct altChunk body child did not parse as one anchor".to_string(),
                    )
                })?,
            )
        },
        PreservedBodyKind::Other => BodyElement::PreservedOther(raw_xml),
    });
    Ok(())
}

/// A body element (paragraph, table, or exact preserved XML).
#[derive(Debug)]
#[allow(clippy::large_enum_variant)] // writer-internal type; variants are moved, not compared
pub(crate) enum BodyElement {
    Paragraph(MutableParagraph),
    Table(MutableTable),
    PreservedParagraph(String),
    PreservedTable(String),
    PreservedSectionProperties(String),
    PreservedAltChunk(String, AltChunk),
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

    #[test]
    fn protection_patching_preserves_unrelated_settings_exactly() {
        let input = br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:extension"><!--before--><q:smartTagType q:namespaceuri="urn:test" q:name="person" q:url="https://example.test"/><q:documentProtection q:edit="readOnly" q:enforcement="1" x:keep="yes"/><x:opaque><![CDATA[a < b]]></x:opaque><q:doNotEmbedSmartTags/></q:settings>"#;
        let mut document = MutableDocument::new();
        document.set_protection_with_password(
            ProtectionType::Comments,
            "hash&\"value".into(),
            "salt<value".into(),
        );

        let output = document.generate_settings_xml(Some(input)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"<q:smartTagType q:namespaceuri="urn:test" q:name="person" q:url="https://example.test"/>"#));
        assert!(output.contains(r#"<x:opaque><![CDATA[a < b]]></x:opaque>"#));
        assert!(output.contains("<q:doNotEmbedSmartTags/>"));
        assert!(output.contains(r#"<q:documentProtection q:edit="comments" q:enforcement="1" q:hash="hash&amp;&quot;value" q:salt="salt&lt;value"/>"#));
        assert_eq!(output.matches("documentProtection").count(), 1);
    }

    #[test]
    fn protection_patching_removes_only_protection_and_handles_empty_roots() {
        let input = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="125"/><w:documentProtection w:edit="forms"/><w:savePreviewPicture/></w:settings>"#;
        let mut document = MutableDocument::new();
        document.remove_protection();
        let output = document.generate_settings_xml(Some(input)).unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="125"/><w:savePreviewPicture/></w:settings>"#
        );

        document.set_protection(ProtectionType::ReadOnly);
        let empty = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"/>"#;
        let output =
            String::from_utf8(document.generate_settings_xml(Some(empty)).unwrap()).unwrap();
        assert_eq!(
            output,
            r#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:documentProtection s:edit="readOnly" s:enforcement="1"/></s:settings>"#
        );

        let default_namespace =
            br#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let output = String::from_utf8(
            document
                .generate_settings_xml(Some(default_namespace))
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            output,
            r#"<settings xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:edit="readOnly" w:enforcement="1"/></settings>"#
        );
    }
}
