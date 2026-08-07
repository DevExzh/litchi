use crate::OfficeMath;
use crate::alt::{Chunk, Conformance};
/// Document writer implementation for DOCX.
use crate::error::{Error, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
use super::super::super::format::ImageFormat;
// Import from other writer modules
use super::super::comment::MutableComment;
use super::super::note::Note;
use super::super::ole_object::MutableOleObject;
use super::super::paragraph::{MutableParagraph, ParagraphElement};
use super::super::section::SectionProperties;
use super::super::smartart::{MAX_SMART_ARTS, MutableSmartArt};
use super::super::table::MutableTable;
use super::super::theme::MutableTheme;
use super::super::toc::TableOfContents;
use super::super::vml_shape::MutableVmlShape;
use super::super::watermark::{ImageWatermark, Watermark};
use std::collections::HashSet;
// Import settings types
use super::super::super::settings::ProtectionType;
use super::codec::{patch_document_protection, write_document_protection};
use super::package::{
    BodyElement, DocumentBody, collect_explicit_section_relationships, collect_section_parts,
    shift_toc_index_on_insert, shift_toc_index_on_remove,
};

/// A mutable Word document for writing and modification.
///
/// Provides methods to add and modify document content including paragraphs,
/// runs, tables, sections, and other elements.
pub struct MutableDocument {
    /// Document body content (paragraphs, tables, etc.)
    pub(crate) body: DocumentBody,
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
    protection: Option<Protection>,
    /// Whether document protection was explicitly changed.
    protection_dirty: bool,
    /// Section properties (page setup, margins, orientation)
    pub(super) section: SectionProperties,
    /// Theme (optional)
    theme: Option<MutableTheme>,
    /// Watermark (optional)
    pub(crate) watermark: Option<Watermark>,
    /// Image watermark (optional)
    pub(crate) image_watermark: Option<ImageWatermark>,
    /// Table of Contents configuration (optional)
    toc_config: Option<(usize, TableOfContents)>, // (insertion index, config)
    /// Whether the document has been modified
    pub(crate) modified: bool,
    /// Exact document/root/body opening XML retained from an existing document.
    pub(super) preserved_prefix: Option<String>,
    /// Exact body/document closing XML retained from an existing document.
    pub(super) preserved_suffix: Option<String>,
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
pub struct Protection {
    /// Type of protection
    pub protection_type: ProtectionType,
    /// Password hash (optional, for actual enforcement)
    pub password_hash: Option<String>,
    /// Salt for password hash (optional)
    pub salt: Option<String>,
}

#[cfg(feature = "fonts")]
use super::super::smart_tag::MutableSmartTag;
#[cfg(feature = "fonts")]
use litchi_fonts::{CollectGlyphs, GlyphMap};

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableDocument {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();

        // Collect from body elements
        for element in &self.body.elements {
            let element_glyphs = match element {
                BodyElement::Paragraph(p) => p.collect_glyphs(),
                BodyElement::Table(t) => t.collect_glyphs(),
                BodyElement::PreservedParagraph(_)
                | BodyElement::PreservedTable(_)
                | BodyElement::PreservedSectionProperties(_)
                | BodyElement::PreservedAlt(_, _)
                | BodyElement::PreservedOther(_) => continue,
            };
            for (font, bitmap) in element_glyphs {
                *glyphs.entry(font).or_default() |= bitmap;
            }
        }

        // Collect from headers
        if let Some(headers) = &self.header {
            for p in headers {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_default() |= bitmap;
                }
            }
        }

        // Collect from footers
        if let Some(footers) = &self.footer {
            for p in footers {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_default() |= bitmap;
                }
            }
        }

        // Collect from footnotes/endnotes
        for note in self.footnotes.iter().chain(self.endnotes.iter()) {
            for p in &note.paragraphs {
                for (font, bitmap) in p.collect_glyphs() {
                    *glyphs.entry(font).or_default() |= bitmap;
                }
            }
        }

        glyphs
    }
}

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableParagraph {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        for element in &self.elements {
            let element_glyphs = match element {
                ParagraphElement::Run(r) => r.collect_glyphs(),
                ParagraphElement::Hyperlink(h) => h.collect_glyphs(),
                ParagraphElement::SmartTag(tag) => tag.collect_glyphs(),
                _ => continue,
            };
            for (font, bitmap) in element_glyphs {
                *glyphs.entry(font).or_default() |= bitmap;
            }
        }
        glyphs
    }
}

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableSmartTag {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        for element in &self.elements {
            let element_glyphs = match element {
                ParagraphElement::Run(run) => run.collect_glyphs(),
                ParagraphElement::Hyperlink(hyperlink) => hyperlink.collect_glyphs(),
                ParagraphElement::SmartTag(tag) => tag.collect_glyphs(),
                _ => continue,
            };
            for (font, bitmap) in element_glyphs {
                *glyphs.entry(font).or_default() |= bitmap;
            }
        }
        glyphs
    }
}

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableTable {
    fn collect_glyphs(&self) -> GlyphMap {
        let mut glyphs = GlyphMap::new();
        for row in &self.rows {
            for cell in &row.cells {
                for p in &cell.paragraphs {
                    for (font, bitmap) in p.collect_glyphs() {
                        *glyphs.entry(font).or_default() |= bitmap;
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

    /// Whether glyph collection covers every text-bearing node in this value.
    #[cfg(feature = "automatic-fonts")]
    pub(crate) fn glyphs_are_complete(&self) -> bool {
        self.preserved_prefix.is_none()
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
    ) -> Result<Vec<(bool, super::super::section::SectionHeaderFooterPart)>> {
        let mut parts = Vec::new();
        collect_section_parts(&self.section, &mut parts)?;
        self.body.collect_section_parts(&mut parts)?;
        let mut unique: Vec<(bool, super::super::section::SectionHeaderFooterPart)> = Vec::new();
        for (header, part) in parts {
            if let Some((existing_header, existing)) =
                unique.iter().find(|(_, existing)| existing.key == part.key)
            {
                if *existing_header != header || existing.xml != part.xml {
                    return Err(Error::InvalidFormat(format!(
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
            return Err(Error::InvalidFormat(
                "Heading level must be 0-9".to_string(),
            ));
        }
        let style = if level == 0 {
            "Title".to_string()
        } else {
            format!("Heading{level}")
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
    /// [`crate::Document::text_boxes`] inventory after save and reopen.
    pub fn add_text_box(
        &mut self,
        text_box: super::super::textbox::MutableTextBox,
    ) -> &mut super::super::textbox::MutableTextBox {
        self.add_paragraph().add_text_box(text_box)
    }

    /// Embed an OLE/package object in a new paragraph at the end of the
    /// document.
    ///
    /// Assigns the object's VML shape identity when unset, rejecting explicit
    /// IDs that collide with shapes already present in the document. The
    /// payload is stored verbatim as an inert `/word/embeddings/oleObjectN.bin`
    /// part and is discoverable through
    /// [`crate::Package::embedded`] after save and reopen.
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
                    Error::InvalidFormat("OLE shape ID space exhausted".to_string())
                })?;
            };
            self.next_ole_shape_number = number.saturating_add(1);
            object.shape_id = shape_id;
            object.object_id = number;
        } else {
            if self.shape_id_in_use(&object.shape_id) {
                return Err(Error::InvalidFormat(format!(
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
            BodyElement::PreservedAlt(raw, _) => raw.contains(shape_id),
            _ => false,
        })
    }

    /// Add a legacy VML shape in a new paragraph at the end of the document.
    ///
    /// Assigns the shape's VML identity (`_x0000_s1025`, …, matching Word's
    /// numbering convention) when unset, skipping IDs already used by
    /// authored shapes or present in preserved document XML. A shape with a
    /// `v:textbox` story is discoverable through
    /// [`crate::Document::text_boxes`] after save and reopen.
    pub fn add_vml_shape(&mut self, mut shape: MutableVmlShape) -> Result<&mut MutableVmlShape> {
        if shape.id.is_empty() {
            let mut number = self.next_vml_shape_number;
            let id = loop {
                let candidate = format!("_x0000_s{number}");
                if !self.shape_id_in_use(&candidate) {
                    break candidate;
                }
                number = number.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("VML shape ID space exhausted".to_string())
                })?;
            };
            self.next_vml_shape_number = number.saturating_add(1);
            shape.id = id;
        } else if self.shape_id_in_use(&shape.id) {
            return Err(Error::InvalidFormat(format!(
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
    /// [`crate::Document::smart_arts`] after save and reopen. The
    /// optional pre-rendered drawing part is not generated; Word and
    /// LibreOffice re-render from the layout and data parts.
    pub fn add_smart_art(&mut self, mut smartart: MutableSmartArt) -> Result<&mut MutableSmartArt> {
        if self.collect_smart_arts().len() >= MAX_SMART_ARTS {
            return Err(Error::InvalidFormat(format!(
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
            .ok_or_else(|| Error::InvalidFormat(format!("footnote ID {id} does not exist")))?;
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
            .ok_or_else(|| Error::InvalidFormat(format!("endnote ID {id} does not exist")))?;
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
            .ok_or_else(|| Error::InvalidFormat(format!("comment ID {id} does not exist")))?;
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
    /// use litchi_docx::settings::ProtectionType;
    ///
    /// // Protect document as read-only
    /// doc.set_protection(ProtectionType::ReadOnly);
    ///
    /// // Allow only comments
    /// doc.set_protection(ProtectionType::Comments);
    /// ```
    pub fn set_protection(&mut self, protection_type: ProtectionType) {
        self.protection = Some(Protection {
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
        self.protection = Some(Protection {
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
    /// use litchi_docx::writer::MutableTheme;
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
    /// use litchi_docx::writer::Watermark;
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
    /// discoverable through [`crate::Document::image_watermarks`] after
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
    /// use litchi_docx::writer::TableOfContents;
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
        use super::super::field::MutableField;
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
            .push(ParagraphElement::Field(MutableField::begin()));
        toc_field_para
            .elements
            .push(ParagraphElement::Field(MutableField::instruction_char(
                instruction,
            )));
        toc_field_para
            .elements
            .push(ParagraphElement::Field(MutableField::separate()));

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
                .push(super::super::paragraph::TabStop {
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
                super::super::hyperlink::MutableHyperlink::new_anchor(bookmark_name.clone());

            let mut text_run = super::super::run::MutableRun::new();
            text_run.set_text(&heading_text);
            text_run.properties.no_proof = true;
            hyperlink.add_run(text_run);

            let mut tab_run = super::super::run::MutableRun::new();
            tab_run.add_tab();
            tab_run.properties.no_proof = true;
            tab_run.properties.web_hidden = true;
            hyperlink.add_run(tab_run);

            hyperlink
                .elements
                .push(super::super::hyperlink::HyperlinkElement::Field(
                    MutableField::begin(),
                ));

            let mut pageref_instr = String::new();
            write!(&mut pageref_instr, " PAGEREF {} \\h ", bookmark_name).unwrap();
            hyperlink
                .elements
                .push(super::super::hyperlink::HyperlinkElement::Field(
                    MutableField::instruction_char(pageref_instr),
                ));

            hyperlink
                .elements
                .push(super::super::hyperlink::HyperlinkElement::Field(
                    MutableField::separate(),
                ));

            let mut page_run = super::super::run::MutableRun::new();
            page_run.set_text("1");
            page_run.properties.no_proof = true;
            page_run.properties.web_hidden = true;
            hyperlink.add_run(page_run);

            hyperlink
                .elements
                .push(super::super::hyperlink::HyperlinkElement::Field(
                    MutableField::end(),
                ));

            toc_entry
                .elements
                .push(ParagraphElement::Hyperlink(hyperlink));
            toc_paragraphs.push(toc_entry);
        }

        // Add field end to the first TOC paragraph
        if let Some(first_para) = toc_paragraphs.first_mut() {
            first_para
                .elements
                .push(ParagraphElement::Field(MutableField::end()));
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
                    if let ParagraphElement::OleObject(object) = para_element {
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
                    if let ParagraphElement::SmartArt(smartart) = para_element {
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
                .map_err(|e| Error::Xml(e.to_string()))?;

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
                .map_err(|e| Error::Xml(e.to_string()))?;

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
    pub fn alts(&self) -> Vec<Chunk> {
        self.body.alts()
    }

    /// Insert an alternative-format anchor at an anchor-relative index.
    pub(crate) fn insert_alt(
        &mut self,
        index: usize,
        chunk: Chunk,
        namespace: Conformance,
    ) -> Result<()> {
        let position = self.body.insert_alt(index, chunk, namespace)?;
        shift_toc_index_on_insert(&mut self.toc_config, position);
        self.modified = true;
        Ok(())
    }

    /// Replace an alternative-format anchor without disturbing adjacent XML.
    pub(crate) fn replace_alt(
        &mut self,
        index: usize,
        chunk: Chunk,
        namespace: Conformance,
    ) -> Result<Chunk> {
        let old = self.body.replace_alt(index, chunk, namespace)?;
        self.modified = true;
        Ok(old)
    }

    /// Remove an alternative-format anchor.
    pub(crate) fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        let (position, old) = self.body.remove_alt(index)?;
        shift_toc_index_on_remove(&mut self.toc_config, position);
        self.modified = true;
        Ok(old)
    }

    /// Move an alternative-format anchor to another anchor-relative index.
    pub(crate) fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        self.body.move_alt(from, to)?;
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
        rel_mapper: &super::super::relmap::RelationshipMapper,
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
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}
