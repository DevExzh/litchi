//! Mutable document structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODT documents that allows
//! for in-place modification of content, styles, and metadata.

use crate::BookmarkTarget;
use crate::core::{
    MetaXmlPatch, OdfMetadata, OdfStructure, OwnedPackage, PackageWriter, patch_meta_xml,
};
use crate::elements::field::{FieldParser, OdfDynamicTextField};
use crate::elements::parser::DocumentOrderElement;
use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, Paragraph};
use crate::odt::Document;
use crate::odt::ReferenceMark;
use crate::odt::TextIndex;
use crate::odt::TextIndexMark;
use crate::odt::header_footer::{
    HeaderFooterKind, MasterPage, add_master_page, parse_master_pages, set_region_text,
    set_region_xml,
};
use crate::odt::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};
use crate::odt::page_sequence::{OdtPageSequence, parse_page_sequence, set_page_sequence_xml};
use crate::{OdfFormProperty, OdfInteractiveControl, OdfSelectionControl, OdfTextControl};
use crate::{
    OdfVariableDeclarationGroup, OdfVariableDeclarations, OdfVariableKind, OdfVariablePart,
    OdfVariableScope,
};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::{ops::Range, path::Path};

/// Document element type for tracking insertion order
#[derive(Debug, Clone)]
enum DocumentElement {
    /// A paragraph element
    Paragraph(Paragraph),
    /// A heading element
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A text list element
    List(List),
    /// A standalone drawing frame (image or text box) at body level
    Frame(crate::elements::element::Element),
}

/// A mutable ODT document that supports in-place modifications.
///
/// This struct wraps an ODT document and provides methods to modify its content,
/// including adding, updating, and removing paragraphs, tables, and other elements.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Document, MutableDocument};
///
/// # fn main() -> litchi_core::Result<()> {
/// // Open an existing document
/// let doc = Document::open("input.odt")?;
/// let mut mutable_doc = MutableDocument::from_document(doc)?;
///
/// // Modify the document
/// mutable_doc.add_paragraph("New paragraph")?;
/// mutable_doc.remove_paragraph(0)?;
///
/// // Save the modified document
/// mutable_doc.save("output.odt")?;
/// # Ok(())
/// # }
/// ```
pub struct MutableDocument {
    /// Document elements in insertion order (paragraphs and tables mixed)
    elements: Vec<DocumentElement>,
    /// Document metadata (mutable)
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is for now)
    styles_xml: Option<String>,
    /// Original package used to retain auxiliary package parts during rewriting.
    source_package: Option<OwnedPackage>,
    /// Authoritative original content XML used by byte-preserving inline mutations.
    content_xml: Option<String>,
    /// Authored picture payloads written into the package on save.
    pending_images: Vec<crate::odt::frame::PendingImage>,
    /// Monotonic counter for authored frame names (1-based).
    next_frame_number: usize,
}

impl MutableDocument {
    /// Create a mutable document from an existing Document.
    ///
    /// This parses the document structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `doc` - The document to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{Document, MutableDocument};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let mut mutable_doc = MutableDocument::from_document(doc)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_document(doc: Document) -> Result<Self> {
        let content_xml = String::from_utf8(doc.get_file("content.xml")?).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!("content.xml is not UTF-8: {error}"))
        })?;
        let source_elements = doc.elements()?;
        let metadata = doc.metadata()?;

        // Get MIME type from package
        let mimetype = "application/vnd.oasis.opendocument.text".to_string();

        // Extract styles XML from the document's package
        let styles_xml = doc
            .get_file("styles.xml")
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok());

        let elements = source_elements
            .into_iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(paragraph) => DocumentElement::Paragraph(paragraph),
                DocumentOrderElement::NumberedParagraph(paragraph) => {
                    DocumentElement::Paragraph(paragraph.into_paragraph())
                },
                DocumentOrderElement::Heading(heading) => DocumentElement::Heading(heading),
                DocumentOrderElement::Table(table) => DocumentElement::Table(table),
                DocumentOrderElement::List(list) => DocumentElement::List(list),
            })
            .collect();
        let source_package = Some(doc.into_package());

        Ok(Self {
            elements,
            metadata,
            mimetype,
            styles_xml,
            source_package,
            content_xml: Some(content_xml),
            pending_images: Vec::new(),
            next_frame_number: 1,
        })
    }

    /// Create a new empty mutable document.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::MutableDocument;
    ///
    /// let doc = MutableDocument::new();
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.text".to_string(),
            styles_xml: None,
            source_package: None,
            content_xml: None,
            pending_images: Vec::new(),
            next_frame_number: 1,
        }
    }

    /// Override the root MIME type written by `to_bytes`.
    ///
    /// Used by the web-template authoring model to emit the legacy
    /// `application/vnd.oasis.opendocument.text-web` MIME type.
    pub(crate) fn set_mimetype(&mut self, mimetype: impl Into<String>) {
        self.mimetype = mimetype.into();
    }

    /// Return typed dynamic fields from the current authoritative content XML.
    pub fn dynamic_text_fields(&self) -> Result<Vec<OdfDynamicTextField>> {
        self.with_content_xml(FieldParser::parse_dynamic_text_fields)
    }

    /// Return semantic footnotes and endnotes from the current content XML.
    pub fn notes(&self) -> Result<Vec<crate::Note>> {
        self.with_content_xml(crate::odt::parse_notes)
    }

    /// Return only footnotes from the current content XML.
    pub fn footnotes(&self) -> Result<Vec<crate::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == crate::NoteClass::Footnote)
            .collect())
    }

    /// Return only endnotes from the current content XML.
    pub fn endnotes(&self) -> Result<Vec<crate::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == crate::NoteClass::Endnote)
            .collect())
    }

    /// Append a validated footnote or endnote to one `text:p` paragraph.
    ///
    /// The paragraph is selected in document order, including paragraphs nested
    /// in lists, tables, and note bodies. Structured note bodies are serialized
    /// from their validated public model; all embedded content remains inert.
    pub fn insert_note(&mut self, paragraph_index: usize, note: &crate::Note) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::insert_note_xml(xml, paragraph_index, note))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one note selected in document order and return its old semantic value.
    ///
    /// Replacement emits the public note model, including validated structured
    /// content when the replacement carries an `OdfNoteBodyContent`.
    pub fn replace_note(
        &mut self,
        note_index: usize,
        replacement: &crate::Note,
    ) -> Result<crate::Note> {
        let old = self.notes()?.get(note_index).cloned().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("note index {note_index} is out of bounds"))
        })?;
        let updated =
            self.with_content_xml(|xml| crate::replace_note_xml(xml, note_index, replacement))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove one note selected in document order and return its old semantic value.
    pub fn remove_note(&mut self, note_index: usize) -> Result<crate::Note> {
        let old = self.notes()?.get(note_index).cloned().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!("note index {note_index} is out of bounds"))
        })?;
        let updated = self.with_content_xml(|xml| crate::remove_note_xml(xml, note_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return structure-preserving ruby annotations from the current content XML.
    pub fn ruby_annotations(&self) -> Result<crate::RubyAnnotations> {
        self.with_content_xml(crate::parse_ruby_annotations)
    }

    /// Append a validated ruby annotation to one `text:p` paragraph.
    ///
    /// The annotation is inserted at the end of the paragraph selected in
    /// document order. It is purely document metadata and never triggers any
    /// external lookup or code execution.
    pub fn insert_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        annotation: &crate::RubyAnnotation,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_ruby_annotation_xml(xml, paragraph_index, annotation)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Wrap one UTF-8 text range in a selected paragraph with ruby.
    ///
    /// The range uses the structural coordinate space accepted by
    /// `wrap_ruby_annotation_xml`: a plain base may span adjacent character
    /// data under one parent, while an XML base may span balanced legal inline
    /// elements. Existing ancestors and ruby annotations are never split.
    /// Ruby insertion is inert and does not execute scripts, macros, links, or
    /// external content.
    pub fn wrap_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        range: Range<usize>,
        annotation: &crate::RubyAnnotation,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::wrap_ruby_annotation_xml(xml, paragraph_index, range, annotation)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a ruby annotation selected in document order and return its old value.
    pub fn replace_ruby_annotation(
        &mut self,
        annotation_index: usize,
        replacement: &crate::RubyAnnotation,
    ) -> Result<crate::RubyAnnotation> {
        let old = self
            .ruby_annotations()?
            .annotations
            .get(annotation_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "ruby annotation index {annotation_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_ruby_annotation_xml(xml, annotation_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a ruby annotation selected in document order and return its old value.
    pub fn remove_ruby_annotation(
        &mut self,
        annotation_index: usize,
    ) -> Result<crate::RubyAnnotation> {
        let old = self
            .ruby_annotations()?
            .annotations
            .get(annotation_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "ruby annotation index {annotation_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_ruby_annotation_xml(xml, annotation_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return typed named ruby styles from the current `styles.xml`.
    pub fn ruby_styles(&self) -> Result<crate::RubyStyles> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::parse_ruby_styles)
    }

    /// Insert or replace one named ruby style definition and return the old value.
    pub fn set_ruby_style(&mut self, style: &crate::RubyStyle) -> Result<Option<crate::RubyStyle>> {
        style.validate()?;
        let old = self.ruby_styles()?.get(&style.name).cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        self.styles_xml = Some(crate::set_ruby_style_xml(&styles, style)?);
        Ok(old)
    }

    /// Remove one named ruby style definition and return the old value.
    ///
    /// Existing `text:ruby` style references are preserved verbatim, so callers
    /// can intentionally manage their lifecycle separately.
    pub fn remove_ruby_style(&mut self, name: &str) -> Result<Option<crate::RubyStyle>> {
        let old = self.ruby_styles()?.get(name).cloned();
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::remove_ruby_style_xml(styles, name)?);
        Ok(old)
    }

    /// Return font-face declarations from the current `content.xml`.
    ///
    /// Linked font resources remain inert metadata. This does not fetch a URI,
    /// load a font, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        self.with_content_xml(crate::font_face::parse_content_font_face_declarations)
    }

    /// Replace content-part font-face declarations and return the old value.
    ///
    /// This edits `content.xml` only. It does not fetch linked font resources,
    /// load a font, or inspect embedded font data.
    pub fn set_content_font_face_declarations(
        &mut self,
        declarations: &crate::OdfFontFaceDeclarations,
    ) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        let (updated, old) = self.with_content_xml(|xml| {
            crate::font_face::set_content_font_face_declarations_xml(xml, declarations)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove content-part font-face declarations and return the old value.
    ///
    /// This edits `content.xml` only. Existing style references remain
    /// verbatim so callers can manage their lifecycle separately.
    pub fn clear_content_font_face_declarations(
        &mut self,
    ) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        let (updated, old) =
            self.with_content_xml(crate::font_face::remove_content_font_face_declarations_xml)?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return font-face declarations from the current `styles.xml`.
    ///
    /// Linked font resources remain inert metadata. This does not fetch a URI,
    /// load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(None),
            crate::font_face::parse_styles_font_face_declarations,
        )
    }

    /// Replace styles-part font-face declarations and return the old value.
    ///
    /// This edits `styles.xml` only. It does not fetch linked font resources,
    /// load a font, or inspect embedded font data.
    pub fn set_styles_font_face_declarations(
        &mut self,
        declarations: &crate::OdfFontFaceDeclarations,
    ) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        let (updated, old) =
            crate::font_face::set_styles_font_face_declarations_xml(&styles, declarations)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Remove styles-part font-face declarations and return the old value.
    ///
    /// This edits `styles.xml` only. Existing style references remain
    /// verbatim so callers can manage their lifecycle separately.
    pub fn clear_styles_font_face_declarations(
        &mut self,
    ) -> Result<Option<crate::OdfFontFaceDeclarations>> {
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        let (updated, old) = crate::font_face::remove_styles_font_face_declarations_xml(styles)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Return named legacy and SVG drawing gradients from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing_gradient::OdfDrawingGradients> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_gradient::parse_drawing_gradients,
        )
    }

    /// Return named drawing hatch resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing_hatch::OdfDrawingHatches> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_hatch::parse_drawing_hatches,
        )
    }

    /// Return named drawing stroke-dash resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing_stroke_dash::OdfDrawingStrokeDashes> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_stroke_dash::parse_drawing_stroke_dashes,
        )
    }

    /// Return named drawing fill-image definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites, follow links, load linked resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing_fill_image::OdfDrawingFillImages> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_fill_image::parse_drawing_fill_images,
        )
    }

    /// Return named drawing marker definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing_marker::OdfDrawingMarkers> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_marker::parse_drawing_markers,
        )
    }

    /// Return named drawing opacity definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing_opacity::OdfDrawingOpacities> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_opacity::parse_drawing_opacities,
        )
    }

    /// Return stored footnote and endnote presentation configurations.
    ///
    /// The result describes style metadata only. It never renumbers, lays out,
    /// or renders notes.
    pub fn notes_configurations(&self) -> Result<crate::OdfNotesConfigurations> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::parse_notes_configurations)
    }

    /// Return stored outline numbering styles from current styles metadata.
    ///
    /// The result does not apply styles to headings, generate labels, or
    /// update tables of contents.
    pub fn outline_styles(&self) -> Result<crate::OdfOutlineStyles> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::parse_outline_styles)
    }

    /// Insert or replace one named outline numbering style.
    ///
    /// This edits `styles.xml` only and returns the previous style with the
    /// same name. It does not alter heading structure or cached index content.
    pub fn set_outline_style(
        &mut self,
        style: &crate::OdfOutlineStyle,
    ) -> Result<Option<crate::OdfOutlineStyle>> {
        style.validate()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        let (updated, old) = crate::set_outline_style_xml(&styles, style)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Remove one named outline numbering style and return its prior value.
    ///
    /// Existing heading references are retained verbatim, allowing callers to
    /// manage those references separately.
    pub fn remove_outline_style(&mut self, name: &str) -> Result<Option<crate::OdfOutlineStyle>> {
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        let (updated, old) = crate::remove_outline_style_xml(styles, name)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Insert or replace one stored footnote or endnote configuration.
    ///
    /// This edits `styles.xml` only and returns the prior configuration for the
    /// same note class. It never changes note anchors, citations, or numbering.
    pub fn set_notes_configuration(
        &mut self,
        configuration: &crate::OdfNotesConfiguration,
    ) -> Result<Option<crate::OdfNotesConfiguration>> {
        configuration.validate()?;
        let old = self
            .notes_configurations()?
            .get(configuration.note_class)
            .cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        self.styles_xml = Some(crate::set_notes_configuration_xml(&styles, configuration)?);
        Ok(old)
    }

    /// Replace both stored note-class configurations and return the old values.
    ///
    /// An absent class is removed from `styles.xml`. This updates metadata only and
    /// never recalculates citations, sequence numbers, or page layout.
    pub fn set_notes_configurations(
        &mut self,
        configurations: &crate::OdfNotesConfigurations,
    ) -> Result<crate::OdfNotesConfigurations> {
        configurations.validate()?;
        let old = self.notes_configurations()?;
        if self.styles_xml.is_none()
            && configurations.footnote.is_none()
            && configurations.endnote.is_none()
        {
            return Ok(old);
        }
        let mut styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        for note_class in crate::OdfNoteClass::ALL {
            styles = match configurations.get(note_class) {
                Some(configuration) => crate::set_notes_configuration_xml(&styles, configuration)?,
                None => crate::remove_notes_configuration_xml(&styles, note_class)?,
            };
        }
        self.styles_xml = Some(styles);
        Ok(old)
    }

    /// Remove one stored note-class configuration and return its prior value.
    ///
    /// This edits style metadata only. Existing notes and their cached citations
    /// are preserved verbatim.
    pub fn clear_notes_configuration(
        &mut self,
        note_class: crate::OdfNoteClass,
    ) -> Result<Option<crate::OdfNotesConfiguration>> {
        let old = self.notes_configurations()?.get(note_class).cloned();
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::remove_notes_configuration_xml(styles, note_class)?);
        Ok(old)
    }

    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is styles metadata only. It is never used to generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(
        &self,
    ) -> Result<Option<crate::OdfBibliographyConfiguration>> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(None),
            crate::bibliography_configuration::parse_bibliography_configuration,
        )
    }

    /// Insert or replace the document-wide bibliography formatting policy.
    ///
    /// This edits `styles.xml` only and returns the prior policy. It does not
    /// regenerate bibliography entries or modify bibliography marks.
    pub fn set_bibliography_configuration(
        &mut self,
        configuration: &crate::OdfBibliographyConfiguration,
    ) -> Result<Option<crate::OdfBibliographyConfiguration>> {
        configuration.validate()?;
        let old = self.bibliography_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        self.styles_xml = Some(
            crate::bibliography_configuration::set_bibliography_configuration_xml(
                &styles,
                configuration,
            )?,
        );
        Ok(old)
    }

    /// Remove the document-wide bibliography formatting policy.
    ///
    /// This edits styles metadata only. Existing bibliography entries and
    /// source marks are preserved verbatim.
    pub fn clear_bibliography_configuration(
        &mut self,
    ) -> Result<Option<crate::OdfBibliographyConfiguration>> {
        let old = self.bibliography_configuration()?;
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml =
            Some(crate::bibliography_configuration::remove_bibliography_configuration_xml(styles)?);
        Ok(old)
    }

    /// Return stored document line-numbering configuration from current styles.
    ///
    /// The result is presentation metadata only. It is never used to paginate
    /// the document or generate line numbers.
    pub fn line_numbering_configuration(
        &self,
    ) -> Result<Option<crate::OdfLineNumberingConfiguration>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(None), crate::parse_line_numbering_configuration)
    }

    /// Insert or replace document line-numbering configuration.
    ///
    /// This updates stored style metadata only. It never calculates page or
    /// line numbers.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: &crate::OdfLineNumberingConfiguration,
    ) -> Result<Option<crate::OdfLineNumberingConfiguration>> {
        configuration.validate()?;
        let old = self.line_numbering_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(OdfStructure::default_styles_xml);
        self.styles_xml = Some(crate::line_numbering::set_line_numbering_configuration_xml(
            &styles,
            configuration,
        )?);
        Ok(old)
    }

    /// Remove document line-numbering configuration and return its old value.
    pub fn clear_line_numbering_configuration(
        &mut self,
    ) -> Result<Option<crate::OdfLineNumberingConfiguration>> {
        let old = self.line_numbering_configuration()?;
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml =
            Some(crate::line_numbering::remove_line_numbering_configuration_xml(styles)?);
        Ok(old)
    }

    /// Return generated indexes from the current authoritative content XML.
    pub fn text_indexes(&self) -> Result<Vec<TextIndex>> {
        self.with_content_xml(crate::odt::index::parse_text_indexes)
    }

    /// Append caller-authored index markup to `office:text` without refreshing its cache.
    pub fn insert_text_index(&mut self, index: &TextIndex) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::odt::insert_text_index_xml(xml, index))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one named index and return its previous inert representation.
    pub fn replace_text_index(&mut self, name: &str, replacement: &TextIndex) -> Result<TextIndex> {
        let old = self
            .text_indexes()?
            .into_iter()
            .find(|index| index.name() == name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!("text index {name:?} was not found"))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::odt::replace_text_index_xml(xml, name, replacement))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove one named index and return its previous inert representation.
    pub fn remove_text_index(&mut self, name: &str) -> Result<TextIndex> {
        let old = self
            .text_indexes()?
            .into_iter()
            .find(|index| index.name() == name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!("text index {name:?} was not found"))
            })?;
        let updated = self.with_content_xml(|xml| crate::odt::remove_text_index_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return typed point and resolved range index marks in document order.
    pub fn text_index_marks(&self) -> Result<Vec<TextIndexMark>> {
        self.with_content_xml(crate::odt::index_mark::parse_text_index_marks)
    }

    /// Insert a point mark at a paragraph end, or wrap the paragraph with a range mark.
    pub fn insert_text_index_mark(
        &mut self,
        paragraph_index: usize,
        mark: &TextIndexMark,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::odt::insert_text_index_mark_xml(xml, paragraph_index, mark)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    pub fn replace_text_index_mark(
        &mut self,
        mark_index: usize,
        replacement: &TextIndexMark,
    ) -> Result<TextIndexMark> {
        let old = self
            .text_index_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "index mark {mark_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::odt::replace_text_index_mark_xml(xml, mark_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    pub fn remove_text_index_mark(&mut self, mark_index: usize) -> Result<TextIndexMark> {
        let old = self
            .text_index_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "index mark {mark_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::odt::remove_text_index_mark_xml(xml, mark_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return point and resolved range reference targets in document order.
    pub fn reference_marks(&self) -> Result<Vec<ReferenceMark>> {
        self.with_content_xml(crate::odt::reference_mark::parse_reference_marks)
    }

    /// Insert a point reference at a paragraph end, or wrap the paragraph with a range reference.
    pub fn insert_reference_mark(
        &mut self,
        paragraph_index: usize,
        mark: &ReferenceMark,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::odt::insert_reference_mark_xml(xml, paragraph_index, mark)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a reference target selected in document order.
    pub fn replace_reference_mark(
        &mut self,
        mark_index: usize,
        replacement: &ReferenceMark,
    ) -> Result<ReferenceMark> {
        let old = self
            .reference_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "reference mark {mark_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::odt::replace_reference_mark_xml(xml, mark_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove marker elements while preserving enclosed text and markup.
    pub fn remove_reference_mark(&mut self, mark_index: usize) -> Result<ReferenceMark> {
        let old = self
            .reference_marks()?
            .get(mark_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "reference mark {mark_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::odt::remove_reference_mark_xml(xml, mark_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return point and resolved range bookmarks in document order.
    pub fn bookmark_targets(&self) -> Result<Vec<BookmarkTarget>> {
        self.with_content_xml(crate::parse_bookmark_targets)
    }

    /// Insert a point bookmark at a paragraph end, or wrap the paragraph with a range bookmark.
    pub fn insert_bookmark_target(
        &mut self,
        paragraph_index: usize,
        target: &BookmarkTarget,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::insert_bookmark_xml(xml, paragraph_index, target))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a bookmark target selected in document order.
    pub fn replace_bookmark_target(
        &mut self,
        target_index: usize,
        replacement: &BookmarkTarget,
    ) -> Result<BookmarkTarget> {
        let old = self
            .bookmark_targets()?
            .get(target_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "bookmark target {target_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::replace_bookmark_xml(xml, target_index, replacement))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove bookmark markers while preserving enclosed text and markup.
    pub fn remove_bookmark_target(&mut self, target_index: usize) -> Result<BookmarkTarget> {
        let old = self
            .bookmark_targets()?
            .get(target_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "bookmark target {target_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| crate::remove_bookmark_xml(xml, target_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return all typed form/control custom properties in document order.
    pub fn form_properties(&self) -> Result<Vec<OdfFormProperty>> {
        self.with_content_xml(crate::form_properties)
    }

    /// Insert a property into a form/control owner selected in document order.
    pub fn insert_form_property(
        &mut self,
        owner_index: usize,
        property: &OdfFormProperty,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::insert_form_property_xml(xml, owner_index, property))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a form property selected in document order.
    pub fn replace_form_property(
        &mut self,
        property_index: usize,
        replacement: &OdfFormProperty,
    ) -> Result<OdfFormProperty> {
        let old = self
            .form_properties()?
            .get(property_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "form property {property_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_form_property_xml(xml, property_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a form property and remove its container when it becomes empty.
    pub fn remove_form_property(&mut self, property_index: usize) -> Result<OdfFormProperty> {
        let old = self
            .form_properties()?
            .get(property_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "form property {property_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_form_property_xml(xml, property_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return text and textarea controls in document order.
    pub fn text_controls(&self) -> Result<Vec<OdfTextControl>> {
        self.with_content_xml(crate::text_controls)
    }

    /// Insert a text or textarea control into a form selected in document order.
    pub fn insert_text_control(
        &mut self,
        form_index: usize,
        control: &OdfTextControl,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::insert_text_control_xml(xml, form_index, control))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a text or textarea control selected in document order.
    pub fn replace_text_control(
        &mut self,
        control_index: usize,
        replacement: &OdfTextControl,
    ) -> Result<OdfTextControl> {
        let old = self
            .text_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "text control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_text_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a text or textarea control selected in document order.
    pub fn remove_text_control(&mut self, control_index: usize) -> Result<OdfTextControl> {
        let old = self
            .text_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "text control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_text_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return button and checkbox controls in document order.
    pub fn interactive_controls(&self) -> Result<Vec<OdfInteractiveControl>> {
        self.with_content_xml(crate::interactive_controls)
    }

    /// Insert a button or checkbox into a form selected in document order.
    pub fn insert_interactive_control(
        &mut self,
        form_index: usize,
        control: &OdfInteractiveControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_interactive_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a button or checkbox selected in document order.
    pub fn replace_interactive_control(
        &mut self,
        control_index: usize,
        replacement: &OdfInteractiveControl,
    ) -> Result<OdfInteractiveControl> {
        let old = self
            .interactive_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "interactive control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_interactive_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a button or checkbox selected in document order.
    pub fn remove_interactive_control(
        &mut self,
        control_index: usize,
    ) -> Result<OdfInteractiveControl> {
        let old = self
            .interactive_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "interactive control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_interactive_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return listbox and combobox controls in document order.
    pub fn selection_controls(&self) -> Result<Vec<OdfSelectionControl>> {
        self.with_content_xml(crate::selection_controls)
    }

    /// Insert a listbox or combobox into a form selected in document order.
    pub fn insert_selection_control(
        &mut self,
        form_index: usize,
        control: &OdfSelectionControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_selection_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a listbox or combobox selected in document order.
    pub fn replace_selection_control(
        &mut self,
        control_index: usize,
        replacement: &OdfSelectionControl,
    ) -> Result<OdfSelectionControl> {
        let old = self
            .selection_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "selection control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_selection_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a listbox or combobox selected in document order.
    pub fn remove_selection_control(
        &mut self,
        control_index: usize,
    ) -> Result<OdfSelectionControl> {
        let old = self
            .selection_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "selection control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_selection_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return radio, frame, and image-button controls in document order.
    pub fn visual_controls(&self) -> Result<Vec<crate::OdfVisualControl>> {
        self.with_content_xml(crate::visual_controls)
    }

    /// Insert a radio, frame, or image-button into a form selected in document order.
    pub fn insert_visual_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfVisualControl,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::insert_visual_control_xml(xml, form_index, control))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a radio, frame, or image-button selected in document order.
    pub fn replace_visual_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfVisualControl,
    ) -> Result<crate::OdfVisualControl> {
        let old = self
            .visual_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "visual control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_visual_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a radio, frame, or image-button selected in document order.
    pub fn remove_visual_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfVisualControl> {
        let old = self
            .visual_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "visual control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_visual_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return fixed-text, hidden, and generic controls in document order.
    pub fn generic_form_controls(&self) -> Result<Vec<crate::OdfGenericFormControl>> {
        self.with_content_xml(crate::generic_form_controls)
    }

    /// Insert a fixed-text, hidden, or generic control into a form selected in document order.
    pub fn insert_generic_form_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfGenericFormControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_generic_form_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a fixed-text, hidden, or generic control selected in document order.
    pub fn replace_generic_form_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfGenericFormControl,
    ) -> Result<crate::OdfGenericFormControl> {
        let old = self
            .generic_form_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "generic form control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_generic_form_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a fixed-text, hidden, or generic control selected in document order.
    pub fn remove_generic_form_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfGenericFormControl> {
        let old = self
            .generic_form_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "generic form control {control_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::remove_generic_form_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return password and file controls in document order.
    pub fn password_file_controls(&self) -> Result<Vec<crate::OdfPasswordFileControl>> {
        self.with_content_xml(crate::password_file_controls)
    }

    /// Insert a password or file control into a form selected in document order.
    pub fn insert_password_file_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfPasswordFileControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_password_file_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a password or file control selected in document order.
    pub fn replace_password_file_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfPasswordFileControl,
    ) -> Result<crate::OdfPasswordFileControl> {
        let old = self
            .password_file_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "password/file control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_password_file_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a password or file control selected in document order.
    pub fn remove_password_file_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfPasswordFileControl> {
        let old = self
            .password_file_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "password/file control {control_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::remove_password_file_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return image-frame controls in document order without resolving image references.
    pub fn image_frame_controls(&self) -> Result<Vec<crate::OdfImageFrameControl>> {
        self.with_content_xml(crate::image_frame_controls)
    }

    /// Insert an image-frame control into a form selected in document order.
    pub fn insert_image_frame_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfImageFrameControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_image_frame_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace an image-frame control selected in document order.
    pub fn replace_image_frame_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfImageFrameControl,
    ) -> Result<crate::OdfImageFrameControl> {
        let old = self
            .image_frame_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "image-frame control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_image_frame_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove an image-frame control selected in document order.
    pub fn remove_image_frame_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfImageFrameControl> {
        let old = self
            .image_frame_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "image-frame control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_image_frame_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return value-range controls in document order without resolving bindings.
    pub fn value_range_controls(&self) -> Result<Vec<crate::OdfValueRangeControl>> {
        self.with_content_xml(crate::value_range_controls)
    }

    /// Insert a value-range control into a form selected in document order.
    pub fn insert_value_range_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfValueRangeControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_value_range_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a value-range control selected in document order.
    pub fn replace_value_range_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfValueRangeControl,
    ) -> Result<crate::OdfValueRangeControl> {
        let old = self
            .value_range_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "value-range control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_value_range_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a value-range control selected in document order.
    pub fn remove_value_range_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfValueRangeControl> {
        let old = self
            .value_range_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "value-range control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_value_range_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return formatted-text, number, date, and time controls in document order.
    pub fn typed_value_controls(&self) -> Result<Vec<crate::OdfTypedValueControl>> {
        self.with_content_xml(crate::typed_value_controls)
    }

    /// Insert a typed value control into a form selected in document order.
    pub fn insert_typed_value_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfTypedValueControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_typed_value_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a typed value control selected in document order.
    pub fn replace_typed_value_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfTypedValueControl,
    ) -> Result<crate::OdfTypedValueControl> {
        let old = self
            .typed_value_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "typed value control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_typed_value_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a typed value control selected in document order.
    pub fn remove_typed_value_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::OdfTypedValueControl> {
        let old = self
            .typed_value_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "typed value control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_typed_value_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    pub fn grid_controls(&self) -> Result<Vec<crate::OdfGridControl>> {
        self.with_content_xml(crate::grid_controls)
    }
    pub fn insert_grid_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfGridControl,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::insert_grid_control_xml(xml, form_index, control))?;
        self.content_xml = Some(updated);
        Ok(())
    }
    pub fn replace_grid_control(
        &mut self,
        control_index: usize,
        replacement: &crate::OdfGridControl,
    ) -> Result<crate::OdfGridControl> {
        let old = self
            .grid_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "grid control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::replace_grid_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }
    pub fn remove_grid_control(&mut self, control_index: usize) -> Result<crate::OdfGridControl> {
        let old = self
            .grid_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "grid control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_grid_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Insert a field at the end of a paragraph selected in document order.
    pub fn insert_dynamic_text_field(
        &mut self,
        paragraph_index: usize,
        field: &OdfDynamicTextField,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::odt::insert_dynamic_text_field_xml(xml, paragraph_index, field)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a dynamic field selected in document order and return its old value.
    pub fn replace_dynamic_text_field(
        &mut self,
        field_index: usize,
        replacement: &OdfDynamicTextField,
    ) -> Result<OdfDynamicTextField> {
        let old = self
            .dynamic_text_fields()?
            .get(field_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "dynamic text field index {field_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::odt::replace_dynamic_text_field_xml(xml, field_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a dynamic field selected in document order and return its old value.
    pub fn remove_dynamic_text_field(&mut self, field_index: usize) -> Result<OdfDynamicTextField> {
        let old = self
            .dynamic_text_fields()?
            .get(field_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "dynamic text field index {field_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::odt::remove_dynamic_text_field_xml(xml, field_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return the explicit page-sequence master-page assignments, if authored.
    ///
    /// The model preserves only the ordered `text:master-page-name` values of
    /// the document's `text:page-sequence` (ODF 1.3 §5.3); litchi does not
    /// paginate or resolve the referenced master pages.
    pub fn page_sequence(&self) -> Result<Option<OdtPageSequence>> {
        self.with_content_xml(parse_page_sequence)
    }

    /// Set, replace, or clear the document's `text:page-sequence`.
    ///
    /// A new sequence is written as the first child of `office:text`, matching
    /// the element order of ODF 1.3 §5.1. Passing `None` removes an existing
    /// sequence and is a no-op when none exists. Master-page names are stored
    /// lexically and never resolved against `styles.xml`.
    pub fn set_page_sequence(&mut self, sequence: Option<&OdtPageSequence>) -> Result<()> {
        let updated = self.with_content_xml(|xml| set_page_sequence_xml(xml, sequence))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Return validated variable, user-field, sequence, and DDE declarations.
    pub fn variable_declarations(&self) -> Result<OdfVariableDeclarations> {
        self.with_content_xml(|content| {
            if let Some(styles) = self.styles_xml.as_deref() {
                crate::variable_declaration::parse_variable_declaration_parts(&[
                    (content, OdfVariablePart::Content),
                    (styles, OdfVariablePart::Styles),
                ])
            } else {
                crate::variable_declaration::parse_variable_declaration_parts(&[(
                    content,
                    OdfVariablePart::Content,
                )])
            }
        })
    }

    /// Atomically insert or replace one declaration container.
    ///
    /// Content and styles are reparsed together before commit, preserving all
    /// cross-part declaration and field-reference invariants.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &OdfVariableDeclarationGroup,
    ) -> Result<Option<OdfVariableDeclarationGroup>> {
        let current = self.variable_declarations()?;
        let old = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == group.part
                    && candidate.scope == group.scope
                    && candidate.kind == group.kind
            })
            .cloned();
        match group.part {
            OdfVariablePart::Content => {
                let updated = self.with_content_xml(|xml| {
                    crate::set_variable_declaration_group_xml(xml, group)
                })?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_variable_declaration_parts(&[
                        (&updated, OdfVariablePart::Content),
                        (styles, OdfVariablePart::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_variable_declaration_parts(&[(
                        &updated,
                        OdfVariablePart::Content,
                    )])?;
                }
                self.content_xml = Some(updated);
            },
            OdfVariablePart::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::set_variable_declaration_group_xml(styles, group)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_variable_declaration_parts(&[
                        (content, OdfVariablePart::Content),
                        (&updated, OdfVariablePart::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            OdfVariablePart::Flat => {
                return Err(litchi_core::Error::InvalidFormat(
                    "MutableDocument cannot edit flat-document declarations".to_string(),
                ));
            },
        }
        Ok(old)
    }

    /// Atomically remove one declaration container, returning its old value.
    ///
    /// Removal fails without mutation when any remaining field references a
    /// declaration from the removed container.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: OdfVariablePart,
        scope: &OdfVariableScope,
        kind: OdfVariableKind,
    ) -> Result<Option<OdfVariableDeclarationGroup>> {
        let current = self.variable_declarations()?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == part && &candidate.scope == scope && candidate.kind == kind
            })
            .cloned()
        else {
            return Ok(None);
        };
        match part {
            OdfVariablePart::Content => {
                let updated = self.with_content_xml(|xml| {
                    crate::remove_variable_declaration_group_xml(xml, scope, kind)
                })?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_variable_declaration_parts(&[
                        (&updated, OdfVariablePart::Content),
                        (styles, OdfVariablePart::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_variable_declaration_parts(&[(
                        &updated,
                        OdfVariablePart::Content,
                    )])?;
                }
                self.content_xml = Some(updated);
            },
            OdfVariablePart::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::remove_variable_declaration_group_xml(styles, scope, kind)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_variable_declaration_parts(&[
                        (content, OdfVariablePart::Content),
                        (&updated, OdfVariablePart::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            OdfVariablePart::Flat => {
                return Err(litchi_core::Error::InvalidFormat(
                    "MutableDocument cannot edit flat-document declarations".to_string(),
                ));
            },
        }
        Ok(Some(old))
    }

    fn with_content_xml<T>(&self, operation: impl FnOnce(&str) -> Result<T>) -> Result<T> {
        if let Some(xml) = self.content_xml.as_deref() {
            operation(xml)
        } else {
            let xml = self.generate_content_xml();
            operation(&xml)
        }
    }

    fn invalidate_content_xml(&mut self) {
        self.content_xml = None;
    }

    /// Return declarations, policy, and marker-correlated content from current XML.
    pub fn tracked_changes(&self) -> Result<crate::odt::TrackedChanges> {
        self.with_content_xml(super::parser::OdtParser::parse_tracked_changes)
    }

    /// Atomically replace the declaration table and policy metadata.
    pub fn set_tracked_changes(&mut self, tracked: crate::odt::TrackedChanges) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::odt::set_tracked_changes_xml(xml, Some(&tracked)))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Atomically update tracking policy and inert protection metadata.
    pub fn set_tracked_change_policy(
        &mut self,
        track_changes: Option<bool>,
        protection_key: Option<String>,
        digest_algorithm: Option<String>,
    ) -> Result<()> {
        let mut tracked = self.tracked_changes()?;
        tracked.track_changes = track_changes;
        tracked.protection_key = protection_key;
        tracked.protection_key_digest_algorithm = digest_algorithm;
        self.set_tracked_changes(tracked)
    }

    /// Atomically append a declaration in insertion order.
    pub fn add_tracked_change(&mut self, change: crate::odt::TrackChange) -> Result<()> {
        let mut tracked = self.tracked_changes()?;
        tracked.changes.push(change);
        self.set_tracked_changes(tracked)
    }

    /// Atomically replace a declaration without changing marker identity.
    pub fn update_tracked_change(
        &mut self,
        id: &str,
        replacement: crate::odt::TrackChange,
    ) -> Result<crate::odt::TrackChange> {
        if replacement.id != id {
            return Err(litchi_core::Error::InvalidFormat(
                "tracked-change update cannot change its stable ID".to_string(),
            ));
        }
        let mut tracked = self.tracked_changes()?;
        let index = tracked
            .changes
            .iter()
            .position(|change| change.id == id)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "tracked-change declaration '{id}' was not found"
                ))
            })?;
        let old = std::mem::replace(&mut tracked.changes[index], replacement);
        self.set_tracked_changes(tracked)?;
        Ok(old)
    }

    /// Remove a declaration and all of its correlated markers atomically.
    pub fn remove_tracked_change(&mut self, id: &str) -> Result<crate::odt::TrackChange> {
        let mut tracked = self.tracked_changes()?;
        let index = tracked
            .changes
            .iter()
            .position(|change| change.id == id)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "tracked-change declaration '{id}' was not found"
                ))
            })?;
        let removed = tracked.changes.remove(index);
        let updated = self.with_content_xml(|xml| {
            let unmarked = crate::odt::unmark_tracked_change_xml(xml, id)?;
            crate::odt::set_tracked_changes_xml(&unmarked, Some(&tracked))
        })?;
        self.content_xml = Some(updated);
        Ok(removed)
    }

    /// Remove all declarations, policy, and correlated markers.
    pub fn clear_tracked_changes(&mut self) -> Result<()> {
        let tracked = self.tracked_changes()?;
        let updated = self.with_content_xml(|xml| {
            let mut candidate = xml.to_string();
            for change in &tracked.changes {
                candidate = crate::odt::unmark_tracked_change_xml(&candidate, &change.id)?;
            }
            crate::odt::set_tracked_changes_xml(&candidate, None)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Mark a live insertion or format-change range using Unicode character offsets.
    pub fn mark_tracked_change_range(
        &mut self,
        change_id: &str,
        start: crate::odt::OdtTrackedPosition,
        end: crate::odt::OdtTrackedPosition,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::odt::mark_tracked_change_range_xml(xml, change_id, &start, &end)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Place a point deletion marker using a Unicode character offset.
    pub fn mark_tracked_deletion(
        &mut self,
        change_id: &str,
        position: crate::odt::OdtTrackedPosition,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::odt::mark_tracked_deletion_xml(xml, change_id, &position)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove every marker for one declaration while retaining its live text.
    pub fn unmark_tracked_change(&mut self, change_id: &str) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::odt::unmark_tracked_change_xml(xml, change_id))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Return typed nested sections from current authoritative content XML.
    pub fn sections(&self) -> Result<Vec<crate::Section>> {
        self.with_content_xml(super::parser::OdtParser::parse_sections)
    }

    /// Append a complete typed section without rewriting existing body content.
    pub fn add_section(&mut self, section: &crate::Section) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::add_section_xml(xml, section))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Atomically update section metadata/source while preserving enclosed XML bytes.
    pub fn update_section(&mut self, name: &str, replacement: &crate::Section) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::update_section_xml(xml, name, replacement))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Delete a named section together with its enclosed content.
    pub fn remove_section(&mut self, name: &str) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::remove_section_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove one section wrapper/source while retaining all enclosed content bytes.
    pub fn unwrap_section(&mut self, name: &str) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::unwrap_section_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove every section wrapper/source while retaining mixed nested content.
    pub fn clear_sections(&mut self) -> Result<()> {
        let updated = self.with_content_xml(crate::clear_sections_xml)?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Wrap an inclusive stable block range in a typed section.
    pub fn wrap_section(
        &mut self,
        section: &crate::Section,
        start: crate::OdtSectionBlock,
        end: crate::OdtSectionBlock,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::wrap_section_xml(xml, section, &start, &end))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Parse the document's master pages and current header/footer regions.
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), parse_master_pages)
    }

    /// Parse automatic page layouts, their properties, and header/footer styles.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), parse_page_layouts)
    }

    /// Replace one page layout with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:page-layout`
    /// element whose `style:name` matches `page_layout_name`. This supports all
    /// page properties and nested header/footer styles while preserving every
    /// unrelated byte in `styles.xml`.
    pub fn set_page_layout_xml(
        &mut self,
        page_layout_name: &str,
        page_layout_xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_page_layout_xml(
            styles,
            page_layout_name,
            page_layout_xml,
        )?);
        Ok(())
    }

    /// Create or replace typed header/footer properties in one page layout.
    pub fn set_page_layout_header_footer_properties(
        &mut self,
        page_layout_name: &str,
        region: crate::PageHeaderFooterRegion,
        properties: &crate::HeaderFooterStyleProperties,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement = crate::header_footer_properties::replace_page_layout_region_properties(
            layout, region, properties,
        )?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace typed columns in one existing page layout.
    pub fn set_page_layout_columns(
        &mut self,
        page_layout_name: &str,
        columns: &crate::StyleColumns,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement = crate::style_columns::replace_page_layout_columns(layout, columns)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace the typed footnote separator in one existing page layout.
    pub fn set_page_layout_footnote_separator(
        &mut self,
        page_layout_name: &str,
        separator: &crate::StyleFootnoteSeparator,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement =
            crate::footnote_separator::replace_page_layout_footnote_separator(layout, separator)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace one existing named list level's modern label alignment.
    pub fn set_list_level_label_alignment(
        &mut self,
        item: &crate::ListStyleLevelLabelAlignment,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml list style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(
            crate::list_label_alignment::replace_list_level_label_alignment_xml(styles, item)?,
        );
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace, insert, or remove one existing paragraph style's direct drop cap.
    pub fn set_paragraph_style_drop_cap(
        &mut self,
        style: &crate::ParagraphStyleDropCap,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml paragraph style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::paragraph_drop_cap::set_paragraph_style_drop_cap_xml(
            styles, style,
        )?);
        Ok(())
    }

    /// Replace, insert, or remove typed row properties on an existing table-row style.
    pub fn set_table_row_style_properties(
        &mut self,
        style: &crate::TableRowStyleProperties,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table-row style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::set_table_row_style_properties_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed properties on an existing table style.
    pub fn set_table_style_properties(
        &mut self,
        style: &crate::TableStyleProperties,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::set_table_style_properties_xml(styles, style)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// A minimal page layout is created in `office:automatic-styles` when a
    /// layout with `page_layout_name` does not already exist.
    pub fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        let styles = self
            .styles_xml
            .get_or_insert_with(OdfStructure::default_styles_xml);
        *styles = add_master_page(styles, name, page_layout_name)?;
        Ok(())
    }

    /// Insert a complete typed master page without rewriting unrelated styles.
    pub fn insert_master_page(&mut self, page: &MasterPage) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self
            .styles_xml
            .get_or_insert_with(OdfStructure::default_styles_xml);
        *styles = crate::insert_master_page_xml(styles, &fragment)?;
        Ok(())
    }

    /// Replace one named master page without rewriting unrelated styles.
    pub fn replace_master_page(&mut self, name: &str, page: &MasterPage) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::replace_master_page_xml(styles, name, &fragment)?);
        Ok(())
    }

    /// Remove one named master page without rewriting unrelated styles.
    pub fn remove_master_page(&mut self, name: &str) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::remove_master_page_xml(styles, name)?);
        Ok(())
    }

    /// Set plain text in one header/footer region of an existing master page.
    ///
    /// Only the selected region is rewritten; all unrelated style XML is preserved.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
        text: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_text(styles, master_page_name, kind, Some(text))?);
        Ok(())
    }

    /// Replace one header/footer region with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:header`,
    /// `style:footer`, or corresponding first/left variant matching `kind`.
    /// This preserves rich text, fields, tables, drawings, and extension content.
    pub fn set_header_footer_xml(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
        xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_xml(styles, master_page_name, kind, xml)?);
        Ok(())
    }

    /// Remove one header/footer region from an existing master page.
    pub fn clear_header_footer(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_text(styles, master_page_name, kind, None)?);
        Ok(())
    }

    /// Get all paragraphs in the document.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all paragraphs (owned) in the document.
    pub fn paragraphs_owned(&self) -> Vec<Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all headings in the document.
    pub fn headings(&self) -> Vec<&Heading> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Heading(heading) => Some(heading),
                _ => None,
            })
            .collect()
    }

    /// Get all lists in the document.
    pub fn lists(&self) -> Vec<&List> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::List(list) => Some(list),
                _ => None,
            })
            .collect()
    }

    /// Get all tables in the document.
    pub fn tables(&self) -> Vec<&Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all tables (owned) in the document.
    pub fn tables_owned(&self) -> Vec<Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get the document metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the document metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a new paragraph to the end of the document.
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello, World!")?;
    /// doc.add_paragraph("Second paragraph")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph(&mut self, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Paragraph(para));
        Ok(())
    }

    /// Append a paragraph containing one simple ODF hyperlink.
    ///
    /// The target is inert metadata and is never followed by the library.
    pub fn add_hyperlink(&mut self, href: impl AsRef<str>, text: impl AsRef<str>) -> Result<()> {
        let hyperlink = Hyperlink::with_href(href, text)?;
        self.add_hyperlink_element(hyperlink)
    }

    /// Append a paragraph containing a fully configured ODF hyperlink.
    pub fn add_hyperlink_element(&mut self, hyperlink: Hyperlink) -> Result<()> {
        let mut paragraph = Paragraph::new();
        paragraph.add_hyperlink(hyperlink)?;
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(())
    }

    /// Add a heading to the end of the document.
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        if !(1..=6).contains(&level) {
            return Err(litchi_core::Error::InvalidFormat(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }

        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Heading(heading));
        Ok(())
    }

    /// Add an existing list element to the end of the document.
    pub fn add_list(&mut self, list: List) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::List(list));
        Ok(())
    }

    /// Insert a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("First")?;
    /// doc.add_paragraph("Third")?;
    /// doc.insert_paragraph(1, "Second")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);

        if index <= self.elements.len() {
            self.invalidate_content_xml();
            self.elements
                .insert(index, DocumentElement::Paragraph(para));
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )))
        }
    }

    /// Insert an image frame as a new paragraph at a specific index.
    ///
    /// The payload is sniffed (PNG, JPEG, and GIF are accepted), stored
    /// verbatim under `Pictures/` in the package, and referenced from a
    /// `draw:frame`/`draw:image` element with the given geometry and anchor.
    /// Returns the allocated package path of the picture part.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{MutableDocument, OdfFrameAnchor, OdfLength};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// let png = b"\x89PNG\r\n\x1a\n".as_slice();
    /// let path = doc.insert_image(
    ///     0,
    ///     png,
    ///     &OdfLength::centimeters(10.0),
    ///     &OdfLength::centimeters(4.0),
    ///     OdfFrameAnchor::AsChar,
    /// )?;
    /// assert!(path.starts_with("Pictures/"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_image(
        &mut self,
        index: usize,
        image: &[u8],
        width: &crate::odt::OdfLength,
        height: &crate::odt::OdfLength,
        anchor: crate::odt::OdfFrameAnchor,
    ) -> Result<String> {
        use crate::odt::frame;
        let format = frame::validate_image_payload(image)?;
        let path = frame::allocate_picture_path(format.extension(), |candidate| {
            // Picture numbering is global: a stem taken by any supported
            // extension blocks the whole index.
            let taken = |path: &str| {
                self.pending_images
                    .iter()
                    .any(|pending| pending.path == path)
                    || self
                        .source_package
                        .as_ref()
                        .is_some_and(|package| package.has_file(path).unwrap_or(false))
            };
            if taken(candidate) {
                return true;
            }
            let stem = candidate.trim_end_matches(format.extension());
            ["png", "jpg", "gif"]
                .iter()
                .any(|extension| taken(&format!("{stem}{extension}")))
        })?;
        let name = format!("Frame {}", self.next_frame_number);
        let frame_element = frame::image_frame_element(&name, width, height, anchor, &path)?;
        let mut paragraph_element = crate::elements::element::Element::new("text:p");
        paragraph_element.add_child(frame_element);
        let paragraph = Paragraph::from_element(paragraph_element)?;

        if index > self.elements.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )));
        }
        self.invalidate_content_xml();
        self.elements
            .insert(index, DocumentElement::Paragraph(paragraph));
        self.pending_images.push(frame::PendingImage {
            path: path.clone(),
            bytes: image.to_vec(),
        });
        self.next_frame_number += 1;
        Ok(path)
    }

    /// Insert a plain-text text-box frame as a new paragraph at a specific index.
    ///
    /// The box is a `draw:frame` wrapping `draw:text-box`; newlines in `text`
    /// become separate paragraphs in the box story. Returns the frame name.
    pub fn insert_text_box(
        &mut self,
        index: usize,
        text: &str,
        width: &crate::odt::OdfLength,
        height: &crate::odt::OdfLength,
        anchor: crate::odt::OdfFrameAnchor,
    ) -> Result<String> {
        use crate::odt::frame;
        let name = format!("Text Box {}", self.next_frame_number);
        let frame_element = frame::text_box_frame_element(&name, width, height, anchor, text)?;

        if index > self.elements.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )));
        }
        self.invalidate_content_xml();
        self.elements
            .insert(index, DocumentElement::Frame(frame_element));
        self.next_frame_number += 1;
        Ok(name)
    }

    /// Remove a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.remove_paragraph(0)?; // Remove first paragraph
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Paragraph(para) = self.elements.remove(idx) {
                Ok(para)
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Update a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to update (0-based)
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Old text")?;
    /// doc.update_paragraph(0, "New text")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Paragraph(ref mut para) = self.elements[idx] {
                para.set_text(text);
                Ok(())
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Clear all paragraphs from the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.clear_paragraphs();
    /// assert_eq!(doc.paragraphs().len(), 0);
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_paragraphs(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Paragraph(_)));
    }

    /// Add a table to the document.
    ///
    /// # Arguments
    ///
    /// * `table` - Table to add
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// let mut table = Table::new();
    /// table.set_name("Table1");
    /// doc.add_table(table)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Table(table));
        Ok(())
    }

    /// Remove a table at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the table to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_table(Table::new())?;
    /// doc.remove_table(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_table(&mut self, index: usize) -> Result<Table> {
        // Find the index of the nth table
        let mut table_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Table(_)) {
                if table_count == index {
                    element_index = Some(i);
                    break;
                }
                table_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Table(table) = self.elements.remove(idx) {
                Ok(table)
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Table index {} out of bounds (found {} tables)",
                index, table_count
            )))
        }
    }

    /// Clear all tables from the document.
    pub fn clear_tables(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Table(_)));
    }

    /// Clear all content (paragraphs and tables) from the document.
    pub fn clear_content(&mut self) {
        self.invalidate_content_xml();
        self.elements.clear();
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> String {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Heading(h) => h.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
                DocumentElement::Frame(_) => 256,
            })
            .sum::<usize>();
        let mut body = String::with_capacity(estimated);

        // Add elements in their insertion order (paragraphs and tables mixed)
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Frame(frame) => {
                    body.push_str(&frame.to_xml_string());
                },
            }
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> Result<String> {
        if let Some(patched) = self.patched_source_meta_xml()? {
            return Ok(patched);
        }
        Ok(self.generate_meta_xml_from_scratch())
    }

    /// Patch the retained source meta.xml so metadata the edit did not change
    /// survives the save, while fields set through the mutable API, the
    /// generator, and the modification date are updated in place.
    fn patched_source_meta_xml(&self) -> Result<Option<String>> {
        let Some(package) = &self.source_package else {
            return Ok(None);
        };
        let Ok(bytes) = package.get_file("meta.xml") else {
            return Ok(None);
        };
        let Ok(source) = String::from_utf8(bytes) else {
            return Ok(None);
        };
        let source_metadata = OdfMetadata::from_xml(&source)?;
        let patch = MetaXmlPatch::preserve_all()
            .with_generator_and_modification_date("Litchi/0.0.1", chrono::Utc::now().to_rfc3339())
            .diff_simple_fields(&source_metadata, &self.metadata);
        patch_meta_xml(&source, &patch)
    }

    /// Generate meta.xml from the mutable metadata model alone.
    fn generate_meta_xml_from_scratch(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.author.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.subject.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self
            .metadata
            .description
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        estimated += self
            .metadata
            .keywords
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        if let Some(ref subject) = self.metadata.subject {
            let escaped_subject = escape_xml(subject);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:subject>{}</dc:subject>"#,
                escaped_subject
            ));
        }

        if let Some(ref description) = self.metadata.description {
            let escaped_description = escape_xml(description);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:description>{}</dc:description>"#,
                escaped_description
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            let escaped_keywords = escape_xml(keywords);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:keyword>{}</meta:keyword>"#,
                escaped_keywords
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date>{}</office:meta></office:document-meta>"#,
            now,
            meta_fields
        )
    }

    /// Save the modified document to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// let bytes = doc.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let generated_content_xml;
        let content_xml = if let Some(content_xml) = self.content_xml.as_deref() {
            content_xml
        } else {
            generated_content_xml = self.generate_content_xml();
            &generated_content_xml
        };
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (patched from the source or regenerated with current metadata)
        let meta_xml = self.generate_meta_xml()?;
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Add authored picture payloads.
        for pending in &self.pending_images {
            writer.add_file(&pending.path, &pending.bytes)?;
        }

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::elements::parser::DocumentOrderElement;
    use crate::elements::table::{TableCell, TableRow};
    use crate::elements::text::{ListItem, Paragraph};
    use crate::odt::{DocumentBuilder, PageUsage};

    const MINIMAL_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Original</text:p></office:text></office:body></office:document-content>"#;

    fn source_document() -> Document {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Before table").unwrap();

        let mut table = Table::new();
        table.set_name("Data");
        let mut row = TableRow::new();
        let mut cell = TableCell::new();
        cell.set_text("Cell content");
        row.add_cell(cell);
        table.add_row(row);
        builder.add_table(table).unwrap();

        builder.add_heading("After table", 2).unwrap();
        builder
            .add_bulleted_list(vec!["First item", "Second item"])
            .unwrap();
        builder.add_paragraph("Document end").unwrap();

        Document::from_bytes(builder.build().unwrap()).unwrap()
    }

    fn rich_note_body() -> crate::OdfNoteBodyContent {
        const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        crate::OdfNoteBodyContent::new(vec![crate::OdfMetaFieldNode::Element(
            crate::OdfMetaFieldElement {
                namespace_uri: TEXT.to_string(),
                local_name: "p".to_string(),
                attributes: Vec::new(),
                children: vec![
                    crate::OdfMetaFieldNode::Text("Styled ".to_string()),
                    crate::OdfMetaFieldNode::Element(crate::OdfMetaFieldElement {
                        namespace_uri: TEXT.to_string(),
                        local_name: "span".to_string(),
                        attributes: vec![crate::OdfMetaFieldAttribute {
                            namespace_uri: TEXT.to_string(),
                            local_name: "style-name".to_string(),
                            value: "Emphasis".to_string(),
                        }],
                        children: vec![crate::OdfMetaFieldNode::Text("body".to_string())],
                    }),
                ],
            },
        )])
        .unwrap()
    }

    fn element_kinds(document: &Document) -> Vec<&'static str> {
        document
            .elements()
            .unwrap()
            .iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(_) => "paragraph",
                DocumentOrderElement::NumberedParagraph(_) => "numbered-paragraph",
                DocumentOrderElement::Heading(_) => "heading",
                DocumentOrderElement::Table(_) => "table",
                DocumentOrderElement::List(_) => "list",
            })
            .collect()
    }

    #[test]
    fn conversion_preserves_top_level_order_without_nested_paragraph_duplicates() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();

        assert_eq!(mutable.paragraphs().len(), 2);
        assert_eq!(mutable.tables().len(), 1);
        assert_eq!(mutable.headings().len(), 1);
        assert_eq!(mutable.lists().len(), 1);
        assert_eq!(mutable.headings()[0].text().unwrap(), "After table");
        assert_eq!(mutable.headings()[0].level(), Some(2));
        assert_eq!(mutable.lists()[0].items().unwrap().len(), 2);
    }

    #[test]
    fn mutable_hyperlink_authoring_round_trips_through_an_odt_package() {
        let mut mutable = MutableDocument::new();
        mutable
            .add_hyperlink("https://example.test/", "External")
            .unwrap();
        let mut internal = crate::Hyperlink::with_href("#bookmark", "Internal").unwrap();
        internal.set_name("bookmark-link");
        mutable.add_hyperlink_element(internal).unwrap();

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.hyperlinks().unwrap(),
            vec![
                ("External".to_string(), "https://example.test/".to_string()),
                ("Internal".to_string(), "#bookmark".to_string()),
            ]
        );
        assert_eq!(document.text().unwrap(), "External\nInternal");
    }

    #[test]
    fn mutable_ruby_annotation_and_style_crud_round_trip_through_an_odt_package() {
        let first_style = crate::RubyStyle::new(
            "RubyAbove",
            Some(crate::RubyProperties {
                position: Some(crate::RubyPosition::Above),
                alignment: Some(crate::RubyAlignment::Center),
            }),
        )
        .unwrap();
        let second_style = crate::RubyStyle::new(
            "RubyAbove",
            Some(crate::RubyProperties {
                position: Some(crate::RubyPosition::Below),
                alignment: Some(crate::RubyAlignment::DistributeLetter),
            }),
        )
        .unwrap();
        let first = crate::RubyAnnotation::new(
            Some(first_style.name.clone()),
            crate::RubyBase::from_text("語").unwrap(),
            "ご",
            None,
        )
        .unwrap();
        let replacement = crate::RubyAnnotation::new(
            Some(first_style.name.clone()),
            crate::RubyBase::from_text("文").unwrap(),
            "ぶん",
            None,
        )
        .unwrap();

        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("Read ").unwrap();
        assert_eq!(mutable.set_ruby_style(&first_style).unwrap(), None);
        assert_eq!(
            mutable.set_ruby_style(&second_style).unwrap(),
            Some(first_style)
        );
        mutable.insert_ruby_annotation(0, &first).unwrap();
        assert_eq!(
            mutable.replace_ruby_annotation(0, &replacement).unwrap(),
            first
        );

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(document.ruby_styles().unwrap().styles, vec![second_style]);
        assert_eq!(
            document.ruby_annotations().unwrap().annotations,
            vec![replacement.clone()]
        );

        assert_eq!(mutable.remove_ruby_annotation(0).unwrap(), replacement);
        assert!(mutable.ruby_annotations().unwrap().annotations.is_empty());
        assert!(mutable.remove_ruby_style("RubyAbove").unwrap().is_some());
        assert!(mutable.ruby_styles().unwrap().styles.is_empty());
    }

    #[test]
    fn mutable_ruby_range_wrapping_round_trips_through_an_odt_package() {
        let annotation =
            crate::RubyAnnotation::new(None, crate::RubyBase::from_text("字").unwrap(), "じ", None)
                .unwrap();
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("Read 漢字").unwrap();
        let start = "Read 漢".len();
        mutable
            .wrap_ruby_annotation(0, start..start + "字".len(), &annotation)
            .unwrap();

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.ruby_annotations().unwrap().annotations,
            vec![annotation]
        );
    }

    #[test]
    fn mutable_line_numbering_configuration_round_trips_without_generation() {
        let first = crate::OdfLineNumberingConfiguration {
            number_lines: Some(true),
            number_format: Some(crate::OdfLineNumberFormat::LowerAlpha),
            letter_sync: Some(true),
            style_name: Some("LineNumbers".to_string()),
            increment: Some(2),
            number_position: Some(crate::OdfLineNumberPosition::Inner),
            offset: Some(crate::OdfNonNegativeLength::new("0.2in").unwrap()),
            count_empty_lines: Some(false),
            count_in_text_boxes: Some(true),
            restart_on_page: Some(false),
            separator: Some(crate::OdfLineNumberingSeparator {
                increment: Some(4),
                text: " · ".to_string(),
            }),
        };
        let replacement = crate::OdfLineNumberingConfiguration {
            number_lines: Some(false),
            number_format: Some(crate::OdfLineNumberFormat::UpperRoman),
            increment: Some(1),
            ..crate::OdfLineNumberingConfiguration::default()
        };

        let mut mutable = MutableDocument::new();
        assert_eq!(mutable.line_numbering_configuration().unwrap(), None);
        assert_eq!(
            mutable.set_line_numbering_configuration(&first).unwrap(),
            None
        );
        assert_eq!(
            mutable.line_numbering_configuration().unwrap(),
            Some(first.clone())
        );
        assert_eq!(
            mutable
                .set_line_numbering_configuration(&replacement)
                .unwrap(),
            Some(first)
        );

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.line_numbering_configuration().unwrap(),
            Some(replacement.clone())
        );
        assert_eq!(
            mutable.clear_line_numbering_configuration().unwrap(),
            Some(replacement)
        );
        assert_eq!(mutable.line_numbering_configuration().unwrap(), None);
        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(document.line_numbering_configuration().unwrap(), None);
    }

    #[test]
    fn read_modify_write_keeps_paragraph_table_heading_and_list_order() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();
        let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();

        assert_eq!(
            element_kinds(&round_trip),
            ["paragraph", "table", "heading", "list", "paragraph"]
        );

        let elements = round_trip.elements().unwrap();
        let DocumentOrderElement::Table(table) = &elements[1] else {
            panic!("second element should remain a table");
        };
        assert_eq!(table.name(), Some("Data"));
        assert_eq!(
            table
                .row(0)
                .unwrap()
                .unwrap()
                .cell(0)
                .unwrap()
                .unwrap()
                .text()
                .unwrap(),
            "Cell content"
        );

        let DocumentOrderElement::List(list) = &elements[3] else {
            panic!("fourth element should remain a list");
        };
        let items = list.items().unwrap();
        assert_eq!(items[0].text().unwrap(), "First item");
        assert_eq!(items[1].text().unwrap(), "Second item");
    }

    #[test]
    fn mutable_document_can_add_headings_and_lists() {
        let mut document = MutableDocument::new();
        assert!(document.add_heading("Invalid", 0).is_err());
        document.add_heading("Title", 1).unwrap();

        let mut list = List::new();
        let mut item = ListItem::new();
        let mut paragraph = Paragraph::new();
        paragraph.set_text("Item");
        item.add_paragraph(paragraph);
        list.add_item(item);
        document.add_list(list).unwrap();

        let round_trip = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
        assert_eq!(element_kinds(&round_trip), ["heading", "list"]);
    }

    #[test]
    fn read_modify_write_preserves_auxiliary_package_parts_and_media_types() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
            .unwrap();
        writer
            .add_file("settings.xml", b"document settings")
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/photo.bin", b"image", "image/x-test")
            .unwrap();
        writer
            .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("Object 1/content.xml", b"embedded object")
            .unwrap();
        writer
            .add_file_with_media_type(
                "custom/data.bin",
                b"custom payload",
                "application/x-litchi-test",
            )
            .unwrap();
        writer
            .add_file("META-INF/documentsignatures.xml", b"stale signature")
            .unwrap();

        let source_bytes = writer.finish_to_bytes().unwrap();
        let source = Document::from_bytes(source_bytes.clone()).unwrap();
        assert_eq!(source.to_bytes().unwrap(), source_bytes);
        let mut mutable = MutableDocument::from_document(source).unwrap();
        mutable.add_paragraph("Modified").unwrap();
        let output = crate::core::OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();

        assert_eq!(
            output.get_file("settings.xml").unwrap(),
            b"document settings"
        );
        assert_eq!(output.get_file("Pictures/photo.bin").unwrap(), b"image");
        assert_eq!(
            output.get_file("Object 1/content.xml").unwrap(),
            b"embedded object"
        );
        assert_eq!(
            output.get_file("custom/data.bin").unwrap(),
            b"custom payload"
        );
        assert!(!output.has_file("META-INF/documentsignatures.xml").unwrap());

        let package = output.package().unwrap();
        assert_eq!(
            package.manifest().get_media_type("Pictures/photo.bin"),
            Some("image/x-test")
        );
        assert_eq!(
            package.manifest().get_media_type("Object 1/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(
            package.manifest().get_media_type("custom/data.bin"),
            Some("application/x-litchi-test")
        );
    }

    #[test]
    fn mutable_note_crud_round_trips_through_an_odt_package() {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("Before").unwrap();
        let first = crate::Note::new(crate::NoteClass::Footnote, "1", "Initial").unwrap();
        mutable.insert_note(0, &first).unwrap();
        assert_eq!(mutable.footnotes().unwrap(), vec![first.clone()]);
        assert!(mutable.endnotes().unwrap().is_empty());

        let replacement = crate::Note::new(crate::NoteClass::Endnote, "i", "Replacement").unwrap();
        assert_eq!(mutable.replace_note(0, &replacement).unwrap(), first);
        assert_eq!(mutable.endnotes().unwrap(), vec![replacement.clone()]);
        assert_eq!(mutable.remove_note(0).unwrap(), replacement);
        assert!(mutable.notes().unwrap().is_empty());

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert!(document.notes().unwrap().is_empty());
    }

    #[test]
    fn mutable_document_round_trips_structured_note_authoring() {
        let mut mutable = MutableDocument::new();
        mutable.add_paragraph("Before").unwrap();
        let note =
            crate::Note::with_rich_body(crate::NoteClass::Footnote, "1", rich_note_body()).unwrap();
        mutable.insert_note(0, &note).unwrap();

        let document = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let notes = document.notes().unwrap();
        assert_eq!(notes, vec![note]);
    }

    #[test]
    fn edits_master_page_regions_through_the_public_mutable_document_api() {
        const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:styles><style:style style:name="preserved"/></office:styles><office:automatic-styles><style:page-layout style:name="pm1" style:page-usage="left"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Old header</text:p></style:header><style:footer><text:p>Old footer</text:p></style:footer></style:master-page></office:master-styles></office:document-styles>"#;

        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
            .unwrap();
        writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
        let source = Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(source).unwrap();
        let layouts = mutable.page_layouts().unwrap();
        assert_eq!(layouts[0].page_usage, PageUsage::Left);
        assert_eq!(
            layouts[0].properties.as_ref().unwrap().attribute(
                Some("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
                "page-width",
            ),
            Some("21cm")
        );

        mutable
            .set_header_footer_text("Standard", HeaderFooterKind::Header, "New & <header>")
            .unwrap();
        mutable
            .clear_header_footer("Standard", HeaderFooterKind::Footer)
            .unwrap();
        let pages = mutable.master_pages().unwrap();
        assert_eq!(
            pages[0].region(HeaderFooterKind::Header).unwrap().text,
            "New & <header>"
        );
        assert!(pages[0].region(HeaderFooterKind::Footer).is_none());

        let output = OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let styles = String::from_utf8(output.get_file("styles.xml").unwrap()).unwrap();
        assert!(styles.contains("<style:style style:name=\"preserved\"/>"));
        assert!(!styles.contains("Old footer"));
        let round_trip = Document::from_bytes(output.as_bytes().to_vec()).unwrap();
        assert_eq!(round_trip.page_layouts().unwrap(), layouts);
        assert_eq!(
            round_trip.master_pages().unwrap()[0]
                .region(HeaderFooterKind::Header)
                .unwrap()
                .text,
            "New & <header>"
        );
    }

    #[test]
    fn creates_a_master_page_and_header_in_a_new_document() {
        let mut mutable = MutableDocument::new();
        mutable.add_master_page("Standard", "pm1").unwrap();
        let layout = r#"<s:page-layout xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" s:name="pm1" s:page-usage="mirrored"><s:page-layout-properties f:page-width="21cm" f:page-height="29.7cm"/></s:page-layout>"#;
        mutable.set_page_layout_xml("pm1", layout).unwrap();
        mutable
            .set_header_footer_text("Standard", HeaderFooterKind::Header, "Created header")
            .unwrap();
        let rich = r#"<s:header xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Page <t:page-number/></t:p></s:header>"#;
        mutable
            .set_header_footer_xml("Standard", HeaderFooterKind::Header, rich)
            .unwrap();

        let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let pages = round_trip.master_pages().unwrap();
        let layouts = round_trip.page_layouts().unwrap();
        assert_eq!(layouts.len(), 1);
        assert_eq!(layouts[0].xml, layout);
        assert_eq!(layouts[0].page_usage, PageUsage::Mirrored);
        assert_eq!(pages.len(), 1);
        assert_eq!(pages[0].name, "Standard");
        assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm1"));
        assert_eq!(
            pages[0].region(HeaderFooterKind::Header).unwrap().text,
            "Page "
        );
        assert_eq!(pages[0].region(HeaderFooterKind::Header).unwrap().xml, rich);
        let styles = String::from_utf8(round_trip.get_file("styles.xml").unwrap()).unwrap();
        assert!(styles.contains(layout));
    }

    #[test]
    fn insert_image_round_trips_through_package_and_read_api() {
        use crate::odt::{OdfFrameAnchor, OdfLength};

        let mut doc = MutableDocument::new();
        doc.add_paragraph("Before image").unwrap();
        let png = minimal_png();
        let path = doc
            .insert_image(
                1,
                &png,
                &OdfLength::centimeters(10.0),
                &OdfLength::centimeters(4.0),
                OdfFrameAnchor::AsChar,
            )
            .unwrap();
        assert_eq!(path, "Pictures/image1.png");
        doc.add_paragraph("After image").unwrap();
        assert!(
            doc.insert_image(
                99,
                &png,
                &OdfLength::points(1.0),
                &OdfLength::points(1.0),
                OdfFrameAnchor::Page
            )
            .is_err()
        );
        assert!(
            doc.insert_image(
                0,
                b"not-an-image",
                &OdfLength::points(1.0),
                &OdfLength::points(1.0),
                OdfFrameAnchor::Page
            )
            .is_err()
        );

        let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
        // Text content survives around the frame.
        let text = round_trip.text().unwrap();
        assert!(text.contains("Before image"));
        assert!(text.contains("After image"));
        // The frame is discoverable with identity, geometry, and anchor.
        let images = round_trip.images().unwrap();
        assert_eq!(images.len(), 1);
        let frame = images[0].frame.as_ref().unwrap();
        assert_eq!(frame.width.as_deref(), Some("10cm"));
        assert_eq!(frame.height.as_deref(), Some("4cm"));
        assert_eq!(frame.anchor_type.as_deref(), Some("as-char"));
        assert_eq!(images[0].package_path(), Some("Pictures/image1.png"));
        // Payload is stored verbatim.
        assert_eq!(round_trip.image_bytes(&images[0]).unwrap(), Some(png));
    }

    #[test]
    fn insert_image_coexists_with_existing_media() {
        use crate::odt::{OdfFrameAnchor, OdfLength};

        let mut doc = MutableDocument::new();
        let first = doc
            .insert_image(
                0,
                &minimal_png(),
                &OdfLength::points(8.0),
                &OdfLength::points(8.0),
                OdfFrameAnchor::Page,
            )
            .unwrap();
        let second = doc
            .insert_image(
                1,
                &minimal_jpeg(),
                &OdfLength::points(8.0),
                &OdfLength::points(8.0),
                OdfFrameAnchor::Page,
            )
            .unwrap();
        assert_eq!(first, "Pictures/image1.png");
        assert_eq!(second, "Pictures/image2.jpg");

        let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
        let images = round_trip.images().unwrap();
        assert_eq!(images.len(), 2);
        let mut paths: Vec<_> = images
            .iter()
            .filter_map(|image| image.package_path())
            .collect();
        paths.sort_unstable();
        assert_eq!(paths, ["Pictures/image1.png", "Pictures/image2.jpg"]);
    }

    #[test]
    fn insert_text_box_round_trips_story_text() {
        use crate::odt::{OdfFrameAnchor, OdfLength};

        let mut doc = MutableDocument::new();
        doc.add_paragraph("Intro").unwrap();
        let name = doc
            .insert_text_box(
                1,
                "boxed <text> & more\nsecond line",
                &OdfLength::inches(2.0),
                &OdfLength::inches(1.0),
                OdfFrameAnchor::Paragraph,
            )
            .unwrap();
        assert_eq!(name, "Text Box 1");

        let round_trip = Document::from_bytes(doc.to_bytes().unwrap()).unwrap();
        let content = String::from_utf8(round_trip.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains("draw:text-box"));
        assert!(content.contains("boxed &lt;text&gt; &amp; more"));
        assert!(content.contains("second line"));
        assert!(content.contains("text:anchor-type=\"paragraph\""));
        assert!(round_trip.text().unwrap().contains("Intro"));
    }

    fn minimal_png() -> Vec<u8> {
        // 1x1 transparent PNG.
        let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
        bytes.extend_from_slice(&[0, 0, 0, 13]);
        bytes.extend_from_slice(b"IHDR");
        bytes.extend_from_slice(&[0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0]);
        bytes.extend_from_slice(&[0x1f, 0x15, 0xc4, 0x89]);
        bytes.extend_from_slice(&[0, 0, 0, 11]);
        bytes.extend_from_slice(b"IDAT");
        bytes.extend_from_slice(&[
            0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d,
        ]);
        bytes.extend_from_slice(&[0x0a, 0x2d, 0xb4]);
        bytes.extend_from_slice(&[0, 0, 0, 0]);
        bytes.extend_from_slice(b"IEND");
        bytes.extend_from_slice(&[0xae, 0x42, 0x60, 0x82]);
        bytes
    }

    fn minimal_jpeg() -> Vec<u8> {
        let mut bytes = b"\xff\xd8\xff\xe0".to_vec();
        bytes.extend_from_slice(&[0, 16]);
        bytes.extend_from_slice(b"JFIF\0");
        bytes.extend_from_slice(&[1, 1, 0, 0, 1, 0, 1, 0, 0, 0xff, 0xd9]);
        bytes
    }
}
