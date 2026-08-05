//! Content and styles XML snapshots and lossless mutable edits.

use super::model::MutableDocument;
use crate::BookmarkTarget;
use crate::ReferenceMark;
use crate::TextIndex;
use crate::TextIndexMark;
use crate::core::Structure;
use crate::elements::field::{DynamicTextField, FieldParser};
use crate::header_footer::{Master, read, set_text, set_xml};
use crate::master_page::{add, insert, remove, replace};
use crate::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};
use crate::page_sequence::{Sequence, parse_page_sequence, set_page_sequence_xml};
use crate::variable_declaration::{Declarations, Group, Kind, Part, Scope};
use litchi_core::Result;
use std::ops::Range;

impl MutableDocument {
    /// Return typed dynamic fields from the current authoritative content XML.
    pub fn dynamic_text_fields(&self) -> Result<Vec<DynamicTextField>> {
        self.with_content_xml(FieldParser::parse_dynamic_text_fields)
    }

    /// Return semantic footnotes and endnotes from the current content XML.
    pub fn notes(&self) -> Result<Vec<crate::Note>> {
        self.with_content_xml(crate::parse_notes)
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
    /// content when the replacement carries an `NoteBodyContent`.
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
    pub fn ruby_annotations(&self) -> Result<crate::ruby_family::Annotations> {
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
        annotation: &crate::ruby_family::Annotation,
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
        annotation: &crate::ruby_family::Annotation,
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
        replacement: &crate::ruby_family::Annotation,
    ) -> Result<crate::ruby_family::Annotation> {
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
    ) -> Result<crate::ruby_family::Annotation> {
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
    pub fn ruby_styles(&self) -> Result<crate::ruby_family::Styles> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::parse_ruby_styles)
    }

    /// Insert or replace one named ruby style definition and return the old value.
    pub fn set_ruby_style(
        &mut self,
        style: &crate::ruby_family::Style,
    ) -> Result<Option<crate::ruby_family::Style>> {
        style.validate()?;
        let old = self.ruby_styles()?.get(&style.name).cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::set_ruby_style_xml(&styles, style)?);
        Ok(old)
    }

    /// Remove one named ruby style definition and return the old value.
    ///
    /// Existing `text:ruby` style references are preserved verbatim, so callers
    /// can intentionally manage their lifecycle separately.
    pub fn remove_ruby_style(&mut self, name: &str) -> Result<Option<crate::ruby_family::Style>> {
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
    pub fn content_font_face_declarations(&self) -> Result<Option<crate::font_face::Declarations>> {
        self.with_content_xml(crate::font_face::parse_content_font_face_declarations)
    }

    /// Replace content-part font-face declarations and return the old value.
    ///
    /// This edits `content.xml` only. It does not fetch linked font resources,
    /// load a font, or inspect embedded font data.
    pub fn set_content_font_face_declarations(
        &mut self,
        declarations: &crate::font_face::Declarations,
    ) -> Result<Option<crate::font_face::Declarations>> {
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
    ) -> Result<Option<crate::font_face::Declarations>> {
        let (updated, old) =
            self.with_content_xml(crate::font_face::remove_content_font_face_declarations_xml)?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return font-face declarations from the current `styles.xml`.
    ///
    /// Linked font resources remain inert metadata. This does not fetch a URI,
    /// load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<crate::font_face::Declarations>> {
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
        declarations: &crate::font_face::Declarations,
    ) -> Result<Option<crate::font_face::Declarations>> {
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
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
    ) -> Result<Option<crate::font_face::Declarations>> {
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
    pub fn drawing_gradients(&self) -> Result<crate::drawing::resources::gradient::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::gradient::parse_drawing_gradients,
        )
    }

    /// Return named drawing hatch resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing::resources::hatch::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::hatch::parse_drawing_hatches,
        )
    }

    /// Return named drawing stroke-dash resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing::resources::stroke_dash::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::stroke_dash::parse_drawing_stroke_dashes,
        )
    }

    /// Return named drawing fill-image definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites, follow links, load linked resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing::resources::fill_image::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::fill_image::parse_drawing_fill_images,
        )
    }

    /// Return named drawing marker definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing::resources::marker::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::marker::parse_drawing_markers,
        )
    }

    /// Return named drawing opacity definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing::resources::opacity::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::opacity::parse_drawing_opacities,
        )
    }

    /// Return stored footnote and endnote presentation configurations.
    ///
    /// The result describes style metadata only. It never renumbers, lays out,
    /// or renders notes.
    pub fn notes_configurations(&self) -> Result<crate::notes_configuration::Configurations> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::notes_configuration::parse)
    }

    /// Return stored outline numbering styles from current styles metadata.
    ///
    /// The result does not apply styles to headings, generate labels, or
    /// update tables of contents.
    pub fn outline_styles(&self) -> Result<crate::outline_style::Styles> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::outline_style::parse_outline_styles,
        )
    }

    /// Insert or replace one named outline numbering style.
    ///
    /// This edits `styles.xml` only and returns the previous style with the
    /// same name. It does not alter heading structure or cached index content.
    pub fn set_outline_style(
        &mut self,
        style: &crate::outline_style::Style,
    ) -> Result<Option<crate::outline_style::Style>> {
        style.validate()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        let (updated, old) = crate::outline_style::set_outline_style_xml(&styles, style)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Remove one named outline numbering style and return its prior value.
    ///
    /// Existing heading references are retained verbatim, allowing callers to
    /// manage those references separately.
    pub fn remove_outline_style(
        &mut self,
        name: &str,
    ) -> Result<Option<crate::outline_style::Style>> {
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        let (updated, old) = crate::outline_style::remove_outline_style_xml(styles, name)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Insert or replace one stored footnote or endnote configuration.
    ///
    /// This edits `styles.xml` only and returns the prior configuration for the
    /// same note class. It never changes note anchors, citations, or numbering.
    pub fn set_notes_configuration(
        &mut self,
        configuration: &crate::notes_configuration::Configuration,
    ) -> Result<Option<crate::notes_configuration::Configuration>> {
        configuration.validate()?;
        let old = self
            .notes_configurations()?
            .get(configuration.note_class)
            .cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::notes_configuration::set_xml(&styles, configuration)?);
        Ok(old)
    }

    /// Replace both stored note-class configurations and return the old values.
    ///
    /// An absent class is removed from `styles.xml`. This updates metadata only and
    /// never recalculates citations, sequence numbers, or page layout.
    pub fn set_notes_configurations(
        &mut self,
        configurations: &crate::notes_configuration::Configurations,
    ) -> Result<crate::notes_configuration::Configurations> {
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
            .unwrap_or_else(Structure::default_styles_xml);
        for note_class in crate::notes_configuration::Class::ALL {
            styles = match configurations.get(note_class) {
                Some(configuration) => crate::notes_configuration::set_xml(&styles, configuration)?,
                None => crate::notes_configuration::remove_xml(&styles, note_class)?,
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
        note_class: crate::notes_configuration::Class,
    ) -> Result<Option<crate::notes_configuration::Configuration>> {
        let old = self.notes_configurations()?.get(note_class).cloned();
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::notes_configuration::remove_xml(styles, note_class)?);
        Ok(old)
    }

    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is styles metadata only. It is never used to generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(
        &self,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
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
        configuration: &crate::bibliography_configuration::Configuration,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
        configuration.validate()?;
        let old = self.bibliography_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
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
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
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
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(None), crate::line_numbering::parse)
    }

    /// Insert or replace document line-numbering configuration.
    ///
    /// This updates stored style metadata only. It never calculates page or
    /// line numbers.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: &crate::line_numbering::Configuration,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        configuration.validate()?;
        let old = self.line_numbering_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::line_numbering::set_xml(&styles, configuration)?);
        Ok(old)
    }

    /// Remove document line-numbering configuration and return its old value.
    pub fn clear_line_numbering_configuration(
        &mut self,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        let old = self.line_numbering_configuration()?;
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::line_numbering::remove_xml(styles)?);
        Ok(old)
    }

    /// Return generated indexes from the current authoritative content XML.
    pub fn text_indexes(&self) -> Result<Vec<TextIndex>> {
        self.with_content_xml(crate::index::parse_text_indexes)
    }

    /// Append caller-authored index markup to `office:text` without refreshing its cache.
    pub fn insert_text_index(&mut self, index: &TextIndex) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::insert_text_index_xml(xml, index))?;
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
        let updated =
            self.with_content_xml(|xml| crate::replace_text_index_xml(xml, name, replacement))?;
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
        let updated = self.with_content_xml(|xml| crate::remove_text_index_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return typed point and resolved range index marks in document order.
    pub fn text_index_marks(&self) -> Result<Vec<TextIndexMark>> {
        self.with_content_xml(crate::index_mark::parse_text_index_marks)
    }

    /// Insert a point mark at a paragraph end, or wrap the paragraph with a range mark.
    pub fn insert_text_index_mark(
        &mut self,
        paragraph_index: usize,
        mark: &TextIndexMark,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_text_index_mark_xml(xml, paragraph_index, mark)
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
            crate::replace_text_index_mark_xml(xml, mark_index, replacement)
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
            self.with_content_xml(|xml| crate::remove_text_index_mark_xml(xml, mark_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return point and resolved range reference targets in document order.
    pub fn reference_marks(&self) -> Result<Vec<ReferenceMark>> {
        self.with_content_xml(crate::reference_mark::parse_reference_marks)
    }

    /// Insert a point reference at a paragraph end, or wrap the paragraph with a range reference.
    pub fn insert_reference_mark(
        &mut self,
        paragraph_index: usize,
        mark: &ReferenceMark,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::insert_reference_mark_xml(xml, paragraph_index, mark))?;
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
            crate::replace_reference_mark_xml(xml, mark_index, replacement)
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
            self.with_content_xml(|xml| crate::remove_reference_mark_xml(xml, mark_index))?;
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
    pub fn form_properties(&self) -> Result<Vec<crate::form::Property>> {
        self.with_content_xml(crate::form::form_properties)
    }

    /// Insert a property into a form/control owner selected in document order.
    pub fn insert_form_property(
        &mut self,
        owner_index: usize,
        property: &crate::form::Property,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_form_property_xml(xml, owner_index, property)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a form property selected in document order.
    pub fn replace_form_property(
        &mut self,
        property_index: usize,
        replacement: &crate::form::Property,
    ) -> Result<crate::form::Property> {
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
            crate::form::replace_form_property_xml(xml, property_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a form property and remove its container when it becomes empty.
    pub fn remove_form_property(&mut self, property_index: usize) -> Result<crate::form::Property> {
        let old = self
            .form_properties()?
            .get(property_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "form property {property_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::form::remove_form_property_xml(xml, property_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return text and textarea controls in document order.
    pub fn text_controls(&self) -> Result<Vec<crate::form::TextControl>> {
        self.with_content_xml(crate::form::text_controls)
    }

    /// Insert a text or textarea control into a form selected in document order.
    pub fn insert_text_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TextControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_text_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a text or textarea control selected in document order.
    pub fn replace_text_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TextControl,
    ) -> Result<crate::form::TextControl> {
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
            crate::form::replace_text_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a text or textarea control selected in document order.
    pub fn remove_text_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::TextControl> {
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
            self.with_content_xml(|xml| crate::form::remove_text_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return button and checkbox controls in document order.
    pub fn interactive_controls(&self) -> Result<Vec<crate::form::InteractiveControl>> {
        self.with_content_xml(crate::form::interactive_controls)
    }

    /// Insert a button or checkbox into a form selected in document order.
    pub fn insert_interactive_control(
        &mut self,
        form_index: usize,
        control: &crate::form::InteractiveControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_interactive_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a button or checkbox selected in document order.
    pub fn replace_interactive_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::InteractiveControl,
    ) -> Result<crate::form::InteractiveControl> {
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
            crate::form::replace_interactive_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a button or checkbox selected in document order.
    pub fn remove_interactive_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::InteractiveControl> {
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
            crate::form::remove_interactive_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return listbox and combobox controls in document order.
    pub fn selection_controls(&self) -> Result<Vec<crate::form::SelectionControl>> {
        self.with_content_xml(crate::form::selection_controls)
    }

    /// Insert a listbox or combobox into a form selected in document order.
    pub fn insert_selection_control(
        &mut self,
        form_index: usize,
        control: &crate::form::SelectionControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_selection_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a listbox or combobox selected in document order.
    pub fn replace_selection_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::SelectionControl,
    ) -> Result<crate::form::SelectionControl> {
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
            crate::form::replace_selection_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a listbox or combobox selected in document order.
    pub fn remove_selection_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::SelectionControl> {
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
            crate::form::remove_selection_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return radio, frame, and image-button controls in document order.
    pub fn visual_controls(&self) -> Result<Vec<crate::form::VisualControl>> {
        self.with_content_xml(crate::form::visual_controls)
    }

    /// Insert a radio, frame, or image-button into a form selected in document order.
    pub fn insert_visual_control(
        &mut self,
        form_index: usize,
        control: &crate::form::VisualControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_visual_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a radio, frame, or image-button selected in document order.
    pub fn replace_visual_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::VisualControl,
    ) -> Result<crate::form::VisualControl> {
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
            crate::form::replace_visual_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a radio, frame, or image-button selected in document order.
    pub fn remove_visual_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::VisualControl> {
        let old = self
            .visual_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "visual control {control_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::form::remove_visual_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return fixed-text, hidden, and generic controls in document order.
    pub fn generic_form_controls(&self) -> Result<Vec<crate::form::GenericFormControl>> {
        self.with_content_xml(crate::form::generic_form_controls)
    }

    /// Insert a fixed-text, hidden, or generic control into a form selected in document order.
    pub fn insert_generic_form_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GenericFormControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_generic_form_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a fixed-text, hidden, or generic control selected in document order.
    pub fn replace_generic_form_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GenericFormControl,
    ) -> Result<crate::form::GenericFormControl> {
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
            crate::form::replace_generic_form_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a fixed-text, hidden, or generic control selected in document order.
    pub fn remove_generic_form_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::GenericFormControl> {
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
            crate::form::remove_generic_form_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return password and file controls in document order.
    pub fn password_file_controls(&self) -> Result<Vec<crate::form::PasswordFileControl>> {
        self.with_content_xml(crate::form::password_file_controls)
    }

    /// Insert a password or file control into a form selected in document order.
    pub fn insert_password_file_control(
        &mut self,
        form_index: usize,
        control: &crate::form::PasswordFileControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_password_file_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a password or file control selected in document order.
    pub fn replace_password_file_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::PasswordFileControl,
    ) -> Result<crate::form::PasswordFileControl> {
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
            crate::form::replace_password_file_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a password or file control selected in document order.
    pub fn remove_password_file_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::PasswordFileControl> {
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
            crate::form::remove_password_file_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return image-frame controls in document order without resolving image references.
    pub fn image_frame_controls(&self) -> Result<Vec<crate::form::ImageFrameControl>> {
        self.with_content_xml(crate::form::image_frame_controls)
    }

    /// Insert an image-frame control into a form selected in document order.
    pub fn insert_image_frame_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ImageFrameControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_image_frame_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace an image-frame control selected in document order.
    pub fn replace_image_frame_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ImageFrameControl,
    ) -> Result<crate::form::ImageFrameControl> {
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
            crate::form::replace_image_frame_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove an image-frame control selected in document order.
    pub fn remove_image_frame_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::ImageFrameControl> {
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
            crate::form::remove_image_frame_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return value-range controls in document order without resolving bindings.
    pub fn value_range_controls(&self) -> Result<Vec<crate::form::ValueRangeControl>> {
        self.with_content_xml(crate::form::value_range_controls)
    }

    /// Insert a value-range control into a form selected in document order.
    pub fn insert_value_range_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ValueRangeControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_value_range_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a value-range control selected in document order.
    pub fn replace_value_range_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ValueRangeControl,
    ) -> Result<crate::form::ValueRangeControl> {
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
            crate::form::replace_value_range_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a value-range control selected in document order.
    pub fn remove_value_range_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::ValueRangeControl> {
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
            crate::form::remove_value_range_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return formatted-text, number, date, and time controls in document order.
    pub fn typed_value_controls(&self) -> Result<Vec<crate::form::TypedValueControl>> {
        self.with_content_xml(crate::form::typed_value_controls)
    }

    /// Insert a typed value control into a form selected in document order.
    pub fn insert_typed_value_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TypedValueControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_typed_value_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a typed value control selected in document order.
    pub fn replace_typed_value_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TypedValueControl,
    ) -> Result<crate::form::TypedValueControl> {
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
            crate::form::replace_typed_value_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a typed value control selected in document order.
    pub fn remove_typed_value_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::TypedValueControl> {
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
            crate::form::remove_typed_value_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    pub fn grid_controls(&self) -> Result<Vec<crate::form::GridControl>> {
        self.with_content_xml(crate::form::grid_controls)
    }
    pub fn insert_grid_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GridControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_grid_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }
    pub fn replace_grid_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GridControl,
    ) -> Result<crate::form::GridControl> {
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
            crate::form::replace_grid_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }
    pub fn remove_grid_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::GridControl> {
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
            self.with_content_xml(|xml| crate::form::remove_grid_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Insert a field at the end of a paragraph selected in document order.
    pub fn insert_dynamic_text_field(
        &mut self,
        paragraph_index: usize,
        field: &DynamicTextField,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::insert_dynamic_text_field_xml(xml, paragraph_index, field)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a dynamic field selected in document order and return its old value.
    pub fn replace_dynamic_text_field(
        &mut self,
        field_index: usize,
        replacement: &DynamicTextField,
    ) -> Result<DynamicTextField> {
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
            crate::replace_dynamic_text_field_xml(xml, field_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a dynamic field selected in document order and return its old value.
    pub fn remove_dynamic_text_field(&mut self, field_index: usize) -> Result<DynamicTextField> {
        let old = self
            .dynamic_text_fields()?
            .get(field_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "dynamic text field index {field_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::remove_dynamic_text_field_xml(xml, field_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return the explicit page-sequence master-page assignments, if authored.
    ///
    /// The model preserves only the ordered `text:master-page-name` values of
    /// the document's `text:page-sequence` (ODF 1.3 §5.3); litchi does not
    /// paginate or resolve the referenced master pages.
    pub fn page_sequence(&self) -> Result<Option<Sequence>> {
        self.with_content_xml(parse_page_sequence)
    }

    /// Set, replace, or clear the document's `text:page-sequence`.
    ///
    /// A new sequence is written as the first child of `office:text`, matching
    /// the element order of ODF 1.3 §5.1. Passing `None` removes an existing
    /// sequence and is a no-op when none exists. Master-page names are stored
    /// lexically and never resolved against `styles.xml`.
    pub fn set_page_sequence(&mut self, sequence: Option<&Sequence>) -> Result<()> {
        let updated = self.with_content_xml(|xml| set_page_sequence_xml(xml, sequence))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Return validated variable, user-field, sequence, and DDE declarations.
    pub fn variable_declarations(&self) -> Result<Declarations> {
        self.with_content_xml(|content| {
            if let Some(styles) = self.styles_xml.as_deref() {
                crate::variable_declaration::parse_parts(&[
                    (content, Part::Content),
                    (styles, Part::Styles),
                ])
            } else {
                crate::variable_declaration::parse_parts(&[(content, Part::Content)])
            }
        })
    }

    /// Atomically insert or replace one declaration container.
    ///
    /// Content and styles are reparsed together before commit, preserving all
    /// cross-part declaration and field-reference invariants.
    pub fn set_variable_declaration_group(&mut self, group: &Group) -> Result<Option<Group>> {
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
            Part::Content => {
                let updated =
                    self.with_content_xml(|xml| crate::variable_declaration::set_xml(xml, group))?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_parts(&[
                        (&updated, Part::Content),
                        (styles, Part::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_parts(&[(&updated, Part::Content)])?;
                }
                self.content_xml = Some(updated);
            },
            Part::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::variable_declaration::set_xml(styles, group)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_parts(&[
                        (content, Part::Content),
                        (&updated, Part::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            Part::Flat => {
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
        part: Part,
        scope: &Scope,
        kind: Kind,
    ) -> Result<Option<Group>> {
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
            Part::Content => {
                let updated = self.with_content_xml(|xml| {
                    crate::variable_declaration::remove_xml(xml, scope, kind)
                })?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_parts(&[
                        (&updated, Part::Content),
                        (styles, Part::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_parts(&[(&updated, Part::Content)])?;
                }
                self.content_xml = Some(updated);
            },
            Part::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::variable_declaration::remove_xml(styles, scope, kind)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_parts(&[
                        (content, Part::Content),
                        (&updated, Part::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            Part::Flat => {
                return Err(litchi_core::Error::InvalidFormat(
                    "MutableDocument cannot edit flat-document declarations".to_string(),
                ));
            },
        }
        Ok(Some(old))
    }

    pub(super) fn with_content_xml<T>(
        &self,
        operation: impl FnOnce(&str) -> Result<T>,
    ) -> Result<T> {
        if let Some(xml) = self.content_xml.as_deref() {
            operation(xml)
        } else {
            let xml = self.generate_content_xml();
            operation(&xml)
        }
    }

    pub(super) fn invalidate_content_xml(&mut self) {
        self.content_xml = None;
    }

    /// Return declarations, policy, and marker-correlated content from current XML.
    pub fn tracked_changes(&self) -> Result<crate::TrackedChanges> {
        self.with_content_xml(crate::parser::Parser::parse_tracked_changes)
    }

    /// Atomically replace the declaration table and policy metadata.
    pub fn set_tracked_changes(&mut self, tracked: crate::TrackedChanges) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::set_tracked_changes_xml(xml, Some(&tracked)))?;
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
    pub fn add_tracked_change(&mut self, change: crate::TrackChange) -> Result<()> {
        let mut tracked = self.tracked_changes()?;
        tracked.changes.push(change);
        self.set_tracked_changes(tracked)
    }

    /// Atomically replace a declaration without changing marker identity.
    pub fn update_tracked_change(
        &mut self,
        id: &str,
        replacement: crate::TrackChange,
    ) -> Result<crate::TrackChange> {
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
    pub fn remove_tracked_change(&mut self, id: &str) -> Result<crate::TrackChange> {
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
            let unmarked = crate::unmark_tracked_change_xml(xml, id)?;
            crate::set_tracked_changes_xml(&unmarked, Some(&tracked))
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
                candidate = crate::unmark_tracked_change_xml(&candidate, &change.id)?;
            }
            crate::set_tracked_changes_xml(&candidate, None)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Mark a live insertion or format-change range using Unicode character offsets.
    pub fn mark_tracked_change_range(
        &mut self,
        change_id: &str,
        start: crate::Position,
        end: crate::Position,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::mark_tracked_change_range_xml(xml, change_id, &start, &end)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Place a point deletion marker using a Unicode character offset.
    pub fn mark_tracked_deletion(
        &mut self,
        change_id: &str,
        position: crate::Position,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::mark_tracked_deletion_xml(xml, change_id, &position))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove every marker for one declaration while retaining its live text.
    pub fn unmark_tracked_change(&mut self, change_id: &str) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::unmark_tracked_change_xml(xml, change_id))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Return typed nested sections from current authoritative content XML.
    pub fn sections(&self) -> Result<Vec<crate::Section>> {
        self.with_content_xml(crate::parser::Parser::parse_sections)
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
        start: crate::Block,
        end: crate::Block,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::wrap_section_xml(xml, section, &start, &end))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Parse the document's master pages and current header/footer regions.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), read)
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
        region: crate::header_footer::properties::Region,
        properties: &crate::header_footer::properties::StyleProperties,
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
        let replacement = crate::header_footer::properties::replace_page_layout_region_properties(
            layout, region, properties,
        )?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace typed columns in one existing page layout.
    pub fn set_page_layout_columns(
        &mut self,
        page_layout_name: &str,
        columns: &crate::style::columns::Columns,
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
        let replacement = crate::style::columns::replace_page_layout_columns(layout, columns)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace the typed footnote separator in one existing page layout.
    pub fn set_page_layout_footnote_separator(
        &mut self,
        page_layout_name: &str,
        separator: &crate::footnote_separator::Separator,
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
        item: &crate::list_label_alignment::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml list style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::list_label_alignment::set_xml(styles, item)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace, insert, or remove one existing paragraph style's direct drop cap.
    pub fn set_paragraph_style_drop_cap(
        &mut self,
        style: &crate::style::paragraph::drop_cap::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml paragraph style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::paragraph::drop_cap::set_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed row properties on an existing table-row style.
    pub fn set_table_row_style_properties(
        &mut self,
        style: &crate::style::table::row::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table-row style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::table::row::set_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed properties on an existing table style.
    pub fn set_table_style_properties(
        &mut self,
        style: &crate::style::table::table::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::table::table::set_xml(styles, style)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// A minimal page layout is created in `office:automatic-styles` when a
    /// layout with `page_layout_name` does not already exist.
    pub fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        let styles = self
            .styles_xml
            .get_or_insert_with(Structure::default_styles_xml);
        *styles = add(styles, name, page_layout_name)?;
        Ok(())
    }

    /// Insert a complete typed master page without rewriting unrelated styles.
    pub fn insert_master_page(&mut self, page: &Master) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self
            .styles_xml
            .get_or_insert_with(Structure::default_styles_xml);
        *styles = insert(styles, &fragment)?;
        Ok(())
    }

    /// Replace one named master page without rewriting unrelated styles.
    pub fn replace_master_page(&mut self, name: &str, page: &Master) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(replace(styles, name, &fragment)?);
        Ok(())
    }

    /// Remove one named master page without rewriting unrelated styles.
    pub fn remove_master_page(&mut self, name: &str) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(remove(styles, name)?);
        Ok(())
    }

    /// Set plain text in one header/footer region of an existing master page.
    ///
    /// Only the selected region is rewritten; all unrelated style XML is preserved.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: &str,
        kind: crate::header_footer::Kind,
        text: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_text(styles, master_page_name, kind, Some(text))?);
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
        kind: crate::header_footer::Kind,
        xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_xml(styles, master_page_name, kind, xml)?);
        Ok(())
    }

    /// Remove one header/footer region from an existing master page.
    pub fn clear_header_footer(
        &mut self,
        master_page_name: &str,
        kind: crate::header_footer::Kind,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_text(styles, master_page_name, kind, None)?);
        Ok(())
    }
}
