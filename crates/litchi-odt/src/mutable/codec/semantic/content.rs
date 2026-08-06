//! Content-facing semantic reads and lossless snapshot edits.

use super::super::super::model::MutableDocument;
use crate::elements::field::{DynamicTextField, FieldParser};
use litchi_core::Result;
use std::ops::Range;

impl MutableDocument {
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
}
