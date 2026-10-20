//! Contextual content access and mutation.

#![deny(clippy::expect_used, clippy::unwrap_used)]

use super::super::model::MutableDocument;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use litchi_core::Result;

/// Read-only view of the mutable document's content layer.
pub(crate) struct Content<'document> {
    pub(super) document: &'document MutableDocument,
}

/// Mutable view of the document's structural and inline content.
pub(crate) struct ContentMut<'document> {
    pub(super) document: &'document mut MutableDocument,
}

impl Content<'_> {
    /// Return top-level paragraphs in document order.
    pub(crate) fn paragraphs(&self) -> Vec<&Paragraph> {
        self.document.paragraphs()
    }

    /// Return top-level headings in document order.
    pub(crate) fn headings(&self) -> Vec<&Heading> {
        self.document.headings()
    }

    /// Return top-level lists in document order.
    pub(crate) fn lists(&self) -> Vec<&List> {
        self.document.lists()
    }

    /// Return top-level tables in document order.
    pub(crate) fn tables(&self) -> Vec<&Table> {
        self.document.tables()
    }

    /// Parse dynamic text fields from the authoritative content snapshot.
    pub(crate) fn dynamic_text_fields(
        &self,
    ) -> Result<Vec<crate::elements::field::DynamicTextField>> {
        self.document.dynamic_text_fields()
    }

    /// Inspect in-content RDFa declarations and inline `text:meta` values.
    pub(crate) fn in_content_metadata(&self) -> Result<crate::ContentMetadata> {
        self.document.in_content_metadata()
    }

    /// Inspect inert XForms model declarations below `office:forms`.
    pub(crate) fn xforms_models(&self) -> Result<Vec<crate::xforms::Model>> {
        self.document.xforms_models()
    }

    /// Parse the document's semantic sections from the authoritative snapshot.
    pub(crate) fn sections(&self) -> Result<Vec<crate::Section>> {
        self.document.sections()
    }

    /// Parse the explicit page sequence, if the document has one.
    pub(crate) fn page_sequence(&self) -> Result<Option<crate::page_sequence::Sequence>> {
        self.document.page_sequence()
    }
}

impl ContentMut<'_> {
    /// Reborrow this editor as a read-only content view.
    pub(crate) fn read(&self) -> Content<'_> {
        Content {
            document: self.document,
        }
    }

    /// Append a plain paragraph to the document body projection.
    pub(crate) fn add_paragraph(&mut self, text: &str) -> Result<()> {
        self.document.add_paragraph(text)
    }

    /// Insert a plain paragraph at a top-level structural position.
    #[deprecated(note = "use insert_paragraph_at with a checked semantic Position")]
    pub(crate) fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        self.document
            .insert_paragraph_at(litchi_core::Position::new(index), text)
    }

    /// Insert a paragraph at a checked semantic paragraph position.
    pub(crate) fn insert_paragraph_at(
        &mut self,
        position: litchi_core::Position,
        text: &str,
    ) -> Result<()> {
        self.document.insert_paragraph_at(position, text)
    }

    /// Replace one top-level paragraph's plain text.
    #[deprecated(note = "use replace_paragraph_at with a checked semantic Position")]
    pub(crate) fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        self.document
            .replace_paragraph_at(litchi_core::Position::new(index), text)
    }

    /// Replace a paragraph at a checked semantic paragraph position.
    pub(crate) fn replace_paragraph_at(
        &mut self,
        position: litchi_core::Position,
        text: &str,
    ) -> Result<()> {
        self.document.replace_paragraph_at(position, text)
    }

    /// Remove one top-level paragraph and return its typed value.
    #[deprecated(note = "use remove_paragraph_at with a checked semantic Position")]
    pub(crate) fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        self.document
            .remove_paragraph_at(litchi_core::Position::new(index))
    }

    /// Remove a paragraph at a checked semantic paragraph position.
    pub(crate) fn remove_paragraph_at(
        &mut self,
        position: litchi_core::Position,
    ) -> Result<Paragraph> {
        self.document.remove_paragraph_at(position)
    }

    /// Remove all top-level paragraphs while retaining other body elements.
    pub(crate) fn clear_paragraphs(&mut self) {
        self.document.clear_paragraphs();
    }

    /// Append a heading to the document body projection.
    pub(crate) fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        self.document.add_heading(text, level)
    }

    /// Append an existing list to the document body projection.
    pub(crate) fn add_list(&mut self, list: List) -> Result<()> {
        self.document.add_list(list)
    }

    /// Append an existing table to the document body projection.
    pub(crate) fn add_table(&mut self, table: Table) -> Result<()> {
        self.document.add_table(table)
    }

    /// Remove one top-level table and return its typed value.
    pub(crate) fn remove_table(&mut self, index: usize) -> Result<Table> {
        self.document.remove_table(index)
    }

    /// Remove all top-level tables while retaining other body elements.
    pub(crate) fn clear_tables(&mut self) {
        self.document.clear_tables();
    }

    /// Remove all projected top-level body content.
    pub(crate) fn clear(&mut self) {
        self.document.clear_content();
    }

    /// Append an ODF `text:line-break` to a paragraph selected in document
    /// order, preserving its existing inline markup.
    #[deprecated(note = "use append_line_break_at with a checked semantic Position")]
    pub(crate) fn append_line_break(&mut self, paragraph_index: usize) -> Result<()> {
        self.document
            .append_line_break_at(litchi_core::Position::new(paragraph_index))
    }

    /// Append a line break at a checked semantic paragraph position.
    pub(crate) fn append_line_break_at(&mut self, paragraph: litchi_core::Position) -> Result<()> {
        self.document.append_line_break_at(paragraph)
    }

    /// Set or clear an explicit page sequence in `content.xml`.
    pub(crate) fn set_page_sequence(
        &mut self,
        sequence: Option<&crate::page_sequence::Sequence>,
    ) -> Result<()> {
        self.document.set_page_sequence(sequence)
    }

    /// Append a validated note to a paragraph selected in document order.
    pub(crate) fn insert_note(
        &mut self,
        paragraph_index: usize,
        note: &crate::note::Note,
    ) -> Result<()> {
        self.document.insert_note(paragraph_index, note)
    }

    /// Set or clear RDFa on one paragraph.
    pub(crate) fn set_paragraph_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        self.document.set_paragraph_rdfa(position, value)
    }

    /// Set or clear RDFa on one uniquely named bookmark start.
    pub(crate) fn set_bookmark_rdfa(
        &mut self,
        name: &str,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        self.document.set_bookmark_rdfa(name, value)
    }

    /// Set or clear RDFa on one inline `text:meta` element.
    pub(crate) fn set_text_meta_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        self.document.set_text_meta_rdfa(position, value)
    }

    /// Insert an inline `text:meta` into a paragraph.
    pub(crate) fn insert_text_meta(
        &mut self,
        paragraph: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        self.document.insert_text_meta(paragraph, value)
    }

    /// Replace one inline `text:meta` element.
    pub(crate) fn replace_text_meta(
        &mut self,
        position: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        self.document.replace_text_meta(position, value)
    }

    /// Remove one inline `text:meta` element.
    pub(crate) fn remove_text_meta(&mut self, position: litchi_core::Position) -> Result<()> {
        self.document.remove_text_meta(position)
    }

    /// Insert a typed XForms model at the end of `office:forms`.
    pub(crate) fn insert_xforms_model(&mut self, model: &crate::xforms::Model) -> Result<()> {
        self.document.insert_xforms_model(model)
    }

    /// Replace one typed XForms model.
    pub(crate) fn replace_xforms_model(
        &mut self,
        position: litchi_core::Position,
        model: &crate::xforms::Model,
    ) -> Result<()> {
        self.document.replace_xforms_model(position, model)
    }

    /// Remove one typed XForms model.
    pub(crate) fn remove_xforms_model(&mut self, position: litchi_core::Position) -> Result<()> {
        self.document.remove_xforms_model(position)
    }
}
