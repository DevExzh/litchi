//! Contextual content access and mutation.

#![deny(clippy::expect_used, clippy::unwrap_used)]

use super::super::model::MutableDocument;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use litchi_core::Result;

/// Read-only view of the mutable document's content layer.
pub struct Content<'document> {
    pub(super) document: &'document MutableDocument,
}

/// Mutable view of the document's structural and inline content.
pub struct ContentMut<'document> {
    pub(super) document: &'document mut MutableDocument,
}

impl Content<'_> {
    /// Return top-level paragraphs in document order.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        self.document.paragraphs()
    }

    /// Return top-level headings in document order.
    pub fn headings(&self) -> Vec<&Heading> {
        self.document.headings()
    }

    /// Return top-level lists in document order.
    pub fn lists(&self) -> Vec<&List> {
        self.document.lists()
    }

    /// Return top-level tables in document order.
    pub fn tables(&self) -> Vec<&Table> {
        self.document.tables()
    }

    /// Parse dynamic text fields from the authoritative content snapshot.
    pub fn dynamic_text_fields(&self) -> Result<Vec<crate::elements::field::DynamicTextField>> {
        self.document.dynamic_text_fields()
    }

    /// Parse the document's semantic sections from the authoritative snapshot.
    pub fn sections(&self) -> Result<Vec<crate::Section>> {
        self.document.sections()
    }

    /// Parse the explicit page sequence, if the document has one.
    pub fn page_sequence(&self) -> Result<Option<crate::page_sequence::Sequence>> {
        self.document.page_sequence()
    }
}

impl ContentMut<'_> {
    /// Reborrow this editor as a read-only content view.
    pub fn read(&self) -> Content<'_> {
        Content {
            document: self.document,
        }
    }

    /// Append a plain paragraph to the document body projection.
    pub fn add_paragraph(&mut self, text: &str) -> Result<()> {
        self.document.add_paragraph(text)
    }

    /// Insert a plain paragraph at a top-level structural position.
    #[deprecated(note = "use insert_paragraph_at with a checked semantic Position")]
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        self.document
            .insert_paragraph_at(litchi_core::Position::new(index), text)
    }

    /// Insert a paragraph at a checked semantic paragraph position.
    pub fn insert_paragraph_at(
        &mut self,
        position: litchi_core::Position,
        text: &str,
    ) -> Result<()> {
        self.document.insert_paragraph_at(position, text)
    }

    /// Replace one top-level paragraph's plain text.
    #[deprecated(note = "use replace_paragraph_at with a checked semantic Position")]
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        self.document
            .replace_paragraph_at(litchi_core::Position::new(index), text)
    }

    /// Replace a paragraph at a checked semantic paragraph position.
    pub fn replace_paragraph_at(
        &mut self,
        position: litchi_core::Position,
        text: &str,
    ) -> Result<()> {
        self.document.replace_paragraph_at(position, text)
    }

    /// Remove one top-level paragraph and return its typed value.
    #[deprecated(note = "use remove_paragraph_at with a checked semantic Position")]
    pub fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        self.document
            .remove_paragraph_at(litchi_core::Position::new(index))
    }

    /// Remove a paragraph at a checked semantic paragraph position.
    pub fn remove_paragraph_at(&mut self, position: litchi_core::Position) -> Result<Paragraph> {
        self.document.remove_paragraph_at(position)
    }

    /// Remove all top-level paragraphs while retaining other body elements.
    pub fn clear_paragraphs(&mut self) {
        self.document.clear_paragraphs();
    }

    /// Append a heading to the document body projection.
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        self.document.add_heading(text, level)
    }

    /// Append an existing list to the document body projection.
    pub fn add_list(&mut self, list: List) -> Result<()> {
        self.document.add_list(list)
    }

    /// Append an existing table to the document body projection.
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        self.document.add_table(table)
    }

    /// Remove one top-level table and return its typed value.
    pub fn remove_table(&mut self, index: usize) -> Result<Table> {
        self.document.remove_table(index)
    }

    /// Remove all top-level tables while retaining other body elements.
    pub fn clear_tables(&mut self) {
        self.document.clear_tables();
    }

    /// Remove all projected top-level body content.
    pub fn clear(&mut self) {
        self.document.clear_content();
    }

    /// Append an ODF `text:line-break` to a paragraph selected in document
    /// order, preserving its existing inline markup.
    #[deprecated(note = "use append_line_break_at with a checked semantic Position")]
    pub fn append_line_break(&mut self, paragraph_index: usize) -> Result<()> {
        self.document
            .append_line_break_at(litchi_core::Position::new(paragraph_index))
    }

    /// Append a line break at a checked semantic paragraph position.
    pub fn append_line_break_at(&mut self, paragraph: litchi_core::Position) -> Result<()> {
        self.document.append_line_break_at(paragraph)
    }

    /// Set or clear an explicit page sequence in `content.xml`.
    pub fn set_page_sequence(
        &mut self,
        sequence: Option<&crate::page_sequence::Sequence>,
    ) -> Result<()> {
        self.document.set_page_sequence(sequence)
    }

    /// Append a validated note to a paragraph selected in document order.
    pub fn insert_note(&mut self, paragraph_index: usize, note: &crate::note::Note) -> Result<()> {
        self.document.insert_note(paragraph_index, note)
    }
}
