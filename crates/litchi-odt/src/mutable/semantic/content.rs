//! Contextual content access and mutation.

#![deny(clippy::expect_used, clippy::unwrap_used)]

use super::super::model::MutableDocument;
use crate::elements::text::{Heading, Paragraph};
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

    /// Replace a paragraph at a checked semantic paragraph position.
    pub(crate) fn replace_paragraph_at(
        &mut self,
        position: litchi_core::Position,
        text: &str,
    ) -> Result<()> {
        self.document.replace_paragraph_at(position, text)
    }

    /// Append a heading to the document body projection.
    pub(crate) fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        self.document.add_heading(text, level)
    }
}
