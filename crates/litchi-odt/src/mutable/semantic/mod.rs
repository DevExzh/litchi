//! Contextual semantic views over a mutable ODT document.
//!
//! `MutableDocument` retains its original flat methods for compatibility. These
//! views provide a small, discoverable layer for code that wants to keep
//! structural content work separate from style/package work:

mod content;
mod styles;
mod text;

pub(crate) use content::{Content, ContentMut};
pub(crate) use styles::{Styles, StylesMut};

use super::model::MutableDocument;

impl MutableDocument {
    /// Borrow the document's content layer for semantic reads.
    pub(crate) fn content(&self) -> Content<'_> {
        Content { document: self }
    }

    /// Borrow the document's content layer for structural and inline edits.
    pub(crate) fn content_mut(&mut self) -> ContentMut<'_> {
        ContentMut { document: self }
    }

    /// Borrow the retained `styles.xml` layer for semantic reads.
    pub(crate) fn styles(&self) -> Styles<'_> {
        Styles { document: self }
    }

    /// Borrow the retained `styles.xml` layer for targeted style edits.
    pub(crate) fn styles_mut(&mut self) -> StylesMut<'_> {
        StylesMut { document: self }
    }
}
