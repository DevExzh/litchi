//! Contextual semantic views over a mutable ODT document.
//!
//! `MutableDocument` retains its original flat methods for compatibility. These
//! views provide a small, discoverable layer for code that wants to keep
//! structural content work separate from style/package work:

#[cfg(test)]
mod content;
#[cfg(test)]
mod styles;
mod text;

#[cfg(test)]
pub(crate) use content::{Content, ContentMut};
#[cfg(test)]
pub(crate) use styles::{Styles, StylesMut};

use super::model::MutableDocument;

impl MutableDocument {
    /// Borrow the document's content layer for semantic reads.
    #[cfg(test)]
    pub(crate) fn content(&self) -> Content<'_> {
        Content { document: self }
    }

    /// Borrow the document's content layer for structural and inline edits.
    #[cfg(test)]
    pub(crate) fn content_mut(&mut self) -> ContentMut<'_> {
        ContentMut { document: self }
    }

    /// Borrow the retained `styles.xml` layer for semantic reads.
    #[cfg(test)]
    pub(crate) fn styles(&self) -> Styles<'_> {
        Styles { document: self }
    }

    /// Borrow the retained `styles.xml` layer for targeted style edits.
    #[cfg(test)]
    pub(crate) fn styles_mut(&mut self) -> StylesMut<'_> {
        StylesMut { document: self }
    }
}
