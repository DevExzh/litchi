//! Contextual semantic views over a mutable ODT document.
//!
//! `MutableDocument` retains its original flat methods for compatibility. These
//! views provide a small, discoverable layer for code that wants to keep
//! structural content work separate from style/package work:
//!
//! ```no_run
//! # use litchi_odt::mutable::MutableDocument;
//! # fn main() -> litchi_core::Result<()> {
//! let mut document = MutableDocument::new();
//! document.content_mut().add_paragraph("body")?;
//! document
//!     .content_mut()
//!     .append_line_break_at(litchi_core::Position::new(0))?;
//! document.styles_mut().add_master_page("Standard", "pm1")?;
//! # Ok(())
//! # }
//! ```

mod content;
mod styles;
mod text;

pub use content::{Content, ContentMut};
pub use styles::{Styles, StylesMut};

use super::model::MutableDocument;

impl MutableDocument {
    /// Borrow the document's content layer for semantic reads.
    pub fn content(&self) -> Content<'_> {
        Content { document: self }
    }

    /// Borrow the document's content layer for structural and inline edits.
    pub fn content_mut(&mut self) -> ContentMut<'_> {
        ContentMut { document: self }
    }

    /// Borrow the retained `styles.xml` layer for semantic reads.
    pub fn styles(&self) -> Styles<'_> {
        Styles { document: self }
    }

    /// Borrow the retained `styles.xml` layer for targeted style edits.
    pub fn styles_mut(&mut self) -> StylesMut<'_> {
        StylesMut { document: self }
    }
}
