//! Field refresh primitives on the package-bound edit (record 0279).

use litchi_core::Position;

use super::super::Result;
use super::super::transaction::fields::{Fields, Outcome, Update};
use super::Edit;

impl<P: Clone> Edit<P> {
    /// Index the paragraphs, fields and bookmarks of the staged main
    /// document.
    ///
    /// ```
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Plain");
    /// let package = builder.build()?;
    /// let edit = package.edit()?;
    /// let fields = edit.fields()?;
    /// assert_eq!(fields.paragraphs()[0].text(), "Plain");
    /// assert!(fields.fields().is_empty());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an XML error for malformed markup.
    pub fn fields(&self) -> Result<Fields> {
        self.inner_ref(|inner| inner.fields())?
    }

    /// Apply `updates` planned over [`Self::fields`] of this edit; see
    /// [`crate::document::fields`].
    ///
    /// # Errors
    ///
    /// As [`super::super::raw::Tree::refresh_fields`]. Source-backed
    /// edits refuse it as an unmanaged-only operation.
    pub fn refresh_fields(&mut self, fields: &Fields, updates: &[Update]) -> Result<Outcome> {
        self.inner_mut()?.refresh_fields(fields, updates)
    }

    /// Insert a bookmark into the paragraph at `paragraph`; see
    /// [`super::super::raw::Tree::insert_bookmark_at`] (record 0325).
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Before");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let id = edit.insert_bookmark_at(&paragraph, None, "Target", Some("here"))?;
    /// assert_eq!(edit.fields()?.bookmarks()[0].id(), Some(id));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`super::super::raw::Tree::insert_bookmark_at`].
    pub fn insert_bookmark_at(
        &mut self,
        paragraph: &super::super::ElementPath,
        at: Option<Position>,
        name: &str,
        text: Option<&str>,
    ) -> Result<i64> {
        self.inner_mut()?
            .insert_bookmark_at(paragraph, at, name, text)
    }

    /// Insert a complex field into the paragraph at `paragraph`; see
    /// [`super::super::raw::Tree::insert_field_at`] (record 0325).
    ///
    /// # Errors
    ///
    /// As [`super::super::raw::Tree::insert_field_at`].
    pub fn insert_field_at(
        &mut self,
        paragraph: &super::super::ElementPath,
        at: Option<Position>,
        instruction: &str,
        result: Option<&str>,
    ) -> Result<Position> {
        self.inner_mut()?
            .insert_field_at(paragraph, at, instruction, result)
    }
}
