//! Field refresh primitives on the package-bound edit (record 0279).

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
        self.inner.fields()
    }

    /// Apply `updates` planned over [`Self::fields`] of this edit; see
    /// [`crate::document::fields`].
    ///
    /// # Errors
    ///
    /// As [`super::super::snapshot::Edit::refresh_fields`]. Source-backed
    /// edits refuse it as an unmanaged-only operation.
    pub fn refresh_fields(&mut self, fields: &Fields, updates: &[Update]) -> Result<Outcome> {
        self.inner.refresh_fields(fields, updates)
    }
}
