//! Field refresh primitives on the package-bound edit (record 0279).

use litchi_core::Position;

use super::super::Result;
use super::super::transaction::fields::{Fields, Outcome, Update, bookmark_markup, field_markup};
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

    /// Insert a bookmark into the paragraph at `paragraph`, before its
    /// direct run `at` (after its last child for `None`): its start marker,
    /// a plain run of `text` when given, and its end marker (record 0325).
    /// The bookmark takes the identifier after the largest the main story
    /// uses, which it returns.
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
    /// Returns a typed refusal for an empty name or one longer than 40
    /// characters, text a document cannot hold, a missing paragraph or run,
    /// or a position inside a field, and an XML error for malformed markup.
    /// The edit is unchanged on failure.
    pub fn insert_bookmark_at(
        &mut self,
        paragraph: &super::super::ElementPath,
        at: Option<Position>,
        name: &str,
        text: Option<&str>,
    ) -> Result<i64> {
        let id = self
            .fields()?
            .bookmarks()
            .iter()
            .filter_map(super::super::fields::Bookmark::id)
            .max()
            .map_or(0, |id| id.saturating_add(1).max(0));
        self.inner.insert_run_markup_at(paragraph, at, |prefix| {
            bookmark_markup(&qualifier(prefix), id, name, text)
        })?;
        Ok(id)
    }

    /// Insert a complex field (`begin`, `instrText`, `separate`, the result
    /// and `end` runs) into the paragraph at `paragraph`, before its direct
    /// run `at` (after its last child for `None`), and return the position
    /// of its begin run among the paragraph's direct runs (record 0325).
    /// `result` is the cached result; the instruction is written as given.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an empty instruction, text a document
    /// cannot hold, a missing paragraph or run, or a position inside a
    /// field. The edit is unchanged on failure.
    pub fn insert_field_at(
        &mut self,
        paragraph: &super::super::ElementPath,
        at: Option<Position>,
        instruction: &str,
        result: Option<&str>,
    ) -> Result<Position> {
        self.inner.insert_run_markup_at(paragraph, at, |prefix| {
            field_markup(&qualifier(prefix), &[], instruction, result)
        })
    }
}

/// A namespace prefix as the markup builders take it: `w:`, or empty.
fn qualifier(prefix: &[u8]) -> String {
    if prefix.is_empty() {
        String::new()
    } else {
        format!("{}:", String::from_utf8_lossy(prefix))
    }
}
