use super::super::prelude::*;
use crate::parts::footnotes::FootnoteReference;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Footnotes / Endnotes
    // ──────────────────────────────────────────────────────────────────

    /// Get all footnotes in the document.
    ///
    /// Each footnote contains its reference position in the main document,
    /// the footnote number, and the footnote text with paragraphs.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for note in doc.footnotes()? {
    ///     println!("Footnote {}: {}", note.number, note.text());
    /// }
    /// ```
    pub fn footnotes(&self) -> Result<Vec<Footnote>> {
        let references = self.footnote_references();
        let mut result = Vec::with_capacity(references.len());
        for reference in references {
            result.push(self.build_note(reference)?);
        }
        Ok(result)
    }

    /// The number of footnotes, in main-document reference order.
    #[must_use]
    pub fn footnote_count(&self) -> usize {
        self.footnote_references().len()
    }

    /// The footnote at zero-based `index` in main-document reference order,
    /// built alone: the value [`Self::footnotes`] returns at that index,
    /// without building the others.
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// for index in 0..document.footnote_count() {
    ///     if let Some(note) = document.footnote(Position::new(index))? {
    ///         println!("{}: {}", note.number, note.text());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the errors [`Self::footnotes`] returns for this footnote.
    pub fn footnote(&self, index: Position) -> Result<Option<Footnote>> {
        self.footnote_references()
            .get(index.get())
            .map(|reference| self.build_note(reference))
            .transpose()
    }

    /// Get all endnotes in the document.
    ///
    /// Endnotes share the same structure as footnotes but are placed
    /// at the end of the document or section.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for note in doc.endnotes()? {
    ///     println!("Endnote {}: {}", note.number, note.text());
    /// }
    /// ```
    pub fn endnotes(&self) -> Result<Vec<Footnote>> {
        let references = self.endnote_references();
        let mut result = Vec::with_capacity(references.len());
        for reference in references {
            result.push(self.build_note(reference)?);
        }
        Ok(result)
    }

    /// The number of endnotes, in main-document reference order.
    #[must_use]
    pub fn endnote_count(&self) -> usize {
        self.endnote_references().len()
    }

    /// The endnote at zero-based `index` in main-document reference order,
    /// built alone: the value [`Self::endnotes`] returns at that index.
    ///
    /// # Errors
    ///
    /// Returns the errors [`Self::endnotes`] returns for this endnote.
    pub fn endnote(&self, index: Position) -> Result<Option<Footnote>> {
        self.endnote_references()
            .get(index.get())
            .map(|reference| self.build_note(reference))
            .transpose()
    }

    fn footnote_references(&self) -> &[FootnoteReference] {
        self.0
            .footnotes_table
            .as_ref()
            .map_or(&[], |table| table.references())
    }

    fn endnote_references(&self) -> &[FootnoteReference] {
        self.0
            .endnotes_table
            .as_ref()
            .map_or(&[], |table| table.references())
    }

    fn build_note(&self, reference: &FootnoteReference) -> Result<Footnote> {
        let text = self
            .0
            .text_extractor
            .text_at_range(reference.text_start_cp, reference.text_end_cp)
            .to_string();
        let paragraphs =
            self.extract_paragraphs_for_range(reference.text_start_cp, reference.text_end_cp)?;
        let mut note = Footnote::new(reference.ref_cp, reference.descriptor.number, text);
        note.paragraphs = paragraphs;
        Ok(note)
    }
}
