use super::super::prelude::*;

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
        let table = match &self.footnotes_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::with_capacity(table.count());
        for reference in table.references() {
            let text = self
                .text_extractor
                .text_at_range(reference.text_start_cp, reference.text_end_cp)
                .to_string();

            let paragraphs =
                self.extract_paragraphs_for_range(reference.text_start_cp, reference.text_end_cp)?;

            let mut note = Footnote::new(reference.ref_cp, reference.descriptor.number, text);
            note.paragraphs = paragraphs;
            result.push(note);
        }

        Ok(result)
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
        let table = match &self.endnotes_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::with_capacity(table.count());
        for reference in table.references() {
            let text = self
                .text_extractor
                .text_at_range(reference.text_start_cp, reference.text_end_cp)
                .to_string();

            let paragraphs =
                self.extract_paragraphs_for_range(reference.text_start_cp, reference.text_end_cp)?;

            let mut note = Footnote::new(reference.ref_cp, reference.descriptor.number, text);
            note.paragraphs = paragraphs;
            result.push(note);
        }

        Ok(result)
    }
}
