use super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Internal helpers for subdocument extraction
    // ──────────────────────────────────────────────────────────────────

    /// Extract paragraphs for a specific character position range.
    ///
    /// Used internally to get paragraphs for subdocuments like
    /// headers, footers, footnotes, and endnotes.
    pub(super) fn extract_paragraphs_for_range(
        &self,
        start_cp: u32,
        end_cp: u32,
    ) -> Result<Vec<Paragraph>> {
        if start_cp >= end_cp {
            return Ok(Vec::new());
        }

        let text = self.text_extractor.text_shared();
        let text_ranges = self.text_extractor.cp_to_byte_shared();

        let para_extractor = ParagraphExtractor::new_with_range_and_stylesheet_and_shared_ranges(
            Arc::clone(&text),
            Arc::clone(&text_ranges),
            self.pap_bin_table.as_ref(),
            self.chp_bin_table.as_ref(),
            (start_cp, end_cp),
            self.stylesheet.as_ref(),
        )?;

        let extracted = para_extractor.extract_paragraphs()?;
        let mut paragraphs = Vec::with_capacity(extracted.len());
        self.convert_to_paragraphs(extracted, &mut paragraphs)?;
        Ok(paragraphs)
    }
}
