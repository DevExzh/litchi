use super::super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Headers / Footers
    // ──────────────────────────────────────────────────────────────────

    /// Get all headers and footers in the document.
    ///
    /// Each section can have up to six stories: first-page header/footer,
    /// even-page header/footer, and odd-page (default) header/footer.
    /// Empty stories (where `start_cp` == `end_cp`) are omitted.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for hf in doc.headers_footers()? {
    ///     println!("{:?}: {}", hf.header_footer_type, hf.text());
    /// }
    /// ```
    pub fn headers_footers(&self) -> Result<Vec<HeaderFooter>> {
        let table = match &self.headers_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::new();
        for story in table.stories() {
            if story.is_empty() {
                continue;
            }

            let text = self
                .text_extractor
                .text_at_range(story.start_cp, story.end_cp)
                .to_string();

            // Extract paragraphs for this header/footer range
            let paragraphs = self.extract_paragraphs_for_range(story.start_cp, story.end_cp)?;

            let mut hf = HeaderFooter::new(story.story_type, text);
            hf.paragraphs = paragraphs;
            result.push(hf);
        }

        Ok(result)
    }

    /// Get only headers (filtering out footers).
    pub fn headers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(HeaderFooter::is_header)
            .collect())
    }

    /// Get only footers (filtering out headers).
    pub fn footers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(HeaderFooter::is_footer)
            .collect())
    }
}
