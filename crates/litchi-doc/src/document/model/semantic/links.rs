use super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Hyperlinks
    // ──────────────────────────────────────────────────────────────────

    /// Get all hyperlinks in the document.
    ///
    /// Hyperlinks are extracted from HYPERLINK fields in the main document.
    /// Each hyperlink includes the legacy destination URL/path, display text,
    /// and type. For stored field metadata from every field story, use
    /// `hyperlink_fields()`.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for link in doc.hyperlinks()? {
    ///     println!("{} -> {}", link.display_text(), link.destination());
    /// }
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<Hyperlink>> {
        let table = match &self.hyperlinks_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        Ok(table
            .hyperlinks()
            .iter()
            .map(Hyperlink::from_internal)
            .collect())
    }

    /// Find hyperlinks at a specific character position in the document.
    pub fn hyperlinks_at_position(&self, cp: u32) -> Vec<Hyperlink> {
        match &self.hyperlinks_table {
            Some(t) => t
                .find_at_position(cp)
                .into_iter()
                .map(Hyperlink::from_internal)
                .collect(),
            None => Vec::new(),
        }
    }
}
