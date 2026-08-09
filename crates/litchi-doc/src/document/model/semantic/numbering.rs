use super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Numbering / Lists
    // ──────────────────────────────────────────────────────────────────

    /// Get the list tables (list definitions and overrides).
    ///
    /// Use this to look up list formatting for individual paragraphs
    /// via their `list_format_override` and `list_level` properties.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(tables) = doc.list_tables() {
    ///     for para in doc.paragraphs()? {
    ///         if let Some(info) = doc.paragraph_list_info(&para) {
    ///             println!("Level {}: {:?}", info.level, info.number_format);
    ///         }
    ///     }
    /// }
    /// ```
    #[must_use]
    pub fn list_tables(&self) -> Option<&ListTables> {
        self.list_tables.as_ref()
    }

    /// Resolve a non-empty `LISTNUM` name by zero-based `PlfLst` definition index.
    ///
    /// Entries beyond the list-definition array are ignored as required by `[MS-DOC]`.
    #[must_use]
    pub fn list_name_for_definition_index(&self, index: usize) -> Option<&str> {
        let definition_count = self.list_tables.as_ref()?.structures().len();
        if index >= definition_count {
            return None;
        }
        self.list_names.as_ref()?.name(index)
    }

    /// Get list/numbering information for a specific paragraph.
    ///
    /// Returns `Some(ListLevel)` if the paragraph is part of a list,
    /// `None` otherwise. Any `LFOLVL` start-at or formatting overrides
    /// attached to the paragraph's LFO are applied to the result.
    #[must_use]
    pub fn paragraph_list_info(
        &self,
        paragraph: &Paragraph,
    ) -> Option<crate::parts::numbering::ListLevel> {
        let binding = self.paragraph_list_binding(paragraph)?;
        let mut level = binding.effective_level().clone();
        level.start_at = binding.effective_start_at();
        Some(level)
    }

    /// Resolve typed list metadata for a paragraph without cloning list data.
    ///
    /// The returned binding exposes the selected `LSTF`, `LFO`, base `LVL`,
    /// optional `LFOLVL`, effective formatting, start value, and the
    /// preserve-indents bit encoded by a negative `sprmPIlfo`.
    #[must_use]
    pub fn paragraph_list_binding(
        &self,
        paragraph: &Paragraph,
    ) -> Option<ParagraphListBinding<'_>> {
        let properties = paragraph.properties();
        let signed_lfo = properties.list_format_override?;
        let level = properties.list_level.unwrap_or(0);
        self.list_tables.as_ref()?.bind_paragraph(signed_lfo, level)
    }
}
