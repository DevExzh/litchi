use super::super::prelude::*;

impl Document {
    /// Get access to the fields table (if parsed).
    ///
    /// Contains information about all fields in the main document,
    /// including embedded objects and hyperlinks.
    #[inline]
    pub fn fields_table(&self) -> Option<&FieldsTable> {
        self.fields_table.as_ref()
    }

    /// Get stored instruction and cached-result text for every field story.
    ///
    /// The returned text follows the field-range rules in MS-DOC section
    /// 2.8.25. It is read from the document's existing text only: fields are
    /// never evaluated or refreshed, DDE conversations are never started,
    /// external paths are never opened, OLE objects are never activated, and
    /// macro instructions are never resolved, loaded, or executed.
    pub fn fields(&self) -> Result<Vec<FieldText>> {
        let Some(fields) = &self.fields_table else {
            return Ok(Vec::new());
        };

        fields.field_texts(|story, start, end| self.field_story_text(story, start, end))
    }
    /// Get stored instruction and cached-result text for one parsed field.
    ///
    /// Field positions are relative to their FieldStory. This method reads only
    /// that stored text range and performs no field evaluation or external
    /// action.
    pub fn field_text(&self, field: &Field) -> Result<FieldText> {
        FieldText::from_field(field, |start, end| {
            self.field_story_text(field.story, start, end)
        })
    }
}
