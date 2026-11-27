use super::super::prelude::*;

impl Document {
    /// field state. This method never parses or evaluates an expression,
    /// resolves field values, or refreshes a field result.
    pub fn if_fields(&self) -> Result<Vec<IfField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::if_field).collect())
    }

    /// Get the number of typed, inert `IF` fields.
    pub fn if_field_count(&self) -> Result<usize> {
        Ok(self.if_fields()?.len())
    }

    /// Get typed, inert `COMPARE` fields in story and source order.
    ///
    /// Returned values expose only stored comparison text, cached results, and
    /// field state. This method never parses or evaluates a comparison,
    /// resolves nested field values, or refreshes a field result.
    pub fn compare_fields(&self) -> Result<Vec<CompareField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::compare_field).collect())
    }

    /// Get the number of typed, inert `COMPARE` fields.
    pub fn compare_field_count(&self) -> Result<usize> {
        Ok(self.compare_fields()?.len())
    }

    /// Get typed, inert `ASK` and `FILLIN` fields in story and source order.
    ///
    /// Returned values expose only stored prompt, bookmark, default-response,
    /// cached-result, and field-state metadata. This method never displays a
    /// prompt, captures a response, creates or updates a bookmark, performs a
    /// merge, or refreshes a field result.
    pub fn prompt_fields(&self) -> Result<Vec<PromptField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::prompt_field).collect())
    }

    /// Get the number of typed, inert `ASK` and `FILLIN` fields.
    pub fn prompt_field_count(&self) -> Result<usize> {
        Ok(self.prompt_fields()?.len())
    }

    /// Get typed, inert user-identity fields in story and source order.
    ///
    /// Returned values expose only stored kind, override, formatting, cached
    /// result, and field state. This method never reads or modifies a host
    /// user's identity, applies formatting, or refreshes a field.
    pub fn user_identity_fields(&self) -> Result<Vec<UserIdentityField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::user_identity_field)
            .collect())
    }

    /// Get the number of typed, inert user-identity fields.
    pub fn user_identity_field_count(&self) -> Result<usize> {
        Ok(self.user_identity_fields()?.len())
    }

    /// Get typed, inert `ADVANCE` fields in story and source order.
    ///
    /// Returned values expose only stored point adjustments, cached results,
    /// and field state. This method never moves text, changes layout, reflows
    /// content, or refreshes a field.
    pub fn advance_fields(&self) -> Result<Vec<AdvanceField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::advance_field).collect())
    }

    /// Get the number of typed, inert `ADVANCE` fields.
    pub fn advance_field_count(&self) -> Result<usize> {
        Ok(self.advance_fields()?.len())
    }
}
