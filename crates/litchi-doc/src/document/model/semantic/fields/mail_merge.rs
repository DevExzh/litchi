use super::super::prelude::*;

impl Document {
    /// Get typed, inert `MERGEREC` and `MERGESEQ` fields in story and source order.
    ///
    /// Returned values expose only stored kinds, cached results, and field
    /// state. This method never selects or counts records, opens a data source,
    /// performs a merge, or refreshes a field result.
    pub fn mail_merge_counters(&self) -> Result<Vec<MailMergeCounterField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_counter)
            .collect())
    }

    /// Get the number of typed, inert mail-merge counter fields.
    pub fn mail_merge_counter_count(&self) -> Result<usize> {
        Ok(self.mail_merge_counters()?.len())
    }

    /// Get typed, inert `NEXT` mail-merge control fields in story and source order.
    ///
    /// Returned values expose only stored cached results and field state. This
    /// method never advances a record, opens a data source, performs a merge,
    /// or refreshes a field result.
    pub fn mail_merge_next_fields(&self) -> Result<Vec<MailMergeNextField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_next)
            .collect())
    }

    /// Get the number of typed, inert `NEXT` mail-merge control fields.
    pub fn mail_merge_next_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_next_fields()?.len())
    }

    /// Get typed, inert `NEXTIF` and `SKIPIF` fields in story and source order.
    ///
    /// Returned values expose only stored comparison text, cached results, and
    /// field state. This method never evaluates a comparison, changes record
    /// selection, opens a data source, performs a merge, or refreshes a field
    /// result.
    pub fn mail_merge_conditional_controls(&self) -> Result<Vec<MailMergeConditionalControlField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_conditional_control)
            .collect())
    }

    /// Get the number of typed, inert conditional mail-merge control fields.
    pub fn mail_merge_conditional_control_count(&self) -> Result<usize> {
        Ok(self.mail_merge_conditional_controls()?.len())
    }

    /// order.
    ///
    /// Returned values expose stored recipient layout, locale, country, fallback,
    /// cached-result, and field-state metadata only. This method never opens a
    /// data source, selects a record, performs a merge, expands placeholders,
    /// generates text, or refreshes a field result.
    pub fn mail_merge_recipient_fields(&self) -> Result<Vec<MailMergeRecipientField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_recipient_field)
            .collect())
    }

    /// Get the number of typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields.
    pub fn mail_merge_recipient_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_recipient_fields()?.len())
    }
}
