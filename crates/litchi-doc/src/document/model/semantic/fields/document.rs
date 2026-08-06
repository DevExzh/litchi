use super::super::prelude::*;

impl Document {
    /// Get typed, inert `MERGEFIELD` fields in story and source order.
    ///
    /// Returned values expose only stored data-column names, switches, cached
    /// results, and field state. This method never opens a data source, resolves
    /// records, performs a merge, or refreshes a field result.
    pub fn merge_fields(&self) -> Result<Vec<MergeField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::merge_field).collect())
    }

    /// Get the number of typed, inert `MERGEFIELD` fields.
    pub fn merge_field_count(&self) -> Result<usize> {
        Ok(self.merge_fields()?.len())
    }

    /// Get typed, inert `DATA` mail-merge source fields in story and source order.
    ///
    /// Returned values expose only stored data-source, header-source, switch,
    /// cached-result, and field-state metadata. This method never opens, reads,
    /// connects to, resolves, or modifies a source; it never selects a record,
    /// performs a merge, or refreshes a field result.
    pub fn mail_merge_data_fields(&self) -> Result<Vec<MailMergeDataField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_data)
            .collect())
    }

    /// Get the number of typed, inert `DATA` mail-merge source fields.
    pub fn mail_merge_data_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_data_fields()?.len())
    }

    /// Get typed, inert `DOCVARIABLE` fields in story and source order.
    ///
    /// Returned values expose only stored variable names, switches, cached
    /// results, and field state. This method never reads document variables,
    /// resolves a value, or refreshes a field result.
    pub fn document_variable_fields(&self) -> Result<Vec<DocumentVariableField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::document_variable)
            .collect())
    }

    /// Get the number of typed, inert `DOCVARIABLE` fields.
    pub fn document_variable_field_count(&self) -> Result<usize> {
        Ok(self.document_variable_fields()?.len())
    }

    /// Get typed, inert `DOCPROPERTY` fields in story and source order.
    ///
    /// Returned values expose only stored property names, switches, cached
    /// results, and field state. This method never reads document properties,
    /// resolves a value, or refreshes a field result.
    pub fn document_property_fields(&self) -> Result<Vec<DocumentPropertyField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::document_property)
            .collect())
    }

    /// Get the number of typed, inert `DOCPROPERTY` fields.
    pub fn document_property_field_count(&self) -> Result<usize> {
        Ok(self.document_property_fields()?.len())
    }

    /// Get typed, inert native `INFO` fields in story and source order.
    ///
    /// Returned values expose only stored property selectors, optional
    /// replacement values, switches, cached results, and field state. This
    /// method never reads, resolves, modifies, or writes document or template
    /// properties, or refreshes a field.
    pub fn info_fields(&self) -> Result<Vec<InfoField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::info_field).collect())
    }

    /// Get the number of typed, inert native `INFO` fields.
    pub fn info_field_count(&self) -> Result<usize> {
        Ok(self.info_fields()?.len())
    }

    /// Get typed, inert built-in document-information fields in story and
    /// source order.
    ///
    /// Returned values expose only the native category, stored switches,
    /// cached results, and field state. This method never reads document
    /// properties or host identity data, calculates dates, revisions, or
    /// statistics, resolves values, or refreshes a field result.
    pub fn document_information_fields(&self) -> Result<Vec<DocumentInformationField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::document_information)
            .collect())
    }

    /// Get the number of typed, inert built-in document-information fields.
    pub fn document_information_field_count(&self) -> Result<usize> {
        Ok(self.document_information_fields()?.len())
    }

    /// Get typed, inert built-in document-context and runtime fields in story
    /// and source order.
    ///
    /// Returned values expose only the native category, stored switches,
    /// cached results, and field state. This method never reads a document
    /// path, attached template, host filesystem state or file size, current
    /// clock, or page and section layout, resolves values, or refreshes a field
    /// result.
    pub fn document_context_fields(&self) -> Result<Vec<DocumentContextField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::document_context)
            .collect())
    }

    /// Get the number of typed, inert built-in document-context and runtime
    /// fields.
    pub fn document_context_field_count(&self) -> Result<usize> {
        Ok(self.document_context_fields()?.len())
    }

    /// Get typed, inert `DDE` and `DDEAUTO` fields in story and source order.
    ///
    /// Returned values expose only stored application, source, item, switch,
    /// cached-result, and field-state metadata. This method never launches an
    /// application, initiates a DDE conversation, opens a source, requests
    /// data, refreshes content, converts content, or executes code.
    pub fn dde_links(&self) -> Result<Vec<DdeField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::dde_link).collect())
    }

    /// Get the number of typed, inert `DDE` and `DDEAUTO` fields.
    pub fn dde_link_count(&self) -> Result<usize> {
        Ok(self.dde_links()?.len())
    }

    /// Get typed, inert `LINK` fields in story and source order.
    ///
    /// Returned values expose only stored application type, source, item,
    /// switch, cached-result, and field-state metadata. This method never
    /// activates an OLE server, launches an application, opens a source,
    /// requests data, refreshes content, converts content, or executes code.
    pub fn link_fields(&self) -> Result<Vec<LinkField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::link_field).collect())
    }

    /// Get the number of typed, inert `LINK` fields.
    pub fn link_field_count(&self) -> Result<usize> {
        Ok(self.link_fields()?.len())
    }

    /// Get typed, inert external-include fields in story and source order.
    ///
    /// Returned values cover `INCLUDETEXT`/`INCLUDEPICTURE` and their historical
    /// `INCLUDE`/`IMPORT` aliases. They expose only stored source, bookmark,
    /// converter, XML-option, cached-result, and field-state metadata. This
    /// method never opens, resolves, imports, fetches, refreshes, transforms,
    /// converts, evaluates, or executes an external source.
    pub fn external_includes(&self) -> Result<Vec<ExternalIncludeField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::external_include)
            .collect())
    }

    /// Get the number of typed, inert external-include fields.
    pub fn external_include_count(&self) -> Result<usize> {
        Ok(self.external_includes()?.len())
    }

    /// Get typed, inert `RD` referenced-document fields in story and source order.
    ///
    /// Native Word omits `RD` marker characters from `Plcfld` metadata, so this
    /// method scans only the stored text of each document story. Returned values
    /// expose stored sources, relative-path requests, switches, cached results,
    /// and source positions. This method never opens, resolves, reads, imports,
    /// refreshes, evaluates, or executes a referenced document.
    pub fn referenced_documents(&self) -> Result<Vec<ReferencedDocumentField>> {
        let mut references = Vec::new();
        for story in FieldStory::ALL {
            let Some((start, end)) = self.field_story_range_if_present(story) else {
                continue;
            };
            let text = self.text_extractor.text_at_range(start, end);
            references.extend(
                non_plcf_field_texts(story, text)
                    .iter()
                    .filter_map(ReferencedDocumentField::from_non_plcf_field),
            );
        }
        Ok(references)
    }

    /// Get the number of typed, inert `RD` referenced-document fields.
    pub fn referenced_document_count(&self) -> Result<usize> {
        Ok(self.referenced_documents()?.len())
    }

    /// Get typed, inert `PRIVATE` conversion-data fields in story and source order.
    ///
    /// Native Word omits `PRIVATE` marker characters from `Plcfld` metadata, so
    /// this method scans only the stored text of each document story. Returned
    /// values expose opaque instructions, cached results, and source positions.
    /// This method never converts a document, interprets field data, reveals
    /// hidden content, changes layout, or refreshes a field. `PRIVATE` is not
    /// treated as a confidentiality mechanism.
    pub fn private_fields(&self) -> Result<Vec<PrivateField>> {
        let mut private_fields = Vec::new();
        for story in FieldStory::ALL {
            let Some((start, end)) = self.field_story_range_if_present(story) else {
                continue;
            };
            let text = self.text_extractor.text_at_range(start, end);
            private_fields.extend(
                non_plcf_field_texts(story, text)
                    .iter()
                    .filter_map(PrivateField::from_non_plcf_field),
            );
        }
        Ok(private_fields)
    }

    /// Get the number of typed, inert `PRIVATE` conversion-data fields.
    pub fn private_field_count(&self) -> Result<usize> {
        Ok(self.private_fields()?.len())
    }
}
