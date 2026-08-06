use super::super::prelude::*;

impl Document {
    /// Get typed, inert `TOC` fields in story and source order.
    ///
    /// Returned values expose only stored configuration, unrecognized switches,
    /// cached results, and field state. This method never scans entries, reads
    /// bookmarks, resolves links, calculates page numbers, paginates,
    /// regenerates a table of contents, or refreshes a field.
    pub fn table_of_contents_fields(&self) -> Result<Vec<TableOfContentsField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::table_of_contents)
            .collect())
    }

    /// Get the number of typed, inert `TOC` fields.
    pub fn table_of_contents_field_count(&self) -> Result<usize> {
        Ok(self.table_of_contents_fields()?.len())
    }

    /// Get typed, inert `TC` table-of-contents entry fields in story and source
    /// order.
    ///
    /// Native Word omits `TC` marker characters from `Plcfld` metadata, so this
    /// method scans only the stored text of each document story. Returned values
    /// expose stored entries, switches, cached results, and source positions.
    /// This method never changes hidden text, calculates page numbers, generates
    /// a table of contents, or refreshes a field.
    pub fn table_of_contents_entries(&self) -> Result<Vec<TableOfContentsEntryField>> {
        let mut entries = Vec::new();
        for story in FieldStory::ALL {
            let Some((start, end)) = self.field_story_range_if_present(story) else {
                continue;
            };
            let text = self.text_extractor.text_at_range(start, end);
            entries.extend(
                non_plcf_field_texts(story, text)
                    .iter()
                    .filter_map(TableOfContentsEntryField::from_non_plcf_field),
            );
        }
        Ok(entries)
    }

    /// Get the number of typed, inert `TC` table-of-contents entry fields.
    pub fn table_of_contents_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_contents_entries()?.len())
    }

    /// Get typed, inert `TOA` fields in story and source order.
    ///
    /// Returned values expose only stored configuration, unrecognized switches,
    /// cached results, and field state. This method never finds citations,
    /// scans hidden text, reads bookmarks, follows links, calculates page
    /// numbers, paginates, regenerates a table of authorities, or refreshes a
    /// field.
    pub fn table_of_authorities_fields(&self) -> Result<Vec<TableOfAuthoritiesField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::table_of_authorities)
            .collect())
    }

    /// Get the number of typed, inert `TOA` fields.
    pub fn table_of_authorities_field_count(&self) -> Result<usize> {
        Ok(self.table_of_authorities_fields()?.len())
    }

    /// Get typed, inert `TA` table-of-authorities entry fields in story and
    /// source order.
    ///
    /// Native Word omits `TA` marker characters from `Plcfld` metadata, so this
    /// method scans only the stored text of each document story. Returned values
    /// expose stored switches, cached results, and source positions. This method
    /// never finds citations, changes hidden text, follows bookmarks, calculates
    /// page numbers, generates a table of authorities, or refreshes a field.
    pub fn table_of_authorities_entries(&self) -> Result<Vec<TableOfAuthoritiesEntryField>> {
        let mut entries = Vec::new();
        for story in FieldStory::ALL {
            let Some((start, end)) = self.field_story_range_if_present(story) else {
                continue;
            };
            let text = self.text_extractor.text_at_range(start, end);
            entries.extend(
                non_plcf_field_texts(story, text)
                    .iter()
                    .filter_map(TableOfAuthoritiesEntryField::from_non_plcf_field),
            );
        }
        Ok(entries)
    }

    /// Get the number of typed, inert `TA` table-of-authorities entry fields.
    pub fn table_of_authorities_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_authorities_entries()?.len())
    }

    /// Get typed, inert generated-index (`INDEX`) fields in story and source order.
    ///
    /// Returned values expose only stored configuration, unrecognized switches,
    /// cached results, and field state. This method never scans index markers,
    /// reads bookmarks, calculates page numbers, sorts entries, paginates,
    /// generates an index, or refreshes a field.
    pub fn indexes(&self) -> Result<Vec<IndexField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::index).collect())
    }

    /// Get the number of typed, inert generated-index (`INDEX`) fields.
    pub fn index_count(&self) -> Result<usize> {
        Ok(self.indexes()?.len())
    }

    /// Get typed, inert `XE` index-entry fields in story and source order.
    ///
    /// Native Word omits `XE` marker characters from `Plcfld` metadata, so this
    /// method scans only the stored text of each document story. Returned values
    /// expose stored entries, switches, cached results, and source positions.
    /// This method never changes hidden text, resolves bookmarks, calculates
    /// page numbers, sorts entries, generates an index, or refreshes a field.
    pub fn index_entries(&self) -> Result<Vec<IndexEntryField>> {
        let mut entries = Vec::new();
        for story in FieldStory::ALL {
            let Some((start, end)) = self.field_story_range_if_present(story) else {
                continue;
            };
            let text = self.text_extractor.text_at_range(start, end);
            entries.extend(
                non_plcf_field_texts(story, text)
                    .iter()
                    .filter_map(IndexEntryField::from_non_plcf_field),
            );
        }
        Ok(entries)
    }

    /// Get the number of typed, inert `XE` index-entry fields.
    pub fn index_entry_count(&self) -> Result<usize> {
        Ok(self.index_entries()?.len())
    }
}
