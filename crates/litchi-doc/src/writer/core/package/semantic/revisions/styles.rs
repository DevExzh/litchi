use crate::writer::core::{Writer, model::*};
impl Writer {
    pub(in crate::writer::core::package) fn validate_style_reference(
        &self,
        index: u16,
        expected_kind: crate::StyleKind,
        context: &str,
    ) -> Result<(), WriteError> {
        let actual_kind = match index {
            0 => Some(crate::StyleKind::Paragraph),
            10 => Some(crate::StyleKind::Character),
            15..=0x0FFC => self
                .styles
                .get(usize::from(index - 15))
                .map(|style| style.kind),
            _ => None,
        };
        let Some(actual_kind) = actual_kind else {
            return Err(WriteError::InvalidData(format!(
                "{context} references undefined DOC style index {index}"
            )));
        };
        if actual_kind != expected_kind {
            return Err(WriteError::InvalidData(format!(
                "{context} references {actual_kind:?} DOC style {index}, expected {expected_kind:?}"
            )));
        }
        Ok(())
    }

    pub(in crate::writer::core::package) fn validate_character_style_references(
        &self,
        formatting: &CharacterFormatting,
        context: &str,
    ) -> Result<(), WriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::StyleKind::Character, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_character_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(in crate::writer::core::package) fn validate_paragraph_style_references(
        &self,
        formatting: &ParagraphFormatting,
        context: &str,
    ) -> Result<(), WriteError> {
        if let Some(index) = formatting.style_index {
            self.validate_style_reference(index, crate::StyleKind::Paragraph, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_paragraph_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(in crate::writer::core::package) fn validate_table_style_references(
        &self,
        formatting: &crate::writer::tap::TableRow,
        context: &str,
    ) -> Result<(), WriteError> {
        if let Some(index) = formatting.table_style_index {
            self.validate_style_reference(index, crate::StyleKind::Table, context)?;
        }
        if let Some(previous) = &formatting.preserved_properties_for_revision {
            self.validate_table_style_references(previous, context)?;
        }
        Ok(())
    }

    pub(in crate::writer::core::package) fn validate_style_references(
        &self,
    ) -> Result<(), WriteError> {
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            self.validate_paragraph_style_references(
                &paragraph.formatting,
                "DOC paragraph formatting",
            )?;
            for run in &paragraph.runs {
                self.validate_character_style_references(
                    &run.formatting,
                    "DOC character formatting",
                )?;
            }
        }
        for table in &self.tables {
            for row in &table.rows {
                self.validate_table_style_references(&row.formatting, "DOC table row formatting")?;
            }
        }
        Ok(())
    }
}
