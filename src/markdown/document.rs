use super::config::MarkdownOptions;
use super::traits::ToMarkdown;
use super::writer::MarkdownWriter;
/// ToMarkdown implementations for Document types.
///
/// This module implements the `ToMarkdown` trait for Word document types,
/// including Document, Paragraph, Run, and Table.
///
/// **Note**: This module is only available when the `ole` or `ooxml` feature is enabled.
use crate::common::Result;
use crate::document::{Document, Paragraph, Run, Table};

impl ToMarkdown for Document {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(*options);

        // Write metadata as YAML front matter if enabled
        if options.include_metadata {
            let metadata = self.metadata()?;
            writer.write_metadata(&metadata)?;
        }

        // Get all paragraphs once (avoid extracting twice)
        // Following Apache POI's design: extract paragraphs, then identify tables from them
        let paragraphs = self.paragraphs()?;

        // Extract tables from the already-extracted paragraphs
        // Note: tables() internally calls paragraphs() again which causes duplication
        // TODO: Refactor tables() to accept pre-extracted paragraphs to avoid double extraction
        let tables = self.tables()?;

        // For performance, pre-allocate based on estimated content size
        let estimated_size = paragraphs.len() * 100 + tables.len() * 500; // Rough estimates
        writer.reserve(estimated_size);

        // Write paragraphs (excluding those that are part of tables)
        // For now, write all paragraphs - table filtering can be added later
        for para in paragraphs {
            writer.write_paragraph(&para)?;
        }

        // Write tables
        for table in tables {
            writer.write_table(&table)?;
        }

        Ok(writer.finish())
    }
}

impl ToMarkdown for Paragraph {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(*options);
        writer.write_paragraph(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

impl ToMarkdown for Run {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(*options);
        writer.write_run(self)?;
        Ok(writer.finish())
    }
}

impl ToMarkdown for Table {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(*options);
        writer.write_table(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}
