/// ToMarkdown implementations for Document types.
///
/// This module implements the `ToMarkdown` trait for Word document types,
/// including Document, Paragraph, Run, and Table.
use crate::common::Result;
use crate::document::{Document, Paragraph, Run, Table};
use super::config::MarkdownOptions;
use super::traits::ToMarkdown;
use super::writer::MarkdownWriter;

impl ToMarkdown for Document {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());

        // Write metadata as YAML front matter if enabled
        if options.include_metadata {
            let metadata = self.metadata()?;
            writer.write_metadata(&metadata)?;
        }

        // Get all paragraphs and tables
        let paragraphs = self.paragraphs()?;
        let tables = self.tables()?;

        // For performance, pre-allocate based on estimated content size
        let estimated_size = paragraphs.len() * 100 + tables.len() * 500; // Rough estimates
        writer.reserve(estimated_size);

        // Write paragraphs
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
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_paragraph(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

impl ToMarkdown for Run {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_run(self)?;
        Ok(writer.finish())
    }
}

impl ToMarkdown for Table {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(options.clone());
        writer.write_table(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

