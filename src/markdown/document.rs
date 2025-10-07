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

        // TODO: Add metadata support when Document metadata API is available
        // if options.include_metadata {
        //     writer.write_metadata(...)?;
        // }

        // Write paragraphs
        for para in self.paragraphs()? {
            writer.write_paragraph(&para)?;
        }

        // Write tables
        for table in self.tables()? {
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

