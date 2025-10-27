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
        use crate::document::DocumentElement;

        // Write metadata first (must be sequential)
        let metadata_md = if options.include_metadata {
            let mut metadata_writer = MarkdownWriter::new(*options);
            let metadata = self.metadata()?;
            metadata_writer.write_metadata(&metadata)?;
            metadata_writer.finish()
        } else {
            String::new()
        };

        // Extract all document elements (paragraphs and tables) in document order
        // This is more efficient than calling paragraphs() and tables() separately,
        // and it maintains the correct order for Markdown conversion
        let elements = self.elements()?;

        // Process elements sequentially to maintain document order
        // Note: Parallel processing is not used here because we need to preserve
        // the exact order of elements in the document for correct Markdown output
        let mut writer = MarkdownWriter::new(*options);
        // Estimate: 100 bytes per paragraph, 500 bytes per table
        let estimated_size = elements.len() * 150; // Rough average
        writer.reserve(estimated_size);

        for element in elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    writer.write_paragraph(&para)?;
                },
                DocumentElement::Table(table) => {
                    writer.write_table(&table)?;
                },
            }
        }
        let content_md = writer.finish();

        // Combine metadata and content
        Ok(format!("{}{}", metadata_md, content_md))
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
