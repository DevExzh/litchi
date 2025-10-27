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
use rayon::prelude::*;

/// Minimum number of elements to justify parallel processing overhead.
const PARALLEL_THRESHOLD: usize = 50;

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
        let elements = self.elements()?;

        // Decide whether to use parallel or sequential processing
        let content_md = if options.use_parallel && elements.len() >= PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process elements in parallel for large documents
            // With Arc-based Send + Sync types, we can now safely parallelize
            let element_strings: Vec<String> = elements
                .par_iter()
                .map(|element| {
                    let mut writer = MarkdownWriter::new(*options);
                    match element {
                        DocumentElement::Paragraph(para) => {
                            let _ = writer.write_paragraph(para);
                        },
                        DocumentElement::Table(table) => {
                            let _ = writer.write_table(table);
                        },
                    }
                    writer.finish()
                })
                .collect();

            // Estimate total size and pre-allocate
            let total_size: usize = element_strings.iter().map(|s| s.len()).sum();
            let mut result = String::with_capacity(total_size);

            // Concatenate in document order
            for s in &element_strings {
                result.push_str(s);
            }

            result
        } else {
            // SEQUENTIAL PATH: Process elements sequentially for small documents
            // This avoids the parallelization overhead when it's not beneficial
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

            writer.finish()
        };

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
