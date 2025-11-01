use super::config::MarkdownOptions;
use super::traits::ToMarkdown;
use super::writer::MarkdownWriter;
/// ToMarkdown implementations for Presentation types.
///
/// This module implements the `ToMarkdown` trait for PowerPoint presentation types,
/// including Presentation and Slide.
///
/// **Note**: This module is only available when the `ole` or `ooxml` feature is enabled.
use crate::common::Result;
use crate::presentation::{Presentation, Slide};
use rayon::prelude::*;

/// Minimum number of slides to justify parallel processing overhead.
const PARALLEL_THRESHOLD: usize = 10;

impl ToMarkdown for Presentation {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        // Write metadata as YAML front matter if available and enabled
        let metadata_md = if options.include_metadata
            && let Some(metadata) = self.metadata()?
        {
            let mut metadata_writer = MarkdownWriter::new(*options);
            metadata_writer.write_metadata(&metadata)?;
            metadata_writer.finish()
        } else {
            String::new()
        };

        // Use optimized fast path that extracts text without shape parsing
        // This is significantly faster for PPT files (3-10x speedup)
        let slide_texts = self.extract_text_for_markdown()?;

        // Decide whether to use parallel or sequential processing
        let content_md = if options.use_parallel && slide_texts.len() >= PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process slides in parallel for large presentations
            let slide_count = slide_texts.len();
            let slide_strings: Vec<String> = slide_texts
                .into_par_iter()
                .map(|(slide_num, text)| {
                    let mut writer = MarkdownWriter::new(*options);

                    // Format slide header with first line as title
                    let first_line = text.lines().next().unwrap_or("");
                    let header_text = if first_line.is_empty() {
                        format!("# Slide {}", slide_num)
                    } else {
                        format!("# Slide {} {}", slide_num, first_line)
                    };

                    writer.push_str(&header_text);
                    writer.push_str("\n\n");

                    // Add slide content
                    if !text.is_empty() {
                        writer.push_str(&text);
                        writer.push_str("\n\n");
                    }

                    writer.finish()
                })
                .collect();

            // Estimate total size and pre-allocate
            let total_size: usize = slide_strings.iter().map(|s| s.len()).sum();
            let separator_size = slide_count.saturating_sub(1) * 8; // "\n\n---\n\n"
            let mut result = String::with_capacity(total_size + separator_size);

            // Concatenate slides in order with separators
            for (i, slide_md) in slide_strings.iter().enumerate() {
                if i > 0 {
                    result.push_str("\n\n---\n\n");
                }
                result.push_str(slide_md);
            }

            result
        } else {
            // SEQUENTIAL PATH: Process slides sequentially for small presentations
            let mut writer = MarkdownWriter::new(*options);

            for (i, (slide_num, text)) in slide_texts.iter().enumerate() {
                if i > 0 {
                    writer.push_str("\n\n---\n\n");
                }

                // Format slide header with first line as title
                let first_line = text.lines().next().unwrap_or("");
                let header_text = if first_line.is_empty() {
                    format!("# Slide {}", slide_num)
                } else {
                    format!("# Slide {} {}", slide_num, first_line)
                };

                writer.push_str(&header_text);
                writer.push_str("\n\n");

                // Add slide content
                if !text.is_empty() {
                    writer.push_str(text);
                    writer.push_str("\n\n");
                }
            }

            writer.finish()
        };

        // Combine metadata and content
        Ok(format!("{}{}", metadata_md, content_md))
    }
}

impl ToMarkdown for Slide {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        // For individual slides, just return the text
        // Formatting is minimal for presentations
        self.text()
    }
}
