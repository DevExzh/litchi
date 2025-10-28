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

        let slides = self.slides()?;

        // Decide whether to use parallel or sequential processing
        let content_md = if options.use_parallel && slides.len() >= PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process slides in parallel for large presentations
            // Use into_par_iter() to consume and own the slides
            let slide_count = slides.len();
            let slide_strings: Vec<Result<String>> = slides
                .into_par_iter()
                .enumerate()
                .map(|(i, slide)| {
                    let mut writer = MarkdownWriter::new(*options);

                    // Format slide header with title placeholder if available
                    let slide_title = extract_slide_title(&slide)?;
                    let header_text = if slide_title.is_empty() {
                        format!("# Slide {}", i + 1)
                    } else {
                        format!("# Slide {} {}", i + 1, slide_title)
                    };

                    writer.write_fmt(format_args!("{}\n", header_text))?;
                    writer.push('\n');

                    // Add slide content with proper markdown formatting
                    write_slide_content(&mut writer, &slide, options)?;

                    Ok(writer.finish())
                })
                .collect();

            // Collect results and handle errors
            let slide_strings: Result<Vec<String>> = slide_strings.into_iter().collect();
            let slide_strings = slide_strings?;

            // Estimate total size and pre-allocate
            let total_size: usize = slide_strings.iter().map(|s| s.len()).sum();
            // Add space for separators between slides
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

            for (i, slide) in slides.iter().enumerate() {
                if i > 0 {
                    // Separate slides with horizontal rule
                    writer.push_str("\n\n---\n\n");
                }

                // Format slide header with title placeholder if available
                let slide_title = extract_slide_title(slide)?;
                let header_text = if slide_title.is_empty() {
                    format!("# Slide {}", i + 1)
                } else {
                    format!("# Slide {} {}", i + 1, slide_title)
                };

                writer.write_fmt(format_args!("{}\n", header_text))?;
                writer.push('\n');

                // Add slide content with proper markdown formatting
                write_slide_content(&mut writer, slide, options)?;
            }

            writer.finish()
        };

        // Combine metadata and content
        Ok(format!("{}{}", metadata_md, content_md))
    }
}

/// Extract the title from a slide by looking for title placeholders.
fn extract_slide_title(slide: &Slide) -> Result<String> {
    match slide {
        Slide::Ppt(_) => {
            // PPT slides don't have structured title extraction yet
            // Just use the first line of text as title
            let text = slide.text()?;
            let first_line = text.lines().next().unwrap_or("");
            Ok(first_line.to_string())
        },
        Slide::Pptx(pptx_data) => {
            // For PPTX slides, use the slide name if available
            // In a full implementation, we'd parse the slide content to find title placeholders
            Ok(pptx_data.name.clone().unwrap_or_default())
        },
        #[cfg(feature = "iwa")]
        Slide::Keynote(keynote_slide) => {
            // For Keynote slides, use the title if available
            Ok(keynote_slide.title.clone().unwrap_or_default())
        },
        #[cfg(feature = "odf")]
        Slide::Odp(_) => {
            // For ODP slides, use the first line of text as title
            let text = slide.text()?;
            let first_line = text.lines().next().unwrap_or("");
            Ok(first_line.to_string())
        },
    }
}

/// Write slide content with proper markdown formatting.
fn write_slide_content(
    writer: &mut MarkdownWriter,
    slide: &Slide,
    _options: &MarkdownOptions,
) -> Result<()> {
    match slide {
        Slide::Ppt(_) => {
            // Write PPT slide text content
            let text = slide.text()?;
            if !text.is_empty() {
                writer.push_str(&text);
                writer.push_str("\n\n");
            }
        },
        Slide::Pptx(_) => {
            // For PPTX slides, we have limited access to structured content
            // Just write the plain text for now
            let text = slide.text()?;
            if !text.is_empty() {
                writer.push_str(&text);
                writer.push_str("\n\n");
            }
        },
        #[cfg(feature = "iwa")]
        Slide::Keynote(_) => {
            // For Keynote slides, write the text content
            let text = slide.text()?;
            if !text.is_empty() {
                writer.push_str(&text);
                writer.push_str("\n\n");
            }
        },
        #[cfg(feature = "odf")]
        Slide::Odp(_) => {
            // For ODP slides, write the text content
            let text = slide.text()?;
            if !text.is_empty() {
                writer.push_str(&text);
                writer.push_str("\n\n");
            }
        },
    }

    Ok(())
}

impl ToMarkdown for Slide {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        // For individual slides, just return the text
        // Formatting is minimal for presentations
        self.text()
    }
}
