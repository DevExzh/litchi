/// ToMarkdown implementations for Presentation types.
///
/// This module implements the `ToMarkdown` trait for PowerPoint presentation types,
/// including Presentation and Slide.
use crate::common::{Error, Result};
use crate::presentation::{Presentation, Slide};
use super::config::MarkdownOptions;
use super::traits::ToMarkdown;
use std::fmt::Write as FmtWrite;

impl ToMarkdown for Presentation {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        let mut output = String::with_capacity(4096);

        // TODO: Add metadata support when Presentation metadata API is available
        // if _options.include_metadata {
        //     output.push_str("---\n");
        //     output.push_str(&format!("slides: {}\n", self.slide_count()?));
        //     output.push_str("---\n\n");
        // }

        let slides = self.slides()?;
        for (i, slide) in slides.iter().enumerate() {
            if i > 0 {
                // Separate slides with horizontal rule
                output.push_str("\n\n---\n\n");
            }

            // Add slide number as heading
            writeln!(output, "# Slide {}", i + 1).map_err(|e| Error::Other(e.to_string()))?;
            output.push('\n');

            // Add slide content
            let text = slide.text()?;
            output.push_str(&text);
        }

        Ok(output)
    }
}

impl ToMarkdown for Slide {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        // For individual slides, just return the text
        // Formatting is minimal for presentations
        self.text()
    }
}

