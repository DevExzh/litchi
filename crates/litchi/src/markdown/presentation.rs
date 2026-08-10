#[cfg(feature = "yaml")]
use crate::MetadataYaml;
use crate::presentation::{Presentation, Slide};
/// ToMarkdown implementations for Presentation types.
///
/// This module implements the `ToMarkdown` trait for PowerPoint presentation types,
/// including Presentation and Slide.
///
/// **Note**: This module is only available when a presentation feature is enabled.
#[cfg(not(feature = "yaml"))]
use litchi_core::Error;
use litchi_core::Result;
use litchi_markdown::{MarkdownOptions, ToMarkdown, escape};

impl ToMarkdown for Presentation {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        // Write metadata as YAML front matter if available and enabled
        let metadata_md = if options.include_metadata {
            #[cfg(feature = "yaml")]
            {
                match self.metadata()? {
                    Some(metadata) => metadata.to_yaml_front_matter()?,
                    None => String::new(),
                }
            }
            #[cfg(not(feature = "yaml"))]
            {
                return Err(Error::FeatureDisabled("yaml".to_owned()));
            }
        } else {
            String::new()
        };

        // Use optimized fast path that extracts text without shape parsing
        // This is significantly faster for PPT files (3-10x speedup)
        let slide_texts = self.extract_text_for_markdown()?;

        // Ordinary conversion is serial, so options never select Rayon’s global
        // worker pool. Iterate in producer order to retain deterministic output.
        let mut content_md = String::new();
        for (i, (slide_num, text)) in slide_texts.iter().enumerate() {
            if i > 0 {
                content_md.push_str("\n\n---\n\n");
            }

            let (first_line, body) = split_title_and_body(text);
            let header_text = if first_line.is_empty() {
                format!("# Slide {}", slide_num)
            } else {
                format!("# Slide {} {}", slide_num, escape::text(first_line))
            };

            content_md.push_str(&header_text);
            content_md.push_str("\n\n");

            if !body.is_empty() {
                content_md.push_str(&escape::text(body));
                content_md.push_str("\n\n");
            }
        }

        // Combine metadata and content
        let mut markdown =
            String::with_capacity(metadata_md.len().saturating_add(content_md.len()));
        markdown.push_str(&metadata_md);
        markdown.push_str(&content_md);
        Ok(markdown)
    }
}

impl ToMarkdown for Slide {
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        // For individual slides, just return the text
        // Formatting is minimal for presentations
        Ok(escape::text(&self.text()?).into_owned())
    }
}

fn split_title_and_body(text: &str) -> (&str, &str) {
    text.split_once('\n').unwrap_or((text, ""))
}
