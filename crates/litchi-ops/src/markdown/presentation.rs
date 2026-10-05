//! Markdown for the presentation family: the facade's
//! `impl ToMarkdown for Presentation`, moved onto [`PresentationView`].

#[cfg(feature = "yaml")]
use litchi_core::Metadata;
use litchi_markdown::{MarkdownOptions, escape};

#[cfg(feature = "yaml")]
use crate::view::presentation::Inner;
use crate::{PresentationView, Result};

/// Convert a presentation view, naming its format on every error.
pub(super) fn convert(view: &PresentationView<'_>, options: &MarkdownOptions) -> Result<String> {
    convert_in_format(view, options).map_err(|error| super::in_format(error, view.format()))
}

fn convert_in_format(view: &PresentationView<'_>, options: &MarkdownOptions) -> Result<String> {
    // Write metadata as YAML front matter if available and enabled
    let metadata_md = if options.include_metadata {
        #[cfg(feature = "yaml")]
        {
            match metadata(view)? {
                Some(metadata) => super::yaml_front_matter(&metadata)?,
                None => String::new(),
            }
        }
        #[cfg(not(feature = "yaml"))]
        {
            return Err(super::yaml_required());
        }
    } else {
        String::new()
    };

    // Use the fast path that extracts text without shape parsing, which is
    // significantly faster for PPT files.
    let slide_texts = view.extract_text_for_markdown()?;

    // Ordinary conversion is serial, so options never select a global worker
    // pool. Iterate in producer order to retain deterministic output.
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
    let mut markdown = String::with_capacity(metadata_md.len().saturating_add(content_md.len()));
    markdown.push_str(&metadata_md);
    markdown.push_str(&content_md);
    Ok(markdown)
}

/// The document properties the front matter renders, when the format has
/// any.
#[cfg(feature = "yaml")]
fn metadata(view: &PresentationView<'_>) -> Result<Option<Metadata>> {
    match view.inner() {
        #[cfg(feature = "pptx")]
        Inner::Pptx(pptx) => pptx.metadata(),
        #[cfg(feature = "ppt")]
        Inner::Ppt(ppt) => ppt.metadata(),
        #[cfg(feature = "odp")]
        Inner::Odp(odp) => odp.metadata(),
    }
}

fn split_title_and_body(text: &str) -> (&str, &str) {
    text.split_once('\n').unwrap_or((text, ""))
}
