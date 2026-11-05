//! Markdown for Pages documents and Keynote presentations (record 0306):
//! the unified `File`'s object tree, rendered the way the document and
//! presentation families render (paragraphs; one `# Slide N` section per
//! slide). Numbers workbooks are refused like every workbook.

use std::borrow::Cow;

use litchi_core::FileFormat;
use litchi_markdown::{MarkdownOptions, escape};

use crate::format::iwork::Iwork;
use crate::{Code, Error, File, Result};

/// Convert a Pages or Keynote file, naming its format on every error.
pub(super) fn convert(file: &File, handle: &Iwork, options: &MarkdownOptions) -> Result<String> {
    let format = handle.format();
    let mut markdown = front_matter(handle, options)?;
    match format {
        FileFormat::Pages => pages(file, &mut markdown)?,
        _ => keynote(file, &mut markdown)?,
    }
    Ok(markdown)
}

/// The YAML front matter, when the options ask for it.
fn front_matter(handle: &Iwork, options: &MarkdownOptions) -> Result<String> {
    if !options.include_metadata {
        return Ok(String::new());
    }
    #[cfg(feature = "yaml")]
    {
        match handle.metadata()? {
            Some(metadata) => super::yaml_front_matter(&metadata),
            None => Ok(String::new()),
        }
    }
    #[cfg(not(feature = "yaml"))]
    {
        let _ = handle;
        Err(super::yaml_required())
    }
}

/// The text of the node at `path`, if it has any.
fn text(file: &File, path: &str) -> Result<Option<String>> {
    Ok(file.get(path)?.view().text()?.map(Cow::into_owned))
}

/// Every body paragraph, escaped, one Markdown paragraph each; an empty
/// paragraph is a blank line in Markdown and is skipped. Body tables cannot
/// be placed among the paragraphs yet, so a document holding one is refused
/// rather than converted without it.
fn pages(file: &File, markdown: &mut String) -> Result<()> {
    if file.query("tbl")?.next().is_some() {
        return Err(Error::new(
            Code::UnsupportedElement,
            "Markdown output of Pages body tables is not supported yet",
        ));
    }
    for node in file.query("p")? {
        let node = node?;
        let Some(paragraph) = node.view().text()? else {
            continue;
        };
        if paragraph.trim().is_empty() {
            continue;
        }
        markdown.push_str(&escape::text(&paragraph));
        markdown.push_str("\n\n");
    }
    Ok(())
}

/// One `# Slide N <title>` section per slide, its body text below, slides
/// separated by a thematic break: the presentation family's layout.
fn keynote(file: &File, markdown: &mut String) -> Result<()> {
    let count = file.query("slide")?.count();
    for index in 1..=count {
        if index > 1 {
            markdown.push_str("\n---\n\n");
        }
        let title = placeholder(file, index, "title")?;
        let body = placeholder(file, index, "body")?;
        match title
            .as_deref()
            .map(str::trim)
            .filter(|title| !title.is_empty())
        {
            Some(title) => {
                markdown.push_str(&format!("# Slide {index} {}\n\n", escape::text(title)));
            },
            None => markdown.push_str(&format!("# Slide {index}\n\n")),
        }
        if let Some(body) = body.as_deref().filter(|body| !body.trim().is_empty()) {
            markdown.push_str(&escape::text(body));
            markdown.push_str("\n\n");
        }
    }
    Ok(())
}

/// The text of a slide's title or body placeholder; a slide without one
/// has none.
fn placeholder(file: &File, slide: usize, kind: &str) -> Result<Option<String>> {
    match text(file, &format!("/slide[{slide}]/placeholder[{kind}]")) {
        Err(error) if error.code() == Code::PathNotFound => Ok(None),
        other => other,
    }
}
