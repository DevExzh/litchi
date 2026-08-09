use super::writer::MarkdownWriter;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use crate::document::Table;
use crate::document::{Document, Paragraph, Run};
/// ToMarkdown implementations for Document types.
///
/// This module implements the `ToMarkdown` trait for Word document types,
/// including Document, Paragraph, Run, and Table.
///
/// **Note**: This module is only available when a document-format feature is enabled.
use litchi_core::{Error, Result};
use litchi_markdown::{MarkdownOptions, ToMarkdown};
use rayon::prelude::*;

/// Minimum number of elements to justify parallel processing overhead.
const PARALLEL_THRESHOLD: usize = 50;

#[cfg(feature = "docx")]
pub(crate) fn resolve_docx_lists(
    document: &crate::docx::Document<'_>,
) -> Result<Vec<Option<super::writer::ListItemInfo>>> {
    use crate::docx::list::ListKind;

    let elements = document
        .elements_with_resolved_list_items()
        .map_err(crate::map_ooxml_error)?;
    let mut resolved = Vec::new();
    resolved
        .try_reserve_exact(elements.len())
        .map_err(|source| Error::Allocation {
            resource: "Markdown DOCX list metadata",
            source,
        })?;
    for (element, item) in elements {
        if let crate::docx::Element::Unknown(block) = element {
            if crate::document::docx_unknown_is_section_properties(&block) {
                continue;
            }
            return Err(Error::Unsupported(
                "Markdown export cannot preserve an active unmodeled DOCX body block".to_owned(),
            ));
        }
        let Some(item) = item else {
            resolved.push(None);
            continue;
        };
        let info = match item.kind {
            ListKind::Unordered => {
                super::writer::ListItemInfo::bullet(usize::from(item.numbering.level))
            },
            ListKind::Ordered => {
                if !(0..=999_999_999).contains(&item.value) {
                    return Err(Error::Unsupported(format!(
                        "DOCX ordered-list value {} is outside CommonMark's marker domain",
                        item.value
                    )));
                }
                super::writer::ListItemInfo::ordered(
                    usize::from(item.numbering.level),
                    format!("{}.", item.value),
                )
            },
            ListKind::Unmarked => {
                return Err(Error::Unsupported(format!(
                    "DOCX numbering format '{}' has no Markdown list marker",
                    item.format
                )));
            },
        };
        resolved.push(Some(info));
    }
    Ok(resolved)
}

#[cfg(feature = "doc")]
pub(crate) fn resolve_doc_lists(
    document: &litchi_doc::Document,
) -> Result<Vec<Option<super::writer::ListItemInfo>>> {
    use litchi_doc::Element;

    let elements = document.elements().map_err(Error::from)?;
    let mut resolved = Vec::new();
    resolved
        .try_reserve_exact(elements.len())
        .map_err(|source| Error::Allocation {
            resource: "Markdown DOC list metadata",
            source,
        })?;
    for element in elements {
        let Element::Paragraph(paragraph) = element else {
            resolved.push(None);
            continue;
        };
        let properties = paragraph.properties();
        if properties.legacy_autonumbering.is_some() {
            return Err(Error::Unsupported(
                "Markdown export cannot preserve legacy DOC autonumber counters and restarts"
                    .to_owned(),
            ));
        }
        if properties.list_level == Some(12)
            || !properties
                .list_format_override
                .is_some_and(|value| value != 0)
        {
            resolved.push(None);
            continue;
        }
        let binding = document.paragraph_list_binding(&paragraph).ok_or_else(|| {
            Error::Unsupported(
                "DOC paragraph references a missing or ambiguous list definition".to_owned(),
            )
        })?;
        let level = binding.effective_level();
        if level.is_bullet() {
            resolved.push(Some(super::writer::ListItemInfo::bullet(usize::from(
                binding.level,
            ))));
        } else if level.is_numbered() {
            return Err(Error::Unsupported(
                "Markdown export cannot preserve DOC ordered-list counters and restarts".to_owned(),
            ));
        } else {
            return Err(Error::Unsupported(
                "DOC list definition has no representable bullet or ordered format".to_owned(),
            ));
        }
    }
    Ok(resolved)
}

impl ToMarkdown for Document {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        use crate::document::DocumentElement;

        self.validate_markdown_projection()?;

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
        #[cfg(feature = "docx")]
        let docx_bundle = self
            .markdown_docx_document()?
            .map(|document| super::docx::Bundle::build(&elements, &document))
            .transpose()?;
        let mut resolved_lists = self.markdown_list_items()?;
        let mut heading_levels = self.markdown_heading_levels()?;
        if resolved_lists.is_empty() {
            resolved_lists
                .try_reserve_exact(elements.len())
                .map_err(|source| Error::Allocation {
                    resource: "Markdown list alignment",
                    source,
                })?;
            resolved_lists.resize(elements.len(), None);
        } else if resolved_lists.len() != elements.len() {
            return Err(Error::InvalidFormat(
                "resolved list metadata is not aligned with document elements".to_owned(),
            ));
        }
        if heading_levels.is_empty() {
            heading_levels
                .try_reserve_exact(elements.len())
                .map_err(|source| Error::Allocation {
                    resource: "Markdown heading alignment",
                    source,
                })?;
            heading_levels.resize(elements.len(), None);
        } else if heading_levels.len() != elements.len() {
            return Err(Error::InvalidFormat(
                "heading metadata is not aligned with document elements".to_owned(),
            ));
        }

        // Decide whether to use parallel or sequential processing
        let content_md = if options.use_parallel && elements.len() >= PARALLEL_THRESHOLD {
            // PARALLEL PATH: Process elements in parallel for large documents
            // With Arc-based Send + Sync types, we can now safely parallelize
            let element_strings: Result<Vec<String>> = elements
                .par_iter()
                .enumerate()
                .zip(resolved_lists.par_iter())
                .zip(heading_levels.par_iter())
                .map(
                    |(((_element_index, element), resolved_list), heading_level)| {
                        let mut writer = MarkdownWriter::new(*options);
                        match element {
                            DocumentElement::Paragraph(para) => {
                                writer.write_paragraph_with_projection(
                                    para,
                                    resolved_list.as_ref(),
                                    *heading_level,
                                    #[cfg(feature = "docx")]
                                    docx_bundle.as_ref().and_then(|bundle| {
                                        bundle.paragraphs[_element_index].as_ref()
                                    }),
                                )?;
                            },
                            #[cfg(any(
                                feature = "doc",
                                feature = "docx",
                                feature = "rtf",
                                feature = "odt"
                            ))]
                            DocumentElement::Table(table) => {
                                writer.write_table(table)?;
                            },
                        }
                        Ok(writer.finish())
                    },
                )
                .collect();
            let element_strings = element_strings?;

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

            for (_element_index, ((element, resolved_list), heading_level)) in elements
                .into_iter()
                .zip(resolved_lists.iter())
                .zip(heading_levels.iter())
                .enumerate()
            {
                match element {
                    DocumentElement::Paragraph(para) => {
                        writer.write_paragraph_with_projection(
                            &para,
                            resolved_list.as_ref(),
                            *heading_level,
                            #[cfg(feature = "docx")]
                            docx_bundle
                                .as_ref()
                                .and_then(|bundle| bundle.paragraphs[_element_index].as_ref()),
                        )?;
                    },
                    #[cfg(any(
                        feature = "doc",
                        feature = "docx",
                        feature = "rtf",
                        feature = "odt"
                    ))]
                    DocumentElement::Table(table) => {
                        writer.write_table(&table)?;
                    },
                }
            }

            writer.finish()
        };

        // Combine metadata and content
        let mut markdown =
            String::with_capacity(metadata_md.len().saturating_add(content_md.len()));
        markdown.push_str(&metadata_md);
        markdown.push_str(&content_md);
        #[cfg(feature = "docx")]
        if let Some(bundle) = docx_bundle {
            markdown.push_str(&bundle.render_note_definitions(options)?);
        }
        Ok(markdown)
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

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
impl ToMarkdown for Table {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = MarkdownWriter::new(*options);
        writer.write_table(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

#[cfg(all(test, feature = "doc"))]
mod tests {
    use super::*;

    #[test]
    fn real_doc_inline_images_are_refused_in_plain_output() -> Result<()> {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/testPictures.doc");
        let document = Document::open(path)?;
        for options in [
            MarkdownOptions::new(),
            MarkdownOptions::new().with_styles(false),
        ] {
            assert!(matches!(
                document.to_markdown_with_options(&options),
                Err(Error::Unsupported(_))
            ));
        }
        Ok(())
    }
}
