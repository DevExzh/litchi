//! Markdown for the word-processing family: the facade's
//! `impl ToMarkdown for Document` and the `Document` helpers it called
//! (`validate_markdown_projection`, `markdown_docx_document`,
//! `markdown_list_items`, `markdown_heading_levels` and
//! `check_source_freshness_after_markdown`), moved onto [`DocumentView`].

use litchi_core::Metadata;
use litchi_markdown::MarkdownOptions;

use super::writer::{ListItemInfo, MarkdownWriter};
use super::{allocation, invariant, unrepresentable};
#[cfg(feature = "docx")]
use crate::format::docx::DocxOwner;
#[cfg(feature = "odt")]
use crate::format::odt::OdtOwner;
use crate::view::document::Inner;
use crate::{DocumentView, Result};

/// Convert a document view, naming its format on every error.
pub(super) fn convert(view: &DocumentView<'_>, options: &MarkdownOptions) -> Result<String> {
    convert_in_format(view, options).map_err(|error| super::in_format(error, view.format()))
}

#[cfg_attr(
    not(feature = "docx"),
    allow(
        clippy::unused_enumerate_index,
        reason = "the element index aligns the DOCX projections only"
    )
)]
fn convert_in_format(view: &DocumentView<'_>, options: &MarkdownOptions) -> Result<String> {
    use crate::compat::word::DocumentElement;

    validate_projection(view)?;

    // Write metadata first (must be sequential)
    let metadata_md = if options.include_metadata {
        let mut metadata_writer = MarkdownWriter::new(*options);
        let metadata = metadata(view)?;
        metadata_writer.write_metadata(&metadata)?;
        metadata_writer.finish()
    } else {
        String::new()
    };

    // Extract all document elements (paragraphs and tables) in document order
    let elements = view.elements()?;
    #[cfg(feature = "docx")]
    let docx_bundle = docx_document(view)?
        .map(|document| super::docx::Bundle::build(&elements, &document))
        .transpose()?;
    let mut resolved_lists = list_items(view)?;
    let mut heading_levels = heading_levels(view)?;
    if resolved_lists.is_empty() {
        resolved_lists
            .try_reserve_exact(elements.len())
            .map_err(|source| allocation("Markdown list alignment", source))?;
        resolved_lists.resize(elements.len(), None);
    } else if resolved_lists.len() != elements.len() {
        return Err(invariant(
            "the resolved list metadata is not aligned with the document elements",
        ));
    }
    if heading_levels.is_empty() {
        heading_levels
            .try_reserve_exact(elements.len())
            .map_err(|source| allocation("Markdown heading alignment", source))?;
        heading_levels.resize(elements.len(), None);
    } else if heading_levels.len() != elements.len() {
        return Err(invariant(
            "the heading metadata is not aligned with the document elements",
        ));
    }

    // Ordinary conversion is serial. The public options value remains part
    // of the rendering contract, but it never selects a hidden global
    // scheduler.
    let mut writer = MarkdownWriter::new(*options);
    // Estimate: 100 bytes per paragraph, 500 bytes per table.
    let estimated_size = elements.len() * 150;
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
            DocumentElement::Table(table) => {
                writer.write_table(&table)?;
            },
        }
    }

    let content_md = writer.finish();

    // Combine metadata and content
    let mut markdown = String::with_capacity(metadata_md.len().saturating_add(content_md.len()));
    markdown.push_str(&metadata_md);
    markdown.push_str(&content_md);
    #[cfg(feature = "docx")]
    if let Some(bundle) = docx_bundle {
        markdown.push_str(&bundle.render_note_definitions(options)?);
    }
    check_source(view)?;
    Ok(markdown)
}

/// The document properties the front matter renders; empty when the format
/// reports none.
fn metadata(view: &DocumentView<'_>) -> Result<Metadata> {
    let metadata = match view.inner() {
        #[cfg(feature = "docx")]
        Inner::Docx(docx) => docx.metadata()?,
        #[cfg(feature = "doc")]
        Inner::Doc(doc) => doc.metadata()?,
        #[cfg(feature = "rtf")]
        Inner::Rtf(rtf) => rtf.metadata()?,
        #[cfg(feature = "odt")]
        Inner::Odt(odt) => odt.metadata()?,
    };
    Ok(metadata.unwrap_or_default())
}

/// Validate the lossless DOCX body stream before Markdown starts rendering.
/// `elements()` intentionally omits alternative-format anchors, so using it
/// here would silently drop active `altChunk` content. Keep unknown body
/// children inert only when they are section properties; every other child
/// must be refused rather than discarded.
#[cfg(feature = "docx")]
fn validate_docx_blocks(blocks: Vec<litchi_docx::Block>) -> Result<()> {
    for block in blocks {
        match block {
            litchi_docx::Block::Paragraph(_) | litchi_docx::Block::Table(_) => {},
            litchi_docx::Block::Sdt(_) => {
                return Err(unrepresentable(
                    "Markdown export cannot preserve DOCX content controls",
                ));
            },
            litchi_docx::Block::Alt(_) => {
                return Err(unrepresentable(
                    "Markdown export cannot preserve active DOCX altChunk body blocks",
                ));
            },
            litchi_docx::Block::Unknown(block)
                if !crate::view::document::docx_unknown_is_section_properties(&block) =>
            {
                return Err(unrepresentable(
                    "Markdown export cannot preserve unmodeled DOCX body blocks",
                ));
            },
            litchi_docx::Block::Unknown(_) => {},
        }
    }
    Ok(())
}

/// Reject source semantics that the Markdown adapter cannot place without
/// guessing. The adapter never follows targets or activates content.
fn validate_projection(view: &DocumentView<'_>) -> Result<()> {
    #[cfg(any(feature = "doc", feature = "rtf", feature = "odt"))]
    let unsupported = |kind: &str| {
        unrepresentable(format!(
            "Markdown export cannot preserve {kind} without its source placement context"
        ))
    };

    match view.inner() {
        #[cfg(feature = "doc")]
        Inner::Doc(doc) => {
            let document = doc.reader()?;
            if !document.hyperlinks()?.is_empty() {
                return Err(unsupported("DOC hyperlinks"));
            }
            if !document.footnotes()?.is_empty() {
                return Err(unsupported("DOC footnotes"));
            }
            for element in document.elements()? {
                let litchi_doc::Element::Paragraph(paragraph) = element else {
                    continue;
                };
                if paragraph.runs()?.iter().any(litchi_doc::Run::has_image) {
                    return Err(unsupported("DOC inline images"));
                }
            }
        },
        #[cfg(feature = "docx")]
        Inner::Docx(docx) => match docx.owner() {
            DocxOwner::Eager(package) => validate_docx_blocks(package.document()?.blocks()?)?,
            owner @ DocxOwner::Source(package) => {
                owner.finish((|| validate_docx_blocks(package.document()?.blocks()?))())?;
            },
        },
        #[cfg(feature = "rtf")]
        Inner::Rtf(rtf) => {
            let document = rtf.document();
            if !document.hyperlinks().is_empty() {
                return Err(unsupported("RTF hyperlinks"));
            }
            if !document.footnotes().is_empty() {
                return Err(unsupported("RTF footnotes"));
            }
            if !document.quote_fields().is_empty() {
                return Err(unsupported("RTF quote fields"));
            }
        },
        #[cfg(feature = "odt")]
        Inner::Odt(odt) => {
            let result = match odt.owner() {
                OdtOwner::Eager(document) => (|| {
                    let odt_error = super::odt_error;
                    if !document.hyperlinks().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT hyperlinks"));
                    }
                    if !document.footnotes().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT footnotes"));
                    }
                    if !document.images().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT images"));
                    }
                    Ok(())
                })(),
                #[cfg(any(unix, windows))]
                OdtOwner::Source(document) => (|| {
                    let odt_error = super::odt_error;
                    if !document.hyperlinks().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT hyperlinks"));
                    }
                    if !document.footnotes().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT footnotes"));
                    }
                    if !document.images().map_err(odt_error)?.is_empty() {
                        return Err(unsupported("ODT images"));
                    }
                    Ok(())
                })(),
            };
            finish_odt(odt.owner(), result)?;
        },
    }
    Ok(())
}

/// Prefer a final ODT source-revision failure over the result of a read
/// that already returned a unified error.
#[cfg(feature = "odt")]
fn finish_odt<T>(owner: &OdtOwner, result: Result<T>) -> Result<T> {
    owner.check_source()?;
    result
}

/// The package-aware DOCX view the Markdown adapter reads relationships and
/// note definitions from.
#[cfg(feature = "docx")]
fn docx_document(view: &DocumentView<'_>) -> Result<Option<litchi_docx::Document>> {
    #[allow(
        unreachable_patterns,
        reason = "DOCX is the only document format in a docx-only feature build"
    )]
    match view.inner() {
        Inner::Docx(docx) => {
            let package = docx.markdown_package()?;
            docx.owner()
                .finish(package.document().map(Some).map_err(crate::Error::from))
        },
        _ => Ok(None),
    }
}

/// Resolve list semantics while the format owner and its definitions remain
/// available.
fn list_items(view: &DocumentView<'_>) -> Result<Vec<Option<ListItemInfo>>> {
    #[allow(
        unreachable_patterns,
        reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
    )]
    match view.inner() {
        #[cfg(feature = "doc")]
        Inner::Doc(doc) => resolve_doc_lists(doc.reader()?),
        #[cfg(feature = "docx")]
        Inner::Docx(docx) => {
            let package = docx.markdown_package()?;
            docx.owner()
                .finish((|| resolve_docx_lists(&package.document()?))())
        },
        _ => Ok(Vec::new()),
    }
}

/// Preserve ODT outline levels while the ODT semantic model is still
/// available. The paragraph wrapper stores only paragraph content, so this
/// sidecar is aligned with [`DocumentView::elements`].
fn heading_levels(view: &DocumentView<'_>) -> Result<Vec<Option<u8>>> {
    #[allow(
        unreachable_patterns,
        reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
    )]
    match view.inner() {
        #[cfg(feature = "odt")]
        Inner::Odt(odt) => match odt.owner() {
            OdtOwner::Eager(document) => {
                odt_heading_levels(document.elements().map_err(super::odt_error)?)
            },
            #[cfg(any(unix, windows))]
            owner @ OdtOwner::Source(document) => {
                let result = document
                    .elements()
                    .map_err(super::odt_error)
                    .and_then(odt_heading_levels);
                finish_odt(owner, result)
            },
        },
        _ => Ok(Vec::new()),
    }
}

/// The outline level of each ODT element that becomes a document element.
#[cfg(feature = "odt")]
fn odt_heading_levels(
    source: Vec<litchi_odt::elements::parser::OrderElement>,
) -> Result<Vec<Option<u8>>> {
    use litchi_odt::elements::parser::OrderElement;

    let mut levels = Vec::new();
    levels
        .try_reserve_exact(source.len())
        .map_err(|source| allocation("Markdown ODT heading metadata", source))?;
    for element in source {
        match element {
            OrderElement::Paragraph(_)
            | OrderElement::NumberedParagraph(_)
            | OrderElement::Table(_) => levels.push(None),
            OrderElement::Heading(heading) => {
                let level = heading.level().ok_or_else(|| {
                    unrepresentable("an ODT heading has no representable outline level")
                })?;
                if !(1..=6).contains(&level) {
                    return Err(unrepresentable(format!(
                        "ODT heading outline level {level} is outside Markdown's range"
                    )));
                }
                levels.push(Some(level));
            },
            // `elements()` deliberately omits ODT lists from the document
            // view, so they must not consume alignment.
            OrderElement::List(_) => {},
        }
    }
    Ok(levels)
}

/// Check the source revision after a complete Markdown conversion. The
/// adapter may read the cached owning package for relationships, notes,
/// images and numbering, so its final return boundary must still be bound to
/// the source-backed owner's exact revision.
fn check_source(view: &DocumentView<'_>) -> Result<()> {
    #[allow(
        unreachable_patterns,
        reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
    )]
    match view.inner() {
        #[cfg(feature = "docx")]
        Inner::Docx(docx) => docx.owner().check_source(),
        #[cfg(feature = "odt")]
        Inner::Odt(odt) => odt.owner().check_source(),
        _ => Ok(()),
    }
}

#[cfg(feature = "docx")]
fn resolve_docx_lists(document: &litchi_docx::Document) -> Result<Vec<Option<ListItemInfo>>> {
    use litchi_docx::list::ListKind;

    let elements = document.elements_with_resolved_list_items()?;
    let mut resolved = Vec::new();
    resolved
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("Markdown DOCX list metadata", source))?;
    for (element, item) in elements {
        if let litchi_docx::Element::Unknown(block) = element {
            if crate::view::document::docx_unknown_is_section_properties(&block) {
                continue;
            }
            return Err(unrepresentable(
                "Markdown export cannot preserve an active unmodeled DOCX body block",
            ));
        }
        let Some(item) = item else {
            resolved.push(None);
            continue;
        };
        let info = match item.kind {
            ListKind::Unordered => ListItemInfo::bullet(usize::from(item.numbering.level)),
            ListKind::Ordered => {
                if !(0..=999_999_999).contains(&item.value) {
                    return Err(unrepresentable(format!(
                        "DOCX ordered-list value {} is outside CommonMark's marker domain",
                        item.value
                    )));
                }
                ListItemInfo::ordered(
                    usize::from(item.numbering.level),
                    format!("{}.", item.value),
                )
            },
            ListKind::Unmarked => {
                return Err(unrepresentable(format!(
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
fn resolve_doc_lists(document: &litchi_doc::Document) -> Result<Vec<Option<ListItemInfo>>> {
    use litchi_doc::Element;

    let elements = document.elements()?;
    let mut resolved = Vec::new();
    resolved
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("Markdown DOC list metadata", source))?;
    for element in elements {
        let Element::Paragraph(paragraph) = element else {
            resolved.push(None);
            continue;
        };
        let properties = paragraph.properties();
        if properties.legacy_autonumbering.is_some() {
            return Err(unrepresentable(
                "Markdown export cannot preserve legacy DOC autonumber counters and restarts",
            ));
        }
        if properties.list_level == Some(12)
            || properties
                .list_format_override
                .is_none_or(|value| value == 0)
        {
            resolved.push(None);
            continue;
        }
        let binding = document.paragraph_list_binding(&paragraph).ok_or_else(|| {
            unrepresentable("a DOC paragraph references a missing or ambiguous list definition")
        })?;
        let level = binding.effective_level();
        if level.is_bullet() {
            resolved.push(Some(ListItemInfo::bullet(usize::from(binding.level))));
        } else if level.is_numbered() {
            return Err(unrepresentable(
                "Markdown export cannot preserve DOC ordered-list counters and restarts",
            ));
        } else {
            return Err(unrepresentable(
                "the DOC list definition has no representable bullet or ordered format",
            ));
        }
    }
    Ok(resolved)
}
