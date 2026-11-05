//! Package-aware DOCX inline projection built on typed source-order APIs,
//! moved from the `litchi` facade's `markdown::docx` and switched to the
//! [`compat`](crate::compat) element wrappers and the typed [`Error`].

use std::collections::BTreeMap;

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_markdown::{MarkdownOptions, escape};

use super::writer::MarkdownWriter;
use super::{allocation, invariant, unrepresentable};
use crate::compat::word::{DocumentElement, Paragraph};
use crate::{Code, Error, Result};

#[derive(Clone, Debug)]
pub(super) struct Link {
    pub(super) start_run: usize,
    pub(super) end_run: usize,
    pub(super) destination: String,
    pub(super) title: Option<String>,
}

#[derive(Clone, Debug)]
pub(super) struct EmbeddedImage {
    pub(super) alt: String,
    pub(super) title: Option<String>,
    pub(super) data_uri: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum NoteKey {
    Footnote(u32),
    Endnote(u32),
}

impl NoteKey {
    pub(super) fn label(self) -> String {
        match self {
            Self::Footnote(id) => format!("fn-{id}"),
            Self::Endnote(id) => format!("en-{id}"),
        }
    }
}

#[derive(Clone, Debug)]
pub(super) enum RunProjection {
    Ordinary,
    Images(Vec<EmbeddedImage>),
    Note(NoteKey),
}

#[derive(Clone, Debug)]
pub(super) struct ParagraphProjection {
    pub(super) runs: Vec<RunProjection>,
    pub(super) links: Vec<Link>,
}

impl ParagraphProjection {
    pub(super) fn has_non_text_content(&self) -> bool {
        !self.links.is_empty()
            || self
                .runs
                .iter()
                .any(|run| !matches!(run, RunProjection::Ordinary))
    }
}

pub(super) struct Bundle {
    pub(super) paragraphs: Vec<Option<ParagraphProjection>>,
    notes: BTreeMap<NoteKey, litchi_docx::footnote::Note>,
    used_notes: Vec<NoteKey>,
}

impl Bundle {
    #[allow(
        unreachable_patterns,
        reason = "DOCX is the only paragraph variant in a docx-only feature build"
    )]
    pub(super) fn build(
        elements: &[DocumentElement],
        document: &litchi_docx::Document,
    ) -> Result<Self> {
        let main_part = document.opc_package().main_document_part()?;
        let relationships = main_part.rels();
        let opc = document.opc_package();
        let mut paragraphs = Vec::new();
        paragraphs
            .try_reserve_exact(elements.len())
            .map_err(|source| allocation("Markdown DOCX inline projection", source))?;
        let mut used_notes = Vec::new();

        for element in elements {
            let projection = match element {
                DocumentElement::Paragraph(paragraph) => match paragraph.as_ref() {
                    Paragraph::Docx(paragraph) => {
                        let projection = project_paragraph(paragraph, relationships, opc)?;
                        for run in &projection.runs {
                            if let RunProjection::Note(key) = run
                                && !used_notes.contains(key)
                            {
                                used_notes.push(*key);
                            }
                        }
                        Some(projection)
                    },
                    _ => None,
                },
                _ => None,
            };
            paragraphs.push(projection);
        }

        let mut notes = BTreeMap::new();
        for note in document.footnotes()? {
            notes.insert(NoteKey::Footnote(note.id()), note);
        }
        for note in document.endnotes()? {
            notes.insert(NoteKey::Endnote(note.id()), note);
        }
        for key in &used_notes {
            if !notes.contains_key(key) {
                return Err(missing_note(*key));
            }
        }

        Ok(Self {
            paragraphs,
            notes,
            used_notes,
        })
    }

    pub(super) fn render_note_definitions(&self, options: &MarkdownOptions) -> Result<String> {
        let mut output = String::new();
        for key in &self.used_notes {
            let note = self.notes.get(key).ok_or_else(|| missing_note(*key))?;
            let paragraphs = note.paragraphs()?;
            if paragraphs.is_empty() {
                return Err(unrepresentable(format!(
                    "DOCX {} has no paragraph content for Markdown",
                    key.label()
                )));
            }
            let mut bodies = Vec::new();
            bodies
                .try_reserve_exact(paragraphs.len())
                .map_err(|source| allocation("Markdown DOCX note paragraphs", source))?;
            for paragraph in paragraphs {
                validate_note_paragraph(&paragraph)?;
                let mut writer = MarkdownWriter::new(*options);
                writer.write_paragraph(&Paragraph::Docx(paragraph))?;
                bodies.push(writer.finish().trim_end().to_owned());
            }
            output.push_str("[^");
            output.push_str(&key.label());
            output.push_str("]: ");
            append_indented_definition(&mut output, &bodies.join("\n\n"));
            output.push('\n');
        }
        Ok(output)
    }
}

/// The body references a note the package does not define.
fn missing_note(key: NoteKey) -> Error {
    Error::new(
        Code::CorruptFile,
        format!("the DOCX body references the missing note {}", key.label()),
    )
}

fn project_paragraph(
    paragraph: &litchi_docx::Paragraph,
    relationships: &litchi_opc::rel::Relationships,
    opc: &litchi_opc::OpcPackage,
) -> Result<ParagraphProjection> {
    let inlines = paragraph.inlines_with_relationships(relationships)?;
    let mut runs = Vec::new();
    let mut links = Vec::new();

    for inline in inlines {
        match inline {
            litchi_docx::Inline::Run(run) => {
                runs.push(project_run(&run, relationships, opc)?);
            },
            litchi_docx::Inline::Hyperlink(hyperlink) => {
                if hyperlink.has_unmodeled_content() {
                    return Err(unrepresentable(
                        "Markdown export cannot preserve unmodeled DOCX hyperlink children",
                    ));
                }
                if hyperlink.target_frame().is_some() {
                    return Err(unrepresentable(
                        "Markdown export cannot preserve DOCX hyperlink attribute w:tgtFrame",
                    ));
                }
                if hyperlink.runs().is_empty() {
                    return Err(unrepresentable(
                        "Markdown export cannot place a DOCX hyperlink with no runs",
                    ));
                }
                let start_run = runs.len();
                for run in hyperlink.runs() {
                    let projected = project_run(run, relationships, opc)?;
                    if matches!(projected, RunProjection::Note(_)) {
                        return Err(unrepresentable(
                            "Markdown export cannot nest a DOCX note reference inside a hyperlink",
                        ));
                    }
                    runs.push(projected);
                }
                let end_run = runs.len();
                links.push(Link {
                    start_run,
                    end_run,
                    destination: hyperlink_destination(
                        hyperlink.link(),
                        hyperlink.document_location(),
                    )?,
                    title: hyperlink.link().tooltip().map(str::to_owned),
                });
            },
            litchi_docx::Inline::BookmarkStart(_) | litchi_docx::Inline::BookmarkEnd(_) => {
                // Bookmark boundaries are invisible anchors with no Markdown
                // representation; dropping them loses no content.
            },
            litchi_docx::Inline::Unknown(_) => {
                return Err(unrepresentable(
                    "Markdown export cannot preserve an unmodeled DOCX paragraph child",
                ));
            },
            _ => {
                return Err(unrepresentable(
                    "Markdown export cannot preserve a future DOCX paragraph child",
                ));
            },
        }
    }

    let wrapper_run_count = paragraph.runs()?.len();
    if wrapper_run_count != runs.len() {
        return Err(invariant(
            "the DOCX typed inline projection is not aligned with the paragraph runs",
        ));
    }
    Ok(ParagraphProjection { runs, links })
}

fn project_run(
    run: &litchi_docx::Run,
    relationships: &litchi_opc::rel::Relationships,
    opc: &litchi_opc::OpcPackage,
) -> Result<RunProjection> {
    let contents = run.contents()?;
    let mut images = Vec::new();
    let mut note = None;
    let mut has_visible_text = false;
    let mut unknown_count = 0usize;

    for content in contents {
        match content {
            litchi_docx::RunContent::Text(text) => has_visible_text |= !text.is_empty(),
            litchi_docx::RunContent::Tab
            | litchi_docx::RunContent::CarriageReturn
            | litchi_docx::RunContent::NoBreakHyphen
            | litchi_docx::RunContent::SoftHyphen => has_visible_text = true,
            litchi_docx::RunContent::Break(run_break) => {
                if !matches!(
                    run_break.break_type,
                    litchi_docx::RunBreakType::TextWrapping
                ) {
                    return Err(unrepresentable(
                        "Markdown export cannot preserve a DOCX page or column break inside a run",
                    ));
                }
                has_visible_text = true;
            },
            litchi_docx::RunContent::Image(image) => {
                let bytes = image.data(opc, relationships)?;
                let format = image.format(opc, relationships)?;
                let mut data_uri = String::from("data:");
                data_uri.push_str(format.mime_type());
                data_uri.push_str(";base64,");
                BASE64.encode_string(&bytes, &mut data_uri);
                images.push(EmbeddedImage {
                    alt: image.description().to_owned(),
                    title: (!image.name().is_empty()).then(|| image.name().to_owned()),
                    data_uri,
                });
            },
            litchi_docx::RunContent::EmbeddedObject(object) => {
                let preview = object.preview_image();
                let resolved = preview.and_then(|preview| {
                    preview
                        .data(opc, relationships)
                        .and_then(|bytes| {
                            preview
                                .format(opc, relationships)
                                .map(|format| (bytes, format))
                        })
                        .ok()
                        .map(|(bytes, format)| {
                            let mut data_uri = String::from("data:");
                            data_uri.push_str(format.mime_type());
                            data_uri.push_str(";base64,");
                            BASE64.encode_string(&bytes, &mut data_uri);
                            EmbeddedImage {
                                alt: preview.description().to_owned(),
                                title: (!preview.name().is_empty())
                                    .then(|| preview.name().to_owned()),
                                data_uri,
                            }
                        })
                });
                // An object whose preview cannot be resolved degrades to an
                // alt-text placeholder rather than failing the export.
                images.push(resolved.unwrap_or_else(|| {
                    EmbeddedImage {
                        alt: preview
                            .map(|preview| preview.description().to_owned())
                            .filter(|alt| !alt.is_empty())
                            .or_else(|| object.prog_id().map(str::to_owned))
                            .unwrap_or_default(),
                        title: None,
                        data_uri: String::new(),
                    }
                }));
            },
            litchi_docx::RunContent::FootnoteReference(id) => {
                set_note(&mut note, NoteKey::Footnote(id))?;
            },
            litchi_docx::RunContent::EndnoteReference(id) => {
                set_note(&mut note, NoteKey::Endnote(id))?;
            },
            litchi_docx::RunContent::Unknown(_) => {
                unknown_count = unknown_count
                    .checked_add(1)
                    .ok_or_else(|| invariant("the DOCX run-content counter overflowed"))?;
            },
            _ => {
                return Err(unrepresentable(
                    "Markdown export cannot preserve a future DOCX run child",
                ));
            },
        }
    }

    if unknown_count > 0 {
        let one_formula = unknown_count == 1 && run.omml_formula()?.is_some();
        if !one_formula {
            return Err(unrepresentable(
                "Markdown export cannot preserve an unmodeled DOCX run child",
            ));
        }
    }
    if (!images.is_empty() || note.is_some()) && has_visible_text {
        return Err(unrepresentable(
            "Markdown export cannot preserve interleaved DOCX text and embedded semantics in one run",
        ));
    }
    if !images.is_empty() && note.is_some() {
        return Err(unrepresentable(
            "Markdown export cannot combine a DOCX note reference and image in one run",
        ));
    }
    if !images.is_empty() {
        Ok(RunProjection::Images(images))
    } else if let Some(note) = note {
        Ok(RunProjection::Note(note))
    } else {
        Ok(RunProjection::Ordinary)
    }
}

fn set_note(slot: &mut Option<NoteKey>, key: NoteKey) -> Result<()> {
    if slot.replace(key).is_some() {
        return Err(unrepresentable(
            "Markdown export cannot preserve multiple DOCX note references in one run",
        ));
    }
    Ok(())
}

fn hyperlink_destination(
    link: &litchi_docx::Hyperlink,
    document_location: Option<&str>,
) -> Result<String> {
    if document_location.is_some() && link.anchor().is_some() {
        return Err(unrepresentable(
            "Markdown export cannot combine DOCX hyperlink anchor and document location",
        ));
    }
    let mut destination = match (link.url(), link.anchor()) {
        (Some(url), _) => url.to_owned(),
        (None, Some(anchor)) => format!("#{anchor}"),
        (None, None) if document_location.is_some() => String::new(),
        (None, None) => {
            return Err(unrepresentable(
                "Markdown export cannot resolve a DOCX hyperlink target",
            ));
        },
    };
    if link.url().is_some()
        && let Some(anchor) = link.anchor()
    {
        destination.push('#');
        destination.push_str(anchor);
    }
    if let Some(location) = document_location {
        destination.push('#');
        destination.push_str(location);
    }
    Ok(escape::link_destination(&destination).into_owned())
}

pub(super) fn requires_package_context(paragraph: &litchi_docx::Paragraph) -> Result<bool> {
    for inline in paragraph.inlines()? {
        match inline {
            litchi_docx::Inline::Run(run) => {
                let contents = run.contents()?;
                let unknowns = contents
                    .iter()
                    .filter(|content| matches!(content, litchi_docx::RunContent::Unknown(_)))
                    .count();
                if contents.iter().any(|content| {
                    matches!(
                        content,
                        litchi_docx::RunContent::Image(_)
                            | litchi_docx::RunContent::EmbeddedObject(_)
                            | litchi_docx::RunContent::FootnoteReference(_)
                            | litchi_docx::RunContent::EndnoteReference(_)
                    )
                }) || (unknowns > 0 && !(unknowns == 1 && run.omml_formula()?.is_some()))
                {
                    return Ok(true);
                }
            },
            litchi_docx::Inline::Unknown(_) | litchi_docx::Inline::Hyperlink(_) => return Ok(true),
            _ => return Ok(true),
        }
    }
    Ok(false)
}

fn validate_note_paragraph(paragraph: &litchi_docx::Paragraph) -> Result<()> {
    for inline in paragraph.inlines()? {
        let litchi_docx::Inline::Run(run) = inline else {
            return Err(unrepresentable(
                "Markdown export cannot preserve a non-run DOCX note child",
            ));
        };
        for content in run.contents()? {
            match content {
                litchi_docx::RunContent::Text(_)
                | litchi_docx::RunContent::Tab
                | litchi_docx::RunContent::CarriageReturn
                | litchi_docx::RunContent::NoBreakHyphen
                | litchi_docx::RunContent::SoftHyphen
                | litchi_docx::RunContent::FootnoteMark
                | litchi_docx::RunContent::EndnoteMark => {},
                litchi_docx::RunContent::Break(run_break)
                    if matches!(
                        run_break.break_type,
                        litchi_docx::RunBreakType::TextWrapping
                    ) => {},
                litchi_docx::RunContent::Unknown(_) => {
                    return Err(unrepresentable(
                        "Markdown export cannot preserve an unmodeled DOCX note run child",
                    ));
                },
                _ => {
                    return Err(unrepresentable(
                        "Markdown export cannot preserve rich DOCX note content",
                    ));
                },
            }
        }
    }
    Ok(())
}

fn append_indented_definition(output: &mut String, body: &str) {
    let mut first = true;
    for line in body.split('\n') {
        if first {
            first = false;
        } else {
            output.push('\n');
            output.push_str("    ");
        }
        output.push_str(line);
    }
}
