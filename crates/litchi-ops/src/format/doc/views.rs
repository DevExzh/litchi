//! Textual views of a legacy DOC document.

use std::fmt;

use super::Doc;
use crate::error::{Code, Error, Result};
use crate::view::Mode;
use serde::Serialize;

/// A compact OfficeCLI-shaped JSON view for the DOC paragraph projection.
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub(crate) enum Json {
    Text(TextJson),
    Annotated(AnnotatedJson),
    Outline(OutlineJson),
    Stats(StatsJson),
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct TextJson {
    #[serde(rename = "totalElements")]
    total_elements: usize,
    elements: Vec<TextElement>,
}

#[derive(Debug, Clone, Serialize)]
struct TextElement {
    path: String,
    #[serde(rename = "type")]
    kind: &'static str,
    text: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct AnnotatedJson {
    view: &'static str,
    content: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct OutlineJson {
    paragraphs: usize,
    elements: Vec<ParagraphJson>,
}

#[derive(Debug, Clone, Serialize)]
struct ParagraphJson {
    number: usize,
    text: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct StatsJson {
    paragraphs: usize,
    words: usize,
    #[serde(rename = "totalCharacters")]
    total_characters: usize,
    #[serde(rename = "emptyParagraphs")]
    empty_paragraphs: usize,
    #[serde(rename = "consecutiveSpaces")]
    consecutive_spaces: usize,
}

/// Build the typed JSON view without an intermediate textual JSON round trip.
pub(crate) fn view_json(doc: &Doc, mode: Mode) -> Result<Json> {
    let paragraphs = doc.body_paragraph_texts()?;
    match mode {
        Mode::Text => Ok(Json::Text(TextJson {
            total_elements: paragraphs.len(),
            elements: paragraphs
                .iter()
                .enumerate()
                .map(|(index, text)| TextElement {
                    path: format!("/body/p[{}]", index + 1),
                    kind: "paragraph",
                    text: text.clone(),
                })
                .collect(),
        })),
        Mode::Annotated => {
            let mut content = String::new();
            view_to(doc, mode, &mut content)?;
            Ok(Json::Annotated(AnnotatedJson {
                view: "annotated",
                content,
            }))
        },
        Mode::Outline => Ok(Json::Outline(OutlineJson {
            paragraphs: paragraphs.len(),
            elements: paragraphs
                .iter()
                .enumerate()
                .map(|(index, text)| ParagraphJson {
                    number: index + 1,
                    text: text.clone(),
                })
                .collect(),
        })),
        Mode::Stats => {
            let words = paragraphs
                .iter()
                .map(|text| text.split_whitespace().count())
                .sum();
            let total_characters = paragraphs.iter().map(|text| text.chars().count()).sum();
            let empty_paragraphs = paragraphs.iter().filter(|text| text.is_empty()).count();
            let consecutive_spaces = paragraphs
                .iter()
                .map(|text| text.matches("  ").count())
                .sum();
            Ok(Json::Stats(StatsJson {
                paragraphs: paragraphs.len(),
                words,
                total_characters,
                empty_paragraphs,
                consecutive_spaces,
            }))
        },
        _ => Err(Error::new(
            Code::UnsupportedCapability,
            format!("the `{}` view is not supported for DOC", mode.name()),
        )
        .with_format(litchi_core::FileFormat::Doc)),
    }
}

/// Render the four document modes supported by the tier-2 DOC adapter.
pub(crate) fn view_to<W: fmt::Write>(doc: &Doc, mode: Mode, out: &mut W) -> Result<()> {
    let paragraphs = doc.body_paragraph_texts()?;
    match mode {
        Mode::Text => {
            for paragraph in paragraphs.iter() {
                writeln!(out, "{paragraph}").map_err(sink)?;
            }
        },
        Mode::Annotated => {
            for (index, paragraph) in paragraphs.iter().enumerate() {
                writeln!(out, "[/body/p[{}]] {paragraph}", index + 1).map_err(sink)?;
            }
        },
        Mode::Outline => {
            writeln!(out, "{} paragraphs", paragraphs.len()).map_err(sink)?;
            for (index, paragraph) in paragraphs.iter().enumerate() {
                writeln!(out, "├── [{}] \"{paragraph}\"", index + 1).map_err(sink)?;
            }
        },
        Mode::Stats => {
            let mut words = 0usize;
            let mut characters = 0usize;
            for paragraph in &paragraphs {
                words = words.saturating_add(paragraph.split_whitespace().count());
                characters = characters.saturating_add(paragraph.chars().count());
            }
            writeln!(out, "paragraphs: {}", paragraphs.len()).map_err(sink)?;
            writeln!(out, "words: {words}").map_err(sink)?;
            writeln!(out, "characters: {characters}").map_err(sink)?;
        },
        _ => {
            return Err(Error::new(
                Code::UnsupportedCapability,
                format!("the `{}` view is not supported for DOC", mode.name()),
            )
            .with_format(litchi_core::FileFormat::Doc));
        },
    }
    Ok(())
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the DOC view output sink refused to write")
        .with_format(litchi_core::FileFormat::Doc)
}
