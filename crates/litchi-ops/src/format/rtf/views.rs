//! Textual views of a RTF document.

use std::fmt;

use super::Rtf;
use crate::error::{Code, Error, Result};
use crate::view::Mode;
use serde::Serialize;

/// A compact OfficeCLI-shaped JSON view for the RTF paragraph projection.
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

/// Build the requested JSON projection directly from native paragraph views.
pub(crate) fn view_json(doc: &Rtf, mode: Mode) -> Result<Json> {
    let paragraphs = doc.0.package.body().paragraphs();
    let count = doc.0.package.paragraph_count();
    match mode {
        Mode::Text => Ok(Json::Text(TextJson {
            total_elements: count,
            elements: paragraphs
                .enumerate()
                .map(|(index, paragraph)| TextElement {
                    path: format!("/body/p[{}]", index + 1),
                    kind: "paragraph",
                    text: paragraph.to_text(),
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
            paragraphs: count,
            elements: paragraphs
                .enumerate()
                .map(|(index, paragraph)| ParagraphJson {
                    number: index + 1,
                    text: paragraph.to_text(),
                })
                .collect(),
        })),
        Mode::Stats => Ok(Json::Stats(stats(doc))),
        _ => Err(unsupported(mode)),
    }
}

fn stats(doc: &Rtf) -> StatsJson {
    let mut stats = StatsJson {
        paragraphs: doc.0.package.paragraph_count(),
        words: 0,
        total_characters: 0,
        empty_paragraphs: 0,
        consecutive_spaces: 0,
    };
    for paragraph in doc.0.package.body().paragraphs() {
        let text = paragraph.as_str().map_or_else(
            || std::borrow::Cow::Owned(paragraph.to_text()),
            std::borrow::Cow::Borrowed,
        );
        stats.words += text.split_whitespace().count();
        stats.total_characters += text.chars().count();
        stats.empty_paragraphs += usize::from(text.is_empty());
        stats.consecutive_spaces += text.matches("  ").count();
    }
    stats
}

/// Stream document views from the immutable native story.
pub(crate) fn view_to<W: fmt::Write>(doc: &Rtf, mode: Mode, out: &mut W) -> Result<()> {
    let paragraphs = doc.0.package.body().paragraphs();
    match mode {
        Mode::Text => {
            for paragraph in paragraphs {
                writeln!(out, "{paragraph}").map_err(sink)?;
            }
        },
        Mode::Annotated => {
            for (index, paragraph) in paragraphs.enumerate() {
                writeln!(out, "[/body/p[{}]] {paragraph}", index + 1).map_err(sink)?;
            }
        },
        Mode::Outline => {
            writeln!(out, "{} paragraphs", doc.0.package.paragraph_count()).map_err(sink)?;
            for (index, paragraph) in paragraphs.enumerate() {
                writeln!(out, "├── [{}] \"{paragraph}\"", index + 1).map_err(sink)?;
            }
        },
        Mode::Stats => {
            let stats = stats(doc);
            writeln!(out, "paragraphs: {}", stats.paragraphs).map_err(sink)?;
            writeln!(out, "words: {}", stats.words).map_err(sink)?;
            writeln!(out, "characters: {}", stats.total_characters).map_err(sink)?;
        },
        _ => return Err(unsupported(mode)),
    }
    Ok(())
}

fn unsupported(mode: Mode) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        format!("the `{}` view is not supported for RTF", mode.name()),
    )
    .with_format(litchi_core::FileFormat::Rtf)
}
fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the RTF view output sink refused to write")
        .with_format(litchi_core::FileFormat::Rtf)
}
