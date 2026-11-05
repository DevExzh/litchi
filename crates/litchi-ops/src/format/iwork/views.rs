//! Textual views of an iWork object tree: text, annotated text, outline and
//! statistics, in the shapes of the ODF object-tree adapter (record 0282).

use std::collections::BTreeMap;
use std::fmt;

use serde::Serialize;

use super::Iwork;
use super::tree::{At, Item};
use crate::error::{Code, Error, Result};
use crate::node::NodePath;
use crate::view::Mode;

/// A JSON view of an object tree.
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
    elements: Vec<OutlineElement>,
}

#[derive(Debug, Clone, Serialize)]
struct OutlineElement {
    path: String,
    #[serde(rename = "type")]
    kind: &'static str,
    depth: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    name: Option<String>,
    #[serde(rename = "childCount")]
    child_count: usize,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct StatsJson {
    elements: usize,
    #[serde(rename = "byType")]
    by_type: BTreeMap<&'static str, usize>,
    words: usize,
    #[serde(rename = "totalCharacters")]
    total_characters: usize,
}

fn path(item: &Item) -> String {
    let path = NodePath::new(item.path()).to_string();
    if path.is_empty() {
        "/".to_owned()
    } else {
        path
    }
}

fn depth(item: &Item) -> usize {
    item.path().len()
}

/// The `mode` view as JSON.
pub(crate) fn view_json(file: &Iwork, mode: Mode) -> Result<Json> {
    file.checked_tree()?;
    let items = file.ordered();
    match mode {
        Mode::Text => {
            let elements: Vec<TextElement> = items
                .iter()
                .filter_map(|(_, item)| {
                    let text = item.text().filter(|text| !text.is_empty())?;
                    Some(TextElement {
                        path: path(item),
                        kind: item.type_name(),
                        text: text.to_owned(),
                    })
                })
                .collect();
            Ok(Json::Text(TextJson {
                total_elements: elements.len(),
                elements,
            }))
        },
        Mode::Annotated => {
            let mut content = String::new();
            view_to(file, mode, &mut content)?;
            Ok(Json::Annotated(AnnotatedJson {
                view: "annotated",
                content,
            }))
        },
        Mode::Outline => Ok(Json::Outline(OutlineJson {
            elements: items
                .iter()
                .map(|(at, item)| OutlineElement {
                    path: path(item),
                    kind: item.type_name(),
                    depth: depth(item),
                    name: item.label().map(str::to_owned),
                    child_count: file.kids(*at).len(),
                })
                .collect(),
        })),
        Mode::Stats => Ok(Json::Stats(stats(&items))),
        _ => Err(unsupported(file, mode)),
    }
}

fn stats(items: &[(At, &Item)]) -> StatsJson {
    let mut stats = StatsJson {
        elements: items.len(),
        by_type: BTreeMap::new(),
        words: 0,
        total_characters: 0,
    };
    for (_, item) in items {
        *stats.by_type.entry(item.type_name()).or_default() += 1;
        if let Some(text) = item.text() {
            stats.words += text.split_whitespace().count();
            stats.total_characters += text.chars().count();
        }
    }
    stats
}

/// Stream the `mode` view as text.
pub(crate) fn view_to<W: fmt::Write>(file: &Iwork, mode: Mode, out: &mut W) -> Result<()> {
    file.checked_tree()?;
    let items = file.ordered();
    let sink = |_: fmt::Error| {
        Error::new(Code::IoError, "the view output sink refused to write")
            .with_format(file.format())
    };
    match mode {
        Mode::Text => {
            for text in items
                .iter()
                .filter_map(|(_, item)| item.text())
                .filter(|text| !text.is_empty())
            {
                writeln!(out, "{text}").map_err(sink)?;
            }
        },
        Mode::Annotated => {
            for (_, item) in &items {
                if let Some(text) = item.text().filter(|text| !text.is_empty()) {
                    writeln!(out, "[{}] {text}", path(item)).map_err(sink)?;
                }
            }
        },
        Mode::Outline => {
            for (_, item) in &items {
                let indent = "  ".repeat(depth(item));
                match item.label() {
                    Some(name) => writeln!(out, "{indent}{} \"{name}\"", path(item)),
                    None => writeln!(out, "{indent}{}", path(item)),
                }
                .map_err(sink)?;
            }
        },
        Mode::Stats => {
            let stats = stats(&items);
            writeln!(out, "elements: {}", stats.elements).map_err(sink)?;
            for (kind, count) in &stats.by_type {
                writeln!(out, "{kind}: {count}").map_err(sink)?;
            }
            writeln!(out, "words: {}", stats.words).map_err(sink)?;
            writeln!(out, "characters: {}", stats.total_characters).map_err(sink)?;
        },
        _ => return Err(unsupported(file, mode)),
    }
    Ok(())
}

fn unsupported(file: &Iwork, mode: Mode) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        format!(
            "the `{}` view is not supported for this format",
            mode.name()
        ),
    )
    .with_format(file.format())
}
