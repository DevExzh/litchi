//! Textual views of an ODT document, built from the leaf body index.

use std::fmt;

use litchi_odt::body::{Attribute, Body, Kind as LeafKind, Node as LeafNode};
use serde::Serialize;

use super::{Odt, core_error};
use crate::error::{Code, Error, Result};
use crate::view::Mode;

/// A compact OfficeCLI-shaped JSON view of the body.
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
    headings: usize,
    tables: usize,
    elements: Vec<HeadingJson>,
}

#[derive(Debug, Clone, Serialize)]
struct HeadingJson {
    path: String,
    level: u32,
    text: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct StatsJson {
    paragraphs: usize,
    headings: usize,
    tables: usize,
    lists: usize,
    words: usize,
    #[serde(rename = "totalCharacters")]
    total_characters: usize,
    #[serde(rename = "emptyParagraphs")]
    empty_paragraphs: usize,
}

/// One top-level block with its positional path.
struct Block<'a> {
    path: String,
    node: LeafNode<'a>,
}

/// The body's direct blocks with their paths (`/body/p[N]`, `/body/h[N]`,
/// `/body/tbl[N]`, ...).
fn blocks(body: Body<'_>) -> Vec<Block<'_>> {
    let mut counters = [0usize; super::Kind::COUNT];
    let mut out = Vec::new();
    for node in body.root().children() {
        let Some(kind) = super::Kind::of(node.kind()) else {
            continue;
        };
        let counter = &mut counters[kind as usize];
        *counter += 1;
        use crate::path::navigate::ElementKind;
        out.push(Block {
            path: format!("/body/{}[{}]", kind.name(), counter),
            node,
        });
    }
    out
}

fn block_text(node: LeafNode<'_>) -> Result<String> {
    match node.kind() {
        LeafKind::Table => {
            let mut rows = Vec::new();
            for row in node
                .children()
                .filter(|child| child.kind() == LeafKind::Row)
            {
                let mut cells = Vec::new();
                for cell in row.children() {
                    cells.push(cell.text().map_err(core_error)?.into_owned());
                }
                rows.push(cells.join("\t"));
            }
            Ok(rows.join("\n"))
        },
        LeafKind::List | LeafKind::Section => {
            let mut lines = Vec::new();
            for child in node.children() {
                lines.push(block_text(child)?);
            }
            Ok(lines.join("\n"))
        },
        LeafKind::ListItem => Ok(node.text().map_err(core_error)?.into_owned()),
        _ => Ok(node.text().map_err(core_error)?.into_owned()),
    }
}

/// Build the requested JSON projection directly from the body index.
pub(crate) fn view_json(doc: &Odt, mode: Mode) -> Result<Json> {
    let body = doc.body()?;
    let result = match mode {
        Mode::Text => {
            let blocks = blocks(body);
            let mut elements = Vec::with_capacity(blocks.len());
            for block in &blocks {
                elements.push(TextElement {
                    path: block.path.clone(),
                    kind: super::Kind::of(block.node.kind())
                        .map_or("element", super::Kind::type_name),
                    text: block_text(block.node)?,
                });
            }
            Json::Text(TextJson {
                total_elements: elements.len(),
                elements,
            })
        },
        Mode::Annotated => {
            let mut content = String::new();
            view_to(doc, mode, &mut content)?;
            Json::Annotated(AnnotatedJson {
                view: "annotated",
                content,
            })
        },
        Mode::Outline => {
            let stats = stats(body)?;
            let mut elements = Vec::new();
            for (path, node) in headings(body) {
                elements.push(HeadingJson {
                    path,
                    level: level(node),
                    text: node.text().map_err(core_error)?.into_owned(),
                });
            }
            Json::Outline(OutlineJson {
                paragraphs: stats.paragraphs,
                headings: stats.headings,
                tables: stats.tables,
                elements,
            })
        },
        Mode::Stats => Json::Stats(stats(body)?),
        _ => return Err(unsupported(mode)),
    };
    doc.check_source()?;
    Ok(result)
}

fn level(node: LeafNode<'_>) -> u32 {
    node.attribute(Attribute::OutlineLevel)
        .and_then(|level| level.parse().ok())
        .unwrap_or(1)
}

/// Every heading in document order with its positional path.
fn headings(body: Body<'_>) -> Vec<(String, LeafNode<'_>)> {
    blocks(body)
        .into_iter()
        .filter(|block| block.node.kind() == LeafKind::Heading)
        .map(|block| (block.path, block.node))
        .collect()
}

fn stats(body: Body<'_>) -> Result<StatsJson> {
    let mut stats = StatsJson {
        paragraphs: 0,
        headings: 0,
        tables: 0,
        lists: 0,
        words: 0,
        total_characters: 0,
        empty_paragraphs: 0,
    };
    for node in body.nodes() {
        match node.kind() {
            LeafKind::Paragraph | LeafKind::Heading => {
                if node.kind() == LeafKind::Heading {
                    stats.headings += 1;
                } else {
                    stats.paragraphs += 1;
                }
                let text = node.text().map_err(core_error)?;
                stats.words += text.split_whitespace().count();
                stats.total_characters += text.chars().count();
                stats.empty_paragraphs += usize::from(text.is_empty());
            },
            LeafKind::Table => stats.tables += 1,
            LeafKind::List => stats.lists += 1,
            _ => {},
        }
    }
    Ok(stats)
}

/// Stream document views from the body index.
pub(crate) fn view_to<W: fmt::Write>(doc: &Odt, mode: Mode, out: &mut W) -> Result<()> {
    let body = doc.body()?;
    match mode {
        Mode::Text => {
            for block in blocks(body) {
                writeln!(out, "{}", block_text(block.node)?).map_err(sink)?;
            }
        },
        Mode::Annotated => {
            for block in blocks(body) {
                writeln!(out, "[{}] {}", block.path, block_text(block.node)?).map_err(sink)?;
            }
        },
        Mode::Outline => {
            let stats = stats(body)?;
            writeln!(
                out,
                "{} paragraphs, {} headings, {} tables",
                stats.paragraphs, stats.headings, stats.tables
            )
            .map_err(sink)?;
            for (path, node) in headings(body) {
                let indent = "  ".repeat(level(node).saturating_sub(1).min(9) as usize);
                writeln!(
                    out,
                    "{indent}├── [{path}] \"{}\"",
                    node.text().map_err(core_error)?
                )
                .map_err(sink)?;
            }
        },
        Mode::Stats => {
            let stats = stats(body)?;
            writeln!(out, "paragraphs: {}", stats.paragraphs).map_err(sink)?;
            writeln!(out, "headings: {}", stats.headings).map_err(sink)?;
            writeln!(out, "tables: {}", stats.tables).map_err(sink)?;
            writeln!(out, "words: {}", stats.words).map_err(sink)?;
            writeln!(out, "characters: {}", stats.total_characters).map_err(sink)?;
        },
        _ => return Err(unsupported(mode)),
    }
    doc.check_source()
}

fn unsupported(mode: Mode) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        format!("the `{}` view is not supported for ODT", mode.name()),
    )
    .with_format(litchi_core::FileFormat::Odt)
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the ODT view output sink refused to write")
        .with_format(litchi_core::FileFormat::Odt)
}
