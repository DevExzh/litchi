//! Typed property values of ODT nodes, read from the leaf on demand.

use std::borrow::Cow;

use litchi_core::Metadata;
use litchi_odt::body::{Attribute, Kind as LeafKind, Node as LeafNode, Style};

use super::node::{Node, Target};
use super::spec::Kind;
use crate::error::Result;
use crate::node::Value;
use crate::value::{Color, Length, LengthContext};

/// Where a node's `format` values come from.
#[derive(Debug)]
pub(crate) enum Props<'a> {
    Body(Kind, LeafNodeRef<'a>),
    Style(StyleRef<'a>),
    Metadata(Box<Metadata>),
    /// A part of record 0328.
    Part(PartRef<'a>),
    Empty,
}

/// A part node's handle and part, with a `Debug` that omits content.
#[derive(Clone, Copy)]
pub(crate) struct PartRef<'a> {
    odt: &'a super::Odt,
    kind: Kind,
    part: super::node::Part,
}

impl std::fmt::Debug for PartRef<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PartRef")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

/// A leaf node with a `Debug` that omits authored content.
#[derive(Clone, Copy)]
pub(crate) struct LeafNodeRef<'a>(pub(crate) LeafNode<'a>);

impl std::fmt::Debug for LeafNodeRef<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("LeafNode")
    }
}

/// A style with a `Debug` that omits authored content.
#[derive(Clone, Copy)]
pub(crate) struct StyleRef<'a>(pub(crate) Style<'a>);

impl std::fmt::Debug for StyleRef<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("Style")
    }
}

impl<'a> Props<'a> {
    pub(super) fn of(node: &Node<'a>) -> Result<Self> {
        Ok(match node.target {
            Target::Body(..) => match node.leaf() {
                Some(leaf) => Self::Body(node.kind, LeafNodeRef(leaf)),
                None => Self::Empty,
            },
            Target::Style(style) => Self::Style(StyleRef(style)),
            Target::Part(part) => Self::Part(PartRef {
                odt: node.odt,
                kind: node.kind,
                part,
            }),
            Target::Synthetic if node.kind == Kind::Metadata => {
                Self::Metadata(Box::new(node.odt.metadata()?.unwrap_or_default()))
            },
            Target::Synthetic => Self::Empty,
        })
    }

    /// The value at `index` of the kind's property table.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match self {
            Self::Body(kind, leaf) => body_value(*kind, leaf.0, index),
            Self::Style(style) => style_value(style.0, index),
            Self::Metadata(metadata) => metadata_value(metadata, index),
            Self::Part(part) => super::parts::value(part.odt, part.kind, part.part, index),
            Self::Empty => None,
        }
    }
}

fn text<'a>(value: Option<Cow<'a, str>>) -> Option<Value<'a>> {
    value.map(Value::Text)
}

/// An ODF length (`0.5cm`, `12pt`) as a unified length.
pub(crate) fn length(raw: Option<&str>) -> Option<Value<'static>> {
    Length::parse(raw?).ok().map(Value::Length)
}

fn font_size(raw: Option<&str>) -> Option<Value<'static>> {
    let raw = raw?;
    Length::parse_in(raw, LengthContext::FontSize)
        .or_else(|_| Length::parse(raw))
        .ok()
        .map(Value::Length)
}

/// An ODF `#rrggbb` colour.
pub(crate) fn color(raw: Option<&str>) -> Option<Value<'static>> {
    let hex = raw?.strip_prefix('#')?;
    if hex.len() != 6 {
        return None;
    }
    let channel = |offset: usize| u8::from_str_radix(hex.get(offset..offset + 2)?, 16).ok();
    Some(Value::Color(Color::rgb(
        channel(0)?,
        channel(2)?,
        channel(4)?,
    )))
}

fn align(raw: Option<&str>) -> Option<Value<'static>> {
    Some(Value::Token(match raw? {
        "left" => "left",
        "right" => "right",
        "center" => "center",
        "justify" => "justify",
        _ => return None,
    }))
}

fn int(raw: Option<Cow<'_, str>>) -> Option<Value<'static>> {
    raw?.parse::<i64>().ok().map(Value::Int)
}

fn count(value: usize) -> Value<'static> {
    Value::Int(i64::try_from(value).unwrap_or(i64::MAX))
}

/// Character properties at `index` of the shared character block.
fn character<'a>(leaf: LeafNode<'a>, index: usize) -> Option<Value<'a>> {
    let properties = leaf.text_properties();
    match index {
        0 => properties.bold().map(Value::Bool),
        1 => properties.italic().map(Value::Bool),
        2 => properties.underline().map(Value::Bool),
        3 => properties.strike().map(Value::Bool),
        4 => font_size(properties.size()),
        5 => color(properties.color()),
        6 => properties
            .font()
            .map(|font| Value::Text(Cow::Owned(font.to_owned()))),
        7 => color(properties.highlight()),
        _ => None,
    }
}

fn body_value(kind: Kind, leaf: LeafNode<'_>, index: usize) -> Option<Value<'_>> {
    match kind {
        Kind::Paragraph | Kind::Heading => paragraph_value(leaf, index),
        Kind::Run => match index {
            0 => leaf.text().ok().map(Value::Text),
            1 => text(leaf.common_style_name()),
            2..=9 => character(leaf, index - 2),
            10 => leaf.text_properties().superscript().map(Value::Bool),
            11 => leaf.text_properties().subscript().map(Value::Bool),
            12 => leaf.text_properties().caps().map(Value::Bool),
            13 => leaf.text_properties().small_caps().map(Value::Bool),
            _ => None,
        },
        Kind::Hyperlink => match index {
            0 => leaf.text().ok().map(Value::Text),
            1 => text(leaf.attribute(Attribute::Href)),
            2 => text(leaf.attribute(Attribute::Title)),
            3 => text(leaf.attribute(Attribute::TargetFrame)),
            _ => None,
        },
        Kind::List => match index {
            0 => text(leaf.style_name()),
            _ => None,
        },
        Kind::ListItem | Kind::TextBox => match index {
            0 => leaf.text().ok().map(Value::Text),
            _ => None,
        },
        Kind::Table => match index {
            0 => text(leaf.attribute(Attribute::Name)),
            1 => Some(count(
                leaf.children()
                    .filter(|child| child.kind() == LeafKind::Row)
                    .count(),
            )),
            2 => Some(count(super::write::table_columns(leaf))),
            3 => text(leaf.style_name()),
            _ => None,
        },
        Kind::Row => match index {
            0 => Some(count(leaf.children().count())),
            _ => None,
        },
        Kind::Cell => match index {
            0 => leaf.text().ok().map(Value::Text),
            1 => int(leaf.attribute(Attribute::ColumnsSpanned)),
            2 => int(leaf.attribute(Attribute::RowsSpanned)),
            3 => Some(Value::Bool(leaf.kind() == LeafKind::CoveredCell)),
            4 => text(leaf.attribute(Attribute::ValueType)),
            5 => color(leaf.cell_properties().shading()),
            6 => leaf
                .cell_properties()
                .vertical_align()
                .map(|value| Value::Text(Cow::Owned(value.to_owned()))),
            _ => None,
        },
        Kind::Section => match index {
            0 => text(leaf.attribute(Attribute::Name)),
            1 => leaf
                .attribute(Attribute::Protected)
                .map(|value| Value::Bool(value == "true")),
            2 => leaf
                .attribute(Attribute::Display)
                .map(|value| Value::Bool(value == "none")),
            _ => None,
        },
        Kind::Frame => match index {
            0 => text(leaf.attribute(Attribute::Name)),
            1 => length(leaf.attribute(Attribute::Width).as_deref()),
            2 => length(leaf.attribute(Attribute::Height).as_deref()),
            3 => length(leaf.attribute(Attribute::X).as_deref()),
            4 => length(leaf.attribute(Attribute::Y).as_deref()),
            5 => leaf.attribute(Attribute::AnchorType).and_then(|anchor| {
                Some(Value::Token(match &*anchor {
                    "paragraph" => "paragraph",
                    "char" => "char",
                    "as-char" => "as-char",
                    "page" => "page",
                    "frame" => "frame",
                    _ => return None,
                }))
            }),
            _ => None,
        },
        Kind::Picture => match index {
            0 => text(leaf.attribute(Attribute::Href)),
            _ => None,
        },
        Kind::Footnote | Kind::Endnote => match index {
            0 => leaf.text().ok().map(Value::Text),
            1 => text(leaf.citation()),
            2 => Some(Value::Token(if leaf.kind() == LeafKind::Endnote {
                "endnote"
            } else {
                "footnote"
            })),
            _ => None,
        },
        Kind::Comment => match index {
            0 => leaf.text().ok().map(Value::Text),
            1 => text(leaf.author()),
            2 => text(leaf.date()),
            _ => None,
        },
        Kind::Bookmark => match index {
            0 => text(leaf.attribute(Attribute::Name)),
            _ => None,
        },
        _ => None,
    }
}

fn paragraph_value(leaf: LeafNode<'_>, index: usize) -> Option<Value<'_>> {
    let paragraph = || leaf.paragraph_properties();
    match index {
        0 => leaf.text().ok().map(Value::Text),
        1 => text(leaf.common_style_name()),
        2 => align(paragraph().align()),
        3 => length(paragraph().space_before()),
        4 => length(paragraph().space_after()),
        5 => length(paragraph().indent()),
        6 => length(paragraph().right_indent()),
        7 => length(paragraph().first_line_indent()),
        8 => paragraph()
            .line_height()
            .map(|value| Value::Text(Cow::Owned(value.to_owned()))),
        9 => paragraph().keep_next().map(Value::Bool),
        10 => paragraph().page_break_before().map(Value::Bool),
        11 => color(paragraph().shading()),
        12..=19 => character(leaf, index - 12),
        20 => int(leaf.attribute(Attribute::OutlineLevel)),
        _ => None,
    }
}

fn style_value(style: Style<'_>, index: usize) -> Option<Value<'_>> {
    match index {
        0 => text(style.name()),
        1 => text(style.display_name()),
        2 => text(style.family()),
        3 => text(style.parent_name()),
        4 => align(style.paragraph_properties().align()),
        5 => style.text_properties().bold().map(Value::Bool),
        6 => style.text_properties().italic().map(Value::Bool),
        7 => font_size(style.text_properties().size()),
        8 => color(style.text_properties().color()),
        9 => style
            .text_properties()
            .font()
            .map(|font| Value::Text(Cow::Owned(font.to_owned()))),
        _ => None,
    }
}

fn metadata_value(metadata: &Metadata, index: usize) -> Option<Value<'static>> {
    let owned = |value: &Option<String>| value.clone().map(|value| Value::Text(Cow::Owned(value)));
    match index {
        0 => owned(&metadata.title),
        1 => owned(&metadata.subject),
        2 => owned(&metadata.description),
        3 => owned(&metadata.author),
        4 => owned(&metadata.last_modified_by),
        5 => owned(&metadata.language),
        6 => Value::owned_list(&metadata.keywords),
        7 => metadata
            .created
            .map(|at| Value::Text(Cow::Owned(at.to_string()))),
        8 => metadata
            .modified
            .map(|at| Value::Text(Cow::Owned(at.to_string()))),
        9 => owned(&metadata.application),
        10 => owned(&metadata.revision),
        _ => None,
    }
}
