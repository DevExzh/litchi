//! The record-0323 element surface of legacy PowerPoint: slide comments,
//! table rows, cells and columns, hyperlinks, text runs, line breaks, the
//! theme and charts.
//!
//! Each element is a value read from the native presentation projection when
//! the node is resolved; the node keeps it so `format` and `text` answer
//! without another parse. Writes go through the native primary edit
//! (`format::ppt::write`).

use std::borrow::Cow;
use std::sync::Arc;

use litchi_ppt::ParsedComment;

use super::node::{Node, ShapeItem, length};
use super::spec::Kind;
use crate::Result;
use crate::node::{Preview, Value};

/// One record-0323 element.
#[derive(Clone)]
pub(crate) enum Element {
    /// A slide comment: its slide, its zero-based position among the slide's
    /// comments, and the parsed comment.
    Comment {
        slide: usize,
        index: usize,
        comment: Arc<ParsedComment>,
    },
    /// A table row: the table shape and the zero-based row.
    Row { shape: ShapeItem, row: usize },
    /// A table cell: the table shape, row and column.
    Cell {
        shape: ShapeItem,
        row: usize,
        column: usize,
    },
    /// A table column: the table shape and the zero-based column.
    Column { shape: ShapeItem, column: usize },
}

/// The values one element's `format` projection answers, in its property
/// table's order.
pub(crate) struct Values<'a> {
    values: Vec<Option<Value<'a>>>,
}

impl<'a> Values<'a> {
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        self.values.get(index).cloned().flatten()
    }
}

impl Element {
    pub(crate) const fn kind(&self) -> Kind {
        match self {
            Self::Comment { .. } => Kind::Comment,
            Self::Row { .. } => Kind::Row,
            Self::Cell { .. } => Kind::Cell,
            Self::Column { .. } => Kind::Col,
        }
    }

    pub(crate) const fn slide(&self) -> Option<usize> {
        match self {
            Self::Comment { slide, .. } => Some(*slide),
            Self::Row { shape, .. } | Self::Cell { shape, .. } | Self::Column { shape, .. } => {
                Some(shape.slide)
            },
        }
    }

    pub(crate) const fn shape_item(&self) -> Option<&ShapeItem> {
        match self {
            Self::Comment { .. } => None,
            Self::Row { shape, .. } | Self::Cell { shape, .. } | Self::Column { shape, .. } => {
                Some(shape)
            },
        }
    }

    pub(crate) fn text<'n>(&'n self, node: &'n Node<'_>) -> Result<Option<Cow<'n, str>>> {
        Ok(match self {
            Self::Comment { comment, .. } => Some(Cow::Borrowed(comment.text.as_str())),
            Self::Cell { row, column, .. } => table(node)?
                .cell(*row, *column)
                .map(|text| Cow::Owned(text.to_owned())),
            Self::Row { .. } | Self::Column { .. } => None,
        })
    }

    pub(crate) fn preview<'n>(&'n self, node: &'n Node<'_>) -> Result<Option<Preview<'n>>> {
        Ok(self.text(node)?.map(Preview::Prefix))
    }

    pub(crate) fn child_count(&self, node: &Node<'_>) -> Result<usize> {
        Ok(match self {
            Self::Row { .. } => table(node)?.columns(),
            Self::Comment { .. } | Self::Cell { .. } | Self::Column { .. } => 0,
        })
    }

    pub(crate) fn values<'n>(&'n self, node: &'n Node<'_>) -> Result<Values<'n>> {
        let props = self.kind().spec().props;
        let values = match self {
            Self::Comment { comment, .. } => props
                .iter()
                .map(|prop| comment_value(comment, prop.name))
                .collect(),
            Self::Row { row, .. } => {
                let heights = table(node)?.row_heights();
                props
                    .iter()
                    .map(|prop| match prop.name {
                        "height" => heights.get(*row).copied().map(length),
                        _ => None,
                    })
                    .collect()
            },
            Self::Column { column, .. } => {
                let widths = table(node)?.column_widths();
                props
                    .iter()
                    .map(|prop| match prop.name {
                        "width" => widths.get(*column).copied().map(length),
                        _ => None,
                    })
                    .collect()
            },
            Self::Cell { .. } => Vec::new(),
        };
        Ok(Values { values })
    }
}

fn comment_value<'a>(comment: &'a ParsedComment, key: &str) -> Option<Value<'a>> {
    match key {
        "text" => Some(Value::Text(Cow::Borrowed(comment.text.as_str()))),
        "author" => Some(Value::Text(Cow::Borrowed(comment.author.as_str()))),
        "initials" => Some(Value::Text(Cow::Borrowed(comment.initials.as_str()))),
        "index" => Some(Value::Int(i64::from(comment.index))),
        "date" => comment_date(comment).map(|date| Value::Text(Cow::Owned(date))),
        "x" => Some(length(comment.x)),
        "y" => Some(length(comment.y)),
        _ => None,
    }
}

/// The comment's `SYSTEMTIME` as `YYYY-MM-DDTHH:MM:SS.fff`; an all-zero date
/// (no creation time) reads as absent.
fn comment_date(comment: &ParsedComment) -> Option<String> {
    (comment.year != 0).then(|| {
        litchi_ppt::comments::DateTime {
            year: comment.year,
            month: comment.month,
            day: comment.day,
            hour: comment.hour,
            minute: comment.minute,
            second: comment.second,
            millisecond: comment.millisecond,
        }
        .to_string()
    })
}

/// A slide's comments, read from the native slide.
pub(crate) fn slide_comments(file: &super::Ppt, slide: usize) -> Result<Vec<ParsedComment>> {
    let selected = file.reader().slide_at(slide)?.ok_or_else(|| {
        crate::Error::new(
            crate::Code::InternalError,
            format!("slide {} vanished from the catalog", slide + 1),
        )
        .with_format(litchi_core::FileFormat::Ppt)
    })?;
    Ok(selected.comments()?)
}

/// The native table behind a table, row, cell or column node.
pub(crate) fn table<'n>(
    node: &'n Node<'_>,
) -> Result<&'n litchi_ppt::shapes::shape_enum::TableShape> {
    node.shape()
        .and_then(litchi_ppt::shapes::ShapeEnum::as_table)
        .ok_or_else(|| {
            crate::Error::new(
                crate::Code::InternalError,
                "table vanished from the native scene",
            )
            .with_format(litchi_core::FileFormat::Ppt)
        })
}
