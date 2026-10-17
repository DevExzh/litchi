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

use super::node::{Node, length};
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
        }
    }

    pub(crate) const fn slide(&self) -> Option<usize> {
        match self {
            Self::Comment { slide, .. } => Some(*slide),
        }
    }

    pub(crate) const fn shape_item(&self) -> Option<&super::node::ShapeItem> {
        match self {
            Self::Comment { .. } => None,
        }
    }

    pub(crate) fn text<'n>(&'n self, _node: &'n Node<'_>) -> Result<Option<Cow<'n, str>>> {
        Ok(match self {
            Self::Comment { comment, .. } => Some(Cow::Borrowed(comment.text.as_str())),
        })
    }

    pub(crate) fn preview<'n>(&'n self, node: &'n Node<'_>) -> Result<Option<Preview<'n>>> {
        Ok(self.text(node)?.map(Preview::Prefix))
    }

    pub(crate) fn child_count(&self, _node: &Node<'_>) -> Result<usize> {
        Ok(0)
    }

    pub(crate) fn values<'n>(&'n self, _node: &'n Node<'_>) -> Result<Values<'n>> {
        let table = self.kind().spec().props;
        let values = table
            .iter()
            .map(|prop| match self {
                Self::Comment { comment, .. } => comment_value(comment, prop.name),
            })
            .collect();
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
