//! `listStyle` on paragraphs (record 0251e).
//!
//! `set <paragraph> listStyle=bullet|ordered|none [start=N]` follows
//! OfficeCLI's `ApplyListStyle`:
//!
//! - `bullet` (also `unordered`, `ul`) or `ordered` (any other value)
//!   continues the list of the immediately preceding paragraph when that
//!   list has the same kind, and otherwise adds a new list definition
//!   (`litchi_docx::numbering::ListKind`);
//! - `start` always adds a new list, restarting level 0 at that value;
//! - `none` (also `remove`, `clear`) removes the paragraph's numbering.
//!
//! The paragraph keeps its list level, or takes level 0.

use std::collections::HashMap;

use litchi_docx::numbering::ListKind;

use super::node::{Node, Value};
use super::spec::Kind;
use crate::error::{Code, Error, Result};

/// The list a `listStyle` assignment asks for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Style {
    List(ListKind),
    None,
}

/// A paragraph's `listStyle` request: the style and an optional restart.
#[derive(Debug, Clone, Copy)]
pub(super) struct Request {
    pub(super) style: Option<Style>,
    pub(super) start: Option<u32>,
}

/// The style a `listStyle` value names.
pub(super) fn style(text: &str) -> Style {
    let text = text.trim();
    if ["bullet", "unordered", "ul"]
        .iter()
        .any(|name| text.eq_ignore_ascii_case(name))
    {
        Style::List(ListKind::Bullet)
    } else if ["none", "remove", "clear"]
        .iter()
        .any(|name| text.eq_ignore_ascii_case(name))
    {
        Style::None
    } else {
        Style::List(ListKind::Decimal)
    }
}

/// What precedes a target paragraph in its container.
enum Prior {
    /// No paragraph, or a paragraph outside any list.
    None,
    /// An earlier target of the same write.
    Target(usize),
    /// A paragraph in the list `num_id`.
    List(u32),
}

/// The numbering each target joins: `(numId, level)`, or `None` to remove
/// its numbering. New definitions are staged on `edit`.
pub(super) fn plan(
    nodes: &[Node<'_>],
    request: Request,
    edit: &mut litchi_docx::document::Edit,
) -> Result<Vec<Option<(u32, u8)>>> {
    let style = match (request.style, request.start) {
        (Some(style), _) => style,
        // `start` alone restarts the paragraph's own list kind.
        (None, Some(_)) => Style::List(ListKind::Decimal),
        (None, None) => return Ok(Vec::new()),
    };
    let mut kinds: HashMap<u32, Option<ListKind>> = HashMap::new();
    let mut planned: Vec<Option<(u32, u8)>> = Vec::with_capacity(nodes.len());
    for (index, node) in nodes.iter().enumerate() {
        let Value::Paragraph { paragraph, .. } = &node.value else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "listStyle applies to paragraphs",
            )
            .with_property("listStyle"));
        };
        let own = paragraph.numbering().map_err(Error::from)?;
        let level = own.map_or(0, |numbering| numbering.level);
        let Style::List(kind) = style else {
            planned.push(None);
            continue;
        };
        let continued = if request.start.is_some() {
            None
        } else {
            match prior(nodes, index)? {
                Prior::None => None,
                Prior::Target(earlier) => planned[earlier].map(|(num_id, _)| num_id),
                Prior::List(num_id) => {
                    let known = match kinds.get(&num_id) {
                        Some(known) => *known,
                        None => {
                            let known = edit.list_kind(num_id).map_err(Error::from)?;
                            kinds.insert(num_id, known);
                            known
                        },
                    };
                    (known == Some(kind)).then_some(num_id)
                },
            }
        };
        let num_id = match continued {
            Some(num_id) => num_id,
            None => {
                let list = edit
                    .add_list(kind, request.start.map(|start| (0, start)))
                    .map_err(Error::from)?;
                kinds.insert(list.num_id(), Some(kind));
                list.num_id()
            },
        };
        planned.push(Some((num_id, level)));
    }
    Ok(planned)
}

/// The paragraph immediately before target `index` in its container.
fn prior(nodes: &[Node<'_>], index: usize) -> Result<Prior> {
    let node = &nodes[index];
    let Some(block) = node.own_block() else {
        return Ok(Prior::None);
    };
    let Some(previous) = block.position.get().checked_sub(1) else {
        return Ok(Prior::None);
    };
    let Some(paragraph) = block
        .container
        .get(litchi_core::Position::new(previous))
        .and_then(|indexed| indexed.paragraph())
    else {
        return Ok(Prior::None);
    };
    // An earlier target at the preceding paragraph address.
    let steps = node.address.steps();
    if let Some((last, parent)) = steps.split_last()
        && last.kind == Kind::Paragraph
        && let Some(ordinal) = last.ordinal.checked_sub(1)
    {
        let earlier = nodes[..index].iter().position(|other| {
            other
                .address
                .steps()
                .split_last()
                .is_some_and(|(other_last, other_parent)| {
                    other_parent == parent
                        && other_last.kind == Kind::Paragraph
                        && other_last.ordinal == ordinal
                })
        });
        if let Some(earlier) = earlier {
            return Ok(Prior::Target(earlier));
        }
    }
    Ok(paragraph
        .numbering()
        .map_err(Error::from)?
        .filter(|numbering| numbering.num_id != 0)
        .map_or(Prior::None, |numbering| Prior::List(numbering.num_id)))
}
