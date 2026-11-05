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
//!
//! `add <container> --type paragraph listStyle=... [start=N]` (record 0342)
//! numbers the new paragraph the same way, after OfficeCLI's add path: it
//! continues the list of the nearest preceding paragraph of the container at
//! the same or a shallower level (deeper items are skipped) when that list
//! has the same kind, and otherwise adds a new list.

use std::collections::HashMap;

use litchi_core::Position;
use litchi_docx::document::{ElementPath, ElementStep};
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

/// The list setter an entry's access names, if any.
pub(super) const fn list_setter(
    access: crate::registry::Access<super::setters::Setter>,
) -> Option<super::setters::Setter> {
    use super::setters::Setter;
    match access {
        crate::registry::Access::Set(setter) | crate::registry::Access::SetAndAdd(setter)
            if matches!(
                setter,
                Setter::ParagraphListStyle | Setter::ParagraphListStart
            ) =>
        {
            Some(setter)
        },
        _ => None,
    }
}

/// A `start` value: a non-negative integer.
pub(super) fn start(value: Option<crate::value::Scalar<'_>>) -> Result<u32> {
    let invalid = || Error::new(Code::InvalidValue, "a list start is a non-negative integer");
    match value {
        Some(crate::value::Scalar::Int(start)) => u32::try_from(start).map_err(|_| invalid()),
        _ => Err(invalid()),
    }
}

/// Number the paragraph just inserted at `position` of `container` as
/// `request` asks (record 0342). Its own level (`numLevel`) is kept, or it
/// takes level 0.
pub(super) fn number_added(
    edit: &mut litchi_docx::document::Edit,
    container: &ElementPath,
    position: Position,
    request: Request,
) -> Result<()> {
    let style = match (request.style, request.start) {
        (Some(style), _) => style,
        (None, Some(_)) => Style::List(ListKind::Decimal),
        (None, None) => return Ok(()),
    };
    // A new paragraph is outside every list already.
    let Style::List(kind) = style else {
        return Ok(());
    };
    let path = container.child(ElementStep::Paragraph(position));
    let level = numbering(edit, &path)?
        .and_then(|numbering| numbering.level)
        .unwrap_or(0);
    let continued = if request.start.is_some() {
        None
    } else {
        continuation(edit, container, position, kind, level)?
    };
    let num_id = match continued {
        Some(num_id) => num_id,
        None => edit
            .add_list(kind, request.start.map(|start| (0, start)))
            .map_err(Error::from)?
            .num_id(),
    };
    edit.format_paragraph_at(&path, |properties| {
        let numbering = properties.numbering.get_or_insert_default();
        numbering.id = Some(num_id);
        numbering.level.get_or_insert(level);
    })
    .map_err(Error::from)?;
    Ok(())
}

/// The list the paragraph at `position` continues: the nearest preceding
/// paragraph of the container at level `level` or shallower, when it is in
/// a list of `kind`. A paragraph outside every list ends the search.
fn continuation(
    edit: &mut litchi_docx::document::Edit,
    container: &ElementPath,
    position: Position,
    kind: ListKind,
    level: u8,
) -> Result<Option<u32>> {
    let mut index = position.get();
    while let Some(previous) = index.checked_sub(1) {
        index = previous;
        let path = container.child(ElementStep::Paragraph(Position::new(previous)));
        let Some(numbering) = numbering(edit, &path)? else {
            return Ok(None);
        };
        let Some(num_id) = numbering.id.filter(|id| *id != 0) else {
            return Ok(None);
        };
        if numbering.level.unwrap_or(0) > level {
            continue;
        }
        let known = edit.list_kind(num_id).map_err(Error::from)?;
        return Ok((known == Some(kind)).then_some(num_id));
    }
    Ok(None)
}

/// The direct numbering of the staged paragraph at `path`. An unchanged
/// format is an exact no-op of the leaf edit, so this only reads.
fn numbering(
    edit: &mut litchi_docx::document::Edit,
    path: &ElementPath,
) -> Result<Option<litchi_docx::properties::Numbering>> {
    let mut found = None;
    edit.format_paragraph_at(path, |properties| found.clone_from(&properties.numbering))
        .map_err(Error::from)?;
    Ok(found)
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
        .get(Position::new(previous))
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
