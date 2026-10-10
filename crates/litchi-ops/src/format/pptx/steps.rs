//! Positional node addresses for the PPTX adapter.
//!
//! Ordinary unified paths use checked positions for shape-tree elements. The
//! producer `cNvPr@id` is a native identifier and therefore never becomes a
//! canonical ordinary path step; callers that need it must opt into the raw
//! path surface. Keeping this module positional also means a node's path is
//! stable under duplicate or malformed native IDs.

use litchi_core::Position;
use litchi_pptx::shape::Shape;
use litchi_pptx::shape::owned::Handle;

use super::spec::Kind;
use crate::node::PathStep;
use crate::path::navigate::ElementKind;

/// The number of path steps retained inline before spilling to a boxed slice.
const INLINE: usize = 4;

const EMPTY: PathStep<'static> = PathStep::new("");

/// The steps of one node address, from the presentation root.
#[derive(Debug, Clone)]
pub(crate) struct Steps {
    inline: [PathStep<'static>; INLINE],
    len: u8,
    spill: Option<Box<[PathStep<'static>]>>,
}

impl Steps {
    /// The steps, from the root.
    pub(crate) fn as_slice(&self) -> &[PathStep<'static>] {
        match &self.spill {
            Some(spill) => spill,
            None => &self.inline[..usize::from(self.len)],
        }
    }

    /// The steps as an owned slice.
    pub(crate) fn into_boxed(self) -> Box<[PathStep<'static>]> {
        match self.spill {
            Some(spill) => spill,
            None => self.inline[..usize::from(self.len)].into(),
        }
    }
}

/// Builds a [`Steps`] without allocating until it outgrows the inline buffer.
pub(crate) struct Builder {
    inline: [PathStep<'static>; INLINE],
    len: usize,
    spill: Vec<PathStep<'static>>,
}

impl Builder {
    pub(crate) const fn new() -> Self {
        Self {
            inline: [EMPTY; INLINE],
            len: 0,
            spill: Vec::new(),
        }
    }

    pub(crate) fn push(&mut self, step: PathStep<'static>) {
        if self.spill.is_empty() && self.len < INLINE {
            self.inline[self.len] = step;
            self.len += 1;
            return;
        }
        if self.spill.is_empty() {
            self.spill.extend_from_slice(&self.inline[..self.len]);
        }
        self.spill.push(step);
    }

    pub(crate) fn finish(self) -> Steps {
        if self.spill.is_empty() {
            Steps {
                inline: self.inline,
                len: u8::try_from(self.len).unwrap_or(u8::MAX),
                spill: None,
            }
        } else {
            Steps {
                inline: [EMPTY; INLINE],
                len: 0,
                spill: Some(self.spill.into_boxed_slice()),
            }
        }
    }
}

/// The slide step for a zero-based catalog position.
pub(crate) fn slide(position: u32) -> PathStep<'static> {
    PathStep::at("slide", Position::new(position as usize))
}

/// The shape-tree kind represented by a detached handle, or `None` when the
/// leaf has no ordinary adapter node for that shape variant.
pub(crate) fn tree_kind(handle: &Handle) -> Option<Kind> {
    Some(match handle.view() {
        Shape::Auto(_) => Kind::Shape,
        Shape::Picture(_) => Kind::Picture,
        Shape::Table(_) => Kind::Table,
        Shape::Chart(_) => Kind::Chart,
        Shape::Diagram(_) => Kind::SmartArt,
        Shape::Ole(_) => Kind::Ole,
        Shape::Group(_) => Kind::Group,
        Shape::Connector(_) => Kind::Connector,
        Shape::Frame(_) | Shape::Content(_) | Shape::Unknown(_) => return None,
        _ => return None,
    })
}

/// The direct siblings of `handle`, in source order.
fn siblings(handle: &Handle) -> litchi_pptx::shape::owned::Handles {
    match handle.parent() {
        Some(group) => group.children(),
        None => handle.scene().roots(),
    }
}

/// The zero-based position of `handle` among its container's elements of
/// `kind`.
pub(crate) fn ordinal(handle: &Handle, kind: Kind) -> u32 {
    let target = handle.position();
    let mut count = 0_u32;
    for sibling in siblings(handle) {
        if sibling.position() == target {
            break;
        }
        if tree_kind(&sibling) == Some(kind) {
            count = count.saturating_add(1);
        }
    }
    count
}

/// Build one shape-tree step. The `canonical` argument is retained at this
/// internal seam so callers can share the path-building code, but native IDs
/// are deliberately never emitted by ordinary node paths.
pub(crate) fn tree_step(
    handle: &Handle,
    kind: Kind,
    ordinal_hint: Option<u32>,
    _canonical: bool,
) -> PathStep<'static> {
    let ordinal = ordinal_hint.unwrap_or_else(|| ordinal(handle, kind));
    PathStep::at(kind.name(), Position::new(ordinal as usize))
}

/// Push a shape-tree node's slide, enclosing groups and own positional step.
pub(crate) fn push_tree(
    out: &mut Builder,
    slide_position: u32,
    handle: &Handle,
    kind: Kind,
    ordinal_hint: u32,
    canonical: bool,
) {
    out.push(slide(slide_position));

    // The scene's bounded depth makes this small in normal documents. The
    // vector keeps the enclosing groups in outermost-to-innermost order.
    let mut groups = Vec::new();
    let mut cursor = handle.parent();
    while let Some(group) = cursor {
        cursor = group.parent();
        groups.push(group);
    }
    for group in groups.iter().rev() {
        out.push(tree_step(group, Kind::Group, None, canonical));
    }
    out.push(tree_step(handle, kind, Some(ordinal_hint), canonical));
}
