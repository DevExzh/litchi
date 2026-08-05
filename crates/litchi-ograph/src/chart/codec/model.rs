//! Parser state for the chart record codec.

use super::super::{axis, format, group};

#[derive(Clone, Copy)]
pub(super) enum PendingLine {
    Axis {
        owner: usize,
        kind: axis::LineKind,
    },
    Group {
        owner: usize,
        kind: crate::record::line::Kind,
    },
}

pub(super) struct PendingDrop {
    pub(super) owner: usize,
    pub(super) depth: usize,
    pub(super) gap: group::Gap,
    pub(super) line: Option<format::Line>,
    pub(super) area: Option<format::Area>,
}
