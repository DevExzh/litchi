//! Bookmark semantics.

use litchi_core::Position;

/// A projected text location.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Anchor {
    block: Position,
    offset: usize,
}

impl Anchor {
    pub(crate) const fn new(block: Position, offset: usize) -> Self {
        Self { block, offset }
    }

    /// Zero-based position in [`crate::TextBody::blocks`].
    #[must_use]
    pub const fn block(self) -> Position {
        self.block
    }

    /// UTF-8 byte offset in the projected block text.
    #[must_use]
    pub const fn offset(self) -> usize {
        self.offset
    }
}

/// A point bookmark or a paired bookmark range.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Bookmark {
    /// A `text:bookmark` point.
    Point { name: String, at: Anchor },
    /// A matched `text:bookmark-start` / `text:bookmark-end` range.
    Range {
        name: String,
        start: Anchor,
        end: Anchor,
    },
}

impl Bookmark {
    /// Producer-visible bookmark name.
    #[must_use]
    pub fn name(&self) -> &str {
        match self {
            Self::Point { name, .. } | Self::Range { name, .. } => name,
        }
    }

    /// Start or point location.
    #[must_use]
    pub const fn start(&self) -> Anchor {
        match self {
            Self::Point { at, .. } => *at,
            Self::Range { start, .. } => *start,
        }
    }

    /// Range end, or `None` for a point bookmark.
    #[must_use]
    pub const fn end(&self) -> Option<Anchor> {
        match self {
            Self::Point { .. } => None,
            Self::Range { end, .. } => Some(*end),
        }
    }
}
