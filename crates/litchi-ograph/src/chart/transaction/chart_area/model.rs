//! Typed chart-area requests and reversible change metadata.

use crate::chart::Rect;

/// One source-checked chart-area rectangle change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    before: Rect,
    after: Rect,
}

impl Change {
    pub(crate) const fn new(before: Rect, after: Rect) -> Self {
        Self { before, after }
    }

    /// Rectangle required before this change can be applied.
    pub const fn before(self) -> Rect {
        self.before
    }

    /// Rectangle produced by this change.
    pub const fn after(self) -> Rect {
        self.after
    }

    pub(crate) const fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }
}

/// One staged chart-area replacement, kept private to the transaction facade.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Request {
    pub(crate) value: Rect,
}
