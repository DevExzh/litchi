//! Per-slide visibility for layout-provided title and body placeholders.

pub use crate::package::slide_placeholder_visibility::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch,
};

/// Layout-provided text placeholder selected on one slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The slide title placeholder.
    Title,
    /// The slide body placeholder.
    Body,
}

impl std::fmt::Display for Kind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Title => "title",
            Self::Body => "body",
        })
    }
}

/// Whether an existing placeholder participates in slide drawing and z-order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum State {
    /// The placeholder is present once in both native ownership lists.
    Visible,
    /// The placeholder remains referenced and retains text, but is absent from
    /// both native ownership lists.
    Hidden,
}

impl State {
    /// Return whether the placeholder participates in drawing.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self, Self::Visible)
    }
}
