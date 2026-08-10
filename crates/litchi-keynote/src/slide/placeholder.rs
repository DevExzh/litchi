//! Per-slide visibility for existing title, body, and slide-number placeholders.
//!
//! A [`Kind::SlideNumber`] edit controls only the selected slide's own number
//! placeholder. It neither reads nor changes the show-wide slide-number
//! preference exposed by [`crate::show::Settings`]. The number can therefore
//! remain unavailable to viewers when that separate show-wide preference is
//! disabled.

pub use crate::package::slide_placeholder_visibility::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch,
};

/// Existing placeholder selected on one slide.
///
/// [`Self::Title`] and [`Self::Body`] select layout-provided text roles.
/// [`Self::SlideNumber`] selects the selected slide's existing number role;
/// it is a visibility-only role and cannot be read or edited through the
/// slide-text APIs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The slide title placeholder.
    Title,
    /// The slide body placeholder.
    Body,
    /// The selected slide's number placeholder.
    ///
    /// This is independent of the show-wide slide-number preference.
    SlideNumber,
}

impl std::fmt::Display for Kind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::SlideNumber => "slide number",
        })
    }
}

/// Whether an existing placeholder participates in the selected slide's display.
///
/// This value models only an existing role. A visibility read returns `None`
/// when that role is absent; `None` is not another hidden state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum State {
    /// The placeholder participates in the selected slide's display.
    ///
    /// For [`Kind::SlideNumber`], this also means the selected slide has its
    /// local number display enabled. It does not imply that the show-wide
    /// slide-number preference is enabled.
    Visible,
    /// The placeholder remains part of the document but does not display on
    /// the selected slide.
    ///
    /// Hidden title and body placeholders retain their text. Hiding a slide
    /// number does not change the show-wide slide-number preference.
    Hidden,
}

impl State {
    /// Return whether this state participates in the selected slide's display.
    ///
    /// This is a constant-time value comparison.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self, Self::Visible)
    }
}
