//! Archive-free selectors for Keynote semantic snapshots.

use litchi_core::Position;

/// Selects one drawable owned by a selected slide in source order.
///
/// The position is resolved against the package's direct-drawable inventory;
/// it is intentionally not a native object identifier.  A selector can be
/// constructed without a package, but package operations validate ownership
/// before they read or mutate anything.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DrawableSelector(Position);

impl DrawableSelector {
    /// Construct a selector from a checked source position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self(position)
    }

    /// Construct a selector from a zero-based source index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Return the selected source position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        self.0
    }

    /// Return the selected source index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.0.get()
    }
}

impl From<Position> for DrawableSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<usize> for DrawableSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

/// Selects one direct reply by its checked ordinal in the comment thread.
///
/// Reply storage identifiers are deliberately excluded.  The ordinal is
/// checked against the complete source thread before a mutation is staged.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ReplySelector(Position);

impl ReplySelector {
    /// Construct a selector from a checked reply ordinal.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self(position)
    }

    /// Construct a selector from a zero-based reply ordinal.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Return the selected reply ordinal.
    #[must_use]
    pub const fn as_position(self) -> Position {
        self.0
    }

    /// Return the selected reply ordinal.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.0.get()
    }
}

impl From<Position> for ReplySelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<usize> for ReplySelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

/// Selects one slide by its exact navigator name or checked source position.
///
/// A navigator name is a developer-facing Keynote property distinct from the
/// slide's visible title content. Selectors never contain native object IDs,
/// component names, or template identifiers.
#[allow(
    clippy::module_name_repetitions,
    reason = "SlideSelector keeps the selected Keynote domain explicit at the public boundary"
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SlideSelector<'a> {
    /// Select the slide with this exact navigator name.
    Name(&'a str),
    /// Select the slide at this zero-based position in source order.
    Position(Position),
}

impl<'a> SlideSelector<'a> {
    /// Create an exact-name selector without allocating.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Create an exact-name selector after checking that the name is usable.
    ///
    /// This checked constructor is useful at input boundaries where an empty
    /// navigator name should be reported as invalid rather than treated as a
    /// valid exact-name lookup. [`Self::name`] remains available for callers
    /// that intentionally preserve the historical empty-name behavior.
    ///
    /// # Errors
    ///
    /// Returns [`SlideSelectorError::EmptySlideName`] when `name` is empty.
    pub const fn try_name(name: &'a str) -> Result<Self, SlideSelectorError> {
        if name.is_empty() {
            Err(SlideSelectorError::EmptySlideName)
        } else {
            Ok(Self::Name(name))
        }
    }

    /// Create a selector from a typed zero-based source position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Create a selector from a zero-based source index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Borrow the selected exact name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Position(_) => None,
        }
    }

    /// Return the selected typed source position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Name(_) => None,
            Self::Position(position) => Some(position),
        }
    }
}

impl<'a> From<&'a str> for SlideSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl<'a> From<&'a String> for SlideSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::name(name)
    }
}

impl From<Position> for SlideSelector<'_> {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl From<usize> for SlideSelector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

/// Errors raised while resolving a semantic slide selector.
#[allow(
    clippy::module_name_repetitions,
    reason = "SlideSelectorError keeps the failing selector domain clear at the crate boundary"
)]
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SlideSelectorError {
    /// A checked selector was requested for an empty navigator name.
    EmptySlideName,
    /// More than one slide carries the requested exact navigator name.
    DuplicateSlideName {
        /// The ambiguous developer-facing name.
        name: Box<str>,
    },
}

impl std::fmt::Display for SlideSelectorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::EmptySlideName => formatter.write_str("Keynote slide name must not be empty"),
            Self::DuplicateSlideName { name } => {
                write!(formatter, "show contains duplicate slide name {name:?}")
            },
        }
    }
}

impl std::error::Error for SlideSelectorError {}

/// Result type for checked semantic slide selection.
pub type SlideSelectorResult<T> = Result<T, SlideSelectorError>;

#[cfg(test)]
mod tests {
    use super::{SlideSelector, SlideSelectorError};
    use litchi_core::Position;

    #[test]
    fn selector_preserves_exact_name_and_typed_source_position() {
        const POSITION: Position = Position::new(3);
        const POSITION_SELECTOR: SlideSelector<'static> = SlideSelector::position(POSITION);
        const NAME_SELECTOR: SlideSelector<'static> = SlideSelector::name("Agenda");

        assert_eq!(POSITION_SELECTOR.as_position(), Some(POSITION));
        assert_eq!(POSITION_SELECTOR.as_name(), None);
        assert_eq!(NAME_SELECTOR.as_name(), Some("Agenda"));
        assert_eq!(NAME_SELECTOR.as_position(), None);
    }

    #[test]
    fn ergonomic_values_convert_without_native_identity() {
        let from_index: SlideSelector<'_> = 5usize.into();
        let from_position: SlideSelector<'_> = Position::new(5).into();
        let from_name: SlideSelector<'_> = "Agenda".into();

        assert_eq!(from_index, from_position);
        assert_eq!(from_name, SlideSelector::Name("Agenda"));
    }

    #[test]
    fn borrowed_owned_names_convert_without_allocating() {
        let slide_name = String::from("Agenda");
        let selector: SlideSelector<'_> = (&slide_name).into();

        assert_eq!(selector, SlideSelector::Name("Agenda"));
    }

    #[test]
    fn checked_name_rejects_empty_input_without_changing_legacy_constructor() {
        assert_eq!(
            SlideSelector::try_name(""),
            Err(SlideSelectorError::EmptySlideName)
        );
        assert_eq!(
            SlideSelector::try_name("Agenda"),
            Ok(SlideSelector::name("Agenda"))
        );
        assert_eq!(SlideSelector::name(""), SlideSelector::Name(""));
    }
}
