//! Archive-free selectors for immutable Pages sections.

use litchi_core::Position;
use thiserror::Error;

/// Selects one section by its exact semantic name or zero-based source
/// position without retaining a native object identifier.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public name identifies the selected Pages object."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionSelector<'a> {
    /// Select by the exact, case-sensitive section name shown by Pages.
    Name(&'a str),
    /// Select by zero-based position in the immutable source snapshot.
    Position(Position),
}

impl<'a> SectionSelector<'a> {
    /// Create an exact-name selector without allocating.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Create a checked zero-based source-index selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Create a selector from a typed zero-based source position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
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

impl<'a> From<&'a str> for SectionSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl From<usize> for SectionSelector<'_> {
    fn from(position: usize) -> Self {
        Self::index(position)
    }
}

impl From<Position> for SectionSelector<'_> {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

/// Errors raised while resolving a section selector.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public name distinguishes selector failures from document construction errors."
)]
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SelectorError {
    /// More than one section has the requested exact name.
    #[error("Pages sections at source positions {first} and {duplicate} share the name {name:?}")]
    AmbiguousSectionName {
        /// The exact section name that resolved ambiguously.
        name: Box<str>,
        /// Source position of the first matching section.
        first: usize,
        /// Source position of the next matching section.
        duplicate: usize,
    },
}

/// Result type for checked semantic section lookup.
#[allow(
    clippy::module_name_repetitions,
    reason = "The alias is re-exported at crate scope beside SelectorError."
)]
pub type SelectorResult<T> = Result<T, SelectorError>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selectors_are_borrowed_copyable_values() {
        let name: SectionSelector<'_> = "Chapter One".into();
        let position: SectionSelector<'_> = 2.into();
        let typed: SectionSelector<'_> = Position::new(2).into();

        assert_eq!(name, SectionSelector::Name("Chapter One"));
        assert_eq!(position, SectionSelector::Position(Position::new(2)));
        assert_eq!(position, typed);
        assert_eq!(name.as_name(), Some("Chapter One"));
        assert_eq!(position.as_position(), Some(Position::new(2)));
        assert_eq!(name, name);
    }
}
