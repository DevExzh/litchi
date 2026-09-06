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

    /// Create an exact-name selector after checking that the name is usable.
    ///
    /// This checked constructor is useful at input boundaries where an empty
    /// section name should be reported as invalid rather than treated as a
    /// valid exact-name lookup. [`Self::name`] remains available for callers
    /// that intentionally preserve an empty authored name.
    ///
    /// # Errors
    ///
    /// Returns [`SelectorError::EmptySectionName`] when `name` is empty.
    pub const fn try_name(name: &'a str) -> Result<Self, SelectorError> {
        if name.is_empty() {
            Err(SelectorError::EmptySectionName)
        } else {
            Ok(Self::Name(name))
        }
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

impl<'a> From<&'a String> for SectionSelector<'a> {
    fn from(name: &'a String) -> Self {
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
    /// A checked selector was requested for an empty section name.
    #[error("Pages section name must not be empty")]
    EmptySectionName,
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

/// Selects one table attached to the main Pages body by its exact visible
/// name or checked zero-based body order.
///
/// Native attachment, drawable, and table-model identities are deliberately
/// not representable by this selector. The package adapter proves those
/// private ownership edges only after resolving this semantic value.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public selector name identifies the Pages body-table domain."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BodyTableSelector<'a> {
    /// Select by the exact table name stored in the Pages table model.
    Name(&'a str),
    /// Select by typed zero-based source order among body tables.
    Position(Position),
}

impl<'a> BodyTableSelector<'a> {
    /// Create a name-first body-table selector without allocating.
    #[must_use]
    pub const fn name(name: &'a str) -> Self {
        Self::Name(name)
    }

    /// Create a zero-based body-table selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::position(Position::new(index))
    }

    /// Create a selector from a typed zero-based collection position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Position(position)
    }

    /// Borrow the selected exact table name, if present.
    #[must_use]
    pub const fn as_name(self) -> Option<&'a str> {
        match self {
            Self::Name(name) => Some(name),
            Self::Position(_) => None,
        }
    }

    /// Return the selected zero-based body-table index, if present.
    #[must_use]
    pub const fn as_index(self) -> Option<usize> {
        match self {
            Self::Name(_) => None,
            Self::Position(position) => Some(position.get()),
        }
    }

    /// Return the selected typed zero-based body-table position, if present.
    #[must_use]
    pub const fn as_position(self) -> Option<Position> {
        match self {
            Self::Name(_) => None,
            Self::Position(position) => Some(position),
        }
    }
}

impl<'a> From<&'a str> for BodyTableSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::name(name)
    }
}

impl<'a> From<&'a String> for BodyTableSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::name(name)
    }
}

impl From<usize> for BodyTableSelector<'_> {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for BodyTableSelector<'_> {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

/// Selects one body anchored image by its zero based source order.
///
/// The order is the order of ordinary `TSD.ImageArchive` attachments in the
/// rooted body storage. Native object identifiers, archive members, and
/// generated protobuf values are deliberately not representable here.
#[allow(
    clippy::module_name_repetitions,
    reason = "The selector name identifies the Pages body-image domain."
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageSelector {
    /// Select the image at this zero-based body source position.
    Index(Position),
}

/// Selects one chart anchored in the rooted Pages body by its zero-based
/// source order.
///
/// Native attachment, drawable, and archive identities are deliberately not
/// representable here. The package owner proves those private graph edges
/// after resolving this semantic selector.
#[allow(
    clippy::module_name_repetitions,
    reason = "The selector name identifies the Pages body-chart domain."
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BodyChartSelector {
    /// Select the chart at this zero-based body source position.
    Index(Position),
}

impl BodyChartSelector {
    /// Create a selector from a zero-based body-chart index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(Position::new(index))
    }

    /// Create a selector from a typed zero-based body-chart position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Index(position)
    }

    /// Return the selected typed body-chart position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Index(position) => position,
        }
    }

    /// Return the selected zero-based body-chart index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.as_position().get()
    }
}

impl From<usize> for BodyChartSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for BodyChartSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

impl ImageSelector {
    /// Create a selector from a zero-based body-image index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(Position::new(index))
    }

    /// Create a selector from a typed zero-based body-image position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Index(position)
    }

    /// Return the selected typed body-image position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Index(position) => position,
        }
    }

    /// Return the selected zero-based body-image index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.as_position().get()
    }
}

impl From<usize> for ImageSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for ImageSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

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

    #[test]
    fn borrowed_owned_names_convert_without_allocating() {
        let section_name = String::from("Chapter One");
        let selector: SectionSelector<'_> = (&section_name).into();

        assert_eq!(selector, SectionSelector::Name("Chapter One"));
    }

    #[test]
    fn checked_name_rejects_empty_input_without_changing_legacy_constructor() {
        assert_eq!(
            SectionSelector::try_name(""),
            Err(SelectorError::EmptySectionName)
        );
        assert_eq!(
            SectionSelector::try_name("Chapter One"),
            Ok(SectionSelector::name("Chapter One"))
        );
        assert_eq!(SectionSelector::name(""), SectionSelector::Name(""));
    }

    #[test]
    fn body_table_selectors_are_borrowed_and_typed() {
        let name: BodyTableSelector<'_> = "Revenue".into();
        let index: BodyTableSelector<'_> = 2usize.into();
        let position = Position::new(2);

        assert_eq!(name, BodyTableSelector::name("Revenue"));
        assert_eq!(index, BodyTableSelector::position(position));
        assert_eq!(name.as_name(), Some("Revenue"));
        assert_eq!(name.as_index(), None);
        assert_eq!(index.as_index(), Some(2));
        assert_eq!(index.as_position(), Some(position));
    }

    #[test]
    fn body_image_selector_is_typed_and_source_ordered() {
        let from_index = ImageSelector::index(3);
        let from_position = ImageSelector::from(Position::new(3));

        assert_eq!(from_index, from_position);
        assert_eq!(from_index.as_index(), 3);
        assert_eq!(from_index.as_position(), Position::new(3));
    }

    #[test]
    fn body_chart_selector_is_typed_and_source_ordered() {
        let from_index = BodyChartSelector::index(4);
        let from_position = BodyChartSelector::from(Position::new(4));

        assert_eq!(from_index, from_position);
        assert_eq!(from_index.as_index(), 4);
        assert_eq!(from_index.as_position(), Position::new(4));
    }
}
