//! Archive-free Pages body-footnote values and selectors.

use thiserror::Error;

/// Maximum UTF-8 bytes retained by one semantic body-footnote text value.
pub const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
/// Maximum UTF-8 bytes retained by one custom marker.
pub const MAX_CUSTOM_MARK_BYTES: usize = 1024;

/// Validation failures for semantic body-footnote values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A UTF-16 position cannot be represented by the semantic position type.
    #[error("Pages footnote UTF-16 position exceeds the semantic range")]
    PositionOutOfRange,
    /// Footnote text exceeds its bounded semantic storage budget.
    #[error("Pages footnote text exceeds the semantic byte budget")]
    TextTooLarge,
    /// A custom marker exceeds its bounded semantic storage budget.
    #[error("Pages footnote custom marker exceeds the semantic byte budget")]
    CustomMarkTooLarge,
}

/// Result type for semantic body-footnote construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A checked UTF-16 code-unit boundary in the Pages body story.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Position(u32);

impl Position {
    /// The beginning of the body story.
    pub const ZERO: Self = Self(0);

    /// Construct a position from a UTF-16 code-unit index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PositionOutOfRange`] when the index cannot be stored
    /// in the compact semantic representation.
    pub fn from_utf16_index(index: usize) -> Result<Self> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_conversion_error| Error::PositionOutOfRange)
    }

    /// Return the UTF-16 code-unit index represented by this position.
    #[must_use]
    pub const fn utf16_index(self) -> u32 {
        self.0
    }
}

/// Select one body footnote without exposing a package or runtime identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Selector {
    /// Select the footnote whose body anchor is at this UTF-16 position.
    At(Position),
    /// Select the footnote at this zero-based source order.
    Index(usize),
}

impl Selector {
    /// Create a position-based selector.
    #[must_use]
    pub const fn at(position: Position) -> Self {
        Self::At(position)
    }

    /// Create a zero-based source-order selector.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }
}

impl From<Position> for Selector {
    fn from(position: Position) -> Self {
        Self::At(position)
    }
}

/// One semantic footnote attached to the Pages body story.
///
/// The body anchor position and user-visible text are retained; native
/// reference, storage, and marker object identifiers remain private to the
/// package adapter.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Footnote {
    /// UTF-16 position of the body anchor.
    pub position: Position,
    /// Footnote content without Pages' internal marker and separator.
    pub text: Box<str>,
    /// Optional custom marker written by Pages instead of automatic numbering.
    pub custom_mark: Option<Box<str>>,
}

impl Footnote {
    /// Construct a footnote without a custom marker.
    ///
    /// # Errors
    ///
    /// Returns an error when the text exceeds [`MAX_TEXT_BYTES`].
    pub fn new(position: Position, text: impl Into<Box<str>>) -> Result<Self> {
        Self::with_custom_mark(position, text, None)
    }

    /// Construct a footnote with an optional custom marker.
    ///
    /// # Errors
    ///
    /// Returns an error when the text or custom marker exceeds its semantic
    /// byte budget.
    pub fn with_custom_mark(
        position: Position,
        text_value: impl Into<Box<str>>,
        custom_mark: Option<Box<str>>,
    ) -> Result<Self> {
        let text = text_value.into();
        if text.len() > MAX_TEXT_BYTES {
            return Err(Error::TextTooLarge);
        }
        if custom_mark
            .as_deref()
            .is_some_and(|mark| mark.len() > MAX_CUSTOM_MARK_BYTES)
        {
            return Err(Error::CustomMarkTooLarge);
        }
        Ok(Self {
            position,
            text,
            custom_mark,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Footnote, MAX_CUSTOM_MARK_BYTES, MAX_TEXT_BYTES, Position, Selector};

    #[test]
    fn positions_and_selectors_are_compact_and_typed() {
        let position = Position::from_utf16_index(17).unwrap();
        assert_eq!(position.utf16_index(), 17);
        assert_eq!(Selector::from(position), Selector::At(position));
        assert_eq!(Selector::at(position), Selector::At(position));
        assert_eq!(Selector::index(2), Selector::Index(2));
        if usize::BITS > u32::BITS {
            assert_eq!(
                Position::from_utf16_index(usize::MAX),
                Err(Error::PositionOutOfRange)
            );
        }
    }

    #[test]
    fn footnote_values_are_bounded_without_native_ids() {
        let footnote = Footnote::with_custom_mark(
            Position::ZERO,
            "body",
            Some("*".to_owned().into_boxed_str()),
        )
        .unwrap();
        assert_eq!(footnote.position, Position::ZERO);
        assert_eq!(footnote.text.as_ref(), "body");
        assert_eq!(footnote.custom_mark.as_deref(), Some("*"));
        assert_eq!(
            Footnote::new(Position::ZERO, "x".repeat(MAX_TEXT_BYTES + 1)),
            Err(Error::TextTooLarge)
        );
        assert_eq!(
            Footnote::with_custom_mark(
                Position::ZERO,
                "x",
                Some("x".repeat(MAX_CUSTOM_MARK_BYTES + 1).into_boxed_str()),
            ),
            Err(Error::CustomMarkTooLarge)
        );
    }
}
