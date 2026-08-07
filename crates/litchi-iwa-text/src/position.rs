//! Strict UTF-16 positions and ranges shared by native text attributes.

/// Validation failures produced while constructing a text position or range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A platform-sized UTF-16 index does not fit the compact native domain.
    PositionOverflow {
        /// The rejected platform-sized index.
        index: usize,
    },
    /// A range is empty or its start follows its end.
    InvalidRange {
        /// The inclusive start boundary.
        start: TextPosition,
        /// The exclusive end boundary.
        end: TextPosition,
    },
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::PositionOverflow { .. } => {
                formatter.write_str("UTF-16 text position exceeds u32")
            },
            Self::InvalidRange { .. } => {
                formatter.write_str("iWork text range must be nonempty and ordered")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for text position and range construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A UTF-16 code-unit boundary in an iWork text storage.
#[allow(
    clippy::module_name_repetitions,
    reason = "TextPosition names the semantic value represented by this position module."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct TextPosition(u32);

impl TextPosition {
    /// The beginning of a text storage.
    pub const ZERO: Self = Self(0);

    /// Construct a position from a platform-sized UTF-16 code-unit index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PositionOverflow`] when `index` cannot be represented
    /// by the compact native `u32` domain.
    pub fn from_utf16_index(index: usize) -> Result<Self> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_error| Error::PositionOverflow { index })
    }

    /// Construct a position from an already compact UTF-16 code-unit index.
    ///
    /// Every `u32` value is representable by the semantic value. The IWA
    /// adapter separately checks that the position is a scalar boundary in
    /// the text storage before using it for a native edit.
    #[must_use]
    pub const fn from_utf16_code_units(index: u32) -> Self {
        Self(index)
    }

    /// Return the native UTF-16 code-unit index.
    #[must_use]
    pub const fn utf16_index(self) -> u32 {
        self.0
    }
}

/// A nonempty half-open range of UTF-16 code units in an iWork text storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextRange {
    start: TextPosition,
    end: TextPosition,
}

impl TextRange {
    /// Construct a nonempty range whose start is strictly before its end.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidRange`] when `start` is not less than `end`.
    pub fn new(start: TextPosition, end: TextPosition) -> Result<Self> {
        if start >= end {
            return Err(Error::InvalidRange { start, end });
        }
        Ok(Self { start, end })
    }

    /// Construct a range from platform-sized UTF-16 indexes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PositionOverflow`] when either index exceeds the
    /// compact native domain, or [`Error::InvalidRange`] when the range is
    /// empty or reversed.
    pub fn from_utf16_indexes(start: usize, end: usize) -> Result<Self> {
        Self::new(
            TextPosition::from_utf16_index(start)?,
            TextPosition::from_utf16_index(end)?,
        )
    }

    /// Return the inclusive start boundary.
    #[must_use]
    pub const fn start(self) -> TextPosition {
        self.start
    }

    /// Return the exclusive end boundary.
    #[must_use]
    pub const fn end(self) -> TextPosition {
        self.end
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn positions_preserve_compact_utf16_indexes() {
        let position = TextPosition::from_utf16_index(7);
        assert_eq!(position, Ok(TextPosition::from_utf16_code_units(7)));
        assert_eq!(position.map(TextPosition::utf16_index), Ok(7));
        assert_eq!(TextPosition::ZERO.utf16_index(), 0);
    }

    #[cfg(target_pointer_width = "64")]
    #[test]
    fn positions_reject_indexes_beyond_the_native_domain() {
        assert_eq!(
            TextPosition::from_utf16_index(usize::MAX),
            Err(Error::PositionOverflow { index: usize::MAX })
        );
    }

    #[test]
    fn ranges_are_nonempty_and_ordered() {
        let range = TextRange::from_utf16_indexes(2, 7);
        assert_eq!(
            range.map(TextRange::start),
            Ok(TextPosition::from_utf16_code_units(2))
        );
        assert_eq!(
            range.map(TextRange::end),
            Ok(TextPosition::from_utf16_code_units(7))
        );
        assert!(matches!(
            TextRange::from_utf16_indexes(2, 2),
            Err(Error::InvalidRange { .. })
        ));
        assert!(matches!(
            TextRange::from_utf16_indexes(7, 2),
            Err(Error::InvalidRange { .. })
        ));
    }
}
