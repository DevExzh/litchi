//! Strict UTF-16 positions and ranges shared by native text attributes.

use crate::{Error, Result};

/// A UTF-16 code-unit boundary in an iWork text storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct TextPosition(u32);

impl TextPosition {
    /// The beginning of a text storage.
    pub const ZERO: Self = Self(0);

    /// Construct a position from a UTF-16 code-unit index.
    pub fn from_utf16_index(index: usize) -> Result<Self> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_| Error::ParseError("UTF-16 text position exceeds u32".to_owned()))
    }

    /// Return the native UTF-16 code-unit index.
    pub const fn utf16_index(self) -> u32 {
        self.0
    }

    pub(crate) const fn from_native(index: u32) -> Self {
        Self(index)
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
    pub fn new(start: TextPosition, end: TextPosition) -> Result<Self> {
        if start >= end {
            return Err(Error::ParseError(
                "iWork text range must be nonempty and ordered".to_owned(),
            ));
        }
        Ok(Self { start, end })
    }

    /// Construct a range from native UTF-16 indexes.
    pub fn from_utf16_indexes(start: usize, end: usize) -> Result<Self> {
        Self::new(
            TextPosition::from_utf16_index(start)?,
            TextPosition::from_utf16_index(end)?,
        )
    }

    /// Return the inclusive start boundary.
    pub const fn start(self) -> TextPosition {
        self.start
    }

    /// Return the exclusive end boundary.
    pub const fn end(self) -> TextPosition {
        self.end
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_are_nonempty_and_ordered() {
        let range = TextRange::from_utf16_indexes(2, 7).unwrap();
        assert_eq!(range.start().utf16_index(), 2);
        assert_eq!(range.end().utf16_index(), 7);
        assert!(TextRange::from_utf16_indexes(2, 2).is_err());
        assert!(TextRange::from_utf16_indexes(7, 2).is_err());
    }
}
