//! Typed paragraph base-writing direction.

use crate::{Error, Result};

/// Base direction used to lay out bidirectional paragraph text.
///
/// `Natural` lets iWork infer the direction from the paragraph contents. The
/// explicit variants are useful for neutral or mixed-direction content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ParagraphWritingDirection {
    /// Infer direction from the paragraph contents.
    #[default]
    Natural,
    /// Lay out the paragraph from left to right.
    LeftToRight,
    /// Lay out the paragraph from right to left.
    RightToLeft,
}

impl ParagraphWritingDirection {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::Natural => -1,
            Self::LeftToRight => 0,
            Self::RightToLeft => 1,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> Result<Self> {
        match value {
            -1 => Ok(Self::Natural),
            0 => Ok(Self::LeftToRight),
            1 => Ok(Self::RightToLeft),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported iWork paragraph writing direction {value}"
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_values_round_trip() {
        for direction in [
            ParagraphWritingDirection::Natural,
            ParagraphWritingDirection::LeftToRight,
            ParagraphWritingDirection::RightToLeft,
        ] {
            assert_eq!(
                ParagraphWritingDirection::from_native_value(direction.native_value()).unwrap(),
                direction
            );
        }
    }

    #[test]
    fn rejects_unknown_native_value() {
        assert!(ParagraphWritingDirection::from_native_value(2).is_err());
    }
}
