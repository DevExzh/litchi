//! IWA-native adapters for archive-free paragraph writing direction.

use crate::{Error, Result};

pub use litchi_iwa_text::paragraph::direction::WritingDirection as ParagraphWritingDirection;

/// Decode native iWork's paragraph writing-direction discriminant.
pub(crate) fn from_native(value: i32) -> Result<ParagraphWritingDirection> {
    match value {
        -1 => Ok(ParagraphWritingDirection::Natural),
        0 => Ok(ParagraphWritingDirection::LeftToRight),
        1 => Ok(ParagraphWritingDirection::RightToLeft),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported native iWork paragraph writing direction {value}"
        ))),
    }
}

/// Encode a semantic writing direction as native iWork's discriminant.
pub(crate) const fn to_native(value: ParagraphWritingDirection) -> i32 {
    match value {
        ParagraphWritingDirection::Natural => -1,
        ParagraphWritingDirection::LeftToRight => 0,
        ParagraphWritingDirection::RightToLeft => 1,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_values_round_trip_at_the_iwa_boundary() {
        for direction in [
            ParagraphWritingDirection::Natural,
            ParagraphWritingDirection::LeftToRight,
            ParagraphWritingDirection::RightToLeft,
        ] {
            assert_eq!(from_native(to_native(direction)).unwrap(), direction);
        }
    }

    #[test]
    fn rejects_unknown_native_value_at_the_iwa_boundary() {
        assert!(from_native(2).is_err());
    }
}
