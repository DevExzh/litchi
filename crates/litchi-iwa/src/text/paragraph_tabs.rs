//! IWA-native adapters for archive-free paragraph ruler values.

use crate::{Error, Result};

pub use litchi_iwa_text::paragraph::tabs::{
    Alignment as ParagraphTabAlignment, DecimalCharacter as ParagraphDecimalTabCharacter,
    DefaultInterval as ParagraphDefaultTabInterval, Leader as ParagraphTabLeader,
    Position as ParagraphTabPosition, Stop as ParagraphTabStop, Stops as ParagraphTabStops,
};

impl From<litchi_iwa_text::paragraph::tabs::Error> for Error {
    fn from(error: litchi_iwa_text::paragraph::tabs::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

/// Decode native iWork's tab-stop alignment discriminant.
pub(crate) fn alignment_from_native(value: i32) -> Result<ParagraphTabAlignment> {
    match value {
        0 => Ok(ParagraphTabAlignment::Left),
        1 => Ok(ParagraphTabAlignment::Center),
        2 => Ok(ParagraphTabAlignment::Right),
        3 => Ok(ParagraphTabAlignment::Decimal),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported native iWork tab-stop alignment {value}"
        ))),
    }
}

/// Encode a semantic tab-stop alignment as native iWork's discriminant.
pub(crate) const fn alignment_to_native(value: ParagraphTabAlignment) -> i32 {
    match value {
        ParagraphTabAlignment::Left => 0,
        ParagraphTabAlignment::Center => 1,
        ParagraphTabAlignment::Right => 2,
        ParagraphTabAlignment::Decimal => 3,
    }
}

/// Decode native iWork's one-scalar decimal-tab string.
pub(crate) fn decimal_character_from_native(value: &str) -> Result<ParagraphDecimalTabCharacter> {
    let mut characters = value.chars();
    let Some(character) = characters.next() else {
        return Err(Error::InvalidFormat(
            "native iWork decimal-tab character is empty".to_owned(),
        ));
    };
    if characters.next().is_some() {
        return Err(Error::InvalidFormat(
            "native iWork decimal-tab character contains multiple Unicode scalars".to_owned(),
        ));
    }
    ParagraphDecimalTabCharacter::new(character).map_err(Into::into)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_tab_alignment_values_round_trip_at_the_iwa_boundary() {
        for alignment in [
            ParagraphTabAlignment::Left,
            ParagraphTabAlignment::Center,
            ParagraphTabAlignment::Right,
            ParagraphTabAlignment::Decimal,
        ] {
            assert_eq!(
                alignment_from_native(alignment_to_native(alignment)).unwrap(),
                alignment
            );
        }
    }

    #[test]
    fn native_decimal_string_validation_stays_in_the_adapter() {
        assert_eq!(decimal_character_from_native("٫").unwrap().character(), '٫');
        assert!(decimal_character_from_native("").is_err());
        assert!(decimal_character_from_native("..").is_err());
        assert!(decimal_character_from_native("\n").is_err());
    }

    #[test]
    fn leaf_errors_map_to_the_iwa_error_at_the_boundary() {
        let error: Error = litchi_iwa_text::paragraph::tabs::Error::LeaderEmpty.into();
        assert!(matches!(error, Error::InvalidFormat(message) if message.contains("tab leader")));
    }
}
