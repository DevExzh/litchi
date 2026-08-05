//! Compact, presence-preserving title settings for Numbers tables.

use std::mem::size_of;

const VISIBLE_PRESENT: u8 = 1 << 0;
const VISIBLE_VALUE: u8 = 1 << 1;
const OUTLINED_PRESENT: u8 = 1 << 2;
const OUTLINED_VALUE: u8 = 1 << 3;

/// Lossless visibility and outline settings for a table title.
///
/// Each native optional boolean uses one presence bit and one value bit, so
/// the semantic value is one byte while still distinguishing an absent field
/// from an explicit `false`. Native object identifiers, protobuf messages,
/// and archive state remain outside this value.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Settings(u8);

impl Settings {
    /// Creates settings while preserving the presence of both native fields.
    #[must_use]
    pub const fn new(visible: Option<bool>, outlined: Option<bool>) -> Self {
        Self(
            encode(visible, VISIBLE_PRESENT, VISIBLE_VALUE)
                | encode(outlined, OUTLINED_PRESENT, OUTLINED_VALUE),
        )
    }

    /// Returns the lossless native visibility field.
    #[must_use]
    pub const fn visible(self) -> Option<bool> {
        decode(self.0, VISIBLE_PRESENT, VISIBLE_VALUE)
    }

    /// Returns the lossless native outline field.
    #[must_use]
    pub const fn outlined(self) -> Option<bool> {
        decode(self.0, OUTLINED_PRESENT, OUTLINED_VALUE)
    }

    /// Returns whether the title is effectively visible.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        matches!(self.visible(), Some(true))
    }

    /// Returns whether the title region is effectively outlined.
    #[must_use]
    pub const fn is_outlined(self) -> bool {
        matches!(self.outlined(), Some(true))
    }
}

const fn encode(value: Option<bool>, present: u8, bit: u8) -> u8 {
    match value {
        Some(true) => present | bit,
        Some(false) => present,
        None => 0,
    }
}

const fn decode(bits: u8, present: u8, value: u8) -> Option<bool> {
    if bits & present == 0 {
        None
    } else {
        Some(bits & value != 0)
    }
}

const _: () = assert!(size_of::<Settings>() == 1);

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_optional_boolean_presence_and_values() {
        for visible in [None, Some(false), Some(true)] {
            for outlined in [None, Some(false), Some(true)] {
                let settings = Settings::new(visible, outlined);
                assert_eq!(settings.visible(), visible);
                assert_eq!(settings.outlined(), outlined);
                assert_eq!(settings.is_visible(), visible == Some(true));
                assert_eq!(settings.is_outlined(), outlined == Some(true));
            }
        }
    }

    #[test]
    fn stays_one_byte_without_archive_state() {
        assert_eq!(size_of::<Settings>(), 1);
    }
}
