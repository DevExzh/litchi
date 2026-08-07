//! Archive-free title settings for native iWork tables.
//!
//! The IWA adapter owns protobuf decoding, optional-field wire validation,
//! unknown-field preservation, and transactional package mutation. This leaf
//! owns only the compact semantic value exchanged at that boundary.

use std::mem::size_of;
use std::num::NonZeroU8;

const VALID: u8 = 1 << 4;
const VISIBLE_PRESENT: u8 = 1 << 0;
const VISIBLE_VALUE: u8 = 1 << 1;
const OUTLINED_PRESENT: u8 = 1 << 2;
const OUTLINED_VALUE: u8 = 1 << 3;

/// Lossless visibility and outline settings for a table title.
///
/// Each native optional Boolean uses one presence bit and one value bit. The
/// complete semantic value therefore occupies one byte while distinguishing
/// an absent field from an explicitly stored `false`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Settings(NonZeroU8);

impl Settings {
    /// Creates settings while preserving the presence of both native fields.
    #[must_use]
    pub const fn new(visible: Option<bool>, outlined: Option<bool>) -> Self {
        let bits = VALID
            | encode(visible, VISIBLE_PRESENT, VISIBLE_VALUE)
            | encode(outlined, OUTLINED_PRESENT, OUTLINED_VALUE);
        let value = match NonZeroU8::new(bits) {
            Some(value) => value,
            // `VALID` is always set above; this fallback keeps the constructor
            // total if the private encoding changes in the future.
            None => NonZeroU8::MIN,
        };
        Self(value)
    }

    /// Returns the lossless native visibility field.
    #[must_use]
    pub const fn visible(self) -> Option<bool> {
        decode(self.0.get(), VISIBLE_PRESENT, VISIBLE_VALUE)
    }

    /// Returns the lossless native outline field.
    #[must_use]
    pub const fn outlined(self) -> Option<bool> {
        decode(self.0.get(), OUTLINED_PRESENT, OUTLINED_VALUE)
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

impl Default for Settings {
    fn default() -> Self {
        Self::new(None, None)
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
    use std::mem::{align_of, size_of};

    use super::Settings;

    #[test]
    fn settings_are_compact_and_archive_free() {
        assert_eq!(size_of::<Settings>(), 1);
        assert_eq!(align_of::<Settings>(), 1);
        assert_eq!(size_of::<Option<Settings>>(), 1);
    }

    #[test]
    fn optional_boolean_presence_and_values_round_trip() {
        for visible in [None, Some(false), Some(true)] {
            for outlined in [None, Some(false), Some(true)] {
                let settings = Settings::new(visible, outlined);
                assert_eq!(settings.visible(), visible);
                assert_eq!(settings.outlined(), outlined);
            }
        }
    }

    #[test]
    fn effective_values_validate_only_explicit_true_as_enabled() {
        let combinations = [(None, false), (Some(false), false), (Some(true), true)];
        for (value, effective) in combinations {
            let settings = Settings::new(value, value);
            assert_eq!(settings.is_visible(), effective);
            assert_eq!(settings.is_outlined(), effective);
        }
    }
}
