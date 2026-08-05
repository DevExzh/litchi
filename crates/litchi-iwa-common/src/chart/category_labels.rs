//! Archive-free category-label values for native charts.
//!
//! The native category-label interval is an int32 domain: zero means
//! automatic fitting, one means every category, and positive values from two
//! onward are explicit intervals. Negative values are retained as unknown
//! native values so a reader can round-trip a future producer state without
//! pretending to understand it. Protobuf decoding and package mutation stay
//! in litchi-iwa.

const MINIMUM_CUSTOM_INTERVAL: u32 = 2;
const MAXIMUM_CUSTOM_INTERVAL: u32 = i32::MAX as u32;

/// Validation failures for an explicit category-label interval.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The interval is reserved for a canonical native frequency or exceeds
    /// the signed native field range.
    #[error(
        "chart category-label interval {value} must be in {MINIMUM_CUSTOM_INTERVAL}..={MAXIMUM_CUSTOM_INTERVAL}"
    )]
    IntervalOutOfRange { value: u32 },
}

/// Result type for category-label value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated number of categories between displayed labels.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Interval(u32);

impl Interval {
    /// The smallest explicit native category-label interval.
    pub const MINIMUM: Self = Self(MINIMUM_CUSTOM_INTERVAL);
    /// The largest explicit native category-label interval.
    pub const MAXIMUM: Self = Self(MAXIMUM_CUSTOM_INTERVAL);

    /// Construct an explicit category-label interval.
    ///
    /// Native values zero and one are reserved for the automatic and all
    /// frequencies, respectively.
    ///
    /// # Errors
    ///
    /// Returns the interval-out-of-range error when the value is zero, one,
    /// or outside the signed native field range.
    #[must_use = "use the validated interval or handle the validation error"]
    pub const fn new(value: u32) -> Result<Self> {
        if value < MINIMUM_CUSTOM_INTERVAL || value > MAXIMUM_CUSTOM_INTERVAL {
            return Err(Error::IntervalOutOfRange { value });
        }
        Ok(Self(value))
    }

    /// Return the number of categories between displayed labels.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for Interval {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// How frequently a native chart displays category labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Frequency {
    /// Hide ordinary category labels. The native interval may remain stored
    /// underneath this visibility state for lossless reset.
    None,
    /// Let iWork choose a readable interval for the available chart width.
    #[default]
    AutoFit,
    /// Display every category label.
    All,
    /// Display labels at one explicit category interval.
    Every(Interval),
    /// Preserve an unrecognized native signed interval value.
    Unsupported(i32),
}

impl Frequency {
    /// Decode the native category-label interval while preserving unknown
    /// signed values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        match value {
            0 => Self::AutoFit,
            1 => Self::All,
            2..=i32::MAX => Self::Every(Interval(value.cast_unsigned())),
            other => Self::Unsupported(other),
        }
    }

    /// Return the native signed interval, or None when labels are hidden.
    #[must_use]
    pub const fn native_value(self) -> Option<i32> {
        match self {
            Self::None => None,
            Self::AutoFit => Some(0),
            Self::All => Some(1),
            Self::Every(interval) => Some(interval.value().cast_signed()),
            Self::Unsupported(value) => Some(value),
        }
    }

    /// Whether ordinary category labels are visible.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        !matches!(self, Self::None)
    }
}

/// Complete native category-label menu state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Layout {
    frequency: Frequency,
    show_last_category: bool,
}

impl Layout {
    /// Construct category-label layout settings.
    #[must_use]
    pub const fn new(frequency: Frequency, show_last_category: bool) -> Self {
        Self {
            frequency,
            show_last_category,
        }
    }

    /// Return how frequently category labels are displayed.
    #[must_use]
    pub const fn frequency(self) -> Frequency {
        self.frequency
    }

    /// Return whether iWork forces the final category label to appear.
    #[must_use]
    pub const fn show_last_category(self) -> bool {
        self.show_last_category
    }
}

impl Default for Layout {
    fn default() -> Self {
        Self::new(Frequency::AutoFit, true)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, Frequency, Interval, Layout};

    #[test]
    fn intervals_are_compact_and_strict() {
        assert_eq!(size_of::<Interval>(), 4);
        assert_eq!(
            Interval::new(0),
            Err(Error::IntervalOutOfRange { value: 0 })
        );
        assert_eq!(
            Interval::new(1),
            Err(Error::IntervalOutOfRange { value: 1 })
        );
        assert_eq!(Interval::new(2).unwrap().value(), 2);
        assert_eq!(Interval::new(i32::MAX as u32).unwrap(), Interval::MAXIMUM);
        assert_eq!(
            Interval::new(i32::MAX as u32 + 1),
            Err(Error::IntervalOutOfRange {
                value: i32::MAX as u32 + 1,
            })
        );
    }

    #[test]
    fn frequencies_preserve_canonical_and_unknown_native_values() {
        assert_eq!(size_of::<Frequency>(), 8);
        for (native, frequency) in [
            (0, Frequency::AutoFit),
            (1, Frequency::All),
            (3, Frequency::Every(Interval::new(3).unwrap())),
            (-7, Frequency::Unsupported(-7)),
        ] {
            assert_eq!(Frequency::from_native(native), frequency);
            assert_eq!(frequency.native_value(), Some(native));
        }
        assert_eq!(Frequency::None.native_value(), None);
        assert!(!Frequency::None.is_visible());
        assert!(Frequency::Unsupported(-7).is_visible());
    }

    #[test]
    fn layouts_are_copyable_and_default_to_native_defaults() {
        assert_eq!(size_of::<Layout>(), 12);
        assert_eq!(Layout::default(), Layout::new(Frequency::AutoFit, true));
        let layout = Layout::new(Frequency::Every(Interval::new(3).unwrap()), false);
        assert_eq!(
            layout.frequency(),
            Frequency::Every(Interval::new(3).unwrap())
        );
        assert!(!layout.show_last_category());
    }
}
