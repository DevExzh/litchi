//! Duration display values.

use std::fmt;

/// Presentation style for a duration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Style {
    /// Display selected units as colon-separated fields.
    Colon,
    /// Display compact unit symbols.
    #[default]
    Abbreviated,
    /// Display complete unit names.
    FullNames,
}

/// A unit supported by a duration display.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Unit {
    /// Weeks.
    Weeks,
    /// Days.
    Days,
    /// Hours.
    Hours,
    /// Minutes.
    Minutes,
    /// Seconds.
    Seconds,
    /// Milliseconds.
    Milliseconds,
}

/// Errors returned by duration range construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The largest unit is finer than the smallest unit.
    ReversedRange { largest: Unit, smallest: Unit },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ReversedRange { largest, smallest } => write!(
                formatter,
                "duration range is reversed: largest unit {largest:?} follows smallest unit {smallest:?}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked duration constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// Inclusive range of displayed duration units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UnitRange {
    largest: Unit,
    smallest: Unit,
}

impl UnitRange {
    /// Validates and constructs a range from largest through smallest unit.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ReversedRange`] when the arguments are not ordered
    /// from coarse to fine.
    pub fn new(largest: Unit, smallest: Unit) -> Result<Self> {
        if (largest as u8) > (smallest as u8) {
            return Err(Error::ReversedRange { largest, smallest });
        }
        Ok(Self { largest, smallest })
    }

    /// The complete range from weeks through milliseconds.
    #[must_use]
    pub const fn all() -> Self {
        Self {
            largest: Unit::Weeks,
            smallest: Unit::Milliseconds,
        }
    }

    /// The common range from hours through milliseconds.
    #[must_use]
    pub const fn hours_to_milliseconds() -> Self {
        Self {
            largest: Unit::Hours,
            smallest: Unit::Milliseconds,
        }
    }

    /// Returns the largest displayed unit.
    #[must_use]
    pub const fn largest(self) -> Unit {
        self.largest
    }

    /// Returns the smallest displayed unit.
    #[must_use]
    pub const fn smallest(self) -> Unit {
        self.smallest
    }
}

impl Default for UnitRange {
    fn default() -> Self {
        Self::all()
    }
}

/// Automatic or explicit duration units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Units {
    /// Let the spreadsheet select visible units while retaining its range.
    Automatic(UnitRange),
    /// Display exactly the selected range.
    Custom(UnitRange),
}

impl Units {
    /// Returns the persisted unit range.
    #[must_use]
    pub const fn range(self) -> UnitRange {
        match self {
            Self::Automatic(range) | Self::Custom(range) => range,
        }
    }

    /// Returns whether the spreadsheet selects visible units automatically.
    #[must_use]
    pub const fn is_automatic(self) -> bool {
        matches!(self, Self::Automatic(_))
    }
}

impl Default for Units {
    fn default() -> Self {
        Self::Automatic(UnitRange::all())
    }
}

/// Duration display format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Duration {
    style: Style,
    units: Units,
}

impl Duration {
    /// Constructs a duration format from a style and unit policy.
    #[must_use]
    pub const fn new(style: Style, units: Units) -> Self {
        Self { style, units }
    }

    /// Constructs an automatic-unit duration format.
    #[must_use]
    pub const fn automatic(style: Style) -> Self {
        Self::new(style, Units::Automatic(UnitRange::all()))
    }

    /// Constructs a fixed-unit duration format.
    #[must_use]
    pub const fn custom(style: Style, range: UnitRange) -> Self {
        Self::new(style, Units::Custom(range))
    }

    /// Returns the presentation style.
    #[must_use]
    pub const fn style(self) -> Style {
        self.style
    }

    /// Returns the automatic or fixed unit policy.
    #[must_use]
    pub const fn units(self) -> Units {
        self.units
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ranges_reject_reverse_order() {
        assert!(matches!(
            UnitRange::new(Unit::Seconds, Unit::Hours),
            Err(Error::ReversedRange {
                largest: Unit::Seconds,
                smallest: Unit::Hours,
            })
        ));
    }

    #[test]
    fn duration_values_round_trip_through_accessors() {
        let range = UnitRange::hours_to_milliseconds();
        let value = Duration::custom(Style::Abbreviated, range);
        assert_eq!(value.style(), Style::Abbreviated);
        assert_eq!(value.units(), Units::Custom(range));
        assert!(!value.units().is_automatic());
        assert_eq!(value.units().range(), range);
    }
}
