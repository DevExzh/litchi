//! Archive-free chart gap vocabulary.
//!
//! The native inspector expresses the space between bars or columns and the
//! space between sets as percentages. Native archive decoding and mutation
//! remain in the concrete IWA adapter; this module only owns the validated,
//! compact semantic values exchanged at that boundary.

/// Validation failures for chart-gap percentages.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The supplied percentage was NaN or infinite.
    #[error("chart gap percentage must be finite")]
    NonFinite,
    /// The supplied percentage was outside the native inclusive domain.
    #[error("chart gap percentage must be in 0.0..=999.0")]
    OutOfRange,
}

/// Result type for chart-gap value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated native chart-gap percentage.
//
// The native archive stores an `f32`; retaining that representation avoids a
// conversion and preserves fractional values exactly at the archive boundary.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Percentage(f32);

impl Percentage {
    /// No space between chart elements.
    pub const ZERO: Self = Self(0.0);

    /// The largest percentage accepted by the native inspector.
    pub const MAXIMUM: Self = Self(999.0);

    /// Construct a percentage in the native inclusive domain.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinity and
    /// [`Error::OutOfRange`] outside `0.0..=999.0`.
    #[must_use = "use the validated percentage or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(0.0..=999.0).contains(&value) {
            return Err(Error::OutOfRange);
        }
        Ok(Self(value))
    }

    /// Return the percentage represented by this value.
    #[must_use]
    pub const fn percent(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for Percentage {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// Native spacing within a set and between adjacent chart sets.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Spacing {
    between_items: Percentage,
    between_sets: Percentage,
}

impl Spacing {
    /// The spacing used by newly inserted native iWork charts.
    pub const DEFAULT: Self = Self::new(Percentage(10.0), Percentage(40.0));

    /// Construct spacing within a set and between adjacent sets.
    #[must_use]
    pub const fn new(between_items: Percentage, between_sets: Percentage) -> Self {
        Self {
            between_items,
            between_sets,
        }
    }

    /// Return the gap between individual bars or columns within a set.
    #[must_use]
    pub const fn between_items(self) -> Percentage {
        self.between_items
    }

    /// Return the gap between adjacent sets of bars or columns.
    #[must_use]
    pub const fn between_sets(self) -> Percentage {
        self.between_sets
    }
}

impl Default for Spacing {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{Error, Percentage, Spacing};

    #[test]
    fn percentage_is_compact_and_validated() {
        assert_eq!(size_of::<Percentage>(), 4);
        assert_eq!(align_of::<Percentage>(), 4);
        assert_eq!(Percentage::new(f32::NAN), Err(Error::NonFinite));
        assert_eq!(Percentage::new(f32::INFINITY), Err(Error::NonFinite));
        assert_eq!(Percentage::new(-0.1), Err(Error::OutOfRange));
        assert_eq!(Percentage::new(999.1), Err(Error::OutOfRange));
        assert_eq!(Percentage::new(12.5).unwrap().percent(), 12.5);
        assert_eq!(Percentage::MAXIMUM.percent(), 999.0);
    }

    #[test]
    fn spacing_is_compact_and_matches_native_defaults() {
        assert_eq!(size_of::<Spacing>(), 8);
        assert_eq!(align_of::<Spacing>(), 4);
        assert_eq!(
            Spacing::DEFAULT,
            Spacing::new(
                Percentage::new(10.0).unwrap(),
                Percentage::new(40.0).unwrap()
            )
        );
        assert_eq!(Spacing::default(), Spacing::DEFAULT);
    }
}
