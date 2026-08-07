//! Checked time values used by Keynote animations.

use crate::{Error, Result};

/// A finite, non-negative duration represented in seconds.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Seconds(f64);

impl Seconds {
    /// The zero duration.
    pub const ZERO: Self = Self(0.0);

    /// Construct a duration from seconds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDuration`] for a negative, infinite, or NaN
    /// value.
    pub const fn new(value: f64) -> Result<Self> {
        if value.is_finite() && value >= 0.0 {
            Ok(Self(value))
        } else {
            Err(Error::InvalidDuration)
        }
    }

    /// Return the duration in seconds.
    #[must_use]
    pub const fn as_f64(self) -> f64 {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use super::Seconds;
    use crate::Error;

    const COMPILE_TIME_SECOND: Seconds = match Seconds::new(1.0) {
        Ok(value) => value,
        Err(_) => Seconds::ZERO,
    };

    #[test]
    fn constructor_accepts_finite_non_negative_values() {
        assert_eq!(COMPILE_TIME_SECOND.as_f64(), 1.0);
        for value in [0.0, f64::MAX] {
            assert_eq!(Seconds::new(value).map(Seconds::as_f64), Ok(value));
        }
    }

    #[test]
    fn constructor_rejects_negative_and_non_finite_values() {
        for value in [-1.0, f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert_eq!(Seconds::new(value), Err(Error::InvalidDuration));
        }
    }
}
