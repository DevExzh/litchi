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
    pub fn new(value: f64) -> Result<Self> {
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
