//! Archive-free rounded-corner values of native charts.

const MAXIMUM_PERCENT: f32 = 100.0;

/// Validation failures of [`Radius`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The value was NaN or infinite.
    #[error("chart corner radius must be finite")]
    NonFinite,
    /// The value was outside `0%..=100%` (fraction `0.0..=1.0`).
    #[error("chart corner radius must be within 0%..=100%")]
    OutOfRange,
}

/// Result type of [`Radius`] construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A chart corner radius (the inspector's `Corner Radius`), stored as a
/// normalized fraction and shown in percent.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Radius(f32);

impl Radius {
    /// Square corners.
    pub const ZERO: Self = Self(0.0);

    /// Construct a radius from the inspector percentage.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `0%..=100%`.
    pub fn new(percent: f32) -> Result<Self> {
        if !percent.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(0.0..=MAXIMUM_PERCENT).contains(&percent) {
            return Err(Error::OutOfRange);
        }
        Ok(Self(percent / MAXIMUM_PERCENT))
    }

    /// Construct a radius from its stored fraction.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `0.0..=1.0`.
    pub fn from_native(fraction: f32) -> Result<Self> {
        if !fraction.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(0.0..=1.0).contains(&fraction) {
            return Err(Error::OutOfRange);
        }
        Ok(Self(fraction))
    }

    /// The percentage displayed by iWork.
    #[must_use]
    pub const fn percent(self) -> f32 {
        self.0 * MAXIMUM_PERCENT
    }

    /// The stored fraction.
    #[must_use]
    pub const fn native_fraction(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for Radius {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::new(percent)
    }
}

/// Rounded-corner settings of a chart: the radius and the inspector's
/// `Outside Corners Only` switch.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rounding {
    radius: Radius,
    outside_corners_only: bool,
}

impl Rounding {
    /// Square corners without the outside-corners override, the native
    /// default.
    pub const NONE: Self = Self::new(Radius::ZERO, false);

    /// Construct rounded-corner settings.
    #[must_use]
    pub const fn new(radius: Radius, outside_corners_only: bool) -> Self {
        Self {
            radius,
            outside_corners_only,
        }
    }

    /// The corner radius.
    #[must_use]
    pub const fn radius(self) -> Radius {
        self.radius
    }

    /// Whether only the outside corners are rounded.
    #[must_use]
    pub const fn outside_corners_only(self) -> bool {
        self.outside_corners_only
    }
}

impl Default for Rounding {
    fn default() -> Self {
        Self::NONE
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Radius, Rounding};

    #[test]
    fn corner_radius_is_a_normalized_percentage() {
        let quarter = Radius::new(25.0).unwrap();
        assert_eq!(quarter.native_fraction(), 0.25);
        assert_eq!(quarter.percent(), 25.0);
        assert_eq!(Radius::new(100.5), Err(Error::OutOfRange));
        assert_eq!(Radius::try_from(f32::NAN), Err(Error::NonFinite));
        assert_eq!(Radius::from_native(1.5), Err(Error::OutOfRange));
        let rounding = Rounding::new(quarter, true);
        assert_eq!(
            (rounding.radius(), rounding.outside_corners_only()),
            (quarter, true)
        );
        assert_eq!(Rounding::default(), Rounding::NONE);
    }
}
