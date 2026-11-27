//! Archive-free geometry of native donut charts.

const MINIMUM_PERCENT: f32 = 1.0;
const DEFAULT_PERCENT: f32 = 75.0;
const MAXIMUM_PERCENT: f32 = 99.0;
const PERCENT_SCALE: f32 = 100.0;

/// Validation failures of [`InnerRadius`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The value was NaN or infinite.
    #[error("donut inner radius must be finite")]
    NonFinite,
    /// The value was outside `1%..=99%` (fraction `0.01..=0.99`).
    #[error("donut inner radius must be within 1%..=99%")]
    OutOfRange,
}

/// Result type of [`InnerRadius`] construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Size of the center hole of a donut chart (the Segments inspector's
/// `Inner Radius`), stored as a fraction and shown in percent.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct InnerRadius(f32);

impl InnerRadius {
    /// The smallest hole the inspector accepts (1%).
    pub const MINIMUM: Self = Self(MINIMUM_PERCENT / PERCENT_SCALE);
    /// The native hole when the property is absent (75%).
    pub const DEFAULT: Self = Self(DEFAULT_PERCENT / PERCENT_SCALE);
    /// The largest hole the inspector accepts (99%).
    pub const MAXIMUM: Self = Self(MAXIMUM_PERCENT / PERCENT_SCALE);

    /// Construct an inner radius from an inspector percentage.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `1%..=99%`.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(MINIMUM_PERCENT..=MAXIMUM_PERCENT).contains(&percent) {
            return Err(Error::OutOfRange);
        }
        Ok(Self(percent / PERCENT_SCALE))
    }

    /// Construct an inner radius from its stored fraction.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside
    /// [`Self::MINIMUM`]`..=`[`Self::MAXIMUM`].
    pub fn from_native(fraction: f32) -> Result<Self> {
        if !fraction.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(Self::MINIMUM.0..=Self::MAXIMUM.0).contains(&fraction) {
            return Err(Error::OutOfRange);
        }
        Ok(Self(fraction))
    }

    /// The percentage displayed by iWork.
    #[must_use]
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_SCALE
    }

    /// The stored fraction.
    #[must_use]
    pub const fn native_fraction(self) -> f32 {
        self.0
    }
}

impl Default for InnerRadius {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<f32> for InnerRadius {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, InnerRadius};

    #[test]
    fn inner_radius_is_an_inclusive_fraction() {
        assert_eq!(InnerRadius::default().native_fraction(), 0.75);
        assert_eq!(InnerRadius::from_percent(1.0), Ok(InnerRadius::MINIMUM));
        assert_eq!(InnerRadius::from_percent(99.0), Ok(InnerRadius::MAXIMUM));
        assert_eq!(
            InnerRadius::from_percent(50.0).map(InnerRadius::native_fraction),
            Ok(0.5)
        );
        assert_eq!(InnerRadius::from_percent(0.5), Err(Error::OutOfRange));
        assert_eq!(InnerRadius::from_percent(f32::NAN), Err(Error::NonFinite));
        assert_eq!(InnerRadius::from_native(0.75), Ok(InnerRadius::DEFAULT));
        assert_eq!(InnerRadius::from_native(1.0), Err(Error::OutOfRange));
    }
}
