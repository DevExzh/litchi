//! Archive-free label angles for native chart axes.

const MINIMUM_DEGREES: f32 = 0.0;
const MAXIMUM_DEGREES_EXCLUSIVE: f32 = 360.0;

/// Validation failures for a chart-axis label angle.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The angle was NaN or infinite.
    #[error("chart axis label angle must be finite")]
    NonFinite,
    /// The angle was outside the native half-open degree range.
    #[error("chart axis label angle must be in [0, 360) degrees")]
    OutOfRange,
}

/// Result type for label-angle construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A normalized chart-axis label angle in degrees.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct LabelAngle(f32);

impl LabelAngle {
    /// Horizontal labels.
    pub const HORIZONTAL: Self = Self(MINIMUM_DEGREES);
    /// Labels rising diagonally toward the left.
    pub const LEFT_DIAGONAL: Self = Self(45.0);
    /// Labels written vertically toward the left.
    pub const LEFT_VERTICAL: Self = Self(90.0);
    /// Labels written vertically toward the right.
    pub const RIGHT_VERTICAL: Self = Self(270.0);
    /// Labels rising diagonally toward the right.
    pub const RIGHT_DIAGONAL: Self = Self(315.0);

    /// Construct a normalized native label angle.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinity and
    /// [`Error::OutOfRange`] outside the native half-open degree range.
    #[must_use = "use the validated angle or handle its validation error"]
    pub fn new(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(MINIMUM_DEGREES..MAXIMUM_DEGREES_EXCLUSIVE).contains(&degrees) {
            return Err(Error::OutOfRange);
        }
        // Canonicalize negative zero so equality and wire output are stable.
        Ok(Self(if degrees == MINIMUM_DEGREES {
            MINIMUM_DEGREES
        } else {
            degrees
        }))
    }

    /// Return the normalized angle in degrees.
    #[must_use]
    pub const fn degrees(self) -> f32 {
        self.0
    }
}

impl Default for LabelAngle {
    fn default() -> Self {
        Self::HORIZONTAL
    }
}

impl TryFrom<f32> for LabelAngle {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, LabelAngle};

    #[test]
    fn angles_are_compact_normalized_and_strict() {
        assert_eq!(size_of::<LabelAngle>(), 4);
        assert_eq!(LabelAngle::default(), LabelAngle::HORIZONTAL);
        assert_eq!(
            LabelAngle::new(-0.0).unwrap().degrees().to_bits(),
            0.0_f32.to_bits()
        );
        assert_eq!(LabelAngle::new(f32::NAN), Err(Error::NonFinite));
        assert_eq!(LabelAngle::new(360.0), Err(Error::OutOfRange));
    }
}
