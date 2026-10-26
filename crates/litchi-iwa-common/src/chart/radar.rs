//! Archive-free geometry of native radar charts.

const FULL_TURN_DEGREES: f32 = 360.0;

/// Validation failures of [`StartAngle`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The value was NaN or infinite.
    #[error("radar start angle must be finite")]
    NonFinite,
    /// The value was outside `0°..360°`.
    #[error("radar start angle must be within 0°..360°")]
    OutOfRange,
}

/// Result type of [`StartAngle`] construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Rotation of the first radar axis (the Radar inspector's `Rotation`), in
/// clockwise degrees in `0°..360°`. Negative zero is stored as zero.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct StartAngle(f32);

impl StartAngle {
    /// No rotation, the native default.
    pub const ZERO: Self = Self(0.0);
    /// A quarter turn.
    pub const QUARTER_TURN: Self = Self(90.0);
    /// A half turn.
    pub const HALF_TURN: Self = Self(180.0);
    /// Three quarters of a turn.
    pub const THREE_QUARTER_TURN: Self = Self(270.0);

    /// Construct a canonical angle accepted by the native inspector.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] outside `0°..360°`.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(0.0..FULL_TURN_DEGREES).contains(&degrees) {
            return Err(Error::OutOfRange);
        }
        Ok(if degrees == 0.0 {
            Self::ZERO
        } else {
            Self(degrees)
        })
    }

    /// The clockwise degrees displayed by iWork (also the stored value).
    #[must_use]
    pub const fn degrees(self) -> f32 {
        self.0
    }
}

impl Default for StartAngle {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for StartAngle {
    type Error = Error;

    fn try_from(degrees: f32) -> Result<Self> {
        Self::from_degrees(degrees)
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, StartAngle};

    #[test]
    fn radar_start_angles_are_canonical() {
        assert_eq!(
            StartAngle::from_degrees(-0.0).map(|a| a.degrees().to_bits()),
            Ok(0)
        );
        assert_eq!(
            StartAngle::from_degrees(270.0),
            Ok(StartAngle::THREE_QUARTER_TURN)
        );
        assert_eq!(StartAngle::from_degrees(360.0), Err(Error::OutOfRange));
        assert_eq!(
            StartAngle::from_degrees(f32::INFINITY),
            Err(Error::NonFinite)
        );
        assert_eq!(StartAngle::default(), StartAngle::ZERO);
    }
}
