//! Validated, deterministic calculation timestamps for date/time formulas.
//!
//! The evaluator never reads a host clock.  A caller that wants `NOW`,
//! `TODAY`, or the omitted-year form of `EASTERSUNDAY` supplies one of these
//! values through `EvaluationOptions`.  The private bit representation keeps
//! the exact finite `f64` fraction while allowing the options structure to
//! remain `Copy` and `Eq`.

use super::calendar::{
    CivilDate, MAX_DATE_SERIAL_EXCLUSIVE, MIN_DATE_SERIAL, civil_from_serial, serial_from_civil,
};
use std::{error::Error, fmt};

/// A rejected calculation-timestamp construction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CalculationTimestampError {
    /// The supplied serial or seconds value was not finite.
    NonFinite,
    /// The serial was outside the profile's half-open date/time domain.
    OutOfRange,
    /// The civil date was not representable by the selected profile.
    InvalidDate,
    /// The clock components were outside one ordinary day.
    InvalidTime,
}

impl fmt::Display for CalculationTimestampError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NonFinite => "calculation timestamp must be finite",
            Self::OutOfRange => "calculation timestamp is outside the supported date range",
            Self::InvalidDate => "calculation timestamp contains an invalid civil date",
            Self::InvalidTime => "calculation timestamp contains an invalid clock time",
        })
    }
}

impl Error for CalculationTimestampError {}

/// An exact, validated serial snapshot used by volatile date/time formulas.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CalculationTimestamp {
    serial_bits: u64,
}

impl CalculationTimestamp {
    /// Construct a timestamp from a finite profile serial.
    ///
    /// The accepted interval is `[-693593, 2958466)`.  Negative zero is
    /// canonicalized so equivalent snapshots compare equal.
    pub fn from_serial(serial: f64) -> Result<Self, CalculationTimestampError> {
        if !serial.is_finite() {
            return Err(CalculationTimestampError::NonFinite);
        }
        if serial < MIN_DATE_SERIAL as f64 || serial >= MAX_DATE_SERIAL_EXCLUSIVE as f64 {
            return Err(CalculationTimestampError::OutOfRange);
        }
        let serial = if serial == 0.0 { 0.0 } else { serial };
        Ok(Self {
            serial_bits: serial.to_bits(),
        })
    }

    /// Construct a timestamp from a Gregorian civil date and clock.
    pub fn from_ymd_hms(
        year: i32,
        month: u32,
        day: u32,
        hour: u32,
        minute: u32,
        second: f64,
    ) -> Result<Self, CalculationTimestampError> {
        let date =
            CivilDate::new(year, month, day).ok_or(CalculationTimestampError::InvalidDate)?;
        if hour >= 24 || minute >= 60 || !second.is_finite() || !(0.0..60.0).contains(&second) {
            return if second.is_finite() {
                Err(CalculationTimestampError::InvalidTime)
            } else {
                Err(CalculationTimestampError::NonFinite)
            };
        }
        let date_serial = serial_from_civil(date).ok_or(CalculationTimestampError::OutOfRange)?;
        let day_fraction =
            (f64::from(hour) * 3_600.0 + f64::from(minute) * 60.0 + second) / 86_400.0;
        let serial = date_serial as f64 + day_fraction;
        // At large serial magnitudes an otherwise valid final subsecond can
        // round through the next integer day when it is added to the date.
        // Refuse that unrepresentable instant instead of silently changing
        // the civil date.
        if serial.floor() != date_serial as f64 {
            return Err(CalculationTimestampError::InvalidTime);
        }
        if !serial.is_finite()
            || serial < MIN_DATE_SERIAL as f64
            || serial >= MAX_DATE_SERIAL_EXCLUSIVE as f64
        {
            return Err(CalculationTimestampError::OutOfRange);
        }
        Self::from_serial(serial)
    }

    /// Return the exact profile serial represented by this timestamp.
    #[must_use]
    pub fn serial(self) -> f64 {
        f64::from_bits(self.serial_bits)
    }

    /// Return the timestamp's integer civil-date serial.
    #[must_use]
    pub fn date_serial(self) -> f64 {
        self.serial().floor()
    }

    /// Return the timestamp's nonnegative fraction of a day.
    #[must_use]
    pub fn time_fraction(self) -> f64 {
        // Reuse the checked splitter so a negative subnormal, whose raw
        // subtraction can round to one, exposes the same next-below-one
        // fraction as the date kernel.
        civil_from_serial(self.serial())
            .map(|(_, fraction)| fraction)
            .unwrap_or(0.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_fraction_and_canonicalizes_negative_zero() {
        let stamp = CalculationTimestamp::from_serial(46_000.5).unwrap();
        assert_eq!(stamp.serial(), 46_000.5);
        assert_eq!(stamp.date_serial(), 46_000.0);
        assert_eq!(stamp.time_fraction(), 0.5);

        let zero = CalculationTimestamp::from_serial(-0.0).unwrap();
        assert_eq!(zero.serial().to_bits(), 0.0f64.to_bits());

        let negative_subnormal = CalculationTimestamp::from_serial(-f64::from_bits(1)).unwrap();
        assert_eq!(
            negative_subnormal.serial().to_bits(),
            (-f64::from_bits(1)).to_bits()
        );
        assert!(negative_subnormal.time_fraction() < 1.0);
    }

    #[test]
    fn enforces_profile_bounds_and_clock_components() {
        assert!(matches!(
            CalculationTimestamp::from_serial(f64::NAN),
            Err(CalculationTimestampError::NonFinite)
        ));
        assert!(matches!(
            CalculationTimestamp::from_serial(2_958_466.0),
            Err(CalculationTimestampError::OutOfRange)
        ));
        assert!(CalculationTimestamp::from_ymd_hms(2000, 2, 29, 23, 59, 59.5).is_ok());
        assert!(matches!(
            CalculationTimestamp::from_ymd_hms(2000, 2, 30, 0, 0, 0.0),
            Err(CalculationTimestampError::InvalidDate)
        ));
        assert!(matches!(
            CalculationTimestamp::from_ymd_hms(2000, 1, 1, 24, 0, 0.0),
            Err(CalculationTimestampError::InvalidTime)
        ));
        let almost_sixty = f64::from_bits(60.0f64.to_bits() - 1);
        assert!(matches!(
            CalculationTimestamp::from_ymd_hms(2020, 1, 1, 23, 59, almost_sixty),
            Err(CalculationTimestampError::InvalidTime)
        ));
    }
}
