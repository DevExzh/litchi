//! Archive-free options for inserting a body-anchored Pages audio clip.

use std::time::Duration;

use litchi_iwa_common::shape::geometry::Point;
use thiserror::Error;

/// Validation failures for Pages audio insertion options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// The audio control position contains a non-finite coordinate.
    #[error("Pages audio position must have finite coordinates")]
    InvalidPosition,
    /// The audio duration is zero or cannot be represented by native `f32` seconds.
    #[error("Pages audio duration must be positive and fit in finite f32 seconds")]
    InvalidDuration,
}

/// Result type for Pages audio option construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated placement and duration for a newly inserted body audio clip.
///
/// The value contains only common geometry and the canonical finite `f32`
/// duration required by the Pages native media record. Archive objects,
/// package records, and native identifiers remain in the IWA adapter.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    position: Point,
    duration_seconds: f32,
}

impl Options {
    /// Validate and canonicalize body-audio placement and duration.
    ///
    /// Both coordinates must be finite. The duration must be positive and
    /// representable in Pages' finite `f32`-seconds field.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidPosition`] for a non-finite coordinate and
    /// [`Error::InvalidDuration`] for zero or non-representable duration.
    pub fn new(position: Point, duration: Duration) -> Result<Self> {
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::InvalidPosition);
        }

        let precise_seconds = duration.as_secs_f64();
        if precise_seconds <= 0.0 || precise_seconds > f64::from(f32::MAX) {
            return Err(Error::InvalidDuration);
        }
        let duration_seconds = duration.as_secs_f32();
        if !duration_seconds.is_finite() || duration_seconds <= 0.0 {
            return Err(Error::InvalidDuration);
        }

        Ok(Self {
            position,
            duration_seconds,
        })
    }

    /// Return the center point of Pages' zero-size audio control.
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    /// Return the canonical duration represented by this value.
    #[must_use]
    pub fn duration(self) -> Duration {
        Duration::from_secs_f32(self.duration_seconds)
    }

    /// Return the canonical duration in Pages' native scalar domain.
    #[must_use]
    pub const fn duration_seconds(self) -> f32 {
        self.duration_seconds
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    const POSITION: Point = Point { x: 180.0, y: 240.0 };

    #[test]
    fn stores_validated_audio_options_without_heap_state() {
        let options = Options::new(POSITION, Duration::from_millis(1_375))
            .unwrap_or_else(|error| panic!("valid Pages audio options: {error}"));

        assert_eq!(options.position(), POSITION);
        assert_eq!(options.duration_seconds().to_bits(), 1.375_f32.to_bits());
        assert_eq!(options.duration(), Duration::from_millis(1_375));
        assert_eq!(size_of::<Options>(), 12);
    }

    #[test]
    fn rejects_non_finite_positions() {
        for position in [
            Point {
                x: f32::NAN,
                y: 0.0,
            },
            Point {
                x: 0.0,
                y: f32::INFINITY,
            },
            Point {
                x: f32::NEG_INFINITY,
                y: 0.0,
            },
        ] {
            assert_eq!(
                Options::new(position, Duration::from_secs(1)),
                Err(Error::InvalidPosition)
            );
        }
    }

    #[test]
    fn rejects_zero_duration() {
        assert_eq!(
            Options::new(POSITION, Duration::ZERO),
            Err(Error::InvalidDuration)
        );
    }

    #[test]
    fn accepts_the_duration_type_full_range_when_f32_representable() {
        let options = Options::new(POSITION, Duration::MAX)
            .unwrap_or_else(|error| panic!("maximum representable Pages audio duration: {error}"));

        assert!(options.duration_seconds().is_finite());
        assert!(options.duration_seconds() > 0.0);
    }
}
