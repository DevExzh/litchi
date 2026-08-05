//! Archive-free values for independently positioned slide audio.

use std::time::Duration;

use litchi_iwa_common::shape::geometry::Point;

use crate::{Error, Result};

/// Validated placement and duration for a newly inserted audio clip.
///
/// The duration is stored in the same finite `f32` seconds domain used by
/// Keynote, so constructing this value performs the only lossy conversion.
/// Archive identifiers and package state deliberately remain outside this
/// semantic value.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    position: Point,
    duration_seconds: f32,
}

impl Options {
    /// Validate and canonicalize slide-audio placement and duration.
    ///
    /// Both coordinates must be finite. The duration must be positive and fit
    /// in Keynote's finite `f32`-seconds representation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidAudioPosition`] for a non-finite coordinate and
    /// [`Error::InvalidAudioDuration`] for zero or non-representable duration.
    pub fn new(position: Point, duration: Duration) -> Result<Self> {
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::InvalidAudioPosition);
        }

        let precise_seconds = duration.as_secs_f64();
        if precise_seconds == 0.0 || precise_seconds > f64::from(f32::MAX) {
            return Err(Error::InvalidAudioDuration);
        }
        let duration_seconds = duration.as_secs_f32();
        if !duration_seconds.is_finite() || duration_seconds <= 0.0 {
            return Err(Error::InvalidAudioDuration);
        }

        Ok(Self {
            position,
            duration_seconds,
        })
    }

    /// Return the center point of Keynote's zero-size audio control.
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    /// Return the canonical duration represented by this value.
    #[must_use]
    pub fn duration(self) -> Duration {
        Duration::from_secs_f32(self.duration_seconds)
    }

    /// Return the canonical duration in Keynote's native scalar domain.
    #[must_use]
    pub const fn duration_seconds(self) -> f32 {
        self.duration_seconds
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    const POSITION: Point = Point { x: 960.0, y: 540.0 };

    #[test]
    fn canonicalizes_duration_once_into_a_compact_value() {
        let options = Options::new(POSITION, Duration::from_millis(1_375)).unwrap();

        assert_eq!(options.position(), POSITION);
        assert_eq!(options.duration_seconds(), 1.375);
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
                Err(Error::InvalidAudioPosition)
            );
        }
    }

    #[test]
    fn rejects_zero_duration() {
        assert_eq!(
            Options::new(POSITION, Duration::ZERO),
            Err(Error::InvalidAudioDuration)
        );
    }

    #[test]
    fn accepts_the_duration_type_full_range_when_f32_representable() {
        let options = Options::new(POSITION, Duration::MAX).unwrap();

        assert!(options.duration_seconds().is_finite());
        assert!(options.duration_seconds() > 0.0);
    }
}
