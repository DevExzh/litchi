//! Archive-free options for inserting a body-anchored Pages movie.

use std::time::Duration;

use litchi_iwa_common::shape::geometry::{Point, Size};
use thiserror::Error;

/// Validation failures for Pages movie insertion options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// The movie position contains a non-finite coordinate.
    #[error("Pages movie position must have finite coordinates")]
    InvalidPosition,
    /// Displayed or natural movie dimensions are not finite and strictly positive.
    #[error("Pages movie dimensions must be finite and greater than zero")]
    InvalidSize,
    /// The movie duration is zero or cannot be represented by native `f32` seconds.
    #[error("Pages movie duration must be positive and fit in finite f32 seconds")]
    InvalidDuration,
}

/// Result type for Pages movie option construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated placement, dimensions, and duration for a new body movie.
///
/// The value contains only common geometry and the canonical finite `f32`
/// duration required by the Pages native media record. Archive objects,
/// package records, and native identifiers remain in the IWA adapter.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    position: Point,
    size: Size,
    natural_size: Size,
    duration_seconds: f32,
}

impl Options {
    /// Validate movie placement, displayed dimensions, and duration.
    ///
    /// The position must have finite coordinates. Displayed dimensions must
    /// be finite and strictly positive. The duration must be positive and
    /// representable in Pages' finite `f32`-seconds field.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the position, dimensions, or duration cannot
    /// be represented by the native movie fields.
    pub fn new(position: Point, size: Size, duration: Duration) -> Result<Self> {
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::InvalidPosition);
        }
        validate_size(size)?;
        let duration_seconds = duration_seconds(duration)?;
        Ok(Self {
            position,
            size,
            natural_size: size,
            duration_seconds,
        })
    }

    /// Return a copy with an independently validated natural media size.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSize`] when either natural dimension is
    /// non-finite or not strictly positive.
    pub fn with_natural_size(mut self, natural_size: Size) -> Result<Self> {
        validate_size(natural_size)?;
        self.natural_size = natural_size;
        Ok(self)
    }

    /// Return the top-left body position in points.
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    /// Return the displayed movie dimensions in points.
    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }

    /// Return the untransformed media dimensions in points.
    #[must_use]
    pub const fn natural_size(self) -> Size {
        self.natural_size
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

fn validate_size(size: Size) -> Result<()> {
    if !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::InvalidSize);
    }
    Ok(())
}

fn duration_seconds(duration: Duration) -> Result<f32> {
    let precise_seconds = duration.as_secs_f64();
    if precise_seconds <= 0.0 || precise_seconds > f64::from(f32::MAX) {
        return Err(Error::InvalidDuration);
    }
    let seconds = duration.as_secs_f32();
    if !seconds.is_finite() || seconds <= 0.0 {
        return Err(Error::InvalidDuration);
    }
    Ok(seconds)
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    const POSITION: Point = Point { x: 96.0, y: 144.0 };
    const SIZE: Size = Size {
        width: 320.0,
        height: 180.0,
    };
    const NATURAL_SIZE: Size = Size {
        width: 640.0,
        height: 360.0,
    };

    #[test]
    fn stores_validated_movie_options_without_heap_state() {
        let options = Options::new(POSITION, SIZE, Duration::from_millis(1_250))
            .unwrap_or_else(|error| panic!("valid Pages movie options: {error}"))
            .with_natural_size(NATURAL_SIZE)
            .unwrap_or_else(|error| panic!("valid Pages movie natural size: {error}"));

        assert_eq!(options.position(), POSITION);
        assert_eq!(options.size(), SIZE);
        assert_eq!(options.natural_size(), NATURAL_SIZE);
        assert_eq!(options.duration_seconds().to_bits(), 1.25_f32.to_bits());
        assert_eq!(options.duration(), Duration::from_millis(1_250));
        assert_eq!(size_of::<Options>(), 28);
    }

    #[test]
    fn rejects_non_finite_or_non_positive_geometry() {
        for position in [
            Point {
                x: f32::NAN,
                y: 0.0,
            },
            Point {
                x: 0.0,
                y: f32::INFINITY,
            },
        ] {
            assert_eq!(
                Options::new(position, SIZE, Duration::from_secs(1)),
                Err(Error::InvalidPosition)
            );
        }

        for size in [
            Size {
                width: 0.0,
                height: 1.0,
            },
            Size {
                width: -1.0,
                height: 1.0,
            },
            Size {
                width: f32::NAN,
                height: 1.0,
            },
            Size {
                width: 1.0,
                height: f32::INFINITY,
            },
        ] {
            assert_eq!(
                Options::new(POSITION, size, Duration::from_secs(1)),
                Err(Error::InvalidSize)
            );
        }

        assert_eq!(
            Options::new(POSITION, SIZE, Duration::from_secs(1))
                .unwrap_or_else(|error| panic!("valid Pages movie options: {error}"))
                .with_natural_size(Size {
                    width: 0.0,
                    height: 360.0,
                }),
            Err(Error::InvalidSize)
        );
    }

    #[test]
    fn rejects_zero_duration_before_package_work() {
        assert_eq!(
            Options::new(POSITION, SIZE, Duration::ZERO),
            Err(Error::InvalidDuration)
        );
    }

    #[test]
    fn accepts_the_duration_type_full_range_when_f32_representable() {
        let options = Options::new(POSITION, SIZE, Duration::MAX)
            .unwrap_or_else(|error| panic!("maximum representable Pages movie duration: {error}"));

        assert!(options.duration_seconds().is_finite());
        assert!(options.duration_seconds() > 0.0);
    }
}
