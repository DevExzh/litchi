//! Archive-free selectors and options for semantic slide movies.

use std::time::Duration;

use litchi_core::Position;
use litchi_iwa_common::shape::geometry::{Point, Size};

use crate::{Error, Result};

/// Selects one movie by its zero-based position in the slide's source order.
///
/// The selector is deliberately archive-free: it contains a typed semantic
/// position only. Native object identifiers, component names, and protobuf
/// values remain in the adapter that resolves the selector against a package.
#[allow(
    clippy::module_name_repetitions,
    reason = "MovieSelector keeps the selected Keynote domain explicit at the public boundary"
)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MovieSelector {
    /// Select the movie at this zero-based source-order position.
    Index(Position),
}

impl MovieSelector {
    /// Create a selector from a zero-based source-order index.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(Position::new(index))
    }

    /// Create a selector from a typed zero-based source-order position.
    #[must_use]
    pub const fn position(position: Position) -> Self {
        Self::Index(position)
    }

    /// Return the selected typed source-order position.
    #[must_use]
    pub const fn as_position(self) -> Position {
        match self {
            Self::Index(position) => position,
        }
    }

    /// Return the selected zero-based source-order index.
    #[must_use]
    pub const fn as_index(self) -> usize {
        self.as_position().get()
    }
}

impl From<usize> for MovieSelector {
    fn from(index: usize) -> Self {
        Self::index(index)
    }
}

impl From<Position> for MovieSelector {
    fn from(position: Position) -> Self {
        Self::position(position)
    }
}

/// Validated placement, dimensions, and duration for a new slide movie.
///
/// The value stores only the finite native scalar representation required by
/// Keynote. Archive objects, package records, and native identifiers remain in
/// the IWA adapter.
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
    /// representable in Keynote's finite `f32`-seconds field.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the position, dimensions, or duration cannot
    /// be represented by the native movie fields.
    pub fn new(position: Point, size: Size, duration: Duration) -> Result<Self> {
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::InvalidMoviePosition);
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
    /// Returns [`Error::InvalidMovieSize`] when either natural dimension is
    /// non-finite or not strictly positive.
    pub fn with_natural_size(mut self, natural_size: Size) -> Result<Self> {
        validate_size(natural_size)?;
        self.natural_size = natural_size;
        Ok(self)
    }

    /// Return the top-left slide position in points.
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

    /// Return the canonical duration in Keynote's native scalar domain.
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
        return Err(Error::InvalidMovieSize);
    }
    Ok(())
}

fn duration_seconds(duration: Duration) -> Result<f32> {
    let precise_seconds = duration.as_secs_f64();
    if precise_seconds <= 0.0 || precise_seconds > f64::from(f32::MAX) {
        return Err(Error::InvalidMovieDuration);
    }
    let seconds = duration.as_secs_f32();
    if !seconds.is_finite() || seconds <= 0.0 {
        return Err(Error::InvalidMovieDuration);
    }
    Ok(seconds)
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    const POSITION: Point = Point { x: 100.0, y: 120.0 };
    const SIZE: Size = Size {
        width: 640.0,
        height: 360.0,
    };
    const NATURAL_SIZE: Size = Size {
        width: 1_280.0,
        height: 720.0,
    };

    #[test]
    fn selector_stays_typed_and_source_ordered() {
        const POSITION: Position = Position::new(3);
        const SELECTOR: MovieSelector = MovieSelector::position(POSITION);

        assert_eq!(SELECTOR, MovieSelector::Index(POSITION));
        assert_eq!(SELECTOR.as_position(), POSITION);
        assert_eq!(SELECTOR.as_index(), 3);
    }

    #[test]
    fn selector_constructors_and_conversions_do_not_introduce_native_identity() {
        let from_index = MovieSelector::index(5);
        let from_position = MovieSelector::from(Position::new(5));
        let from_usize = MovieSelector::from(5usize);

        assert_eq!(from_index, from_position);
        assert_eq!(from_index, from_usize);
        assert_eq!(from_index.as_position(), Position::new(5));
    }

    #[test]
    fn stores_validated_movie_options_without_heap_state() {
        let options = Options::new(POSITION, SIZE, Duration::from_millis(1_250))
            .unwrap()
            .with_natural_size(NATURAL_SIZE)
            .unwrap();

        assert_eq!(options.position(), POSITION);
        assert_eq!(options.size(), SIZE);
        assert_eq!(options.natural_size(), NATURAL_SIZE);
        assert_eq!(options.duration_seconds(), 1.25);
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
                Err(Error::InvalidMoviePosition)
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
                Err(Error::InvalidMovieSize)
            );
        }

        assert_eq!(
            Options::new(POSITION, SIZE, Duration::from_secs(1))
                .unwrap()
                .with_natural_size(Size {
                    width: 0.0,
                    height: 720.0,
                }),
            Err(Error::InvalidMovieSize)
        );
    }

    #[test]
    fn rejects_zero_duration_before_package_work() {
        assert_eq!(
            Options::new(POSITION, SIZE, Duration::ZERO),
            Err(Error::InvalidMovieDuration)
        );
    }

    #[test]
    fn accepts_the_duration_type_full_range_when_f32_representable() {
        let options = Options::new(POSITION, SIZE, Duration::MAX).unwrap();

        assert!(options.duration_seconds().is_finite());
        assert!(options.duration_seconds() > 0.0);
    }
}
