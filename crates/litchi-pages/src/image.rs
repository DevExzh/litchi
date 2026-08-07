//! Archive-free options for inserting a body-anchored Pages image.

use litchi_iwa_common::shape::geometry::{Point, Size};
use thiserror::Error;

/// Validation failures for Pages image insertion options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// The image position contains a non-finite coordinate.
    #[error("Pages image position must have finite coordinates")]
    InvalidPosition,
    /// Displayed or natural image dimensions are not finite and strictly positive.
    #[error("Pages image dimensions must be finite and greater than zero")]
    InvalidSize,
}

/// Result type for image option construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated placement and dimensions for a new body image.
///
/// The value contains only common geometry. Archive objects, package records,
/// media identifiers, and native graph relationships remain in the IWA
/// adapter.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    position: Point,
    size: Size,
    natural_size: Size,
}

impl Options {
    /// Validate image placement and displayed dimensions.
    ///
    /// The natural media dimensions initially match the displayed dimensions.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the position or dimensions cannot be
    /// represented by the native image fields.
    pub fn new(position: Point, size: Size) -> Result<Self> {
        validate_position(position)?;
        validate_size(size)?;
        Ok(Self {
            position,
            size,
            natural_size: size,
        })
    }

    /// Return a copy with independently validated natural media dimensions.
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

    /// Return the displayed image dimensions in points.
    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }

    /// Return the untransformed media dimensions in points.
    #[must_use]
    pub const fn natural_size(self) -> Size {
        self.natural_size
    }
}

fn validate_position(position: Point) -> Result<()> {
    if !position.x.is_finite() || !position.y.is_finite() {
        return Err(Error::InvalidPosition);
    }
    Ok(())
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

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    const POSITION: Point = Point { x: 96.0, y: 144.0 };
    const SIZE: Size = Size {
        width: 240.0,
        height: 180.0,
    };
    const NATURAL_SIZE: Size = Size {
        width: 640.0,
        height: 480.0,
    };

    #[test]
    fn stores_validated_image_options_without_heap_state() {
        let options = Options::new(POSITION, SIZE)
            .unwrap_or_else(|error| panic!("valid Pages image options: {error}"))
            .with_natural_size(NATURAL_SIZE)
            .unwrap_or_else(|error| panic!("valid Pages natural image size: {error}"));

        assert_eq!(options.position(), POSITION);
        assert_eq!(options.size(), SIZE);
        assert_eq!(options.natural_size(), NATURAL_SIZE);
        assert_eq!(size_of::<Options>(), 24);
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
            assert_eq!(Options::new(position, SIZE), Err(Error::InvalidPosition));
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
            assert_eq!(Options::new(POSITION, size), Err(Error::InvalidSize));
        }

        let options = Options::new(POSITION, SIZE)
            .unwrap_or_else(|error| panic!("valid Pages image options: {error}"));
        assert_eq!(
            options.with_natural_size(Size {
                width: 0.0,
                height: NATURAL_SIZE.height,
            }),
            Err(Error::InvalidSize)
        );
    }
}
