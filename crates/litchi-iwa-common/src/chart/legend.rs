//! Archive-free chart-legend frame values (record 0375).

/// Validation failures of the legend-frame values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A coordinate was not finite.
    #[error("chart legend coordinate must be finite")]
    CoordinateNonFinite,
    /// An extent was not finite and non-negative.
    #[error("chart legend extent must be finite and non-negative")]
    ExtentOutOfRange,
}

/// Result type of the legend-frame values.
pub type Result<T> = std::result::Result<T, Error>;

/// A finite legend-frame coordinate in native chart units.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Coordinate(f32);

impl Coordinate {
    /// Coordinate zero.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite coordinate.
    ///
    /// # Errors
    ///
    /// [`Error::CoordinateNonFinite`].
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::CoordinateNonFinite);
        }
        Ok(Self(points))
    }

    /// The coordinate in native chart units.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// A finite, non-negative legend-frame extent; zero asks iWork to measure
/// that axis.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Extent(f32);

impl Extent {
    /// Automatic measurement.
    pub const AUTOMATIC: Self = Self(0.0);

    /// Construct a finite, non-negative extent.
    ///
    /// # Errors
    ///
    /// [`Error::ExtentOutOfRange`].
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ExtentOutOfRange);
        }
        Ok(Self(points))
    }

    /// The extent in native chart units.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }

    /// Whether iWork measures this axis.
    #[must_use]
    pub const fn is_automatic(self) -> bool {
        self.0 == 0.0
    }
}

/// An explicit legend rectangle.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rect {
    x: Coordinate,
    y: Coordinate,
    width: Extent,
    height: Extent,
}

impl Rect {
    /// Construct a rectangle from typed parts.
    #[must_use]
    pub const fn new(x: Coordinate, y: Coordinate, width: Extent, height: Extent) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Validate and construct a rectangle from native chart units.
    ///
    /// # Errors
    ///
    /// [`Error::CoordinateNonFinite`] or [`Error::ExtentOutOfRange`].
    pub fn from_points(x: f32, y: f32, width: f32, height: f32) -> Result<Self> {
        Ok(Self::new(
            Coordinate::new(x)?,
            Coordinate::new(y)?,
            Extent::new(width)?,
            Extent::new(height)?,
        ))
    }

    /// The horizontal origin.
    #[must_use]
    pub const fn x(self) -> Coordinate {
        self.x
    }

    /// The vertical origin.
    #[must_use]
    pub const fn y(self) -> Coordinate {
        self.y
    }

    /// The width.
    #[must_use]
    pub const fn width(self) -> Extent {
        self.width
    }

    /// The height.
    #[must_use]
    pub const fn height(self) -> Extent {
        self.height
    }

    /// The stored `[x, y, width, height]`.
    #[must_use]
    pub const fn native(self) -> [f32; 4] {
        [self.x.0, self.y.0, self.width.0, self.height.0]
    }
}

/// The legend placement of a chart.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Frame {
    /// Placed and measured by the chart layout engine.
    #[default]
    Automatic,
    /// An explicit origin and optional fixed extents.
    Frame(Rect),
}

impl Frame {
    /// The frame a stored rectangle (`None` for automatic) describes.
    ///
    /// # Errors
    ///
    /// The [`Rect::from_points`] errors.
    pub fn from_native(stored: Option<[f32; 4]>) -> Result<Self> {
        stored.map_or(Ok(Self::Automatic), |[x, y, width, height]| {
            Rect::from_points(x, y, width, height).map(Self::Frame)
        })
    }

    /// The stored rectangle, `None` for automatic placement.
    #[must_use]
    pub const fn native(self) -> Option<[f32; 4]> {
        match self {
            Self::Automatic => None,
            Self::Frame(rect) => Some(rect.native()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Extent, Frame, Rect};

    #[test]
    fn legend_frames_validate_and_round_trip() {
        let rect = Rect::from_points(10.0, 20.0, 0.0, 30.0).unwrap();
        assert!(rect.width().is_automatic());
        let frame = Frame::Frame(rect);
        assert_eq!(Frame::from_native(frame.native()), Ok(frame));
        assert_eq!(Frame::from_native(None), Ok(Frame::Automatic));
        assert_eq!(Extent::new(-1.0), Err(Error::ExtentOutOfRange));
        assert_eq!(
            Rect::from_points(f32::NAN, 0.0, 0.0, 0.0),
            Err(Error::CoordinateNonFinite)
        );
    }
}
