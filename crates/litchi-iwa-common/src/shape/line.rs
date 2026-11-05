//! Archive-independent line values shared by iWork owners: endpoint
//! decorations and the straight-line [`Segment`] with its native geometry
//! math (record 0375).

use super::geometry::{Point, Size};

const DEGREES_PER_RADIAN: f32 = 180.0 / std::f32::consts::PI;
const FULL_ROTATION_DEGREES: f32 = 360.0;
const COMPARISON_EPSILON: f32 = 0.001;

/// Validation failures of [`Segment`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// An endpoint coordinate was NaN or infinite.
    #[error("line endpoints must be finite")]
    NonFinite,
    /// The endpoints coincide (or their distance is not finite).
    #[error("line endpoints must have a finite positive distance")]
    Degenerate,
    /// A stored line geometry has no positive width or a non-zero height.
    #[error("stored straight-line geometry must have positive width and zero height")]
    InvalidGeometry,
}

/// A non-degenerate straight line between two document-space points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Segment {
    start: Point,
    end: Point,
}

impl Segment {
    /// Validate and construct a line from its endpoints.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::Degenerate`].
    pub fn new(start: Point, end: Point) -> Result<Self, Error> {
        if !start.x.is_finite() || !start.y.is_finite() || !end.x.is_finite() || !end.y.is_finite()
        {
            return Err(Error::NonFinite);
        }
        let length = (end.x - start.x).hypot(end.y - start.y);
        if !length.is_finite() || length <= 0.0 {
            return Err(Error::Degenerate);
        }
        Ok(Self { start, end })
    }

    /// The first endpoint.
    #[must_use]
    pub const fn start(self) -> Point {
        self.start
    }

    /// The second endpoint.
    #[must_use]
    pub const fn end(self) -> Point {
        self.end
    }

    /// The Euclidean length in document points.
    #[must_use]
    pub fn length(self) -> f32 {
        (self.end.x - self.start.x).hypot(self.end.y - self.start.y)
    }

    /// The same line in the opposite direction.
    #[must_use]
    pub const fn reversed(self) -> Self {
        Self {
            start: self.end,
            end: self.start,
        }
    }

    /// Whether both endpoints agree within 0.001 points.
    #[must_use]
    pub fn matches(self, other: Self) -> bool {
        let near = |left: f32, right: f32| (left - right).abs() <= COMPARISON_EPSILON;
        near(self.start.x, other.start.x)
            && near(self.start.y, other.start.y)
            && near(self.end.x, other.end.x)
            && near(self.end.y, other.end.y)
    }

    /// The native drawable geometry of this line: the unrotated position,
    /// the `(length, 0)` size and the rotation angle in `0°..360°`.
    #[must_use]
    pub fn geometry(self) -> (Point, Size, f32) {
        let delta_x = self.end.x - self.start.x;
        let delta_y = self.end.y - self.start.y;
        let length = delta_x.hypot(delta_y);
        let center_x = self.start.x + delta_x / 2.0;
        let center_y = self.start.y + delta_y / 2.0;
        let angle = (delta_y.atan2(delta_x) * DEGREES_PER_RADIAN).rem_euclid(FULL_ROTATION_DEGREES);
        (
            Point {
                x: center_x - length / 2.0,
                y: center_y,
            },
            Size {
                width: length,
                height: 0.0,
            },
            angle,
        )
    }

    /// Recover a line from its stored geometry and the path's natural width.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidGeometry`] for a non-positive width or a non-zero
    /// height, else the [`Self::new`] errors.
    pub fn from_geometry(
        position: Point,
        size: Size,
        angle: f32,
        natural_width: f32,
    ) -> Result<Self, Error> {
        if !size.width.is_finite()
            || size.width <= 0.0
            || !natural_width.is_finite()
            || natural_width <= 0.0
            || (size.height - 0.0).abs() > COMPARISON_EPSILON
        {
            return Err(Error::InvalidGeometry);
        }
        let radians = angle / DEGREES_PER_RADIAN;
        let half_x = radians.cos() * natural_width / 2.0;
        let half_y = radians.sin() * natural_width / 2.0;
        let center = Point {
            x: position.x + natural_width / 2.0,
            y: position.y,
        };
        Self::new(
            Point {
                x: center.x - half_x,
                y: center.y - half_y,
            },
            Point {
                x: center.x + half_x,
                y: center.y + half_y,
            },
        )
    }
}

/// A native endpoint decoration supported by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Endpoint {
    /// No endpoint decoration.
    #[default]
    None,
    /// A compact filled triangular arrowhead.
    SimpleArrow,
    /// A filled circular endpoint.
    FilledCircle,
    /// A filled diamond endpoint.
    FilledDiamond,
    /// An outlined arrowhead with a short center stem.
    OpenArrow,
    /// A broad filled arrowhead with an inset base.
    FilledArrow,
    /// A filled square endpoint.
    FilledSquare,
    /// An outlined square endpoint.
    OpenSquare,
    /// An outlined circular endpoint.
    OpenCircle,
    /// A filled arrowhead pointing toward the line segment.
    InvertedArrow,
    /// A perpendicular bar endpoint.
    Line,
}

/// Decorations at the directed start and end of a straight line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Endpoints {
    /// Decoration at the directed start point.
    pub start: Endpoint,
    /// Decoration at the directed end point.
    pub end: Endpoint,
}

impl Endpoints {
    /// Construct independently typed start and end decorations.
    #[must_use]
    pub const fn new(start: Endpoint, end: Endpoint) -> Self {
        Self { start, end }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Endpoint, Endpoints};

    #[test]
    fn endpoint_values_are_compact_and_copyable() {
        assert_eq!(size_of::<Endpoint>(), 1);
        assert_eq!(size_of::<Endpoints>(), 2);
        assert_eq!(
            Endpoints::new(Endpoint::None, Endpoint::SimpleArrow).start,
            Endpoint::None
        );
    }
}

#[cfg(test)]
mod segment_tests {
    use super::{Error, Point, Segment};

    #[test]
    fn segments_round_trip_through_their_geometry() {
        let segment = Segment::new(Point { x: 10.0, y: 20.0 }, Point { x: 40.0, y: 60.0 }).unwrap();
        assert_eq!(segment.length(), 50.0);
        let (position, size, angle) = segment.geometry();
        assert_eq!(size.width, 50.0);
        let back = Segment::from_geometry(position, size, angle, size.width).unwrap();
        assert!(back.matches(segment));
        assert!(segment.reversed().reversed().matches(segment));
        let point = Point { x: 1.0, y: 1.0 };
        assert_eq!(Segment::new(point, point), Err(Error::Degenerate));
        assert_eq!(
            Segment::new(
                Point {
                    x: f32::NAN,
                    y: 0.0
                },
                point
            ),
            Err(Error::NonFinite)
        );
        assert_eq!(
            Segment::from_geometry(
                point,
                super::Size {
                    width: 0.0,
                    height: 0.0
                },
                0.0,
                1.0
            ),
            Err(Error::InvalidGeometry)
        );
    }
}
