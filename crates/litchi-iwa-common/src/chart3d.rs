//! Archive-free values for three-dimensional charts.

use crate::chart::kind::Kind;

const RECTANGLE_NATIVE_VALUE: i32 = 0;
const CYLINDER_NATIVE_VALUE: i32 = 1;

/// The bar or column geometry used by a native three-dimensional chart.
///
/// The native integer is retained directly so the value remains four bytes
/// and unrecognized future values survive a read-modify-write cycle without
/// allocation or lossy fallback.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct BarShape(i32);

#[allow(
    non_upper_case_globals,
    reason = "PascalCase associated constants are the focused BarShape API"
)]
impl BarShape {
    /// Rectangular prisms, the native default.
    pub const Rectangle: Self = Self(RECTANGLE_NATIVE_VALUE);
    /// Circular cylinders.
    pub const Cylinder: Self = Self(CYLINDER_NATIVE_VALUE);

    /// Decode the integer stored by a native iWork archive.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the integer used by a native iWork archive.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Whether this value is not one of the known native bar shapes.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        self.0 != RECTANGLE_NATIVE_VALUE && self.0 != CYLINDER_NATIVE_VALUE
    }
}

impl std::fmt::Debug for BarShape {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match *self {
            Self::Rectangle => formatter.write_str("Rectangle"),
            Self::Cylinder => formatter.write_str("Cylinder"),
            Self(value) => formatter.debug_tuple("Unsupported").field(&value).finish(),
        }
    }
}

const PERCENT_SCALE: f32 = 100.0;
const MAXIMUM_DEPTH_PERCENT: f32 = 100.0;
const DEFAULT_DEPTH_PERCENT: f32 = 11.020_468;
const DEPTH_PERCENT_EQUALITY_EPSILON: f32 = 0.001;
// The native default scene dimensions of a chart that has not been opened in
// an iWork application yet; later layout passes may adjust them by a few ULPs.
const DEFAULT_SCALE_X: f32 = 1.5;
const DEFAULT_SCALE_Y: f32 = 2.2;
const DEFAULT_SCALE_W: f32 = f32::from_bits(0x3f22_b0c9);
const MINIMUM_ROTATION_DEGREES: f32 = -40.0;
const MAXIMUM_ROTATION_DEGREES: f32 = 40.0;
const MAXIMUM_SERIES_GAP_PERCENT: u16 = 200;
const DEFAULT_SERIES_GAP_PERCENT: u16 = 100;
const WHOLE_PERCENT_TOLERANCE: f32 = 0.001;

/// Validation failures of the three-dimensional scene values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A value was NaN or infinite.
    #[error("3D chart value must be finite")]
    NonFinite,
    /// A value was outside the native inspector's domain.
    #[error("3D chart value is outside the native inspector range")]
    OutOfRange,
    /// A stored scale vector cannot carry a depth (a non-finite component or
    /// a non-positive depth denominator).
    #[error("stored 3D chart scale vector is invalid")]
    InvalidScale,
}

/// Result type of the three-dimensional scene values.
pub type Result<T> = std::result::Result<T, Error>;

/// The per-kind native scale vector that stores a 3D chart's depth.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Scale {
    /// Area and stacked-area 3D charts.
    Area,
    /// Bar and stacked-bar 3D charts.
    Bar,
    /// Column and stacked-column 3D charts.
    Column,
    /// Line 3D charts.
    Line,
    /// Pie and donut 3D charts.
    Pie,
}

impl Scale {
    /// The scale of a chart kind, or `None` for a kind without 3D depth.
    #[must_use]
    pub const fn for_kind(kind: Kind) -> Option<Self> {
        match kind {
            Kind::Area3d | Kind::StackedArea3d => Some(Self::Area),
            Kind::Bar3d | Kind::StackedBar3d => Some(Self::Bar),
            Kind::Column3d | Kind::StackedColumn3d => Some(Self::Column),
            Kind::Line3d => Some(Self::Line),
            Kind::Pie3d | Kind::Donut3d => Some(Self::Pie),
            _ => None,
        }
    }

    const fn ratio_bounds(self) -> (f32, f32) {
        match self {
            Self::Area | Self::Line => (1.0 / 3.0, 3.0),
            Self::Bar => (1.0 / 6.0, 4.0),
            Self::Column => (1.0 / 6.0, 2.5),
            Self::Pie => (0.25, 4.0),
        }
    }

    /// Whether the chart is horizontal (bar charts): the depth ratio is then
    /// relative to the vector's `y` component instead of `x`.
    #[must_use]
    pub const fn is_horizontal(self) -> bool {
        matches!(self, Self::Bar)
    }

    /// The vector component the depth ratio is relative to.
    const fn denominator(self, x: f32, y: f32) -> f32 {
        if self.is_horizontal() { y } else { x }
    }
}

/// Depth of a 3D chart, in Chart-inspector percent (`0%..=100%`).
///
/// The native archive stores the depth as the `z` component of a per-kind
/// scale vector, relative to the chart's horizontal dimension. Two depths
/// compare equal within 0.001 percentage points, which absorbs the
/// round-trip through that ratio.
#[derive(Debug, Clone, Copy)]
pub struct Depth {
    percent: f32,
}

impl PartialEq for Depth {
    fn eq(&self, other: &Self) -> bool {
        (self.percent - other.percent).abs() <= DEPTH_PERCENT_EQUALITY_EPSILON
    }
}

impl Depth {
    /// The depth a chart has when no scale vector is stored.
    pub const DEFAULT: Self = Self {
        percent: DEFAULT_DEPTH_PERCENT,
    };

    /// Construct a depth in the inspector's inclusive `0%..=100%` range.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`].
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite() {
            return Err(Error::NonFinite);
        }
        if !(0.0..=MAXIMUM_DEPTH_PERCENT).contains(&percent) {
            return Err(Error::OutOfRange);
        }
        Ok(Self { percent })
    }

    /// The Chart-inspector percentage.
    #[must_use]
    pub const fn percent(self) -> f32 {
        self.percent
    }

    /// The stored depth ratio: `z` divided by the scale vector's
    /// denominator component ([`Scale::is_horizontal`]).
    #[must_use]
    pub const fn native_ratio(self, scale: Scale) -> f32 {
        let (minimum, maximum) = scale.ratio_bounds();
        minimum + (maximum - minimum) * (self.percent / MAXIMUM_DEPTH_PERCENT)
    }

    /// Read a depth from a stored scale vector `(x, y, z, w)`.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidScale`] when `x`, `y` or `w` is not finite or the
    /// denominator is not positive; [`Error::NonFinite`] when the ratio is
    /// not finite.
    pub fn from_native(scale: Scale, vector: [f32; 4]) -> Result<Self> {
        let [x, y, z, w] = vector;
        if !x.is_finite() || !y.is_finite() || !w.is_finite() {
            return Err(Error::InvalidScale);
        }
        let denominator = scale.denominator(x, y);
        if !denominator.is_finite() || denominator <= 0.0 {
            return Err(Error::InvalidScale);
        }
        let ratio = z / denominator;
        if !ratio.is_finite() {
            return Err(Error::NonFinite);
        }
        let (minimum, maximum) = scale.ratio_bounds();
        let normalized = ((ratio - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
        Self::from_percent(normalized * MAXIMUM_DEPTH_PERCENT)
    }

    /// The `z` component that stores this depth in an existing scale vector
    /// with components `x` and `y`.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidScale`] when the denominator is not finite and
    /// positive.
    pub fn native_z(self, scale: Scale, x: f32, y: f32) -> Result<f32> {
        let denominator = scale.denominator(x, y);
        if !denominator.is_finite() || denominator <= 0.0 {
            return Err(Error::InvalidScale);
        }
        Ok(self.native_ratio(scale) * denominator)
    }

    /// The whole scale vector `(x, y, z, w)` to store when none exists: the
    /// native default scene dimensions with this depth.
    #[must_use]
    pub const fn native_vector(self, scale: Scale) -> [f32; 4] {
        let z = self.native_ratio(scale) * scale.denominator(DEFAULT_SCALE_X, DEFAULT_SCALE_Y);
        [DEFAULT_SCALE_X, DEFAULT_SCALE_Y, z, DEFAULT_SCALE_W]
    }
}

impl Default for Depth {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// X/Y orientation of a 3D chart scene, in inspector degrees
/// (`-40°..=40°` each).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rotation {
    x_degrees: f32,
    y_degrees: f32,
}

impl Rotation {
    /// The orientation used when no rotation vector is stored.
    pub const DEFAULT: Self = Self {
        x_degrees: 18.5,
        y_degrees: -18.125,
    };

    /// Construct a scene orientation.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`] when either angle is
    /// outside `-40°..=40°`.
    pub fn from_degrees(x_degrees: f32, y_degrees: f32) -> Result<Self> {
        for degrees in [x_degrees, y_degrees] {
            if !degrees.is_finite() {
                return Err(Error::NonFinite);
            }
            if !(MINIMUM_ROTATION_DEGREES..=MAXIMUM_ROTATION_DEGREES).contains(&degrees) {
                return Err(Error::OutOfRange);
            }
        }
        Ok(Self {
            x_degrees,
            y_degrees,
        })
    }

    /// Read an orientation from a stored rotation vector `(x, y, z, w)`.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidScale`] when `z` or `w` is not finite, else the
    /// [`Self::from_degrees`] errors.
    pub fn from_native(vector: [f32; 4]) -> Result<Self> {
        let [x, y, z, w] = vector;
        if !z.is_finite() || !w.is_finite() {
            return Err(Error::InvalidScale);
        }
        Self::from_degrees(x, y)
    }

    /// The rotation about the scene's X axis, in degrees.
    #[must_use]
    pub const fn x_degrees(self) -> f32 {
        self.x_degrees
    }

    /// The rotation about the scene's Y axis, in degrees.
    #[must_use]
    pub const fn y_degrees(self) -> f32 {
        self.y_degrees
    }

    /// The whole rotation vector `(x, y, 0, 0)` to store when none exists.
    #[must_use]
    pub const fn native_vector(self) -> [f32; 4] {
        [self.x_degrees, self.y_degrees, 0.0, 0.0]
    }
}

impl Default for Rotation {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Spacing between series along the depth axis of a 3D line or area chart,
/// in whole percent (`0%..=200%`); stored as a ratio.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct SeriesGap(u16);

impl SeriesGap {
    /// No spacing between adjacent series.
    pub const ZERO: Self = Self(0);
    /// The value used when storage is absent.
    pub const NATIVE_DEFAULT: Self = Self(DEFAULT_SERIES_GAP_PERCENT);
    /// The largest spacing the inspector accepts.
    pub const MAXIMUM: Self = Self(MAXIMUM_SERIES_GAP_PERCENT);

    /// Construct a series gap.
    ///
    /// # Errors
    ///
    /// [`Error::OutOfRange`] above 200%.
    pub const fn new(percent: u16) -> Result<Self> {
        if percent > MAXIMUM_SERIES_GAP_PERCENT {
            return Err(Error::OutOfRange);
        }
        Ok(Self(percent))
    }

    /// Read a series gap from its stored ratio, which must be a whole
    /// percentage within 0.001.
    ///
    /// # Errors
    ///
    /// [`Error::NonFinite`] or [`Error::OutOfRange`].
    pub fn from_native_ratio(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::NonFinite);
        }
        let percent = ratio * PERCENT_SCALE;
        let rounded = percent.round();
        if !(0.0..=f32::from(MAXIMUM_SERIES_GAP_PERCENT)).contains(&rounded)
            || (percent - rounded).abs() > WHOLE_PERCENT_TOLERANCE
        {
            return Err(Error::OutOfRange);
        }
        #[allow(
            clippy::cast_possible_truncation,
            clippy::cast_sign_loss,
            reason = "rounded is a whole number in 0.0..=200.0"
        )]
        Ok(Self(rounded as u16))
    }

    /// The whole percentage displayed by iWork.
    #[must_use]
    pub const fn percent(self) -> u16 {
        self.0
    }

    /// The stored ratio.
    #[must_use]
    pub fn native_ratio(self) -> f32 {
        f32::from(self.0) / PERCENT_SCALE
    }
}

impl TryFrom<u16> for SeriesGap {
    type Error = Error;

    fn try_from(percent: u16) -> Result<Self> {
        Self::new(percent)
    }
}

impl Default for SeriesGap {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::BarShape;

    #[test]
    fn bar_shape_is_a_compact_lossless_value() {
        assert_eq!(size_of::<BarShape>(), size_of::<i32>());
        assert_eq!(align_of::<BarShape>(), align_of::<i32>());
        assert_eq!(BarShape::default(), BarShape::Rectangle);
        assert!(!BarShape::Rectangle.is_unsupported());
        assert!(!BarShape::Cylinder.is_unsupported());
    }

    #[test]
    fn known_values_have_stable_native_identifiers() {
        assert_eq!(BarShape::Rectangle.native_value(), 0);
        assert_eq!(BarShape::Cylinder.native_value(), 1);
        assert_eq!(BarShape::from_native(0), BarShape::Rectangle);
        assert_eq!(BarShape::from_native(1), BarShape::Cylinder);
    }

    #[test]
    fn unknown_values_round_trip_without_fallback() {
        for value in [i32::MIN, -1, 2, 9_001, i32::MAX] {
            let shape = BarShape::from_native(value);
            assert!(shape.is_unsupported());
            assert_eq!(shape.native_value(), value);
            assert_eq!(BarShape::from_native(shape.native_value()), shape);
        }
    }
}

#[cfg(test)]
mod scene_tests {
    use super::{Depth, Error, Kind, Rotation, Scale, SeriesGap};

    #[test]
    fn scales_follow_the_3d_kinds() {
        assert_eq!(Scale::for_kind(Kind::StackedBar3d), Some(Scale::Bar));
        assert_eq!(Scale::for_kind(Kind::Donut3d), Some(Scale::Pie));
        assert_eq!(Scale::for_kind(Kind::Line3d), Some(Scale::Line));
        assert_eq!(Scale::for_kind(Kind::Column2d), None);
    }

    #[test]
    fn depth_round_trips_through_its_native_vector() {
        for scale in [
            Scale::Area,
            Scale::Bar,
            Scale::Column,
            Scale::Line,
            Scale::Pie,
        ] {
            for percent in [0.0, 11.020_468, 37.5, 100.0] {
                let depth = Depth::from_percent(percent).unwrap();
                let vector = depth.native_vector(scale);
                assert_eq!(Depth::from_native(scale, vector), Ok(depth));
                let [x, y, _, w] = vector;
                let z = depth.native_z(scale, x * 2.0, y * 2.0).unwrap();
                assert_eq!(
                    Depth::from_native(scale, [x * 2.0, y * 2.0, z, w]),
                    Ok(depth)
                );
            }
        }
        assert_eq!(Depth::default(), Depth::DEFAULT);
        assert_eq!(Depth::from_percent(100.5), Err(Error::OutOfRange));
        assert_eq!(
            Depth::from_native(Scale::Bar, [1.0, 0.0, 1.0, 1.0]),
            Err(Error::InvalidScale)
        );
        assert_eq!(
            Depth::from_native(Scale::Pie, [1.0, 1.0, 1.0, f32::NAN]),
            Err(Error::InvalidScale)
        );
        assert_eq!(
            Depth::DEFAULT.native_z(Scale::Pie, -1.0, 1.0),
            Err(Error::InvalidScale)
        );
    }

    #[test]
    fn rotation_and_series_gap_are_strict() {
        let rotation = Rotation::from_degrees(12.0, -40.0).unwrap();
        assert_eq!(rotation.native_vector(), [12.0, -40.0, 0.0, 0.0]);
        assert_eq!(
            Rotation::from_native(rotation.native_vector()),
            Ok(rotation)
        );
        assert_eq!(Rotation::from_degrees(40.5, 0.0), Err(Error::OutOfRange));
        assert_eq!(
            Rotation::from_native([0.0, 0.0, f32::NAN, 0.0]),
            Err(Error::InvalidScale)
        );
        assert_eq!(Rotation::default(), Rotation::DEFAULT);

        assert_eq!(SeriesGap::new(150).map(SeriesGap::native_ratio), Ok(1.5));
        assert_eq!(SeriesGap::from_native_ratio(1.5), SeriesGap::new(150));
        assert_eq!(SeriesGap::from_native_ratio(1.505), Err(Error::OutOfRange));
        assert_eq!(SeriesGap::new(201), Err(Error::OutOfRange));
        assert_eq!(SeriesGap::default(), SeriesGap::NATIVE_DEFAULT);
    }
}
