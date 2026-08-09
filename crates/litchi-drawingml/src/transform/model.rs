//! Semantic values for the `DrawingML` `a:xfrm` subtree.

use std::{fmt, str::FromStr};

use crate::{
    Error, Result,
    coordinate::{Coordinate, Extent},
};

/// A signed `ST_Angle` value in 60,000ths of a degree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
#[must_use]
pub struct Angle(i32);

impl Angle {
    /// The schema default for an omitted `rot` attribute.
    pub const ZERO: Self = Self(0);

    /// Construct a checked `DrawingML` angle.
    ///
    /// Every `i32` is valid for `ST_Angle`; keeping the constructor explicit
    /// makes the unit (60,000ths of a degree) visible at call sites.
    #[inline]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    /// Parse the XML Schema integer lexical form for `ST_Angle`.
    pub fn parse(value: &str) -> Result<Self> {
        value
            .trim()
            .parse::<i32>()
            .map(Self)
            .map_err(|_| Error::Invalid(format!("invalid DrawingML transform angle '{value}'")))
    }

    /// Return the raw 60,000ths-of-a-degree value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> i32 {
        self.0
    }
}

impl Default for Angle {
    fn default() -> Self {
        Self::ZERO
    }
}

impl From<i32> for Angle {
    fn from(value: i32) -> Self {
        Self::new(value)
    }
}

impl FromStr for Angle {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::parse(value)
    }
}

impl fmt::Display for Angle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

/// A checked `CT_Point2D` offset in `DrawingML` coordinates.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Point {
    x: Coordinate,
    y: Coordinate,
}

impl Point {
    /// Construct a point from already checked coordinates.
    #[inline]
    pub const fn new(x: Coordinate, y: Coordinate) -> Self {
        Self { x, y }
    }

    /// Construct an EMU point with both `ST_Coordinate` bounds checked.
    pub fn emu(x: i64, y: i64) -> Result<Self> {
        Ok(Self::new(
            Coordinate::emu(x).map_err(|error| coordinate_error("x", error))?,
            Coordinate::emu(y).map_err(|error| coordinate_error("y", error))?,
        ))
    }

    /// Borrow the horizontal coordinate.
    #[inline]
    pub const fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Borrow the vertical coordinate.
    #[inline]
    pub const fn y(&self) -> &Coordinate {
        &self.y
    }

    /// Decompose the point without cloning its exact coordinate values.
    #[inline]
    pub fn into_parts(self) -> (Coordinate, Coordinate) {
        (self.x, self.y)
    }
}

/// A checked `CT_PositiveSize2D` extent in EMUs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[must_use]
pub struct Size {
    width: Extent,
    height: Extent,
}

impl Size {
    /// Construct a size from already checked positive coordinates.
    #[inline]
    pub const fn new(width: Extent, height: Extent) -> Self {
        Self { width, height }
    }

    /// Construct an EMU size with both `ST_PositiveCoordinate` bounds checked.
    pub fn emu(width: i64, height: i64) -> Result<Self> {
        Ok(Self::new(
            Extent::emu(width).map_err(|error| coordinate_error("cx", error))?,
            Extent::emu(height).map_err(|error| coordinate_error("cy", error))?,
        ))
    }

    /// Return the checked width.
    #[inline]
    pub const fn width(self) -> Extent {
        self.width
    }

    /// Return the checked height.
    #[inline]
    pub const fn height(self) -> Extent {
        self.height
    }
}

/// The shared `DrawingML` two-dimensional transform (`a:CT_Transform2D`).
///
/// `None` on an optional scalar preserves authored absence, even when the
/// schema supplies an effective default. The accessors [`Self::rotation`],
/// [`Self::flip_horizontal`], and [`Self::flip_vertical`] expose those
/// effective defaults for ordinary callers; the corresponding `authored_*`
/// methods retain the wire-level distinction for source-aware editors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Transform {
    pub(super) offset: Option<Point>,
    pub(super) extent: Option<Size>,
    pub(super) child_offset: Option<Point>,
    pub(super) child_extent: Option<Size>,
    pub(super) rotation: Option<Angle>,
    pub(super) flip_horizontal: Option<bool>,
    pub(super) flip_vertical: Option<bool>,
}

impl Transform {
    /// Construct an empty transform using all `DrawingML` defaults.
    #[inline]
    pub const fn new() -> Self {
        Self {
            offset: None,
            extent: None,
            child_offset: None,
            child_extent: None,
            rotation: None,
            flip_horizontal: None,
            flip_vertical: None,
        }
    }

    /// Borrow the object offset (`a:off`).
    #[inline]
    #[must_use]
    pub const fn offset(&self) -> Option<&Point> {
        self.offset.as_ref()
    }

    /// Borrow the object extent (`a:ext`).
    #[inline]
    #[must_use]
    pub const fn extent(&self) -> Option<Size> {
        self.extent
    }

    /// Borrow the group child-coordinate offset (`a:chOff`).
    #[inline]
    #[must_use]
    pub const fn child_offset(&self) -> Option<&Point> {
        self.child_offset.as_ref()
    }

    /// Return the group child-coordinate extent (`a:chExt`).
    #[inline]
    #[must_use]
    pub const fn child_extent(&self) -> Option<Size> {
        self.child_extent
    }

    /// Return the effective rotation, whose omitted value is zero.
    #[inline]
    pub const fn rotation(&self) -> Angle {
        match self.rotation {
            Some(value) => value,
            None => Angle::ZERO,
        }
    }

    /// Return the authored `rot` value, preserving omission versus zero.
    #[inline]
    #[must_use]
    pub const fn authored_rotation(&self) -> Option<Angle> {
        self.rotation
    }

    /// Return the effective horizontal flip, whose omitted value is false.
    #[inline]
    #[must_use]
    pub const fn flip_horizontal(&self) -> bool {
        match self.flip_horizontal {
            Some(value) => value,
            None => false,
        }
    }

    /// Return the authored `flipH` value, preserving omission versus false.
    #[inline]
    #[must_use]
    pub const fn authored_flip_horizontal(&self) -> Option<bool> {
        self.flip_horizontal
    }

    /// Return the effective vertical flip, whose omitted value is false.
    #[inline]
    #[must_use]
    pub const fn flip_vertical(&self) -> bool {
        match self.flip_vertical {
            Some(value) => value,
            None => false,
        }
    }

    /// Return the authored `flipV` value, preserving omission versus false.
    #[inline]
    #[must_use]
    pub const fn authored_flip_vertical(&self) -> Option<bool> {
        self.flip_vertical
    }

    /// Set or clear the object offset.
    #[inline]
    pub fn set_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.offset = value;
        self
    }

    /// Set the object offset.
    #[inline]
    pub fn with_offset(mut self, value: Point) -> Self {
        self.set_offset(Some(value));
        self
    }

    /// Set or clear the object extent.
    #[inline]
    pub fn set_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.extent = value;
        self
    }

    /// Set the object extent.
    #[inline]
    pub fn with_extent(mut self, value: Size) -> Self {
        self.set_extent(Some(value));
        self
    }

    /// Set or clear the group child-coordinate offset.
    #[inline]
    pub fn set_child_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.child_offset = value;
        self
    }

    /// Set the group child-coordinate offset.
    #[inline]
    pub fn with_child_offset(mut self, value: Point) -> Self {
        self.set_child_offset(Some(value));
        self
    }

    /// Set or clear the group child-coordinate extent.
    #[inline]
    pub fn set_child_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.child_extent = value;
        self
    }

    /// Set the group child-coordinate extent.
    #[inline]
    pub fn with_child_extent(mut self, value: Size) -> Self {
        self.set_child_extent(Some(value));
        self
    }

    /// Set or clear the authored rotation.
    #[inline]
    pub fn set_rotation(&mut self, value: Option<Angle>) -> &mut Self {
        self.rotation = value;
        self
    }

    /// Set an explicitly authored rotation.
    #[inline]
    pub fn with_rotation(mut self, value: Angle) -> Self {
        self.set_rotation(Some(value));
        self
    }

    /// Set or clear the authored horizontal flip.
    #[inline]
    pub fn set_flip_horizontal(&mut self, value: Option<bool>) -> &mut Self {
        self.flip_horizontal = value;
        self
    }

    /// Set an explicitly authored horizontal flip.
    #[inline]
    pub fn with_flip_horizontal(mut self, value: bool) -> Self {
        self.set_flip_horizontal(Some(value));
        self
    }

    /// Set or clear the authored vertical flip.
    #[inline]
    pub fn set_flip_vertical(&mut self, value: Option<bool>) -> &mut Self {
        self.flip_vertical = value;
        self
    }

    /// Set an explicitly authored vertical flip.
    #[inline]
    pub fn with_flip_vertical(mut self, value: bool) -> Self {
        self.set_flip_vertical(Some(value));
        self
    }
}

fn coordinate_error(field: &'static str, error: impl fmt::Display) -> Error {
    Error::Invalid(format!("invalid DrawingML transform {field}: {error}"))
}
