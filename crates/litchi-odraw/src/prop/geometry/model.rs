use super::super::model::Array;

/// One coordinate in `OfficeArt` geometry space.
///
/// Values in the `0x80000000..=0x8000007F` range are guide references rather
/// than literal coordinates.  Keeping that distinction typed prevents a
/// caller from accidentally rendering a guide index as a huge negative point.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Coordinate {
    /// A literal signed geometry-space coordinate.
    Value(i32),
    /// A zero-based index into the shape's `pGuides_complex` array.
    Guide(u8),
}

impl Coordinate {
    /// Decodes the coordinate marker defined by `[MS-ODRAW]` section 2.3.6.7.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the range guard bounds the difference to 0..=0x7F; const fn cannot use u8::try_from"
    )]
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        let bits = raw.cast_unsigned();
        if bits >= 0x8000_0000 && bits <= 0x8000_007F {
            Self::Guide((bits - 0x8000_0000) as u8)
        } else {
            Self::Value(raw)
        }
    }

    /// Returns the exact signed wire value represented by this coordinate.
    #[must_use]
    pub const fn raw(self) -> i32 {
        match self {
            Self::Value(value) => value,
            Self::Guide(index) => (0x8000_0000u32 | index as u32).cast_signed(),
        }
    }

    /// Returns the guide index when this coordinate is guide-driven.
    #[must_use]
    pub const fn guide(self) -> Option<u8> {
        match self {
            Self::Guide(index) => Some(index),
            Self::Value(_) => None,
        }
    }
}

/// A typed `POINT` from an `OfficeArt` geometry array.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Point {
    x: Coordinate,
    y: Coordinate,
}

impl Point {
    /// Creates a point from typed geometry-space coordinates.
    #[must_use]
    pub const fn new(x: Coordinate, y: Coordinate) -> Self {
        Self { x, y }
    }

    /// Returns the x-coordinate.
    #[must_use]
    pub const fn x(self) -> Coordinate {
        self.x
    }

    /// Returns the y-coordinate.
    #[must_use]
    pub const fn y(self) -> Coordinate {
        self.y
    }

    /// Returns the exact pair of signed wire coordinates.
    #[must_use]
    pub const fn raw(self) -> (i32, i32) {
        (self.x.raw(), self.y.raw())
    }

    pub(crate) fn from_bytes(data: &[u8]) -> Option<Self> {
        let bytes: [u8; 8] = data.try_into().ok()?;
        Some(Self::new(
            Coordinate::from_raw(i32::from_le_bytes(bytes[..4].try_into().ok()?)),
            Coordinate::from_raw(i32::from_le_bytes(bytes[4..].try_into().ok()?)),
        ))
    }
}

/// The `shapePath` (`MSOSHAPEPATH`) semantic value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PathKind {
    /// Open straight-line path.
    Lines,
    /// Closed straight-line path.
    LinesClosed,
    /// Open cubic-Bezier path.
    Curves,
    /// Closed cubic-Bezier path.
    CurvesClosed,
    /// Path whose instructions come from `pSegmentInfo_complex`.
    Complex,
    /// A producer-defined future value, retaining its raw enumeration value.
    Unknown(u32),
}

impl PathKind {
    /// Decodes an `MSOSHAPEPATH` value without discarding extensions.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Self {
        match raw {
            0 => Self::Lines,
            1 => Self::LinesClosed,
            2 => Self::Curves,
            3 => Self::CurvesClosed,
            4 => Self::Complex,
            value => Self::Unknown(value),
        }
    }

    /// Returns the exact wire enumeration value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Lines => 0,
            Self::LinesClosed => 1,
            Self::Curves => 2,
            Self::CurvesClosed => 3,
            Self::Complex => 4,
            Self::Unknown(value) => value,
        }
    }
}

/// The standard `MSOPATHTYPE` instruction encoded by one path-info element.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Instruction {
    /// Consume one point per segment and draw lines.
    LineTo,
    /// Consume three points per segment and draw cubic Bezier curves.
    CurveTo,
    /// Start a sub-path at one point.
    MoveTo,
    /// Close the current sub-path.
    Close,
    /// End the current path.
    End,
    /// An escape instruction with a standard or future escape code.
    Escape(EscapeKind),
    /// A client-specific escape instruction.
    ClientEscape(EscapeKind),
    /// A future three-bit path instruction.
    Unknown(u8),
}

impl Instruction {
    /// Returns the three-bit path-type value.
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::LineTo => 0,
            Self::CurveTo => 1,
            Self::MoveTo => 2,
            Self::Close => 3,
            Self::End => 4,
            Self::Escape(_) => 5,
            Self::ClientEscape(_) => 6,
            Self::Unknown(value) => value,
        }
    }

    /// Returns the escape code when this instruction carries one.
    #[must_use]
    pub const fn escape(self) -> Option<EscapeKind> {
        match self {
            Self::Escape(value) | Self::ClientEscape(value) => Some(value),
            Self::LineTo
            | Self::CurveTo
            | Self::MoveTo
            | Self::Close
            | Self::End
            | Self::Unknown(_) => None,
        }
    }
}

/// The `MSOPATHESCAPE` value carried by an escape instruction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum EscapeKind {
    /// Extension point that delegates interpretation to a following escape.
    Extension,
    /// Angle-defined ellipse, continuing the current path.
    AngleEllipseTo,
    /// Angle-defined ellipse starting a new path.
    AngleEllipse,
    /// Counter-clockwise arc, continuing the current path.
    ArcTo,
    /// Counter-clockwise arc starting a new path.
    Arc,
    /// Clockwise arc, continuing the current path.
    ClockwiseArcTo,
    /// Clockwise arc starting a new path.
    ClockwiseArc,
    /// Elliptical quadrant tangent to the x-axis.
    EllipticalQuadrantX,
    /// Elliptical quadrant tangent to the y-axis.
    EllipticalQuadrantY,
    /// Quadratic Bezier control points.
    QuadraticBezier,
    /// Suppress fill for the current path.
    NoFill,
    /// Suppress line drawing for the current path.
    NoLine,
    /// Automatic line editing hint.
    AutoLine,
    /// Automatic curve editing hint.
    AutoCurve,
    /// Corner-line editing hint.
    CornerLine,
    /// Corner-curve editing hint.
    CornerCurve,
    /// Smooth-line editing hint.
    SmoothLine,
    /// Smooth-curve editing hint.
    SmoothCurve,
    /// Symmetric-line editing hint.
    SymmetricLine,
    /// Symmetric-curve editing hint.
    SymmetricCurve,
    /// Freeform editing hint.
    Freeform,
    /// Change the fill colors using one point.
    FillColor,
    /// Change the line colors using one point.
    LineColor,
    /// A future escape value, retaining its five-bit code.
    Unknown(u8),
}

impl EscapeKind {
    /// Decodes an `MSOPATHESCAPE` value without discarding extensions.
    #[must_use]
    pub const fn from_raw(raw: u8) -> Self {
        match raw {
            0 => Self::Extension,
            1 => Self::AngleEllipseTo,
            2 => Self::AngleEllipse,
            3 => Self::ArcTo,
            4 => Self::Arc,
            5 => Self::ClockwiseArcTo,
            6 => Self::ClockwiseArc,
            7 => Self::EllipticalQuadrantX,
            8 => Self::EllipticalQuadrantY,
            9 => Self::QuadraticBezier,
            10 => Self::NoFill,
            11 => Self::NoLine,
            12 => Self::AutoLine,
            13 => Self::AutoCurve,
            14 => Self::CornerLine,
            15 => Self::CornerCurve,
            16 => Self::SmoothLine,
            17 => Self::SmoothCurve,
            18 => Self::SymmetricLine,
            19 => Self::SymmetricCurve,
            20 => Self::Freeform,
            21 => Self::FillColor,
            22 => Self::LineColor,
            value => Self::Unknown(value),
        }
    }

    /// Returns the exact five-bit escape code.
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::Extension => 0,
            Self::AngleEllipseTo => 1,
            Self::AngleEllipse => 2,
            Self::ArcTo => 3,
            Self::Arc => 4,
            Self::ClockwiseArcTo => 5,
            Self::ClockwiseArc => 6,
            Self::EllipticalQuadrantX => 7,
            Self::EllipticalQuadrantY => 8,
            Self::QuadraticBezier => 9,
            Self::NoFill => 10,
            Self::NoLine => 11,
            Self::AutoLine => 12,
            Self::AutoCurve => 13,
            Self::CornerLine => 14,
            Self::CornerCurve => 15,
            Self::SmoothLine => 16,
            Self::SmoothCurve => 17,
            Self::SymmetricLine => 18,
            Self::SymmetricCurve => 19,
            Self::Freeform => 20,
            Self::FillColor => 21,
            Self::LineColor => 22,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn point_count(self, segments: u16) -> Option<usize> {
        match self {
            Self::AngleEllipseTo
            | Self::AngleEllipse
            | Self::ArcTo
            | Self::Arc
            | Self::ClockwiseArcTo
            | Self::ClockwiseArc
            | Self::EllipticalQuadrantX
            | Self::EllipticalQuadrantY
            | Self::QuadraticBezier => Some(segments as usize),
            Self::FillColor | Self::LineColor => Some(1),
            Self::Extension
            | Self::NoFill
            | Self::NoLine
            | Self::AutoLine
            | Self::AutoCurve
            | Self::CornerLine
            | Self::CornerCurve
            | Self::SmoothLine
            | Self::SmoothCurve
            | Self::SymmetricLine
            | Self::SymmetricCurve
            | Self::Freeform
            | Self::Unknown(_) => None,
        }
    }
}

/// One `MSOPATHINFO` element from `pSegmentInfo_complex`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PathInfo {
    instruction: Instruction,
    segments: u16,
    raw: u16,
}

impl PathInfo {
    /// Decodes the exact 16-bit `MSOPATHINFO` value.
    #[must_use]
    pub const fn from_raw(raw: u16) -> Self {
        let kind = (raw & 0x0007) as u8;
        let instruction = match kind {
            0 => Instruction::LineTo,
            1 => Instruction::CurveTo,
            2 => Instruction::MoveTo,
            3 => Instruction::Close,
            4 => Instruction::End,
            5 => Instruction::Escape(EscapeKind::from_raw(((raw >> 3) & 0x001F) as u8)),
            6 => Instruction::ClientEscape(EscapeKind::from_raw(((raw >> 3) & 0x001F) as u8)),
            value => Instruction::Unknown(value),
        };
        let segments = if kind == 5 || kind == 6 {
            raw >> 8
        } else {
            raw >> 3
        };
        Self {
            instruction,
            segments,
            raw,
        }
    }

    /// Returns the typed path instruction.
    #[must_use]
    pub const fn instruction(self) -> Instruction {
        self.instruction
    }

    /// Returns the number of segments encoded by this element.
    #[must_use]
    pub const fn segments(self) -> u16 {
        self.segments
    }

    /// Returns the exact wire element.
    #[must_use]
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Returns the known number of points consumed by this instruction.
    ///
    /// `None` is intentional for extension and editing-hint instructions whose
    /// point consumption is client-defined by the specification.
    pub(crate) const fn point_count(self) -> Option<usize> {
        match self.instruction {
            Instruction::LineTo => Some(self.segments as usize),
            Instruction::CurveTo => (self.segments as usize).checked_mul(3),
            Instruction::MoveTo => Some(1),
            Instruction::Close | Instruction::End => Some(0),
            Instruction::Escape(kind) | Instruction::ClientEscape(kind) => {
                kind.point_count(self.segments)
            },
            Instruction::Unknown(_) => None,
        }
    }
}

/// A borrowed typed view over `pVertices_complex`.
#[derive(Debug, Clone)]
pub struct Points<'data> {
    array: Option<Array<'data>>,
    index: usize,
}

impl<'data> Points<'data> {
    pub(crate) const fn new(array: Option<Array<'data>>) -> Self {
        Self { array, index: 0 }
    }
}

impl Iterator for Points<'_> {
    type Item = Point;

    fn next(&mut self) -> Option<Self::Item> {
        let array = self.array?;
        let data = array.get_element(self.index)?;
        self.index += 1;
        Point::from_bytes(data)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.array.map_or(0, |array| {
            usize::from(array.element_count()).saturating_sub(self.index)
        });
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for Points<'_> {}

/// A borrowed typed view over `pSegmentInfo_complex`.
#[derive(Debug, Clone)]
pub struct PathInfos<'data> {
    array: Option<Array<'data>>,
    index: usize,
}

impl<'data> PathInfos<'data> {
    pub(crate) const fn new(array: Option<Array<'data>>) -> Self {
        Self { array, index: 0 }
    }
}

impl Iterator for PathInfos<'_> {
    type Item = PathInfo;

    fn next(&mut self) -> Option<Self::Item> {
        let array = self.array?;
        let data = array.get_element(self.index)?;
        self.index += 1;
        let bytes: [u8; 2] = data.try_into().ok()?;
        Some(PathInfo::from_raw(u16::from_le_bytes(bytes)))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.array.map_or(0, |array| {
            usize::from(array.element_count()).saturating_sub(self.index)
        });
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for PathInfos<'_> {}

/// A zero-copy semantic view over the custom path properties of one shape.
#[derive(Debug, Clone, Copy)]
pub struct Geometry<'data> {
    shape_path: PathKind,
    vertices: Option<Array<'data>>,
    segment_info: Option<Array<'data>>,
}

impl<'data> Geometry<'data> {
    pub(crate) const fn from_parts(
        shape_path: PathKind,
        vertices: Option<Array<'data>>,
        segment_info: Option<Array<'data>>,
    ) -> Self {
        Self {
            shape_path,
            vertices,
            segment_info,
        }
    }

    /// Returns the `shapePath` semantic value.
    #[must_use]
    pub const fn path_kind(self) -> PathKind {
        self.shape_path
    }

    /// Iterates over typed vertices without allocating or copying the array.
    #[must_use]
    pub const fn vertices(self) -> Points<'data> {
        Points::new(self.vertices)
    }

    /// Iterates over typed path instructions without allocating or copying.
    #[must_use]
    pub const fn segment_info(self) -> PathInfos<'data> {
        PathInfos::new(self.segment_info)
    }

    /// Returns the number of vertices in the geometry.
    #[must_use]
    pub fn vertex_count(self) -> usize {
        match self.vertices {
            Some(array) => usize::from(array.element_count()),
            None => 0,
        }
    }

    /// Returns the number of path-info instructions.
    #[must_use]
    pub fn segment_count(self) -> usize {
        match self.segment_info {
            Some(array) => usize::from(array.element_count()),
            None => 0,
        }
    }

    /// Returns the exact encoded vertex array, including its `IMsoArray` header.
    #[must_use]
    pub fn raw_vertices(self) -> Option<&'data [u8]> {
        match self.vertices {
            Some(array) => Some(array.raw_data()),
            None => None,
        }
    }

    /// Returns the exact encoded path-info array, including its `IMsoArray` header.
    #[must_use]
    pub fn raw_segment_info(self) -> Option<&'data [u8]> {
        match self.segment_info {
            Some(array) => Some(array.raw_data()),
            None => None,
        }
    }
}
