//! Typed DrawingML custom geometry (`a:custGeom`) for XLSX worksheet
//! drawings.
//!
//! ECMA-376 part 1 §20.1.9.8 (CT_CustomGeometry2D) describes a shape outline
//! as an adjust-value list (`a:avLst`), a guide list (`a:gdLst`), adjust
//! handles (`a:ahLst`), connection sites (`a:cxnLst`), an optional text
//! rectangle (`a:rect`), and a path list (`a:pathLst`) whose paths draw with
//! move/line/arc/quadratic-Bezier/cubic-Bezier/close commands
//! (§20.1.9.9–§20.1.9.20). [`XlsxCustomGeometry`] models that structure with
//! typed path commands ([`XlsxPathCommand`]), typed guide formulas
//! ([`XlsxGeometryFormula`]), and adjustable values ([`XlsxAdjustValue`])
//! that carry either a literal or a guide reference, mirroring the
//! ST_AdjCoordinate/ST_AdjAngle unions.
//!
//! The model is shared by the shape inventory reader
//! ([`crate::xlsx::shapes`], which fills `XlsxShape::custom_geometry`) and
//! the shape authoring pipeline ([`crate::xlsx::writer::shape`], through
//! `XlsxShapeSpec::custom_geometry`), so authored geometry round-trips
//! through the inventory with identical semantics. Everything here is inert:
//! no rendering and no formula evaluation.

mod formula;
pub(crate) mod parse;
mod validate;
pub(crate) mod write;

#[cfg(test)]
mod tests;

use std::fmt;
use std::str::FromStr;

use crate::error::{OoxmlError, Result};

pub use formula::XlsxGeometryFormula;
pub(crate) use validate::validate_custom_geometry;

/// Smallest ST_Coordinate literal (ECMA-376 §20.1.10.16).
pub(crate) const MIN_COORDINATE: i64 = -27_273_042_329_600;
/// Largest ST_Coordinate literal (ECMA-376 §20.1.10.16).
pub(crate) const MAX_COORDINATE: i64 = 27_273_042_316_900;
/// Largest ST_PositiveCoordinate literal (ECMA-376 §20.1.10.42).
pub(crate) const MAX_POSITIVE_COORDINATE: i64 = 27_273_042_316_900;
/// ST_Angle literals are 60000ths of a degree within `xsd:int` bounds.
pub(crate) const MIN_ANGLE: i64 = i32::MIN as i64;
/// Largest ST_Angle literal.
pub(crate) const MAX_ANGLE: i64 = i32::MAX as i64;

/// Maximum entries in the adjust-value list or the guide list.
pub(crate) const MAX_GEOMETRY_GUIDES: usize = 4096;
/// Maximum adjust handles per custom geometry.
pub(crate) const MAX_ADJUST_HANDLES: usize = 4096;
/// Maximum connection sites per custom geometry.
pub(crate) const MAX_CONNECTION_SITES: usize = 4096;
/// Maximum paths per custom geometry.
pub(crate) const MAX_GEOMETRY_PATHS: usize = 1024;
/// Maximum drawing commands per geometry path.
pub(crate) const MAX_PATH_COMMANDS: usize = 65_536;
/// Maximum byte length of a geometry guide name.
pub(crate) const MAX_GUIDE_NAME_BYTES: usize = 255;

/// An adjustable geometry value (ST_AdjCoordinate or ST_AdjAngle): either a
/// literal or a reference to a geometry guide by name.
///
/// Coordinate literals are EMUs when the owning path declares no coordinate
/// space, or path-space units otherwise; angle literals are 60000ths of a
/// degree.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum XlsxAdjustValue {
    /// A literal value.
    Value(i64),
    /// A reference to a geometry guide (ST_GeomGuideName).
    Guide(String),
}

impl XlsxAdjustValue {
    /// A guide reference by name.
    pub fn guide(name: impl Into<String>) -> Self {
        Self::Guide(name.into())
    }
}

impl Default for XlsxAdjustValue {
    fn default() -> Self {
        Self::Value(0)
    }
}

impl From<i64> for XlsxAdjustValue {
    fn from(value: i64) -> Self {
        Self::Value(value)
    }
}

impl fmt::Display for XlsxAdjustValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Value(value) => write!(formatter, "{value}"),
            Self::Guide(name) => formatter.write_str(name),
        }
    }
}

impl FromStr for XlsxAdjustValue {
    type Err = OoxmlError;

    /// Numeric tokens become literals; any other non-empty token is kept as
    /// a guide reference, matching the ST_AdjCoordinate/ST_AdjAngle unions.
    fn from_str(text: &str) -> Result<Self> {
        let token = text.trim();
        if token.is_empty() {
            return Err(invalid("custom geometry value is empty"));
        }
        if let Ok(value) = token.parse::<i64>() {
            return Ok(Self::Value(value));
        }
        Ok(Self::Guide(token.to_string()))
    }
}

/// A geometry point (`a:pt`, CT_AdjPoint2D) with adjustable coordinates.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxGeometryPoint {
    /// Horizontal coordinate (`@x`).
    pub x: XlsxAdjustValue,
    /// Vertical coordinate (`@y`).
    pub y: XlsxAdjustValue,
}

impl XlsxGeometryPoint {
    /// A point from two adjustable coordinates.
    pub fn new(x: impl Into<XlsxAdjustValue>, y: impl Into<XlsxAdjustValue>) -> Self {
        Self {
            x: x.into(),
            y: y.into(),
        }
    }
}

/// One geometry guide (`a:gd`, CT_GeomGuide): a named, formula-derived value.
///
/// Guides in the adjust-value list (`a:avLst`) hold the shape's adjustable
/// parameters; guides in the guide list (`a:gdLst`) derive further values
/// from them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsxGeometryGuide {
    /// Guide name (`@name`, ST_GeomGuideName).
    pub name: String,
    /// Guide formula (`@fmla`).
    pub formula: XlsxGeometryFormula,
}

impl XlsxGeometryGuide {
    /// A guide with the given name and formula.
    pub fn new(name: impl Into<String>, formula: XlsxGeometryFormula) -> Self {
        Self {
            name: name.into(),
            formula,
        }
    }
}

/// An XY adjust handle (`a:ahXY`, CT_XYAdjustHandle).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxXyAdjustHandle {
    /// Guide updated by horizontal movement (`@gdRefX`), when declared.
    pub horizontal_guide: Option<String>,
    /// Minimum horizontal position (`@minX`), when declared.
    pub minimum_x: Option<XlsxAdjustValue>,
    /// Maximum horizontal position (`@maxX`), when declared.
    pub maximum_x: Option<XlsxAdjustValue>,
    /// Guide updated by vertical movement (`@gdRefY`), when declared.
    pub vertical_guide: Option<String>,
    /// Minimum vertical position (`@minY`), when declared.
    pub minimum_y: Option<XlsxAdjustValue>,
    /// Maximum vertical position (`@maxY`), when declared.
    pub maximum_y: Option<XlsxAdjustValue>,
    /// Handle position (`a:pos`).
    pub position: XlsxGeometryPoint,
}

/// A polar adjust handle (`a:ahPolar`, CT_PolarAdjustHandle).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxPolarAdjustHandle {
    /// Guide updated by radial movement (`@gdRefR`), when declared.
    pub radius_guide: Option<String>,
    /// Minimum radius (`@minR`), when declared.
    pub minimum_radius: Option<XlsxAdjustValue>,
    /// Maximum radius (`@maxR`), when declared.
    pub maximum_radius: Option<XlsxAdjustValue>,
    /// Guide updated by angular movement (`@gdRefAng`), when declared.
    pub angle_guide: Option<String>,
    /// Minimum angle (`@minAng`), when declared.
    pub minimum_angle: Option<XlsxAdjustValue>,
    /// Maximum angle (`@maxAng`), when declared.
    pub maximum_angle: Option<XlsxAdjustValue>,
    /// Handle position (`a:pos`).
    pub position: XlsxGeometryPoint,
}

/// One adjust handle of the geometry (`a:ahLst` entry).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsxAdjustHandle {
    /// An XY handle (`a:ahXY`).
    Xy(XlsxXyAdjustHandle),
    /// A polar handle (`a:ahPolar`).
    Polar(XlsxPolarAdjustHandle),
}

/// One connection site (`a:cxn`, CT_ConnectionSite).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxConnectionSite {
    /// Site angle (`@ang`, ST_AdjAngle).
    pub angle: XlsxAdjustValue,
    /// Site position (`a:pos`).
    pub position: XlsxGeometryPoint,
}

/// The geometry text rectangle (`a:rect`, CT_GeomRect).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxGeometryRectangle {
    /// Left edge (`@l`).
    pub left: XlsxAdjustValue,
    /// Top edge (`@t`).
    pub top: XlsxAdjustValue,
    /// Right edge (`@r`).
    pub right: XlsxAdjustValue,
    /// Bottom edge (`@b`).
    pub bottom: XlsxAdjustValue,
}

/// How a geometry path is filled (`a:path@fill`, ST_PathFillMode).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum XlsxPathFillMode {
    /// The path is not filled (`none`).
    None,
    /// The path is filled normally (`norm`, the ECMA-376 default).
    #[default]
    Normal,
    /// The path is filled with a lightened version of the shape fill (`lighten`).
    Lighten,
    /// The path is filled with a slightly lightened shape fill (`lightenLess`).
    LightenLess,
    /// The path is filled with a darkened version of the shape fill (`darken`).
    Darken,
    /// The path is filled with a slightly darkened shape fill (`darkenLess`).
    DarkenLess,
}

impl XlsxPathFillMode {
    /// Parse an ST_PathFillMode token.
    pub fn from_token(token: &str) -> Option<Self> {
        match token {
            "none" => Some(Self::None),
            "norm" => Some(Self::Normal),
            "lighten" => Some(Self::Lighten),
            "lightenLess" => Some(Self::LightenLess),
            "darken" => Some(Self::Darken),
            "darkenLess" => Some(Self::DarkenLess),
            _ => None,
        }
    }

    /// The ST_PathFillMode token for this mode.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Normal => "norm",
            Self::Lighten => "lighten",
            Self::LightenLess => "lightenLess",
            Self::Darken => "darken",
            Self::DarkenLess => "darkenLess",
        }
    }
}

/// One drawing command inside a geometry path (`a:path` children,
/// ECMA-376 §20.1.9.10–§20.1.9.20).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsxPathCommand {
    /// `a:moveTo` — start a new sub-path at the given point.
    MoveTo(XlsxGeometryPoint),
    /// `a:lnTo` — draw a straight line to the given point.
    LineTo(XlsxGeometryPoint),
    /// `a:arcTo` — draw an elliptical arc from the current point.
    ArcTo {
        /// Ellipse width radius (`@wR`).
        width_radius: XlsxAdjustValue,
        /// Ellipse height radius (`@hR`).
        height_radius: XlsxAdjustValue,
        /// Start angle (`@stAng`), in 60000ths of a degree.
        start_angle: XlsxAdjustValue,
        /// Swing angle (`@swAng`), in 60000ths of a degree.
        swing_angle: XlsxAdjustValue,
    },
    /// `a:quadBezTo` — draw a quadratic Bezier curve.
    QuadraticBezierTo {
        /// Control point (first `a:pt`).
        control: XlsxGeometryPoint,
        /// End point (second `a:pt`).
        end: XlsxGeometryPoint,
    },
    /// `a:cubicBezTo` — draw a cubic Bezier curve.
    CubicBezierTo {
        /// First control point.
        control1: XlsxGeometryPoint,
        /// Second control point.
        control2: XlsxGeometryPoint,
        /// End point.
        end: XlsxGeometryPoint,
    },
    /// `a:close` — close the current sub-path.
    Close,
}

/// One geometry path (`a:path`, CT_Path2D) with its coordinate space and
/// fill/stroke/extrusion attributes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsxGeometryPath {
    /// Width of the path coordinate space (`@w`; 0, the ECMA-376 default,
    /// means path coordinates are EMUs within the shape extent).
    pub width: i64,
    /// Height of the path coordinate space (`@h`; 0 when absent).
    pub height: i64,
    /// Fill mode of the path (`@fill`).
    pub fill_mode: XlsxPathFillMode,
    /// Whether the path outline is stroked (`@stroke`, default true).
    pub stroked: bool,
    /// Whether 3D extrusion is allowed on the path (`@extrusionOk`,
    /// default true).
    pub extrusion_allowed: bool,
    /// Drawing commands in path order.
    pub commands: Vec<XlsxPathCommand>,
}

impl Default for XlsxGeometryPath {
    fn default() -> Self {
        Self {
            width: 0,
            height: 0,
            fill_mode: XlsxPathFillMode::default(),
            stroked: true,
            extrusion_allowed: true,
            commands: Vec::new(),
        }
    }
}

impl XlsxGeometryPath {
    /// An empty path over a `width` by `height` coordinate space with the
    /// ECMA-376 default fill/stroke/extrusion attributes.
    pub fn new(width: i64, height: i64) -> Self {
        Self {
            width,
            height,
            ..Self::default()
        }
    }

    /// Append a drawing command to the path.
    pub fn with_command(mut self, command: XlsxPathCommand) -> Self {
        self.commands.push(command);
        self
    }
}

/// A custom shape geometry (`a:custGeom`, CT_CustomGeometry2D).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxCustomGeometry {
    /// Adjustable shape parameters (`a:avLst`).
    pub adjust_values: Vec<XlsxGeometryGuide>,
    /// Derived geometry guides (`a:gdLst`).
    pub guides: Vec<XlsxGeometryGuide>,
    /// Adjust handles (`a:ahLst`).
    pub adjust_handles: Vec<XlsxAdjustHandle>,
    /// Connection sites (`a:cxnLst`).
    pub connection_sites: Vec<XlsxConnectionSite>,
    /// Text rectangle (`a:rect`), when declared.
    pub text_rectangle: Option<XlsxGeometryRectangle>,
    /// Geometry paths (`a:pathLst`) in drawing order.
    pub paths: Vec<XlsxGeometryPath>,
}

impl XlsxCustomGeometry {
    /// An empty custom geometry.
    pub fn new() -> Self {
        Self::default()
    }

    /// Append an adjustable parameter to the adjust-value list.
    pub fn with_adjust_value(mut self, guide: XlsxGeometryGuide) -> Self {
        self.adjust_values.push(guide);
        self
    }

    /// Append a derived guide to the guide list.
    pub fn with_guide(mut self, guide: XlsxGeometryGuide) -> Self {
        self.guides.push(guide);
        self
    }

    /// Append a path to the path list.
    pub fn with_path(mut self, path: XlsxGeometryPath) -> Self {
        self.paths.push(path);
        self
    }
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
