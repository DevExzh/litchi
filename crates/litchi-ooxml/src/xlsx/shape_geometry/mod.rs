//! Typed DrawingML custom geometry (`a:custGeom`) for XLSX worksheet
//! drawings.
//!
//! ECMA-376 part 1 §20.1.9.8 (CT_CustomGeometry2D) describes a shape outline
//! as an adjust-value list (`a:avLst`), a guide list (`a:gdLst`), adjust
//! handles (`a:ahLst`), connection sites (`a:cxnLst`), an optional text
//! rectangle (`a:rect`), and a path list (`a:pathLst`) whose paths draw with
//! move/line/arc/quadratic-Bezier/cubic-Bezier/close commands
//! (§20.1.9.9–§20.1.9.20). [`CustomGeometry`] models that structure with
//! typed path commands ([`PathCommand`]), typed guide formulas
//! ([`Formula`]), and adjustable values ([`AdjustValue`])
//! that carry either a literal or a guide reference, mirroring the
//! ST_AdjCoordinate/ST_AdjAngle unions.
//!
//! The model is shared by the shape inventory reader
//! ([`crate::xlsx::shapes`], which fills `Shape::custom_geometry`) and
//! the shape authoring pipeline ([`crate::xlsx::writer::shape`], through
//! [`crate::xlsx::writer::Geometry::Custom`]), so authored geometry round-trips
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

pub use formula::Formula;
pub(crate) use validate::{validate_custom_geometry, validate_parsed_custom_geometry};

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
pub enum AdjustValue {
    /// A literal value.
    Value(i64),
    /// A reference to a geometry guide (ST_GeomGuideName).
    Guide(String),
}

impl AdjustValue {
    /// A guide reference by name.
    pub fn guide(name: impl Into<String>) -> Self {
        Self::Guide(normalize_xsd_token(&name.into()))
    }
}

impl Default for AdjustValue {
    fn default() -> Self {
        Self::Value(0)
    }
}

impl From<i64> for AdjustValue {
    fn from(value: i64) -> Self {
        Self::Value(value)
    }
}

impl fmt::Display for AdjustValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Value(value) => write!(formatter, "{value}"),
            Self::Guide(name) => formatter.write_str(name),
        }
    }
}

impl FromStr for AdjustValue {
    type Err = OoxmlError;

    /// Numeric tokens become literals; any other non-empty token is kept as
    /// a guide reference, matching the ST_AdjCoordinate/ST_AdjAngle unions.
    fn from_str(text: &str) -> Result<Self> {
        let token = normalize_xsd_token(text);
        if token.is_empty() {
            return Err(invalid("custom geometry value is empty"));
        }
        if let Ok(value) = token.parse::<i64>() {
            return Ok(Self::Value(value));
        }
        Ok(Self::Guide(token))
    }
}

/// A geometry point (`a:pt`, CT_AdjPoint2D) with adjustable coordinates.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Point {
    /// Horizontal coordinate (`@x`).
    pub x: AdjustValue,
    /// Vertical coordinate (`@y`).
    pub y: AdjustValue,
}

impl Point {
    /// A point from two adjustable coordinates.
    pub fn new(x: impl Into<AdjustValue>, y: impl Into<AdjustValue>) -> Self {
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
pub struct Guide {
    /// Guide name (`@name`, ST_GeomGuideName).
    pub name: String,
    /// Guide formula (`@fmla`).
    pub formula: Formula,
}

impl Guide {
    /// A guide with the given name and formula.
    pub fn new(name: impl Into<String>, formula: Formula) -> Self {
        Self {
            name: normalize_xsd_token(&name.into()),
            formula,
        }
    }
}

pub(crate) fn normalize_xsd_token(value: &str) -> String {
    value
        .split([' ', '\t', '\r', '\n'])
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
}

/// An XY adjust handle (`a:ahXY`, CT_XYAdjustHandle).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XyAdjustHandle {
    /// Guide updated by horizontal movement (`@gdRefX`), when declared.
    pub horizontal_guide: Option<String>,
    /// Minimum horizontal position (`@minX`), when declared.
    pub minimum_x: Option<AdjustValue>,
    /// Maximum horizontal position (`@maxX`), when declared.
    pub maximum_x: Option<AdjustValue>,
    /// Guide updated by vertical movement (`@gdRefY`), when declared.
    pub vertical_guide: Option<String>,
    /// Minimum vertical position (`@minY`), when declared.
    pub minimum_y: Option<AdjustValue>,
    /// Maximum vertical position (`@maxY`), when declared.
    pub maximum_y: Option<AdjustValue>,
    /// Handle position (`a:pos`).
    pub position: Point,
}

/// A polar adjust handle (`a:ahPolar`, CT_PolarAdjustHandle).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PolarAdjustHandle {
    /// Guide updated by radial movement (`@gdRefR`), when declared.
    pub radius_guide: Option<String>,
    /// Minimum radius (`@minR`), when declared.
    pub minimum_radius: Option<AdjustValue>,
    /// Maximum radius (`@maxR`), when declared.
    pub maximum_radius: Option<AdjustValue>,
    /// Guide updated by angular movement (`@gdRefAng`), when declared.
    pub angle_guide: Option<String>,
    /// Minimum angle (`@minAng`), when declared.
    pub minimum_angle: Option<AdjustValue>,
    /// Maximum angle (`@maxAng`), when declared.
    pub maximum_angle: Option<AdjustValue>,
    /// Handle position (`a:pos`).
    pub position: Point,
}

/// One adjust handle of the geometry (`a:ahLst` entry).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AdjustHandle {
    /// An XY handle (`a:ahXY`).
    Xy(XyAdjustHandle),
    /// A polar handle (`a:ahPolar`).
    Polar(PolarAdjustHandle),
}

/// One connection site (`a:cxn`, CT_ConnectionSite).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ConnectionSite {
    /// Site angle (`@ang`, ST_AdjAngle).
    pub angle: AdjustValue,
    /// Site position (`a:pos`).
    pub position: Point,
}

/// The geometry text rectangle (`a:rect`, CT_GeomRect).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Rectangle {
    /// Left edge (`@l`).
    pub left: AdjustValue,
    /// Top edge (`@t`).
    pub top: AdjustValue,
    /// Right edge (`@r`).
    pub right: AdjustValue,
    /// Bottom edge (`@b`).
    pub bottom: AdjustValue,
}

/// How a geometry path is filled (`a:path@fill`, ST_PathFillMode).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum PathFillMode {
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

impl PathFillMode {
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

impl FromStr for PathFillMode {
    type Err = OoxmlError;

    fn from_str(token: &str) -> Result<Self> {
        match token {
            "none" => Ok(Self::None),
            "norm" => Ok(Self::Normal),
            "lighten" => Ok(Self::Lighten),
            "lightenLess" => Ok(Self::LightenLess),
            "darken" => Ok(Self::Darken),
            "darkenLess" => Ok(Self::DarkenLess),
            _ => Err(invalid(format!("invalid ST_PathFillMode token '{token}'"))),
        }
    }
}

impl fmt::Display for PathFillMode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// One drawing command inside a geometry path (`a:path` children,
/// ECMA-376 §20.1.9.10–§20.1.9.20).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PathCommand {
    /// `a:moveTo` — start a new sub-path at the given point.
    MoveTo(Point),
    /// `a:lnTo` — draw a straight line to the given point.
    LineTo(Point),
    /// `a:arcTo` — draw an elliptical arc from the current point.
    ArcTo {
        /// Ellipse width radius (`@wR`).
        width_radius: AdjustValue,
        /// Ellipse height radius (`@hR`).
        height_radius: AdjustValue,
        /// Start angle (`@stAng`), in 60000ths of a degree.
        start_angle: AdjustValue,
        /// Swing angle (`@swAng`), in 60000ths of a degree.
        swing_angle: AdjustValue,
    },
    /// `a:quadBezTo` — draw a quadratic Bezier curve.
    QuadraticBezierTo {
        /// Control point (first `a:pt`).
        control: Point,
        /// End point (second `a:pt`).
        end: Point,
    },
    /// `a:cubicBezTo` — draw a cubic Bezier curve.
    CubicBezierTo {
        /// First control point.
        control1: Point,
        /// Second control point.
        control2: Point,
        /// End point.
        end: Point,
    },
    /// `a:close` — close the current sub-path.
    Close,
}

/// One geometry path (`a:path`, CT_Path2D) with its coordinate space and
/// fill/stroke/extrusion attributes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Path {
    /// Width of the path coordinate space (`@w`; 0, the ECMA-376 default,
    /// means path coordinates are EMUs within the shape extent).
    pub width: i64,
    /// Height of the path coordinate space (`@h`; 0 when absent).
    pub height: i64,
    /// Fill mode of the path (`@fill`).
    pub fill_mode: PathFillMode,
    /// Whether the path outline is stroked (`@stroke`, default true).
    pub stroked: bool,
    /// Whether 3D extrusion is allowed on the path (`@extrusionOk`,
    /// default true).
    pub extrusion_allowed: bool,
    /// Drawing commands in path order.
    pub commands: Vec<PathCommand>,
}

impl Default for Path {
    fn default() -> Self {
        Self {
            width: 0,
            height: 0,
            fill_mode: PathFillMode::default(),
            stroked: true,
            extrusion_allowed: true,
            commands: Vec::new(),
        }
    }
}

impl Path {
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
    pub fn with_command(mut self, command: PathCommand) -> Self {
        self.commands.push(command);
        self
    }
}

/// A custom shape geometry (`a:custGeom`, CT_CustomGeometry2D).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CustomGeometry {
    /// Adjustable shape parameters (`a:avLst`).
    pub adjust_values: Vec<Guide>,
    /// Derived geometry guides (`a:gdLst`).
    pub guides: Vec<Guide>,
    /// Adjust handles (`a:ahLst`).
    pub adjust_handles: Vec<AdjustHandle>,
    /// Connection sites (`a:cxnLst`).
    pub connection_sites: Vec<ConnectionSite>,
    /// Text rectangle (`a:rect`), when declared.
    pub text_rectangle: Option<Rectangle>,
    /// Geometry paths (`a:pathLst`) in drawing order.
    pub paths: Vec<Path>,
}

impl CustomGeometry {
    /// An empty custom geometry.
    pub fn new() -> Self {
        Self::default()
    }

    /// Append an adjustable parameter to the adjust-value list.
    pub fn with_adjust_value(mut self, guide: Guide) -> Self {
        self.adjust_values.push(guide);
        self
    }

    /// Append a derived guide to the guide list.
    pub fn with_guide(mut self, guide: Guide) -> Self {
        self.guides.push(guide);
        self
    }

    /// Append a path to the path list.
    pub fn with_path(mut self, path: Path) -> Self {
        self.paths.push(path);
        self
    }
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
