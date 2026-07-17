//! Inert Word 6/95 RTF drawing primitives.

use crate::{LegacyTextBox, RtfError, RtfResult};

pub const MAX_LEGACY_DRAWINGS: usize = 16_384;
pub const MAX_LEGACY_DRAWING_DEPTH: usize = 64;
pub const MAX_LEGACY_DRAWING_PRIMITIVES: usize = 65_536;
pub const MAX_LEGACY_DRAWING_POINTS: usize = 65_536;
pub const MAX_LEGACY_DRAWING_TOTAL_POINTS: usize = 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingGeometry {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

impl LegacyDrawingGeometry {
    pub fn validate(self) -> RtfResult<()> {
        if self.width < 0 || self.height < 0 {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing size cannot be negative".to_string(),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingPoint {
    pub x: i32,
    pub y: i32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyDrawingColor {
    Gray(u8),
    Rgb {
        red: u8,
        green: u8,
        blue: u8,
        palette: bool,
    },
}

impl LegacyDrawingColor {
    pub fn gray_half_percent(value: i32) -> RtfResult<Self> {
        let value = u8::try_from(value).map_err(|_| {
            RtfError::MalformedDocument("RTF legacy drawing grayscale is outside 0..=200".to_string())
        })?;
        if value > 200 {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing grayscale is outside 0..=200".to_string(),
            ));
        }
        Ok(Self::Gray(value))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyDrawingLineStyle {
    Solid,
    Hollow,
    Dashed,
    Dotted,
    DashDot,
    DashDotDot,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingLine {
    pub style: LegacyDrawingLineStyle,
    pub color: LegacyDrawingColor,
    pub width: i32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum LegacyDrawingFillPattern {
    Clear = 0,
    Solid = 1,
    Percent5 = 2,
    Percent10 = 3,
    Percent20 = 4,
    Percent25 = 5,
    Percent30 = 6,
    Percent40 = 7,
    Percent50 = 8,
    Percent60 = 9,
    Percent70 = 10,
    Percent75 = 11,
    Percent80 = 12,
    Percent90 = 13,
    DarkHorizontal = 14,
    DarkVertical = 15,
    DarkLeftDiagonal = 16,
    DarkRightDiagonal = 17,
    DarkGrid = 18,
    DarkTrellis = 19,
    LightHorizontal = 20,
    LightVertical = 21,
    LightLeftDiagonal = 22,
    LightRightDiagonal = 23,
    LightGrid = 24,
    LightTrellis = 25,
}

impl TryFrom<i32> for LegacyDrawingFillPattern {
    type Error = RtfError;

    fn try_from(value: i32) -> Result<Self, Self::Error> {
        Ok(match value {
            0 => Self::Clear,
            1 => Self::Solid,
            2 => Self::Percent5,
            3 => Self::Percent10,
            4 => Self::Percent20,
            5 => Self::Percent25,
            6 => Self::Percent30,
            7 => Self::Percent40,
            8 => Self::Percent50,
            9 => Self::Percent60,
            10 => Self::Percent70,
            11 => Self::Percent75,
            12 => Self::Percent80,
            13 => Self::Percent90,
            14 => Self::DarkHorizontal,
            15 => Self::DarkVertical,
            16 => Self::DarkLeftDiagonal,
            17 => Self::DarkRightDiagonal,
            18 => Self::DarkGrid,
            19 => Self::DarkTrellis,
            20 => Self::LightHorizontal,
            21 => Self::LightVertical,
            22 => Self::LightLeftDiagonal,
            23 => Self::LightRightDiagonal,
            24 => Self::LightGrid,
            25 => Self::LightTrellis,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF legacy drawing fill pattern is outside 0..=25".to_string(),
                ));
            },
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingFill {
    pub foreground: LegacyDrawingColor,
    pub background: LegacyDrawingColor,
    pub pattern: LegacyDrawingFillPattern,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyDrawingArrowFill {
    Solid,
    Hollow,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum LegacyDrawingArrowSize {
    Small = 1,
    Medium = 2,
    Large = 3,
}

impl TryFrom<i32> for LegacyDrawingArrowSize {
    type Error = RtfError;

    fn try_from(value: i32) -> Result<Self, Self::Error> {
        match value {
            1 => Ok(Self::Small),
            2 => Ok(Self::Medium),
            3 => Ok(Self::Large),
            _ => Err(RtfError::MalformedDocument(
                "RTF legacy drawing arrow size is outside 1..=3".to_string(),
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingArrow {
    pub fill: LegacyDrawingArrowFill,
    pub length: LegacyDrawingArrowSize,
    pub width: LegacyDrawingArrowSize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LegacyDrawingShadow {
    pub x_offset: i32,
    pub y_offset: i32,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct LegacyDrawingProperties {
    pub line: Option<LegacyDrawingLine>,
    pub fill: Option<LegacyDrawingFill>,
    pub start_arrow: Option<LegacyDrawingArrow>,
    pub end_arrow: Option<LegacyDrawingArrow>,
    pub shadow: Option<LegacyDrawingShadow>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyCalloutType {
    RightAngle,
    Single,
    Double,
    Triple,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyCalloutAttachment {
    Top,
    Center,
    Bottom,
    Absolute,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyCallout<'a> {
    pub callout_type: LegacyCalloutType,
    pub angle: Option<u8>,
    pub accent: bool,
    pub smart_attach: bool,
    pub best_fit: bool,
    pub minus_x: bool,
    pub minus_y: bool,
    pub border: bool,
    pub attachment: Option<LegacyCalloutAttachment>,
    pub descent: Option<i32>,
    pub offset: i32,
    pub length: i32,
    pub polyline: Box<LegacyDrawingPrimitive<'a>>,
    pub text_box: Box<LegacyDrawingPrimitive<'a>>,
    pub geometry: LegacyDrawingGeometry,
    pub properties: LegacyDrawingProperties,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LegacyDrawingPrimitive<'a> {
    Group {
        geometry: LegacyDrawingGeometry,
        children: Vec<LegacyDrawingPrimitive<'a>>,
        end_geometry: LegacyDrawingGeometry,
    },
    Callout(LegacyCallout<'a>),
    Line {
        start: LegacyDrawingPoint,
        end: LegacyDrawingPoint,
        geometry: LegacyDrawingGeometry,
        properties: LegacyDrawingProperties,
    },
    Rectangle {
        rounded: bool,
        geometry: LegacyDrawingGeometry,
        properties: LegacyDrawingProperties,
    },
    TextBox {
        text_box: LegacyTextBox<'a>,
        properties: LegacyDrawingProperties,
    },
    Ellipse {
        geometry: LegacyDrawingGeometry,
        properties: LegacyDrawingProperties,
    },
    Polyline {
        closed: bool,
        points: Vec<LegacyDrawingPoint>,
        geometry: LegacyDrawingGeometry,
        properties: LegacyDrawingProperties,
    },
    Arc {
        flip_x: bool,
        flip_y: bool,
        geometry: LegacyDrawingGeometry,
        properties: LegacyDrawingProperties,
    },
}

impl LegacyDrawingPrimitive<'_> {
    fn validate_at(&self, depth: usize, primitives: &mut usize, points: &mut usize) -> RtfResult<()> {
        if depth > MAX_LEGACY_DRAWING_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing nesting exceeds the safety limit".to_string(),
            ));
        }
        *primitives = primitives.checked_add(1).ok_or_else(|| {
            RtfError::MalformedDocument("RTF legacy drawing primitive count overflow".to_string())
        })?;
        if *primitives > MAX_LEGACY_DRAWING_PRIMITIVES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing primitive count exceeds the safety limit".to_string(),
            ));
        }
        match self {
            Self::Group { geometry, children, end_geometry } => {
                geometry.validate()?;
                end_geometry.validate()?;
                if children.is_empty() {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing group cannot be empty".to_string(),
                    ));
                }
                for child in children {
                    child.validate_at(depth + 1, primitives, points)?;
                }
            },
            Self::Callout(callout) => {
                callout.geometry.validate()?;
                if !matches!(*callout.polyline, Self::Polyline { .. })
                    || !matches!(*callout.text_box, Self::TextBox { .. })
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy callout requires a polyline followed by a text box".to_string(),
                    ));
                }
                if callout.angle.is_some_and(|angle| !matches!(angle, 0 | 30 | 45 | 60 | 90)) {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy callout angle is invalid".to_string(),
                    ));
                }
                callout.polyline.validate_at(depth + 1, primitives, points)?;
                callout.text_box.validate_at(depth + 1, primitives, points)?;
            },
            Self::Line { geometry, .. }
            | Self::Rectangle { geometry, .. }
            | Self::Ellipse { geometry, .. }
            | Self::Arc { geometry, .. } => geometry.validate()?,
            Self::TextBox { text_box, .. } => text_box.validate()?,
            Self::Polyline { points: item_points, geometry, .. } => {
                geometry.validate()?;
                if item_points.is_empty() || item_points.len() > MAX_LEGACY_DRAWING_POINTS {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing polyline point count is invalid".to_string(),
                    ));
                }
                *points = points.checked_add(item_points.len()).ok_or_else(|| {
                    RtfError::MalformedDocument("RTF legacy drawing point count overflow".to_string())
                })?;
                if *points > MAX_LEGACY_DRAWING_TOTAL_POINTS {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing aggregate point count exceeds the safety limit".to_string(),
                    ));
                }
            },
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LegacyDrawingPrimitive<'static> {
        match self {
            Self::Group { geometry, children, end_geometry } => LegacyDrawingPrimitive::Group {
                geometry,
                children: children.into_iter().map(Self::into_owned).collect(),
                end_geometry,
            },
            Self::Callout(callout) => LegacyDrawingPrimitive::Callout(LegacyCallout {
                callout_type: callout.callout_type,
                angle: callout.angle,
                accent: callout.accent,
                smart_attach: callout.smart_attach,
                best_fit: callout.best_fit,
                minus_x: callout.minus_x,
                minus_y: callout.minus_y,
                border: callout.border,
                attachment: callout.attachment,
                descent: callout.descent,
                offset: callout.offset,
                length: callout.length,
                polyline: Box::new(callout.polyline.into_owned()),
                text_box: Box::new(callout.text_box.into_owned()),
                geometry: callout.geometry,
                properties: callout.properties,
            }),
            Self::Line { start, end, geometry, properties } => LegacyDrawingPrimitive::Line { start, end, geometry, properties },
            Self::Rectangle { rounded, geometry, properties } => LegacyDrawingPrimitive::Rectangle { rounded, geometry, properties },
            Self::TextBox { text_box, properties } => LegacyDrawingPrimitive::TextBox { text_box: text_box.into_owned(), properties },
            Self::Ellipse { geometry, properties } => LegacyDrawingPrimitive::Ellipse { geometry, properties },
            Self::Polyline { closed, points, geometry, properties } => LegacyDrawingPrimitive::Polyline { closed, points, geometry, properties },
            Self::Arc { flip_x, flip_y, geometry, properties } => LegacyDrawingPrimitive::Arc { flip_x, flip_y, geometry, properties },
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyDrawing<'a> {
    pub position: usize,
    pub horizontal_anchor: crate::LegacyHorizontalAnchor,
    pub vertical_anchor: crate::LegacyVerticalAnchor,
    pub z_order: i32,
    pub locked: bool,
    pub primitive: LegacyDrawingPrimitive<'a>,
}

impl LegacyDrawing<'_> {
    pub fn validate(&self) -> RtfResult<()> {
        let mut primitives = 0;
        let mut points = 0;
        self.primitive.validate_at(1, &mut primitives, &mut points)
    }

    pub(crate) fn into_owned(self) -> LegacyDrawing<'static> {
        LegacyDrawing {
            position: self.position,
            horizontal_anchor: self.horizontal_anchor,
            vertical_anchor: self.vertical_anchor,
            z_order: self.z_order,
            locked: self.locked,
            primitive: self.primitive.into_owned(),
        }
    }
}
