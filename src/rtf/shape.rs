//! RTF shape and drawing object support.
//!
//! This module provides support for shapes, text boxes, and drawing objects
//! in RTF documents.

use super::border::Border;
use super::types::{ColorRef, Formatting};
use std::borrow::Cow;

/// Shape type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShapeType {
    /// Rectangle
    #[default]
    Rectangle,
    /// Rounded rectangle
    RoundRectangle,
    /// Ellipse/circle
    Ellipse,
    /// Line
    Line,
    /// Polygon
    Polygon,
    /// Curve/arc
    Arc,
    /// Text box
    TextBox,
    /// Picture frame
    PictureFrame,
    /// Group of shapes
    Group,
    /// Unknown or custom shape
    Unknown,
}

/// Shape fill type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FillType {
    /// No fill
    #[default]
    None,
    /// Solid color fill
    Solid,
    /// Gradient fill
    Gradient,
    /// Pattern fill
    Pattern,
    /// Texture/image fill
    Texture,
}

/// Gradient direction
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum GradientDirection {
    /// Horizontal (left to right)
    #[default]
    Horizontal,
    /// Vertical (top to bottom)
    Vertical,
    /// Diagonal (top-left to bottom-right)
    DiagonalDown,
    /// Diagonal (bottom-left to top-right)
    DiagonalUp,
    /// From center
    FromCenter,
}

/// Shape fill properties
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Fill {
    /// Fill type
    pub fill_type: FillType,
    /// Primary fill color
    pub color: ColorRef,
    /// Secondary color (for gradients)
    pub color2: Option<ColorRef>,
    /// Gradient direction
    pub gradient_direction: GradientDirection,
    /// Fill opacity (0-100%)
    pub opacity: u8,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            fill_type: FillType::default(),
            color: 0,
            color2: None,
            gradient_direction: GradientDirection::default(),
            opacity: 100,
        }
    }
}

impl Fill {
    /// Create a solid fill
    #[inline]
    pub fn solid(color: ColorRef) -> Self {
        Self {
            fill_type: FillType::Solid,
            color,
            ..Default::default()
        }
    }

    /// Create a gradient fill
    #[inline]
    pub fn gradient(color1: ColorRef, color2: ColorRef, direction: GradientDirection) -> Self {
        Self {
            fill_type: FillType::Gradient,
            color: color1,
            color2: Some(color2),
            gradient_direction: direction,
            ..Default::default()
        }
    }
}

/// Shape position and size
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ShapeGeometry {
    /// X position (in twips)
    pub x: i32,
    /// Y position (in twips)
    pub y: i32,
    /// Width (in twips)
    pub width: i32,
    /// Height (in twips)
    pub height: i32,
    /// Rotation angle (in degrees, 0-360)
    pub rotation: i32,
    /// Z-order (stacking order)
    pub z_order: i32,
}

impl ShapeGeometry {
    /// Create a new geometry
    #[inline]
    pub fn new(x: i32, y: i32, width: i32, height: i32) -> Self {
        Self {
            x,
            y,
            width,
            height,
            rotation: 0,
            z_order: 0,
        }
    }
}

/// Text wrapping mode for shapes
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WrapMode {
    /// Text does not wrap around shape
    #[default]
    None,
    /// Text wraps around shape on both sides
    Square,
    /// Text wraps tightly around shape
    Tight,
    /// Text wraps through transparent regions
    Through,
    /// Text appears above shape
    TopAndBottom,
    /// Text appears behind shape
    Behind,
    /// Text appears in front of shape
    InFront,
}

/// RTF shape/drawing object
#[derive(Debug, Clone)]
pub struct Shape<'a> {
    /// Shape type
    pub shape_type: ShapeType,
    /// Geometry (position and size)
    pub geometry: ShapeGeometry,
    /// Fill properties
    pub fill: Fill,
    /// Border
    pub border: Border,
    /// Text content (for text boxes)
    pub text: Cow<'a, str>,
    /// Text formatting (for text boxes)
    pub text_formatting: Option<Formatting>,
    /// Text wrapping mode
    pub wrap_mode: WrapMode,
    /// Whether shape is behind text
    pub behind_doc: bool,
    /// Whether shape is locked (cannot be moved/resized)
    pub locked: bool,
    /// Shape name/identifier
    pub name: Cow<'a, str>,
}

impl<'a> Shape<'a> {
    /// Create a new shape
    #[inline]
    pub fn new(shape_type: ShapeType) -> Self {
        Self {
            shape_type,
            geometry: ShapeGeometry::default(),
            fill: Fill::default(),
            border: Border::default(),
            text: Cow::Borrowed(""),
            text_formatting: None,
            wrap_mode: WrapMode::default(),
            behind_doc: false,
            locked: false,
            name: Cow::Borrowed(""),
        }
    }

    /// Create a text box
    #[inline]
    pub fn text_box(text: Cow<'a, str>) -> Self {
        Self {
            shape_type: ShapeType::TextBox,
            text,
            ..Self::new(ShapeType::TextBox)
        }
    }

    /// Check if this is a text box
    #[inline]
    pub fn is_text_box(&self) -> bool {
        self.shape_type == ShapeType::TextBox
    }
}

/// Group of shapes
#[derive(Debug, Clone)]
pub struct ShapeGroup<'a> {
    /// Group name
    pub name: Cow<'a, str>,
    /// Shapes in the group
    pub shapes: Vec<Shape<'a>>,
    /// Group geometry (bounding box)
    pub geometry: ShapeGeometry,
}

impl<'a> ShapeGroup<'a> {
    /// Create a new shape group
    #[inline]
    pub fn new() -> Self {
        Self {
            name: Cow::Borrowed(""),
            shapes: Vec::new(),
            geometry: ShapeGeometry::default(),
        }
    }

    /// Add a shape to the group
    #[inline]
    pub fn add_shape(&mut self, shape: Shape<'a>) {
        self.shapes.push(shape);
    }

    /// Get all shapes in the group
    #[inline]
    pub fn shapes(&self) -> &[Shape<'a>] {
        &self.shapes
    }
}

impl<'a> Default for ShapeGroup<'a> {
    fn default() -> Self {
        Self::new()
    }
}
