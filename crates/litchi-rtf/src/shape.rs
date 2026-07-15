//! RTF shape and drawing object support.
//!
//! This module provides support for shapes, text boxes, and drawing objects
//! in RTF documents.

use super::border::Border;
use super::types::Formatting;
use std::borrow::Cow;

/// Raw 32-bit OfficeArt color value used by RTF shape properties.
///
/// The low three bytes contain blue, green, and red respectively. The high
/// byte contains the palette, scheme, and system-color flags defined by
/// `OfficeArtCOLORREF` in MS-ODRAW.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct OfficeArtColor(pub u32);

impl OfficeArtColor {
    /// Default OfficeArt white.
    pub const WHITE: Self = Self(0x00FF_FFFF);
    /// Default OfficeArt black.
    pub const BLACK: Self = Self(0);

    /// Construct a direct RGB color.
    #[inline]
    pub const fn from_rgb(red: u8, green: u8, blue: u8) -> Self {
        Self((blue as u32) | ((green as u32) << 8) | ((red as u32) << 16))
    }

    /// Return the unmodified OfficeArt value, including its high-byte flags.
    #[inline]
    pub const fn raw(self) -> u32 {
        self.0
    }

    /// Whether the high byte marks this color as ignored.
    #[inline]
    pub const fn is_ignored(self) -> bool {
        self.0 >> 24 == 0xFF
    }

    /// Red component of a direct RGB value.
    #[inline]
    pub const fn red(self) -> u8 {
        (self.0 >> 16) as u8
    }

    /// Green component of a direct RGB value.
    #[inline]
    pub const fn green(self) -> u8 {
        (self.0 >> 8) as u8
    }

    /// Blue component of a direct RGB value.
    #[inline]
    pub const fn blue(self) -> u8 {
        self.0 as u8
    }
}

impl From<u32> for OfficeArtColor {
    fn from(value: u32) -> Self {
        Self(value)
    }
}

impl From<u16> for OfficeArtColor {
    fn from(value: u16) -> Self {
        Self(u32::from(value))
    }
}

/// Unsigned 16.16 fixed-point value used by OfficeArt opacity properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OfficeArtOpacity(pub u32);

impl OfficeArtOpacity {
    /// Fully transparent.
    pub const TRANSPARENT: Self = Self(0);
    /// Fully opaque.
    pub const OPAQUE: Self = Self(0x0001_0000);

    /// Return the unmodified unsigned 16.16 value.
    #[inline]
    pub const fn raw(self) -> u32 {
        self.0
    }

    /// Convert the fixed-point value to a fraction.
    #[inline]
    pub fn as_fraction(self) -> f64 {
        f64::from(self.0) / 65_536.0
    }
}

impl Default for OfficeArtOpacity {
    fn default() -> Self {
        Self::OPAQUE
    }
}

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
    /// A valid OfficeArt shape type not represented by a named variant
    Custom(i32),
    /// No usable shape type was specified
    Unknown,
}

/// Shape fill type
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FillType {
    /// No fill
    None,
    /// Solid color fill
    #[default]
    Solid,
    /// Gradient fill
    Gradient,
    /// Pattern fill
    Pattern,
    /// Texture/image fill
    Texture,
    /// Picture fill
    Picture,
    /// Inherit the application background fill
    Background,
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
    pub color: OfficeArtColor,
    /// Secondary color (for gradients)
    pub color2: Option<OfficeArtColor>,
    /// Gradient direction
    pub gradient_direction: GradientDirection,
    /// Exact OfficeArt 16.16 fill opacity
    pub opacity: OfficeArtOpacity,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            fill_type: FillType::default(),
            color: OfficeArtColor::WHITE,
            color2: None,
            gradient_direction: GradientDirection::default(),
            opacity: OfficeArtOpacity::OPAQUE,
        }
    }
}

impl Fill {
    /// Create a solid fill
    #[inline]
    pub fn solid(color: impl Into<OfficeArtColor>) -> Self {
        Self {
            fill_type: FillType::Solid,
            color: color.into(),
            ..Default::default()
        }
    }

    /// Create a gradient fill
    #[inline]
    pub fn gradient<C1, C2>(color1: C1, color2: C2, direction: GradientDirection) -> Self
    where
        C1: Into<OfficeArtColor>,
        C2: Into<OfficeArtColor>,
    {
        Self {
            fill_type: FillType::Gradient,
            color: color1.into(),
            color2: Some(color2.into()),
            gradient_direction: direction,
            ..Default::default()
        }
    }
}

/// Exact OfficeArt line properties used by a shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeLine {
    /// Whether the line is enabled (`fLine`)
    pub visible: bool,
    /// Raw OfficeArt foreground color
    pub color: OfficeArtColor,
    /// Width in English Metric Units (EMUs)
    pub width_emu: i32,
}

impl Default for ShapeLine {
    fn default() -> Self {
        Self {
            visible: true,
            color: OfficeArtColor::BLACK,
            width_emu: 0x2535,
        }
    }
}

/// A scalar OfficeArt property retained from an RTF `sp` destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeProperty<'a> {
    /// OfficeArt property name from the `sn` destination
    pub name: Cow<'a, str>,
    /// Text value from the `sv` destination
    pub value: Cow<'a, str>,
}

impl<'a> ShapeProperty<'a> {
    /// Construct a scalar shape property.
    #[inline]
    pub fn new(name: Cow<'a, str>, value: Cow<'a, str>) -> Self {
        Self { name, value }
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
    /// Exact OfficeArt line properties
    pub line: ShapeLine,
    /// Text content (for text boxes)
    pub text: Cow<'a, str>,
    /// Text formatting (for text boxes)
    pub text_formatting: Option<Formatting>,
    /// Text wrapping mode
    pub wrap_mode: WrapMode,
    /// Whether shape is behind text
    pub behind_doc: bool,
    /// Whether the shape is a document/page background (`fBackground`)
    pub is_background: bool,
    /// Whether shape is locked (cannot be moved/resized)
    pub locked: bool,
    /// Shape name/identifier
    pub name: Cow<'a, str>,
    /// All scalar OfficeArt properties, including properties unknown to this crate
    pub properties: Vec<ShapeProperty<'a>>,
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
            line: ShapeLine::default(),
            text: Cow::Borrowed(""),
            text_formatting: None,
            wrap_mode: WrapMode::default(),
            behind_doc: false,
            is_background: false,
            locked: false,
            name: Cow::Borrowed(""),
            properties: Vec::new(),
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

    /// Return the last scalar OfficeArt property with the requested name.
    #[inline]
    pub fn property(&self, name: &str) -> Option<&str> {
        self.properties
            .iter()
            .rev()
            .find(|property| property.name == name)
            .map(|property| property.value.as_ref())
    }
}

/// Group of shapes
#[derive(Debug, Clone)]
pub struct ShapeGroup<'a> {
    /// Group name
    pub name: Cow<'a, str>,
    /// Shapes in the group
    pub shapes: Vec<Shape<'a>>,
    /// Nested shape groups in document order
    pub groups: Vec<ShapeGroup<'a>>,
    /// Group geometry (bounding box)
    pub geometry: ShapeGeometry,
    /// All scalar OfficeArt properties attached to the group
    pub properties: Vec<ShapeProperty<'a>>,
}

impl<'a> ShapeGroup<'a> {
    /// Create a new shape group
    #[inline]
    pub fn new() -> Self {
        Self {
            name: Cow::Borrowed(""),
            shapes: Vec::new(),
            groups: Vec::new(),
            geometry: ShapeGeometry::default(),
            properties: Vec::new(),
        }
    }

    /// Add a shape to the group
    #[inline]
    pub fn add_shape(&mut self, shape: Shape<'a>) {
        self.shapes.push(shape);
    }

    /// Add a nested shape group.
    #[inline]
    pub fn add_group(&mut self, group: ShapeGroup<'a>) {
        self.groups.push(group);
    }

    /// Get all shapes in the group
    #[inline]
    pub fn shapes(&self) -> &[Shape<'a>] {
        &self.shapes
    }

    /// Get directly nested shape groups.
    #[inline]
    pub fn groups(&self) -> &[ShapeGroup<'a>] {
        &self.groups
    }

    /// Return the last scalar OfficeArt property with the requested name.
    #[inline]
    pub fn property(&self, name: &str) -> Option<&str> {
        self.properties
            .iter()
            .rev()
            .find(|property| property.name == name)
            .map(|property| property.value.as_ref())
    }
}

impl<'a> Default for ShapeGroup<'a> {
    fn default() -> Self {
        Self::new()
    }
}
