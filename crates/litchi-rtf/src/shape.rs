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

/// A signed RTF drawing coordinate measured in twentieths of a point.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Default)]
pub struct ShapeTwips(i32);

impl ShapeTwips {
    pub const ZERO: Self = Self(0);

    #[inline]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    #[inline]
    pub const fn get(self) -> i32 {
        self.0
    }
}

impl From<i32> for ShapeTwips {
    fn from(value: i32) -> Self {
        Self(value)
    }
}

impl From<ShapeTwips> for i32 {
    fn from(value: ShapeTwips) -> Self {
        value.0
    }
}

/// A clockwise shape rotation measured in whole degrees.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ShapeRotationDegrees(i32);

impl ShapeRotationDegrees {
    #[inline]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    #[inline]
    pub const fn get(self) -> i32 {
        self.0
    }
}

/// A signed root-shape stacking order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Default)]
pub struct ShapeZOrder(i32);

impl ShapeZOrder {
    #[inline]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    #[inline]
    pub const fn get(self) -> i32 {
        self.0
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

/// Maximum decoded binary bytes retained by one OfficeArt property.
pub const MAX_SHAPE_PROPERTY_BINARY_BYTES: usize = 64 * 1024 * 1024;

/// Theme accent selector retained from an RTF `hsv` destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeThemeColor {
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Background1,
    Background2,
    Text1,
    Text2,
}

/// Inert theme metadata attached to a scalar OfficeArt color property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeThemeValue {
    pub color: ShapeThemeColor,
    pub tint: u8,
    pub shade: u8,
}

impl ShapeThemeValue {
    /// Validate the RTF rule that a color may be tinted or shaded, but not both.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.tint < 255 && self.shade < 255 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF hsv cannot apply tint and shade simultaneously".to_string(),
            ));
        }
        Ok(())
    }
}

/// Typed inert Word 6/95 drawing fallback from a shape's `shprslt` destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeResult<'a> {
    /// Legacy drawing fallback. Its body position is always zero while nested in a shape.
    pub drawing: crate::LegacyDrawing<'a>,
}

impl ShapeResult<'_> {
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.drawing.position != 0 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-result drawing must not have a document-body position".to_string(),
            ));
        }
        self.drawing.validate()
    }

    pub(crate) fn into_owned(self) -> ShapeResult<'static> {
        ShapeResult {
            drawing: self.drawing.into_owned(),
        }
    }
}

/// Upper bound for one inert shape-hyperlink string, in bytes.
pub const MAX_SHAPE_HYPERLINK_BYTES: usize = 65_536;

/// Inert hyperlink metadata from the `\hl` group of a shape property.
///
/// The RTF specification ("Hyperlink Property for Shapes") defines the
/// `{\hl {\hlloc …} {\hlsrc …} {\hlfr …}}` group, which may occur inside a
/// shape property (`\sp`) destination: `\hlloc` is the location string,
/// `\hlsrc` the source string, and `\hlfr` the friendly name; the three
/// groups may appear in any order.
///
/// The strings are passive metadata only: they are never resolved, opened,
/// fetched, validated as references, or activated.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ShapeHyperlink<'a> {
    /// Location string (`\hlloc`).
    pub location: Option<Cow<'a, str>>,
    /// Source string (`\hlsrc`).
    pub source: Option<Cow<'a, str>>,
    /// Friendly name (`\hlfr`).
    pub friendly_name: Option<Cow<'a, str>>,
}

impl<'a> ShapeHyperlink<'a> {
    fn validate_string(kind: &str, value: &str) -> crate::RtfResult<()> {
        if value.len() > MAX_SHAPE_HYPERLINK_BYTES {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF shape-hyperlink {kind} string exceeds the safety limit"
            )));
        }
        if value.contains(['\0', '\r', '\n']) {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF shape-hyperlink {kind} string contains a forbidden control character"
            )));
        }
        Ok(())
    }

    /// Validate presence and resource constraints.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.location.is_none() && self.source.is_none() && self.friendly_name.is_none() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape hyperlink must carry at least one string".to_string(),
            ));
        }
        if let Some(location) = &self.location {
            Self::validate_string("location", location)?;
        }
        if let Some(source) = &self.source {
            Self::validate_string("source", source)?;
        }
        if let Some(friendly_name) = &self.friendly_name {
            Self::validate_string("friendly name", friendly_name)?;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> ShapeHyperlink<'static> {
        ShapeHyperlink {
            location: self.location.map(|value| Cow::Owned(value.into_owned())),
            source: self.source.map(|value| Cow::Owned(value.into_owned())),
            friendly_name: self
                .friendly_name
                .map(|value| Cow::Owned(value.into_owned())),
        }
    }
}

/// A scalar or binary OfficeArt property retained from an RTF `sp` destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeProperty<'a> {
    /// OfficeArt property name from the `sn` destination
    pub name: Cow<'a, str>,
    /// Text value from the `sv` destination
    pub value: Cow<'a, str>,
    /// Optional decoded bytes from a nested starred `svb` destination
    pub binary_value: Option<Cow<'a, [u8]>>,
    /// Optional inert theme metadata from the starred `hsv` destination after `sv`
    pub theme_value: Option<ShapeThemeValue>,
    /// Optional inert hyperlink metadata from the `hl` group
    pub hyperlink: Option<ShapeHyperlink<'a>>,
}

impl<'a> ShapeProperty<'a> {
    /// Construct a scalar shape property.
    #[inline]
    pub fn new(name: Cow<'a, str>, value: Cow<'a, str>) -> Self {
        Self {
            name,
            value,
            binary_value: None,
            theme_value: None,
            hyperlink: None,
        }
    }

    /// Construct a scalar color property with inert theme metadata.
    #[inline]
    pub fn new_themed(
        name: Cow<'a, str>,
        value: Cow<'a, str>,
        theme_value: ShapeThemeValue,
    ) -> Self {
        Self {
            name,
            value,
            binary_value: None,
            theme_value: Some(theme_value),
            hyperlink: None,
        }
    }

    /// Construct a property whose value is an inert binary `svb` payload.
    #[inline]
    pub fn new_binary(name: Cow<'a, str>, value: Cow<'a, [u8]>) -> Self {
        Self {
            name,
            value: Cow::Borrowed(""),
            binary_value: Some(value),
            theme_value: None,
            hyperlink: None,
        }
    }

    /// Validate scalar/binary exclusivity and resource constraints.
    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.name.is_empty() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-property name must not be empty".to_string(),
            ));
        }
        if let Some(value) = &self.binary_value {
            if !self.value.is_empty() {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF shape property cannot contain scalar and binary values".to_string(),
                ));
            }
            if value.is_empty() || value.len() > MAX_SHAPE_PROPERTY_BINARY_BYTES {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF svb payload is empty or exceeds the safety limit".to_string(),
                ));
            }
        }
        if let Some(theme_value) = self.theme_value {
            if self.binary_value.is_some() || self.value.is_empty() {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF hsv requires a nonempty scalar shape-property value".to_string(),
                ));
            }
            theme_value.validate()?;
        }
        if let Some(hyperlink) = &self.hyperlink {
            hyperlink.validate()?;
        }
        Ok(())
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

    /// Construct geometry using named RTF coordinate units.
    #[inline]
    pub const fn from_twips(
        x: ShapeTwips,
        y: ShapeTwips,
        width: ShapeTwips,
        height: ShapeTwips,
    ) -> Self {
        Self {
            x: x.0,
            y: y.0,
            width: width.0,
            height: height.0,
            rotation: 0,
            z_order: 0,
        }
    }

    #[inline]
    pub const fn x_twips(self) -> ShapeTwips {
        ShapeTwips(self.x)
    }

    #[inline]
    pub const fn y_twips(self) -> ShapeTwips {
        ShapeTwips(self.y)
    }

    #[inline]
    pub const fn width_twips(self) -> ShapeTwips {
        ShapeTwips(self.width)
    }

    #[inline]
    pub const fn height_twips(self) -> ShapeTwips {
        ShapeTwips(self.height)
    }

    #[inline]
    pub const fn rotation_degrees(self) -> ShapeRotationDegrees {
        ShapeRotationDegrees(self.rotation)
    }

    #[inline]
    pub const fn z_order_value(self) -> ShapeZOrder {
        ShapeZOrder(self.z_order)
    }

    #[inline]
    pub fn set_rotation_degrees(&mut self, value: ShapeRotationDegrees) {
        self.rotation = value.0;
    }

    #[inline]
    pub fn set_z_order(&mut self, value: ShapeZOrder) {
        self.z_order = value.0;
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

/// Horizontal anchoring control used by `shpbx*` shape metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeHorizontalAnchor {
    Page,
    Margin,
    Column,
    ShapeProperty,
}

/// Vertical anchoring control used by `shpby*` shape metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeVerticalAnchor {
    Page,
    Margin,
    Paragraph,
    ShapeProperty,
}

/// Typed value of the RTF `shpwr` control.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeWrapStyle {
    Around,
    None,
    Square,
    Through,
    Tight,
    ThroughLargest,
    Other(i32),
}

impl ShapeWrapStyle {
    pub const fn from_rtf(value: i32) -> Self {
        match value {
            0 => Self::Around,
            1 => Self::None,
            2 => Self::Square,
            3 => Self::Through,
            4 => Self::Tight,
            5 => Self::ThroughLargest,
            other => Self::Other(other),
        }
    }

    pub const fn to_rtf(self) -> i32 {
        match self {
            Self::Around => 0,
            Self::None => 1,
            Self::Square => 2,
            Self::Through => 3,
            Self::Tight => 4,
            Self::ThroughLargest => 5,
            Self::Other(value) => value,
        }
    }
}

/// Typed value of the RTF `shpwrk` wrap-side control.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeWrapSide {
    Both,
    Left,
    Right,
    Largest,
    Other(i32),
}

impl ShapeWrapSide {
    pub const fn from_rtf(value: i32) -> Self {
        match value {
            0 => Self::Both,
            1 => Self::Left,
            2 => Self::Right,
            3 => Self::Largest,
            other => Self::Other(other),
        }
    }

    pub const fn to_rtf(self) -> i32 {
        match self {
            Self::Both => 0,
            Self::Left => 1,
            Self::Right => 2,
            Self::Largest => 3,
            Self::Other(value) => value,
        }
    }
}

/// RTF shape/drawing object
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shape<'a> {
    /// UTF-8 body offset for a standalone root shape.
    pub position: usize,
    /// Whether the shape owns a normal starred `shpinst` destination.
    pub instruction_present: bool,
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
    /// Whether an explicit `shptxt` destination was present or requested, including empty text
    pub text_destination_present: bool,
    /// Text formatting (for text boxes)
    pub text_formatting: Option<Formatting>,
    /// Positional root shapes owned by this text-box story.
    pub text_shapes: Vec<Shape<'a>>,
    /// Positional root shape groups owned by this text-box story.
    pub text_shape_groups: Vec<ShapeGroup<'a>>,
    /// Exact source order of root drawings in the text-box story.
    pub text_drawing_order: Vec<StoryDrawing>,
    /// Exact source order of drawings, fields, and page breaks in the text-box story.
    pub text_story_events: Vec<crate::StoryEvent>,
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
    /// Optional typed legacy drawing fallback from the root-only `shprslt` destination
    pub result: Option<ShapeResult<'a>>,
    /// Additional normative shape-info controls in producer order.
    pub info: Vec<ShapeGroupInfo>,
}

impl<'a> Shape<'a> {
    /// Create a new shape
    #[inline]
    pub fn new(shape_type: ShapeType) -> Self {
        Self {
            position: 0,
            instruction_present: true,
            shape_type,
            geometry: ShapeGeometry::default(),
            fill: Fill::default(),
            border: Border::default(),
            line: ShapeLine::default(),
            text: Cow::Borrowed(""),
            text_destination_present: false,
            text_formatting: None,
            text_shapes: Vec::new(),
            text_shape_groups: Vec::new(),
            text_drawing_order: Vec::new(),
            text_story_events: Vec::new(),
            wrap_mode: WrapMode::default(),
            behind_doc: false,
            is_background: false,
            locked: false,
            name: Cow::Borrowed(""),
            properties: Vec::new(),
            result: None,
            info: Vec::new(),
        }
    }

    /// Create a text box
    #[inline]
    pub fn text_box(text: Cow<'a, str>) -> Self {
        let mut shape = Self::new(ShapeType::TextBox);
        shape.text = text;
        shape.text_destination_present = true;
        shape
    }

    /// Set explicit inert text-frame content and return the previous text.
    pub fn set_text(&mut self, text: Cow<'a, str>) -> Cow<'a, str> {
        self.text_destination_present = true;
        std::mem::replace(&mut self.text, text)
    }

    /// Remove the explicit text-frame destination and return its previous text.
    pub fn clear_text(&mut self) -> Cow<'a, str> {
        self.text_destination_present = false;
        self.text_shapes.clear();
        self.text_shape_groups.clear();
        self.text_drawing_order.clear();
        self.text_story_events.clear();
        std::mem::replace(&mut self.text, Cow::Borrowed(""))
    }

    /// Append a validated positional root shape to this text-box story.
    pub fn push_text_shape(&mut self, shape: Shape<'a>) -> crate::RtfResult<()> {
        validate_story_shape(
            self.text.as_ref(),
            &shape,
            self.text_shapes.last(),
            "shape text",
        )?;
        shape.validate()?;
        if self.text_shapes.len() >= 65_536 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-text shape count exceeds the safety limit".to_string(),
            ));
        }
        self.text_destination_present = true;
        self.text_drawing_order
            .push(StoryDrawing::Shape(self.text_shapes.len()));
        self.text_story_events
            .push(crate::StoryEvent::Drawing(StoryDrawing::Shape(
                self.text_shapes.len(),
            )));
        self.text_shapes.push(shape);
        Ok(())
    }

    /// Append a validated positional root shape group to this text-box story.
    pub fn push_text_shape_group(&mut self, group: ShapeGroup<'a>) -> crate::RtfResult<()> {
        validate_story_group(
            self.text.as_ref(),
            &group,
            self.text_shape_groups.last(),
            "shape text",
        )?;
        group.validate()?;
        if self.text_shape_groups.len() >= 16_384 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-text shape-group count exceeds the safety limit".to_string(),
            ));
        }
        self.text_destination_present = true;
        self.text_drawing_order
            .push(StoryDrawing::ShapeGroup(self.text_shape_groups.len()));
        self.text_story_events
            .push(crate::StoryEvent::Drawing(StoryDrawing::ShapeGroup(
                self.text_shape_groups.len(),
            )));
        self.text_shape_groups.push(group);
        Ok(())
    }

    /// Clear drawings owned by this text-box story without removing its text.
    pub fn clear_text_drawings(&mut self) {
        self.text_shapes.clear();
        self.text_shape_groups.clear();
        self.text_drawing_order.clear();
        self.text_story_events
            .retain(|event| !matches!(event, crate::StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.text_story_events
            .iter()
            .filter_map(|event| match event {
                crate::StoryEvent::PageBreak(page_break) => Some(page_break),
                _ => None,
            })
    }

    pub fn push_page_break(&mut self, position: usize) -> crate::RtfResult<()> {
        self.text_destination_present = true;
        crate::field::push_story_page_break(
            &mut self.text_story_events,
            self.text.as_ref(),
            position,
            "shape text",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.text_story_events
            .retain(|event| !matches!(event, crate::StoryEvent::PageBreak(_)));
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

    /// Return the last declared `shplid`, if present.
    pub fn shape_id(&self) -> Option<i32> {
        self.info.iter().rev().find_map(|info| match info {
            ShapeGroupInfo::ShapeId(value) => Some(*value),
            _ => None,
        })
    }

    /// Recursively find this shape or one owned by its text story by name.
    pub fn find_by_name(&self, name: &str) -> Option<&Shape<'a>> {
        if self.name == name || self.property("wzName") == Some(name) {
            return Some(self);
        }
        self.text_shapes
            .iter()
            .find_map(|shape| shape.find_by_name(name))
            .or_else(|| {
                self.text_shape_groups
                    .iter()
                    .find_map(|group| group.find_shape_by_name(name))
            })
    }

    /// Recursively find this shape or one owned by its text story by `shplid`.
    pub fn find_by_id(&self, id: i32) -> Option<&Shape<'a>> {
        if self.shape_id() == Some(id) {
            return Some(self);
        }
        self.text_shapes
            .iter()
            .find_map(|shape| shape.find_by_id(id))
            .or_else(|| {
                self.text_shape_groups
                    .iter()
                    .find_map(|group| group.find_shape_by_id(id))
            })
    }

    /// Insert or replace the last property with the same name atomically.
    pub fn set_property(
        &mut self,
        property: ShapeProperty<'a>,
    ) -> crate::RtfResult<Option<ShapeProperty<'a>>> {
        property.validate()?;
        if property.name.len().saturating_add(property.value.len()) > 1_048_576 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape property exceeds the safety limit".to_string(),
            ));
        }
        if let Some(index) = self
            .properties
            .iter()
            .rposition(|current| current.name == property.name)
        {
            return Ok(Some(std::mem::replace(
                &mut self.properties[index],
                property,
            )));
        }
        if self.properties.len() >= 65_536 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-property count exceeds the safety limit".to_string(),
            ));
        }
        self.properties.push(property);
        Ok(None)
    }

    /// Remove the last property with the requested name.
    pub fn remove_property(&mut self, name: &str) -> Option<ShapeProperty<'a>> {
        let index = self
            .properties
            .iter()
            .rposition(|property| property.name == name)?;
        Some(self.properties.remove(index))
    }

    /// Validate a standalone shape independently of its document position.
    pub fn validate(&self) -> crate::RtfResult<()> {
        self.validate_at_depth(0)
    }

    pub(crate) fn validate_at_depth(&self, depth: usize) -> crate::RtfResult<()> {
        if depth >= 64 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape story nesting exceeds the safety limit".to_string(),
            ));
        }
        if self.properties.len() > 65_536
            || self.text.len() > 16 * 1_048_576
            || self.info.len() > 32
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape exceeds the safety limit".to_string(),
            ));
        }
        self.geometry
            .x
            .checked_add(self.geometry.width)
            .ok_or_else(|| {
                crate::RtfError::MalformedDocument(
                    "RTF shape horizontal geometry overflows".to_string(),
                )
            })?;
        self.geometry
            .y
            .checked_add(self.geometry.height)
            .ok_or_else(|| {
                crate::RtfError::MalformedDocument(
                    "RTF shape vertical geometry overflows".to_string(),
                )
            })?;
        for property in &self.properties {
            property.validate()?;
            if property.name.len().saturating_add(property.value.len()) > 1_048_576 {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF shape property exceeds the safety limit".to_string(),
                ));
            }
        }
        if let Some(result) = &self.result {
            result.validate()?;
        }
        if !self.instruction_present
            && (self.result.is_none()
                || !self.properties.is_empty()
                || !self.info.is_empty()
                || self.text_destination_present
                || !self.text.is_empty()
                || !self.text_shapes.is_empty()
                || !self.text_shape_groups.is_empty()
                || !self.text_story_events.is_empty()
                || self.geometry != ShapeGeometry::default()
                || self.shape_type != ShapeType::Unknown)
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF fallback-only shape cannot contain shape instructions".to_string(),
            ));
        }
        validate_story_drawings_at_depth(
            self.text.as_ref(),
            &self.text_shapes,
            &self.text_shape_groups,
            &self.text_drawing_order,
            "shape text",
            depth + 1,
        )?;
        crate::field::validate_story_events(
            self.text.as_ref(),
            &self.text_shapes,
            &self.text_shape_groups,
            &self.text_drawing_order,
            &self.text_story_events,
            "shape text",
        )?;
        Ok(())
    }

    /// Convert this shape and all text-story drawings to owned storage.
    pub fn into_owned(self) -> Shape<'static> {
        Shape {
            position: self.position,
            instruction_present: self.instruction_present,
            shape_type: self.shape_type,
            geometry: self.geometry,
            fill: self.fill,
            border: self.border,
            line: self.line,
            text: Cow::Owned(self.text.into_owned()),
            text_destination_present: self.text_destination_present,
            text_formatting: self.text_formatting,
            text_shapes: self
                .text_shapes
                .into_iter()
                .map(Shape::into_owned)
                .collect(),
            text_shape_groups: self
                .text_shape_groups
                .into_iter()
                .map(ShapeGroup::into_owned)
                .collect(),
            text_drawing_order: self.text_drawing_order,
            text_story_events: self.text_story_events,
            wrap_mode: self.wrap_mode,
            behind_doc: self.behind_doc,
            is_background: self.is_background,
            locked: self.locked,
            name: Cow::Owned(self.name.into_owned()),
            properties: self
                .properties
                .into_iter()
                .map(|property| ShapeProperty {
                    name: Cow::Owned(property.name.into_owned()),
                    value: Cow::Owned(property.value.into_owned()),
                    binary_value: property
                        .binary_value
                        .map(|value| Cow::Owned(value.into_owned())),
                    theme_value: property.theme_value,
                    hyperlink: property.hyperlink.map(ShapeHyperlink::into_owned),
                })
                .collect(),
            result: self.result.map(ShapeResult::into_owned),
            info: self.info,
        }
    }
}

/// Reference to a shape-group child in bottom-to-top order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeGroupChild {
    /// Index into [`ShapeGroup::shapes`].
    Shape(usize),
    /// Index into [`ShapeGroup::groups`].
    Group(usize),
}

/// Exact source order of root shapes and shape groups in an owning story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StoryDrawing {
    /// Index into the story's shape collection.
    Shape(usize),
    /// Index into the story's root shape-group collection.
    ShapeGroup(usize),
}

/// Shape-info control retained on a group instance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeGroupInfo {
    ShapeId(i32),
    InHeader(bool),
    HorizontalPage,
    HorizontalMargin,
    HorizontalColumn,
    IgnoreHorizontal,
    VerticalPage,
    VerticalMargin,
    VerticalParagraph,
    IgnoreVertical,
    Wrap(i32),
    WrapSide(i32),
    BelowText(bool),
    LockAnchor,
}

impl ShapeGroupInfo {
    pub const fn horizontal_anchor(value: ShapeHorizontalAnchor) -> Self {
        match value {
            ShapeHorizontalAnchor::Page => Self::HorizontalPage,
            ShapeHorizontalAnchor::Margin => Self::HorizontalMargin,
            ShapeHorizontalAnchor::Column => Self::HorizontalColumn,
            ShapeHorizontalAnchor::ShapeProperty => Self::IgnoreHorizontal,
        }
    }

    pub const fn vertical_anchor(value: ShapeVerticalAnchor) -> Self {
        match value {
            ShapeVerticalAnchor::Page => Self::VerticalPage,
            ShapeVerticalAnchor::Margin => Self::VerticalMargin,
            ShapeVerticalAnchor::Paragraph => Self::VerticalParagraph,
            ShapeVerticalAnchor::ShapeProperty => Self::IgnoreVertical,
        }
    }

    pub const fn wrap(value: ShapeWrapStyle) -> Self {
        Self::Wrap(value.to_rtf())
    }
    pub const fn wrap_side(value: ShapeWrapSide) -> Self {
        Self::WrapSide(value.to_rtf())
    }

    pub const fn as_horizontal_anchor(self) -> Option<ShapeHorizontalAnchor> {
        match self {
            Self::HorizontalPage => Some(ShapeHorizontalAnchor::Page),
            Self::HorizontalMargin => Some(ShapeHorizontalAnchor::Margin),
            Self::HorizontalColumn => Some(ShapeHorizontalAnchor::Column),
            Self::IgnoreHorizontal => Some(ShapeHorizontalAnchor::ShapeProperty),
            _ => None,
        }
    }

    pub const fn as_vertical_anchor(self) -> Option<ShapeVerticalAnchor> {
        match self {
            Self::VerticalPage => Some(ShapeVerticalAnchor::Page),
            Self::VerticalMargin => Some(ShapeVerticalAnchor::Margin),
            Self::VerticalParagraph => Some(ShapeVerticalAnchor::Paragraph),
            Self::IgnoreVertical => Some(ShapeVerticalAnchor::ShapeProperty),
            _ => None,
        }
    }

    pub const fn as_wrap(self) -> Option<ShapeWrapStyle> {
        match self {
            Self::Wrap(value) => Some(ShapeWrapStyle::from_rtf(value)),
            _ => None,
        }
    }
    pub const fn as_wrap_side(self) -> Option<ShapeWrapSide> {
        match self {
            Self::WrapSide(value) => Some(ShapeWrapSide::from_rtf(value)),
            _ => None,
        }
    }
}

/// Group of shapes
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeGroup<'a> {
    /// UTF-8 offset in the owning body or header/footer story.
    pub position: usize,
    /// Group name
    pub name: Cow<'a, str>,
    /// Shapes in the group
    pub shapes: Vec<Shape<'a>>,
    /// Nested shape groups in document order
    pub groups: Vec<ShapeGroup<'a>>,
    /// Bottom-to-top order of shapes and nested groups.
    pub child_order: Vec<ShapeGroupChild>,
    /// Additional normative shape-info controls in producer order.
    pub info: Vec<ShapeGroupInfo>,
    /// Group geometry (bounding box)
    pub geometry: ShapeGeometry,
    /// All scalar OfficeArt properties attached to the group
    pub properties: Vec<ShapeProperty<'a>>,
    /// Optional legacy fallback for the complete root group.
    pub result: Option<ShapeResult<'a>>,
}

impl<'a> ShapeGroup<'a> {
    /// Create a new shape group
    #[inline]
    pub fn new() -> Self {
        Self {
            position: 0,
            name: Cow::Borrowed(""),
            shapes: Vec::new(),
            groups: Vec::new(),
            child_order: Vec::new(),
            info: Vec::new(),
            geometry: ShapeGeometry::default(),
            properties: Vec::new(),
            result: None,
        }
    }

    /// Add a shape to the group
    #[inline]
    pub fn add_shape(&mut self, shape: Shape<'a>) {
        self.child_order
            .push(ShapeGroupChild::Shape(self.shapes.len()));
        self.shapes.push(shape);
    }

    /// Add a nested shape group.
    #[inline]
    pub fn add_group(&mut self, group: ShapeGroup<'a>) {
        self.child_order
            .push(ShapeGroupChild::Group(self.groups.len()));
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

    /// Return child indices in bottom-to-top z-order.
    #[inline]
    pub fn child_order(&self) -> &[ShapeGroupChild] {
        &self.child_order
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

    /// Recursively find a shape by its `wzName` property or typed name.
    pub fn find_shape_by_name(&self, name: &str) -> Option<&Shape<'a>> {
        self.shapes
            .iter()
            .find_map(|shape| shape.find_by_name(name))
            .or_else(|| {
                self.groups
                    .iter()
                    .find_map(|group| group.find_shape_by_name(name))
            })
    }

    /// Recursively find a shape by `shplid`.
    pub fn find_shape_by_id(&self, id: i32) -> Option<&Shape<'a>> {
        self.shapes
            .iter()
            .find_map(|shape| shape.find_by_id(id))
            .or_else(|| {
                self.groups
                    .iter()
                    .find_map(|group| group.find_shape_by_id(id))
            })
    }

    /// Atomically replace a directly contained shape.
    pub fn replace_shape(
        &mut self,
        index: usize,
        replacement: Shape<'a>,
    ) -> crate::RtfResult<Shape<'a>> {
        if index >= self.shapes.len() {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF grouped shape index {index} is out of bounds"
            )));
        }
        let mut staged = self.clone();
        let old = std::mem::replace(&mut staged.shapes[index], replacement);
        staged.validate_at_depth(0, true)?;
        *self = staged;
        Ok(old)
    }

    /// Atomically remove a directly contained shape and repair child indices.
    pub fn remove_shape(&mut self, index: usize) -> crate::RtfResult<Shape<'a>> {
        if index >= self.shapes.len() {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF grouped shape index {index} is out of bounds"
            )));
        }
        let mut staged = self.clone();
        let old = staged.shapes.remove(index);
        staged
            .child_order
            .retain(|child| !matches!(child, ShapeGroupChild::Shape(value) if *value == index));
        for child in &mut staged.child_order {
            if let ShapeGroupChild::Shape(value) = child
                && *value > index
            {
                *value -= 1;
            }
        }
        staged.validate_at_depth(0, true)?;
        *self = staged;
        Ok(old)
    }

    /// Atomically reorder direct children in bottom-to-top order.
    pub fn move_child(&mut self, from: usize, to: usize) -> crate::RtfResult<()> {
        if from >= self.child_order.len() || to >= self.child_order.len() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape-group child reorder index is out of bounds".to_string(),
            ));
        }
        if from == to {
            return Ok(());
        }
        let mut staged = self.clone();
        let child = staged.child_order.remove(from);
        staged.child_order.insert(to, child);
        staged.validate_at_depth(0, true)?;
        *self = staged;
        Ok(())
    }

    /// Validate root-group structure and resource limits.
    pub fn validate(&self) -> crate::RtfResult<()> {
        self.validate_at_depth(0, true)
    }

    /// Convert this group and its complete drawing tree to owned storage.
    pub fn into_owned(self) -> ShapeGroup<'static> {
        ShapeGroup {
            position: self.position,
            name: Cow::Owned(self.name.into_owned()),
            shapes: self.shapes.into_iter().map(Shape::into_owned).collect(),
            groups: self
                .groups
                .into_iter()
                .map(ShapeGroup::into_owned)
                .collect(),
            child_order: self.child_order,
            info: self.info,
            geometry: self.geometry,
            properties: self
                .properties
                .into_iter()
                .map(|property| ShapeProperty {
                    name: Cow::Owned(property.name.into_owned()),
                    value: Cow::Owned(property.value.into_owned()),
                    binary_value: property
                        .binary_value
                        .map(|value| Cow::Owned(value.into_owned())),
                    theme_value: property.theme_value,
                    hyperlink: property.hyperlink.map(ShapeHyperlink::into_owned),
                })
                .collect(),
            result: self.result.map(ShapeResult::into_owned),
        }
    }

    pub(crate) fn validate_at_depth(&self, depth: usize, root: bool) -> crate::RtfResult<()> {
        if depth >= 64 {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape group nesting exceeds the safety limit".to_string(),
            ));
        }
        if self.shapes.len() > 65_536
            || self.groups.len() > 16_384
            || self.child_order.len() > 65_536
            || self.info.len() > 32
            || self.properties.len() > 65_536
        {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape group count exceeds the safety limit".to_string(),
            ));
        }
        if !root && self.result.is_some() {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shprslt is allowed only on a root shape group".to_string(),
            ));
        }
        if let Some(result) = &self.result {
            result.validate()?;
        }
        self.geometry
            .x
            .checked_add(self.geometry.width)
            .ok_or_else(|| {
                crate::RtfError::MalformedDocument(
                    "RTF shape group horizontal geometry overflows".to_string(),
                )
            })?;
        self.geometry
            .y
            .checked_add(self.geometry.height)
            .ok_or_else(|| {
                crate::RtfError::MalformedDocument(
                    "RTF shape group vertical geometry overflows".to_string(),
                )
            })?;
        for property in &self.properties {
            property.validate()?;
            if property.name.len().saturating_add(property.value.len()) > 1_048_576 {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF shape group property exceeds the safety limit".to_string(),
                ));
            }
        }
        let mut saw_shapes = vec![false; self.shapes.len()];
        let mut saw_groups = vec![false; self.groups.len()];
        for child in &self.child_order {
            match *child {
                ShapeGroupChild::Shape(index) if index < saw_shapes.len() && !saw_shapes[index] => {
                    saw_shapes[index] = true;
                },
                ShapeGroupChild::Group(index) if index < saw_groups.len() && !saw_groups[index] => {
                    saw_groups[index] = true;
                },
                _ => {
                    return Err(crate::RtfError::MalformedDocument(
                        "RTF shape group child order is invalid".to_string(),
                    ));
                },
            }
        }
        if saw_shapes.iter().any(|seen| !seen) || saw_groups.iter().any(|seen| !seen) {
            return Err(crate::RtfError::MalformedDocument(
                "RTF shape group child order is incomplete".to_string(),
            ));
        }
        for shape in &self.shapes {
            if shape.position != 0 {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF grouped shape position must be zero".to_string(),
                ));
            }
            if !shape.instruction_present {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF grouped shape must contain shpinst".to_string(),
                ));
            }
            if shape.result.is_some() {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF grouped shape cannot contain shprslt".to_string(),
                ));
            }
            if shape.properties.len() > 65_536 || shape.text.len() > 16 * 1_048_576 {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF grouped shape exceeds the safety limit".to_string(),
                ));
            }
            for property in &shape.properties {
                property.validate()?;
                if property.name.len().saturating_add(property.value.len()) > 1_048_576 {
                    return Err(crate::RtfError::MalformedDocument(
                        "RTF grouped shape property exceeds the safety limit".to_string(),
                    ));
                }
            }
            shape.validate_at_depth(depth + 1)?;
        }
        for group in &self.groups {
            if group.position != 0 {
                return Err(crate::RtfError::MalformedDocument(
                    "RTF nested shape-group position must be zero".to_string(),
                ));
            }
            group.validate_at_depth(depth + 1, false)?;
        }
        Ok(())
    }
}

pub(crate) fn validate_story_drawings(
    text: &str,
    shapes: &[Shape<'_>],
    groups: &[ShapeGroup<'_>],
    order: &[StoryDrawing],
    story: &str,
) -> crate::RtfResult<()> {
    validate_story_drawings_at_depth(text, shapes, groups, order, story, 0)
}

fn validate_story_drawings_at_depth(
    text: &str,
    shapes: &[Shape<'_>],
    groups: &[ShapeGroup<'_>],
    order: &[StoryDrawing],
    story: &str,
    depth: usize,
) -> crate::RtfResult<()> {
    if depth >= 64
        || shapes.len() > 65_536
        || groups.len() > 16_384
        || order.len() != shapes.len().saturating_add(groups.len())
    {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {story} drawing nesting or count exceeds the safety limit"
        )));
    }
    let mut previous_shape = None;
    for shape in shapes {
        validate_story_shape(text, shape, previous_shape, story)?;
        shape.validate_at_depth(depth)?;
        previous_shape = Some(shape);
    }
    let mut previous_group = None;
    for group in groups {
        validate_story_group(text, group, previous_group, story)?;
        group.validate_at_depth(depth, true)?;
        previous_group = Some(group);
    }
    let mut saw_shapes = vec![false; shapes.len()];
    let mut saw_groups = vec![false; groups.len()];
    let mut previous_position = None;
    for drawing in order {
        let position = match *drawing {
            StoryDrawing::Shape(index) if index < shapes.len() && !saw_shapes[index] => {
                saw_shapes[index] = true;
                shapes[index].position
            },
            StoryDrawing::ShapeGroup(index) if index < groups.len() && !saw_groups[index] => {
                saw_groups[index] = true;
                groups[index].position
            },
            _ => {
                return Err(crate::RtfError::MalformedDocument(format!(
                    "RTF {story} drawing order is invalid"
                )));
            },
        };
        if previous_position.is_some_and(|previous| previous > position) {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF {story} drawing order moves backwards"
            )));
        }
        previous_position = Some(position);
    }
    Ok(())
}

fn validate_story_shape(
    text: &str,
    shape: &Shape<'_>,
    previous: Option<&Shape<'_>>,
    story: &str,
) -> crate::RtfResult<()> {
    if shape.is_background
        || text.get(shape.position..shape.position).is_none()
        || previous.is_some_and(|previous| previous.position > shape.position)
    {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {story} shapes are outside or out of story order"
        )));
    }
    Ok(())
}

fn validate_story_group(
    text: &str,
    group: &ShapeGroup<'_>,
    previous: Option<&ShapeGroup<'_>>,
    story: &str,
) -> crate::RtfResult<()> {
    if text.get(group.position..group.position).is_none()
        || previous.is_some_and(|previous| previous.position > group.position)
    {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {story} shape groups are outside or out of story order"
        )));
    }
    Ok(())
}

impl<'a> Default for ShapeGroup<'a> {
    fn default() -> Self {
        Self::new()
    }
}
