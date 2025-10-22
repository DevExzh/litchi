use super::super::package::PptError;
/// Base shape trait and common shape functionality.
///
/// This module defines the core Shape trait that all shape types implement,
/// along with common properties and methods for working with shapes.
use std::fmt;

/// Types of shapes in PowerPoint presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Text box shape
    TextBox,
    /// Placeholder shape (title, content, etc.)
    Placeholder,
    /// Auto shape (rectangle, oval, etc.)
    AutoShape,
    /// Picture shape
    Picture,
    /// Group shape (container for other shapes)
    Group,
    /// Line shape
    Line,
    /// Connector shape
    Connector,
    /// Object shape (embedded objects)
    Object,
    /// Table shape
    Table,
    /// Unknown shape type
    Unknown(u16),
}

impl From<u16> for ShapeType {
    fn from(value: u16) -> Self {
        match value {
            1 => ShapeType::TextBox,
            2 => ShapeType::Placeholder,
            3 => ShapeType::AutoShape,
            4 => ShapeType::Picture,
            5 => ShapeType::Group,
            6 => ShapeType::Line,
            7 => ShapeType::Connector,
            8 => ShapeType::Object,
            9 => ShapeType::Table,
            other => ShapeType::Unknown(other),
        }
    }
}

impl From<crate::ole::consts::EscherShapeType> for ShapeType {
    fn from(escher_type: crate::ole::consts::EscherShapeType) -> Self {
        match escher_type {
            crate::ole::consts::EscherShapeType::NotPrimitive => ShapeType::Unknown(0),
            crate::ole::consts::EscherShapeType::Rectangle => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::RoundRectangle => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Oval => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Diamond => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Triangle => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::RightTriangle => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Parallelogram => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Trapezoid => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Hexagon => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Octagon => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Plus => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Star => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Arrow => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::ThickArrow => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::HomePlate => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Cube => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Balloon => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Seal => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Arc => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Line => ShapeType::Line,
            crate::ole::consts::EscherShapeType::Plaque => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Can => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Donut => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::TextSimple => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextOctagon => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextHexagon => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextCurve => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextWave => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextRing => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextOnCurve => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::TextOnRing => ShapeType::TextBox,
            crate::ole::consts::EscherShapeType::StraightConnector1 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::BentConnector2 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::BentConnector3 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::BentConnector4 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::BentConnector5 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::CurvedConnector2 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::CurvedConnector3 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::CurvedConnector4 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::CurvedConnector5 => ShapeType::Connector,
            crate::ole::consts::EscherShapeType::Callout1 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Callout2 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Callout3 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentCallout1 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentCallout2 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentCallout3 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::BorderCallout1 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::BorderCallout2 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::BorderCallout3 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentBorderCallout1 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentBorderCallout2 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::AccentBorderCallout3 => ShapeType::AutoShape,
            crate::ole::consts::EscherShapeType::Custom => ShapeType::AutoShape,
        }
    }
}

impl fmt::Display for ShapeType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ShapeType::TextBox => write!(f, "TextBox"),
            ShapeType::Placeholder => write!(f, "Placeholder"),
            ShapeType::AutoShape => write!(f, "AutoShape"),
            ShapeType::Picture => write!(f, "Picture"),
            ShapeType::Group => write!(f, "Group"),
            ShapeType::Line => write!(f, "Line"),
            ShapeType::Connector => write!(f, "Connector"),
            ShapeType::Object => write!(f, "Object"),
            ShapeType::Table => write!(f, "Table"),
            ShapeType::Unknown(id) => write!(f, "Unknown({})", id),
        }
    }
}

/// Common properties shared by all shape types.
#[derive(Debug, Clone)]
pub struct ShapeProperties {
    /// Shape ID
    pub id: u32,
    /// Shape type
    pub shape_type: ShapeType,
    /// X position in EMUs (English Metric Units)
    pub x: i32,
    /// Y position in EMUs
    pub y: i32,
    /// Width in EMUs
    pub width: i32,
    /// Height in EMUs
    pub height: i32,
    /// Rotation angle (0-360 degrees)
    pub rotation: u16,
    /// Fill color (RGB)
    pub fill_color: Option<u32>,
    /// Line color (RGB)
    pub line_color: Option<u32>,
    /// Line width in points
    pub line_width: Option<u16>,
    /// Is the shape hidden?
    pub hidden: bool,
    /// Z-order (drawing order)
    pub z_order: u16,
}

impl Default for ShapeProperties {
    fn default() -> Self {
        Self {
            id: 0,
            shape_type: ShapeType::Unknown(0),
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            rotation: 0,
            fill_color: None,
            line_color: None,
            line_width: None,
            hidden: false,
            z_order: 0,
        }
    }
}

/// Base trait for all shape types in PowerPoint presentations.
///
/// This trait defines the common interface that all shape implementations
/// must provide, including access to properties and basic operations.
pub trait Shape: std::any::Any {
    /// Get the shape's properties.
    fn properties(&self) -> &ShapeProperties;

    /// Get the shape's properties as mutable reference.
    fn properties_mut(&mut self) -> &mut ShapeProperties;

    /// Get the shape type.
    fn shape_type(&self) -> ShapeType {
        self.properties().shape_type
    }

    /// Get the shape ID.
    fn id(&self) -> u32 {
        self.properties().id
    }

    /// Get the shape's text content, if any.
    fn text(&self) -> Result<String, PptError>;

    /// Get the shape's position and size.
    fn bounds(&self) -> (i32, i32, i32, i32) {
        let props = self.properties();
        (props.x, props.y, props.width, props.height)
    }

    /// Check if the shape is a placeholder.
    fn is_placeholder(&self) -> bool {
        matches!(self.shape_type(), ShapeType::Placeholder)
    }

    /// Check if the shape has text content.
    fn has_text(&self) -> bool;

    /// Clone the shape as a boxed trait object.
    fn clone_box(&self) -> Box<dyn Shape>;

    /// Get the shape as an Any reference for downcasting.
    fn as_any(&self) -> &dyn std::any::Any;
}

impl Clone for Box<dyn Shape> {
    fn clone(&self) -> Self {
        self.clone_box()
    }
}

/// Shape container that holds shape data and provides common operations.
///
/// This container includes Escher text properties extracted from the shape's
/// OfficeArtFOPT records, following Apache POI's EscherProperties model.
#[derive(Clone)]
pub struct ShapeContainer {
    /// Shape properties
    pub properties: ShapeProperties,
    /// Raw shape data (for parsing)
    pub raw_data: Vec<u8>,
    /// Text content (if applicable)
    pub text_content: Option<String>,
    /// Child shapes (for group shapes)
    pub children: Vec<Box<dyn Shape>>,

    // Escher text properties (from OfficeArtFOPT records)
    /// Text left margin in master units (1/576 inch)
    /// Property ID: 0x0081 (TEXT_LEFT)
    pub text_left: Option<i32>,

    /// Text top margin in master units (1/576 inch)
    /// Property ID: 0x0082 (TEXT_TOP)
    pub text_top: Option<i32>,

    /// Text right margin in master units (1/576 inch)
    /// Property ID: 0x0083 (TEXT_RIGHT)
    pub text_right: Option<i32>,

    /// Text bottom margin in master units (1/576 inch)
    /// Property ID: 0x0084 (TEXT_BOTTOM)
    pub text_bottom: Option<i32>,

    /// Text flow direction
    /// Property ID: 0x0085 (TEXT_FLOW)
    /// Values: 0=horizontal, 1=vertical, 2=vertical rotated, 3=word art vertical
    pub text_flow: Option<u16>,

    /// Wrap text in text box
    /// Property ID: 0x0086 (WRAP_TEXT)
    pub wrap_text: Option<bool>,

    /// Text anchor (vertical alignment)
    /// Property ID: 0x0087 (ANCHOR_TEXT)
    /// Values: 0=top, 1=middle, 2=bottom, 3=top centered, 4=middle centered,
    ///         5=bottom centered, 6=top baseline, 7=bottom baseline, 8=top centered baseline
    pub anchor_text: Option<u16>,

    /// Rotate text with shape
    /// Property ID: 0x00BF (ROTATE_TEXT)
    pub rotate_text: Option<bool>,

    /// Text ID (identifier for the text)
    /// Property ID: 0x0080 (TEXT_ID)
    pub text_id: Option<u32>,

    /// Scale text to fit shape
    /// Property ID: 0x0089 (SCALE_TEXT)
    pub scale_text: Option<bool>,

    /// Size text to fit shape bounds
    /// Property ID: 0x008A (SIZE_TEXT_TO_FIT_SHAPE)
    pub size_text_to_fit_shape: Option<bool>,

    /// Size shape to fit text content
    /// Property ID: 0x008B (SIZE_SHAPE_TO_FIT_TEXT)
    pub size_shape_to_fit_text: Option<bool>,

    /// Font rotation angle (16.16 fixed-point degrees)
    /// Property ID: 0x008D (FONT_ROTATION)
    pub font_rotation: Option<u32>,

    /// Bidirectional text flag
    /// Property ID: 0x0088 (BIDI)
    pub bidi: Option<bool>,

    /// Use host margins (use container's margins)
    /// Property ID: 0x008E (USE_HOST_MARGINS)
    pub use_host_margins: Option<bool>,

    /// Single click selects text
    /// Property ID: 0x008F (SINGLE_CLICK_SELECTS)
    pub single_click_selects: Option<bool>,

    /// ID of next shape in sequence
    /// Property ID: 0x0082 (ID_OF_NEXT_SHAPE) - Note: different context than TEXT_TOP
    pub id_of_next_shape: Option<u32>,
}

impl std::fmt::Debug for ShapeContainer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut debug = f.debug_struct("ShapeContainer");
        debug
            .field("properties", &self.properties)
            .field("raw_data_len", &self.raw_data.len())
            .field("text_content", &self.text_content)
            .field("children_count", &self.children.len());

        // Only show Escher text properties if they're set
        if self.text_left.is_some()
            || self.text_top.is_some()
            || self.text_right.is_some()
            || self.text_bottom.is_some()
        {
            debug.field(
                "text_margins",
                &format_args!(
                    "L:{:?} T:{:?} R:{:?} B:{:?}",
                    self.text_left, self.text_top, self.text_right, self.text_bottom
                ),
            );
        }
        if let Some(flow) = self.text_flow {
            debug.field("text_flow", &flow);
        }
        if let Some(anchor) = self.anchor_text {
            debug.field("anchor_text", &anchor);
        }
        if let Some(wrap) = self.wrap_text {
            debug.field("wrap_text", &wrap);
        }

        debug.finish()
    }
}

impl ShapeContainer {
    /// Create a new shape container.
    ///
    /// All Escher text properties are initialized to `None` and can be
    /// populated later by parsing OfficeArtFOPT records.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            properties,
            raw_data,
            text_content: None,
            children: Vec::new(),
            // Initialize all Escher text properties to None
            text_left: None,
            text_top: None,
            text_right: None,
            text_bottom: None,
            text_flow: None,
            wrap_text: None,
            anchor_text: None,
            rotate_text: None,
            text_id: None,
            scale_text: None,
            size_text_to_fit_shape: None,
            size_shape_to_fit_text: None,
            font_rotation: None,
            bidi: None,
            use_host_margins: None,
            single_click_selects: None,
            id_of_next_shape: None,
        }
    }

    /// Add a child shape to this shape (for group shapes).
    pub fn add_child(&mut self, shape: Box<dyn Shape>) {
        self.children.push(shape);
    }

    /// Get all child shapes.
    pub fn children(&self) -> &[Box<dyn Shape>] {
        &self.children
    }

    /// Set the text content of this shape.
    pub fn set_text(&mut self, text: String) {
        self.text_content = Some(text);
    }

    /// Set text margins from a 4-value tuple (left, top, right, bottom).
    ///
    /// # Arguments
    ///
    /// * `margins` - Tuple of (left, top, right, bottom) margins in master units
    ///
    /// # Example
    ///
    /// ```ignore
    /// container.set_text_margins(Some((91440, 45720, 91440, 45720)));
    /// ```
    pub fn set_text_margins(&mut self, margins: Option<(i32, i32, i32, i32)>) {
        if let Some((left, top, right, bottom)) = margins {
            self.text_left = Some(left);
            self.text_top = Some(top);
            self.text_right = Some(right);
            self.text_bottom = Some(bottom);
        }
    }

    /// Get text margins as a 4-value tuple.
    ///
    /// # Returns
    ///
    /// `Some((left, top, right, bottom))` if all four margins are set, `None` otherwise
    pub fn text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        match (
            self.text_left,
            self.text_top,
            self.text_right,
            self.text_bottom,
        ) {
            (Some(l), Some(t), Some(r), Some(b)) => Some((l, t, r, b)),
            _ => None,
        }
    }

    /// Set text flow direction.
    ///
    /// # Values
    ///
    /// - 0: Horizontal (left to right)
    /// - 1: Vertical (top to bottom)
    /// - 2: Vertical rotated
    /// - 3: Word art vertical
    pub fn set_text_flow(&mut self, flow: Option<u16>) {
        self.text_flow = flow;
    }

    /// Set text anchor (vertical alignment).
    ///
    /// # Values
    ///
    /// - 0: Top
    /// - 1: Middle
    /// - 2: Bottom
    /// - 3: Top centered
    /// - 4: Middle centered
    /// - 5: Bottom centered
    /// - 6: Top baseline
    /// - 7: Bottom baseline
    /// - 8: Top centered baseline
    pub fn set_anchor_text(&mut self, anchor: Option<u16>) {
        self.anchor_text = anchor;
    }

    /// Set wrap text flag.
    pub fn set_wrap_text(&mut self, wrap: Option<bool>) {
        self.wrap_text = wrap;
    }

    /// Set rotate text with shape flag.
    pub fn set_rotate_text(&mut self, rotate: Option<bool>) {
        self.rotate_text = rotate;
    }

    /// Set font rotation angle (16.16 fixed-point degrees).
    pub fn set_font_rotation(&mut self, rotation: Option<u32>) {
        self.font_rotation = rotation;
    }

    /// Get font rotation in degrees as a float.
    ///
    /// Converts from 16.16 fixed-point to f32.
    pub fn font_rotation_degrees(&self) -> Option<f32> {
        self.font_rotation
            .map(|rot| (rot >> 16) as f32 + ((rot & 0xFFFF) as f32 / 65536.0))
    }
}

impl Shape for ShapeContainer {
    fn properties(&self) -> &ShapeProperties {
        &self.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.properties
    }

    fn text(&self) -> Result<String, PptError> {
        Ok(self.text_content.clone().unwrap_or_default())
    }

    fn has_text(&self) -> bool {
        self.text_content.is_some()
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}
