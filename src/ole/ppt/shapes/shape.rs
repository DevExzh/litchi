/// Base shape trait and common shape functionality.
///
/// This module defines the core Shape trait that all shape types implement,
/// along with common properties and methods for working with shapes.
use std::fmt;
use super::super::package::PptError;

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
}

impl std::fmt::Debug for ShapeContainer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ShapeContainer")
            .field("properties", &self.properties)
            .field("raw_data", &self.raw_data.len())
            .field("text_content", &self.text_content)
            .field("children", &self.children.len())
            .finish()
    }
}

impl ShapeContainer {
    /// Create a new shape container.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            properties,
            raw_data,
            text_content: None,
            children: Vec::new(),
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

