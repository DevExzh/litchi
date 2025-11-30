//! Extended shape types for PPT files
//!
//! This module provides comprehensive shape type support including
//! basic shapes, lines, arrows, connectors, and more.
//!
//! Reference: [MS-ODRAW] Section 2.4.6 - MSOSPT

use super::shape_style::{ArrowStyle, FillStyle, LineStyleConfig, ShapeColor, ShapeStyle};
use super::text_format::Paragraph;

// =============================================================================
// Shape Type Constants (MS-ODRAW 2.4.6 MSOSPT)
// =============================================================================

/// Escher shape type values
pub mod shape_type {
    /// Not a primitive shape
    pub const NOT_PRIMITIVE: u16 = 0;
    /// Rectangle
    pub const RECTANGLE: u16 = 1;
    /// Round rectangle
    pub const ROUND_RECTANGLE: u16 = 2;
    /// Ellipse
    pub const ELLIPSE: u16 = 3;
    /// Diamond
    pub const DIAMOND: u16 = 4;
    /// Isoceles triangle
    pub const ISOCELES_TRIANGLE: u16 = 5;
    /// Right triangle
    pub const RIGHT_TRIANGLE: u16 = 6;
    /// Parallelogram
    pub const PARALLELOGRAM: u16 = 7;
    /// Trapezoid
    pub const TRAPEZOID: u16 = 8;
    /// Hexagon
    pub const HEXAGON: u16 = 9;
    /// Octagon
    pub const OCTAGON: u16 = 10;
    /// Plus sign
    pub const PLUS: u16 = 11;
    /// 5-pointed star
    pub const STAR: u16 = 12;
    /// Arrow right
    pub const ARROW: u16 = 13;
    /// Thick arrow right
    pub const THICK_ARROW: u16 = 14;
    /// Home plate / pentagon
    pub const HOME_PLATE: u16 = 15;
    /// Cube
    pub const CUBE: u16 = 16;
    /// Balloon
    pub const BALLOON: u16 = 17;
    /// Seal (burst)
    pub const SEAL: u16 = 18;
    /// Arc
    pub const ARC: u16 = 19;
    /// Line
    pub const LINE: u16 = 20;
    /// Plaque
    pub const PLAQUE: u16 = 21;
    /// Can (cylinder)
    pub const CAN: u16 = 22;
    /// Donut
    pub const DONUT: u16 = 23;
    /// Straight connector
    pub const STRAIGHT_CONNECTOR: u16 = 32;
    /// Bent connector 2
    pub const BENT_CONNECTOR_2: u16 = 33;
    /// Bent connector 3
    pub const BENT_CONNECTOR_3: u16 = 34;
    /// Bent connector 4
    pub const BENT_CONNECTOR_4: u16 = 35;
    /// Bent connector 5
    pub const BENT_CONNECTOR_5: u16 = 36;
    /// Curved connector 2
    pub const CURVED_CONNECTOR_2: u16 = 37;
    /// Curved connector 3
    pub const CURVED_CONNECTOR_3: u16 = 38;
    /// Curved connector 4
    pub const CURVED_CONNECTOR_4: u16 = 39;
    /// Curved connector 5
    pub const CURVED_CONNECTOR_5: u16 = 40;
    /// Callout 1
    pub const CALLOUT_1: u16 = 41;
    /// Callout 2
    pub const CALLOUT_2: u16 = 42;
    /// Callout 3
    pub const CALLOUT_3: u16 = 43;
    /// Accent callout 1
    pub const ACCENT_CALLOUT_1: u16 = 44;
    /// Accent callout 2
    pub const ACCENT_CALLOUT_2: u16 = 45;
    /// Accent callout 3
    pub const ACCENT_CALLOUT_3: u16 = 46;
    /// Border callout 1
    pub const BORDER_CALLOUT_1: u16 = 47;
    /// Border callout 2
    pub const BORDER_CALLOUT_2: u16 = 48;
    /// Border callout 3
    pub const BORDER_CALLOUT_3: u16 = 49;
    /// Left arrow
    pub const LEFT_ARROW: u16 = 66;
    /// Up arrow
    pub const UP_ARROW: u16 = 67;
    /// Down arrow
    pub const DOWN_ARROW: u16 = 68;
    /// Left-right arrow
    pub const LEFT_RIGHT_ARROW: u16 = 69;
    /// Up-down arrow
    pub const UP_DOWN_ARROW: u16 = 70;
    /// Irregular seal 1
    pub const IRREGULAR_SEAL_1: u16 = 71;
    /// Irregular seal 2
    pub const IRREGULAR_SEAL_2: u16 = 72;
    /// Lightning bolt
    pub const LIGHTNING_BOLT: u16 = 73;
    /// Heart
    pub const HEART: u16 = 74;
    /// Quad arrow
    pub const QUAD_ARROW: u16 = 76;
    /// Left arrow callout
    pub const LEFT_ARROW_CALLOUT: u16 = 77;
    /// Right arrow callout
    pub const RIGHT_ARROW_CALLOUT: u16 = 78;
    /// Up arrow callout
    pub const UP_ARROW_CALLOUT: u16 = 79;
    /// Down arrow callout
    pub const DOWN_ARROW_CALLOUT: u16 = 80;
    /// Striped right arrow
    pub const STRIPED_RIGHT_ARROW: u16 = 93;
    /// Notched right arrow
    pub const NOTCHED_RIGHT_ARROW: u16 = 94;
    /// Block arc
    pub const BLOCK_ARC: u16 = 95;
    /// Smiley face
    pub const SMILEY_FACE: u16 = 96;
    /// Vertical scroll
    pub const VERTICAL_SCROLL: u16 = 97;
    /// Horizontal scroll
    pub const HORIZONTAL_SCROLL: u16 = 98;
    /// Circular arrow
    pub const CIRCULAR_ARROW: u16 = 99;
    /// Uturn arrow
    pub const UTURN_ARROW: u16 = 101;
    /// Curved right arrow
    pub const CURVED_RIGHT_ARROW: u16 = 102;
    /// Curved left arrow
    pub const CURVED_LEFT_ARROW: u16 = 103;
    /// Curved up arrow
    pub const CURVED_UP_ARROW: u16 = 104;
    /// Curved down arrow
    pub const CURVED_DOWN_ARROW: u16 = 105;
    /// Cloud callout
    pub const CLOUD_CALLOUT: u16 = 106;
    /// Ellipse ribbon
    pub const ELLIPSE_RIBBON: u16 = 107;
    /// Ellipse ribbon 2
    pub const ELLIPSE_RIBBON_2: u16 = 108;
    /// Flowchart process
    pub const FLOWCHART_PROCESS: u16 = 109;
    /// Flowchart decision
    pub const FLOWCHART_DECISION: u16 = 110;
    /// Flowchart input/output
    pub const FLOWCHART_INPUT_OUTPUT: u16 = 111;
    /// Flowchart predefined process
    pub const FLOWCHART_PREDEFINED_PROCESS: u16 = 112;
    /// Flowchart internal storage
    pub const FLOWCHART_INTERNAL_STORAGE: u16 = 113;
    /// Flowchart document
    pub const FLOWCHART_DOCUMENT: u16 = 114;
    /// Flowchart multi-document
    pub const FLOWCHART_MULTI_DOCUMENT: u16 = 115;
    /// Flowchart terminator
    pub const FLOWCHART_TERMINATOR: u16 = 116;
    /// Flowchart preparation
    pub const FLOWCHART_PREPARATION: u16 = 117;
    /// Flowchart manual input
    pub const FLOWCHART_MANUAL_INPUT: u16 = 118;
    /// Flowchart manual operation
    pub const FLOWCHART_MANUAL_OPERATION: u16 = 119;
    /// Flowchart connector
    pub const FLOWCHART_CONNECTOR: u16 = 120;
    /// Flowchart off-page connector
    pub const FLOWCHART_OFFPAGE_CONNECTOR: u16 = 177;
    /// Action button blank
    pub const ACTION_BUTTON_BLANK: u16 = 189;
    /// Action button home
    pub const ACTION_BUTTON_HOME: u16 = 190;
    /// Action button help
    pub const ACTION_BUTTON_HELP: u16 = 191;
    /// Action button information
    pub const ACTION_BUTTON_INFORMATION: u16 = 192;
    /// Action button forward/next
    pub const ACTION_BUTTON_FORWARD_NEXT: u16 = 193;
    /// Action button back/previous
    pub const ACTION_BUTTON_BACK_PREVIOUS: u16 = 194;
    /// Action button end
    pub const ACTION_BUTTON_END: u16 = 195;
    /// Action button beginning
    pub const ACTION_BUTTON_BEGINNING: u16 = 196;
    /// Action button return
    pub const ACTION_BUTTON_RETURN: u16 = 197;
    /// Action button document
    pub const ACTION_BUTTON_DOCUMENT: u16 = 198;
    /// Action button sound
    pub const ACTION_BUTTON_SOUND: u16 = 199;
    /// Action button movie
    pub const ACTION_BUTTON_MOVIE: u16 = 200;
    /// Text box
    pub const TEXT_BOX: u16 = 202;
    /// Picture frame
    pub const PICTURE_FRAME: u16 = 75;
}

// =============================================================================
// User-Friendly Shape Types
// =============================================================================

/// High-level shape type enumeration
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeKind {
    // Basic shapes
    /// Rectangle
    Rectangle,
    /// Rounded rectangle
    RoundRectangle,
    /// Ellipse (oval)
    Ellipse,
    /// Diamond
    Diamond,
    /// Triangle
    Triangle,
    /// Right triangle
    RightTriangle,
    /// Parallelogram
    Parallelogram,
    /// Trapezoid
    Trapezoid,
    /// Hexagon
    Hexagon,
    /// Octagon
    Octagon,
    /// Plus sign
    Plus,
    /// Star (5-pointed)
    Star,

    // Lines
    /// Straight line
    Line,
    /// Straight connector
    Connector,
    /// Curved connector
    CurvedConnector,
    /// Bent connector
    BentConnector,

    // Arrows
    /// Right arrow
    ArrowRight,
    /// Left arrow
    ArrowLeft,
    /// Up arrow
    ArrowUp,
    /// Down arrow
    ArrowDown,
    /// Left-right arrow
    ArrowLeftRight,
    /// Up-down arrow
    ArrowUpDown,
    /// Quad arrow (4-way)
    ArrowQuad,
    /// Circular arrow
    CircularArrow,
    /// U-turn arrow
    UturnArrow,
    /// Striped arrow
    StripedArrow,
    /// Notched arrow
    NotchedArrow,

    // Callouts
    /// Rectangular callout
    CalloutRect,
    /// Rounded callout
    CalloutRound,
    /// Oval callout
    CalloutOval,
    /// Cloud callout
    CalloutCloud,

    // Flowchart
    /// Flowchart process
    FlowchartProcess,
    /// Flowchart decision
    FlowchartDecision,
    /// Flowchart terminator
    FlowchartTerminator,
    /// Flowchart document
    FlowchartDocument,
    /// Flowchart connector
    FlowchartConnector,

    // Special
    /// Text box
    TextBox,
    /// Picture frame
    PictureFrame,
    /// Heart
    Heart,
    /// Lightning bolt
    LightningBolt,
    /// Smiley face
    SmileyFace,
    /// Donut
    Donut,
    /// Arc
    Arc,
    /// Cube
    Cube,
    /// Can (cylinder)
    Can,
}

impl ShapeKind {
    /// Convert to Escher shape type code
    pub fn to_escher_type(&self) -> u16 {
        match self {
            ShapeKind::Rectangle => shape_type::RECTANGLE,
            ShapeKind::RoundRectangle => shape_type::ROUND_RECTANGLE,
            ShapeKind::Ellipse => shape_type::ELLIPSE,
            ShapeKind::Diamond => shape_type::DIAMOND,
            ShapeKind::Triangle => shape_type::ISOCELES_TRIANGLE,
            ShapeKind::RightTriangle => shape_type::RIGHT_TRIANGLE,
            ShapeKind::Parallelogram => shape_type::PARALLELOGRAM,
            ShapeKind::Trapezoid => shape_type::TRAPEZOID,
            ShapeKind::Hexagon => shape_type::HEXAGON,
            ShapeKind::Octagon => shape_type::OCTAGON,
            ShapeKind::Plus => shape_type::PLUS,
            ShapeKind::Star => shape_type::STAR,
            ShapeKind::Line => shape_type::LINE,
            ShapeKind::Connector => shape_type::STRAIGHT_CONNECTOR,
            ShapeKind::CurvedConnector => shape_type::CURVED_CONNECTOR_3,
            ShapeKind::BentConnector => shape_type::BENT_CONNECTOR_3,
            ShapeKind::ArrowRight => shape_type::ARROW,
            ShapeKind::ArrowLeft => shape_type::LEFT_ARROW,
            ShapeKind::ArrowUp => shape_type::UP_ARROW,
            ShapeKind::ArrowDown => shape_type::DOWN_ARROW,
            ShapeKind::ArrowLeftRight => shape_type::LEFT_RIGHT_ARROW,
            ShapeKind::ArrowUpDown => shape_type::UP_DOWN_ARROW,
            ShapeKind::ArrowQuad => shape_type::QUAD_ARROW,
            ShapeKind::CircularArrow => shape_type::CIRCULAR_ARROW,
            ShapeKind::UturnArrow => shape_type::UTURN_ARROW,
            ShapeKind::StripedArrow => shape_type::STRIPED_RIGHT_ARROW,
            ShapeKind::NotchedArrow => shape_type::NOTCHED_RIGHT_ARROW,
            ShapeKind::CalloutRect => shape_type::CALLOUT_1,
            ShapeKind::CalloutRound => shape_type::CALLOUT_2,
            ShapeKind::CalloutOval => shape_type::CALLOUT_3,
            ShapeKind::CalloutCloud => shape_type::CLOUD_CALLOUT,
            ShapeKind::FlowchartProcess => shape_type::FLOWCHART_PROCESS,
            ShapeKind::FlowchartDecision => shape_type::FLOWCHART_DECISION,
            ShapeKind::FlowchartTerminator => shape_type::FLOWCHART_TERMINATOR,
            ShapeKind::FlowchartDocument => shape_type::FLOWCHART_DOCUMENT,
            ShapeKind::FlowchartConnector => shape_type::FLOWCHART_CONNECTOR,
            ShapeKind::TextBox => shape_type::TEXT_BOX,
            ShapeKind::PictureFrame => shape_type::PICTURE_FRAME,
            ShapeKind::Heart => shape_type::HEART,
            ShapeKind::LightningBolt => shape_type::LIGHTNING_BOLT,
            ShapeKind::SmileyFace => shape_type::SMILEY_FACE,
            ShapeKind::Donut => shape_type::DONUT,
            ShapeKind::Arc => shape_type::ARC,
            ShapeKind::Cube => shape_type::CUBE,
            ShapeKind::Can => shape_type::CAN,
        }
    }

    /// Check if this shape type is a line/connector (no fill by default)
    pub fn is_line(&self) -> bool {
        matches!(
            self,
            ShapeKind::Line
                | ShapeKind::Connector
                | ShapeKind::CurvedConnector
                | ShapeKind::BentConnector
        )
    }

    /// Check if this shape type is an arrow
    pub fn is_arrow(&self) -> bool {
        matches!(
            self,
            ShapeKind::ArrowRight
                | ShapeKind::ArrowLeft
                | ShapeKind::ArrowUp
                | ShapeKind::ArrowDown
                | ShapeKind::ArrowLeftRight
                | ShapeKind::ArrowUpDown
                | ShapeKind::ArrowQuad
                | ShapeKind::CircularArrow
                | ShapeKind::UturnArrow
                | ShapeKind::StripedArrow
                | ShapeKind::NotchedArrow
        )
    }

    /// Check if this shape can contain text
    pub fn can_contain_text(&self) -> bool {
        !self.is_line()
    }
}

// =============================================================================
// Shape Definition
// =============================================================================

/// Complete shape definition with geometry and styling
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape kind
    pub kind: ShapeKind,
    /// X position in EMUs
    pub x: i32,
    /// Y position in EMUs
    pub y: i32,
    /// Width in EMUs
    pub width: i32,
    /// Height in EMUs
    pub height: i32,
    /// Visual style (fill, line, shadow)
    pub style: ShapeStyle,
    /// Text content (if applicable)
    pub text: Option<Vec<Paragraph>>,
    /// Rotation angle in degrees
    pub rotation: f32,
    /// Flip horizontal
    pub flip_h: bool,
    /// Flip vertical
    pub flip_v: bool,
    /// Adjust values (shape-specific parameters)
    pub adjust_values: Vec<i32>,
    /// Associated picture BLIP index (for picture frames)
    pub picture_index: Option<u32>,
}

impl Shape {
    /// Create a new shape
    pub fn new(kind: ShapeKind, x: i32, y: i32, width: i32, height: i32) -> Self {
        let style = if kind.is_line() {
            ShapeStyle::no_fill()
        } else {
            ShapeStyle::default()
        };

        Self {
            kind,
            x,
            y,
            width,
            height,
            style,
            text: None,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
            adjust_values: Vec::new(),
            picture_index: None,
        }
    }

    /// Create rectangle
    pub fn rectangle(x: i32, y: i32, width: i32, height: i32) -> Self {
        Self::new(ShapeKind::Rectangle, x, y, width, height)
    }

    /// Create ellipse
    pub fn ellipse(x: i32, y: i32, width: i32, height: i32) -> Self {
        Self::new(ShapeKind::Ellipse, x, y, width, height)
    }

    /// Create line from (x1,y1) to (x2,y2)
    pub fn line(x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        let x = x1.min(x2);
        let y = y1.min(y2);
        let width = (x2 - x1).abs();
        let height = (y2 - y1).abs();

        let mut shape = Self::new(ShapeKind::Line, x, y, width, height);

        // Set flip flags if needed
        shape.flip_h = x2 < x1;
        shape.flip_v = y2 < y1;

        shape
    }

    /// Create arrow from (x1,y1) to (x2,y2)
    pub fn arrow(x1: i32, y1: i32, x2: i32, y2: i32) -> Self {
        let mut shape = Self::line(x1, y1, x2, y2);
        // Add end arrow
        shape.style.line = shape.style.line.end_arrow(ArrowStyle::Triangle);
        shape
    }

    /// Create text box
    pub fn text_box(x: i32, y: i32, width: i32, height: i32, text: &str) -> Self {
        let mut shape = Self::new(ShapeKind::TextBox, x, y, width, height);
        shape.text = Some(vec![Paragraph::new(text)]);
        shape.style.fill = FillStyle::none();
        shape
    }

    /// Create picture frame
    pub fn picture(x: i32, y: i32, width: i32, height: i32, blip_index: u32) -> Self {
        let mut shape = Self::new(ShapeKind::PictureFrame, x, y, width, height);
        shape.picture_index = Some(blip_index);
        shape.style.fill = FillStyle::picture(blip_index);
        shape
    }

    /// Set fill style
    pub fn with_fill(mut self, fill: FillStyle) -> Self {
        self.style.fill = fill;
        self
    }

    /// Set solid fill color
    pub fn with_fill_color(mut self, color: ShapeColor) -> Self {
        self.style.fill = FillStyle::solid(color);
        self
    }

    /// Set fill color from RGB
    pub fn with_fill_rgb(mut self, r: u8, g: u8, b: u8) -> Self {
        self.style.fill = FillStyle::solid_rgb(r, g, b);
        self
    }

    /// Set no fill
    pub fn with_no_fill(mut self) -> Self {
        self.style.fill = FillStyle::none();
        self
    }

    /// Set line style
    pub fn with_line(mut self, line: LineStyleConfig) -> Self {
        self.style.line = line;
        self
    }

    /// Set line color and width
    pub fn with_line_color(mut self, color: ShapeColor, width_pt: f32) -> Self {
        self.style.line = LineStyleConfig::with_color_and_width(color, width_pt);
        self
    }

    /// Set no line
    pub fn with_no_line(mut self) -> Self {
        self.style.line = LineStyleConfig::none();
        self
    }

    /// Set text content
    pub fn with_text(mut self, paragraphs: Vec<Paragraph>) -> Self {
        self.text = Some(paragraphs);
        self
    }

    /// Set rotation
    pub fn with_rotation(mut self, degrees: f32) -> Self {
        self.rotation = degrees;
        self
    }

    /// Set flip
    pub fn with_flip(mut self, horizontal: bool, vertical: bool) -> Self {
        self.flip_h = horizontal;
        self.flip_v = vertical;
        self
    }

    /// Set adjust value (shape-specific parameter)
    pub fn with_adjust(mut self, index: usize, value: i32) -> Self {
        if index >= self.adjust_values.len() {
            self.adjust_values.resize(index + 1, 0);
        }
        self.adjust_values[index] = value;
        self
    }

    /// Build the Escher properties for this shape
    pub fn build_escher_properties(&self) -> Vec<(u16, u32)> {
        let mut props = self.style.build_properties();

        // Add rotation if non-zero
        if self.rotation.abs() > 0.001 {
            // Rotation is in 1/65536th of a degree
            let rot_value = ((self.rotation * 65536.0) as i32) as u32;
            props.push((0x0004, rot_value)); // rotation
        }

        // Add adjust values
        for (i, &value) in self.adjust_values.iter().enumerate() {
            let prop_id = 0x0080 + i as u16; // adjustValue, adjust2Value, etc.
            props.push((prop_id, value as u32));
        }

        // Add picture BLIP reference if present
        if let Some(blip_idx) = self.picture_index {
            props.push((0x4104, blip_idx)); // pib
        }

        props
    }

    /// Get shape flags
    pub fn get_shape_flags(&self) -> u32 {
        let mut flags = 0u32;

        // fHaveAnchor (0x0200)
        flags |= 0x0200;

        // fHaveSpt (0x0800)
        flags |= 0x0800;

        // fFlipH (0x0040)
        if self.flip_h {
            flags |= 0x0040;
        }

        // fFlipV (0x0080)
        if self.flip_v {
            flags |= 0x0080;
        }

        flags
    }

    /// Get bounds as (left, top, right, bottom)
    pub fn bounds(&self) -> (i32, i32, i32, i32) {
        (self.x, self.y, self.x + self.width, self.y + self.height)
    }
}

// =============================================================================
// Shape Collection
// =============================================================================

/// Collection of shapes for a slide
#[derive(Debug, Clone, Default)]
pub struct ShapeCollection {
    shapes: Vec<Shape>,
}

impl ShapeCollection {
    /// Create new empty collection
    pub fn new() -> Self {
        Self { shapes: Vec::new() }
    }

    /// Add a shape
    pub fn add(&mut self, shape: Shape) -> usize {
        let idx = self.shapes.len();
        self.shapes.push(shape);
        idx
    }

    /// Get shape by index
    pub fn get(&self, index: usize) -> Option<&Shape> {
        self.shapes.get(index)
    }

    /// Get mutable shape by index
    pub fn get_mut(&mut self, index: usize) -> Option<&mut Shape> {
        self.shapes.get_mut(index)
    }

    /// Get number of shapes
    pub fn len(&self) -> usize {
        self.shapes.len()
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.shapes.is_empty()
    }

    /// Iterate over shapes
    pub fn iter(&self) -> impl Iterator<Item = &Shape> {
        self.shapes.iter()
    }

    /// Iterate mutably over shapes
    pub fn iter_mut(&mut self) -> impl Iterator<Item = &mut Shape> {
        self.shapes.iter_mut()
    }

    /// Calculate bounding box of all shapes
    pub fn bounds(&self) -> Option<(i32, i32, i32, i32)> {
        if self.shapes.is_empty() {
            return None;
        }

        let mut min_x = i32::MAX;
        let mut min_y = i32::MAX;
        let mut max_x = i32::MIN;
        let mut max_y = i32::MIN;

        for shape in &self.shapes {
            let (x1, y1, x2, y2) = shape.bounds();
            min_x = min_x.min(x1);
            min_y = min_y.min(y1);
            max_x = max_x.max(x2);
            max_y = max_y.max(y2);
        }

        Some((min_x, min_y, max_x, max_y))
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shape_kind_escher_type() {
        assert_eq!(ShapeKind::Rectangle.to_escher_type(), shape_type::RECTANGLE);
        assert_eq!(ShapeKind::Ellipse.to_escher_type(), shape_type::ELLIPSE);
        assert_eq!(ShapeKind::Line.to_escher_type(), shape_type::LINE);
    }

    #[test]
    fn test_shape_creation() {
        let rect = Shape::rectangle(100, 200, 300, 400);
        assert_eq!(rect.kind, ShapeKind::Rectangle);
        assert_eq!(rect.x, 100);
        assert_eq!(rect.width, 300);
    }

    #[test]
    fn test_line_creation() {
        let line = Shape::line(0, 0, 100, 100);
        assert_eq!(line.kind, ShapeKind::Line);
        assert!(!line.flip_h);
        assert!(!line.flip_v);

        // Line going backwards
        let line2 = Shape::line(100, 100, 0, 0);
        assert!(line2.flip_h);
        assert!(line2.flip_v);
    }

    #[test]
    fn test_shape_collection() {
        let mut shapes = ShapeCollection::new();
        shapes.add(Shape::rectangle(0, 0, 100, 100));
        shapes.add(Shape::ellipse(50, 50, 100, 100));

        assert_eq!(shapes.len(), 2);

        let bounds = shapes.bounds().unwrap();
        assert_eq!(bounds, (0, 0, 150, 150));
    }
}
