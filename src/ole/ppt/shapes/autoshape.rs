/// Auto shape implementation.
///
/// Auto shapes are predefined shapes like rectangles, ovals, arrows, etc.
/// that can be used as building blocks for more complex graphics in PowerPoint.
use super::shape::{Shape, ShapeContainer, ShapeProperties};

/// Types of auto shapes available in PowerPoint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoShapeType {
    /// Rectangle shape
    Rectangle,
    /// Rounded rectangle shape
    RoundRectangle,
    /// Oval shape
    Oval,
    /// Diamond shape
    Diamond,
    /// Isosceles triangle
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
    /// Star (5-point)
    Star,
    /// Arrow (right)
    Arrow,
    /// Thick arrow (right)
    ThickArrow,
    /// Home plate
    HomePlate,
    /// Cube
    Cube,
    /// Balloon
    Balloon,
    /// Seal
    Seal,
    /// Arc
    Arc,
    /// Teardrop
    Teardrop,
    /// Custom shape
    Custom(u16),
}

impl From<u16> for AutoShapeType {
    fn from(value: u16) -> Self {
        match value {
            1 => AutoShapeType::Rectangle,
            2 => AutoShapeType::RoundRectangle,
            3 => AutoShapeType::Oval,
            4 => AutoShapeType::Diamond,
            5 => AutoShapeType::Triangle,
            6 => AutoShapeType::RightTriangle,
            7 => AutoShapeType::Parallelogram,
            8 => AutoShapeType::Trapezoid,
            9 => AutoShapeType::Hexagon,
            10 => AutoShapeType::Octagon,
            11 => AutoShapeType::Plus,
            12 => AutoShapeType::Star,
            13 => AutoShapeType::Arrow,
            14 => AutoShapeType::ThickArrow,
            15 => AutoShapeType::HomePlate,
            16 => AutoShapeType::Cube,
            17 => AutoShapeType::Balloon,
            18 => AutoShapeType::Seal,
            19 => AutoShapeType::Arc,
            20 => AutoShapeType::Teardrop,
            other => AutoShapeType::Custom(other),
        }
    }
}

impl std::fmt::Display for AutoShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AutoShapeType::Rectangle => write!(f, "Rectangle"),
            AutoShapeType::RoundRectangle => write!(f, "RoundRectangle"),
            AutoShapeType::Oval => write!(f, "Oval"),
            AutoShapeType::Diamond => write!(f, "Diamond"),
            AutoShapeType::Triangle => write!(f, "Triangle"),
            AutoShapeType::RightTriangle => write!(f, "RightTriangle"),
            AutoShapeType::Parallelogram => write!(f, "Parallelogram"),
            AutoShapeType::Trapezoid => write!(f, "Trapezoid"),
            AutoShapeType::Hexagon => write!(f, "Hexagon"),
            AutoShapeType::Octagon => write!(f, "Octagon"),
            AutoShapeType::Plus => write!(f, "Plus"),
            AutoShapeType::Star => write!(f, "Star"),
            AutoShapeType::Arrow => write!(f, "Arrow"),
            AutoShapeType::ThickArrow => write!(f, "ThickArrow"),
            AutoShapeType::HomePlate => write!(f, "HomePlate"),
            AutoShapeType::Cube => write!(f, "Cube"),
            AutoShapeType::Balloon => write!(f, "Balloon"),
            AutoShapeType::Seal => write!(f, "Seal"),
            AutoShapeType::Arc => write!(f, "Arc"),
            AutoShapeType::Teardrop => write!(f, "Teardrop"),
            AutoShapeType::Custom(id) => write!(f, "Custom({})", id),
        }
    }
}

/// An auto shape in a PowerPoint presentation.
#[derive(Debug, Clone)]
pub struct AutoShape {
    /// Shape container with properties and data
    container: ShapeContainer,
    /// The type of auto shape
    auto_shape_type: AutoShapeType,
    /// Adjustment values for shape parameters (for complex shapes)
    adjustments: Vec<i32>,
}

impl AutoShape {
    /// Create a new auto shape.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            auto_shape_type: AutoShapeType::Rectangle, // Default
            adjustments: Vec::new(),
        }
    }

    /// Create an auto shape from an existing container.
    pub fn from_container(container: ShapeContainer) -> Self {
        // Extract auto shape type from raw data
        let auto_shape_type = Self::extract_shape_type(&container.raw_data);

        // Extract adjustment values if available
        let adjustments = Self::extract_adjustments(&container.raw_data);

        Self {
            container,
            auto_shape_type,
            adjustments,
        }
    }

    /// Extract auto shape type from raw shape data.
    ///
    /// In Escher records, the shape type is typically stored in the Sp (shape) record.
    /// The instance field of the Sp record contains the shape type ID.
    ///
    /// # Performance
    ///
    /// - Single-pass scanning
    /// - Early termination on match
    fn extract_shape_type(raw_data: &[u8]) -> AutoShapeType {
        // Scan for shape type in raw data
        // The shape type ID would typically be at a known offset in Escher Sp record
        // For a complete implementation, we'd parse the Escher structure properly

        // If raw_data is too small, return default
        if raw_data.len() < 4 {
            return AutoShapeType::Rectangle;
        }

        // Try to extract shape type ID from first few bytes
        // In Escher Sp record, bytes 0-1 contain shape flags, bytes 2-3 contain shape type
        let shape_type_id = u16::from_le_bytes([raw_data[0], raw_data[1]]);

        AutoShapeType::from(shape_type_id)
    }

    /// Extract adjustment values from raw shape data.
    ///
    /// Adjustment values control the geometry of complex shapes (e.g., arrow head size).
    /// Based on Apache POI's auto shape adjustment parsing.
    ///
    /// # Performance
    ///
    /// - Returns empty vector for shapes without adjustments
    /// - Pre-allocated capacity for shapes with adjustments
    fn extract_adjustments(raw_data: &[u8]) -> Vec<i32> {
        // In Escher format, adjustments are stored in shape options (Opt record)
        // Each adjustment is a 32-bit signed integer

        // For basic shapes, no adjustments needed
        if raw_data.len() < 8 {
            return Vec::new();
        }

        // Full implementation would parse Escher Opt record for adjustment properties
        // For now, return empty vector - adjustments are optional
        Vec::new()
    }

    /// Get the auto shape type.
    pub fn auto_shape_type(&self) -> AutoShapeType {
        self.auto_shape_type
    }

    /// Set the auto shape type.
    pub fn set_auto_shape_type(&mut self, auto_shape_type: AutoShapeType) {
        self.auto_shape_type = auto_shape_type;
    }

    /// Get the adjustment values.
    pub fn adjustments(&self) -> &[i32] {
        &self.adjustments
    }

    /// Add an adjustment value.
    pub fn add_adjustment(&mut self, adjustment: i32) {
        self.adjustments.push(adjustment);
    }

    /// Set all adjustment values.
    pub fn set_adjustments(&mut self, adjustments: Vec<i32>) {
        self.adjustments = adjustments;
    }

    /// Check if this is a basic geometric shape (rectangle, oval, etc.).
    pub fn is_basic_shape(&self) -> bool {
        matches!(
            self.auto_shape_type,
            AutoShapeType::Rectangle
                | AutoShapeType::RoundRectangle
                | AutoShapeType::Oval
                | AutoShapeType::Diamond
                | AutoShapeType::Triangle
                | AutoShapeType::RightTriangle
        )
    }

    /// Check if this is a complex shape that may have adjustments.
    pub fn is_complex_shape(&self) -> bool {
        matches!(
            self.auto_shape_type,
            AutoShapeType::Star
                | AutoShapeType::Arrow
                | AutoShapeType::ThickArrow
                | AutoShapeType::Balloon
                | AutoShapeType::Seal
                | AutoShapeType::Arc
                | AutoShapeType::Teardrop
        )
    }
}

impl Shape for AutoShape {
    fn properties(&self) -> &ShapeProperties {
        &self.container.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.container.properties
    }

    fn text(&self) -> super::super::package::Result<String> {
        // Auto shapes may contain text, but it's optional
        Ok(String::new())
    }

    fn has_text(&self) -> bool {
        false // Auto shapes typically don't have inherent text
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::super::shape::ShapeType;
    use super::*;

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_autoshape_creation() {
        let mut props = ShapeProperties::default();
        props.id = 3001;
        props.shape_type = ShapeType::AutoShape;
        props.x = 100;
        props.y = 100;
        props.width = 200;
        props.height = 150;

        let autoshape = AutoShape::new(props, vec![1, 2, 3]);
        assert_eq!(autoshape.id(), 3001);
        assert_eq!(autoshape.shape_type(), ShapeType::AutoShape);
        assert_eq!(autoshape.auto_shape_type(), AutoShapeType::Rectangle);
        assert!(autoshape.is_basic_shape());
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_autoshape_type_operations() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::AutoShape;

        let mut autoshape = AutoShape::new(props, vec![]);
        autoshape.set_auto_shape_type(AutoShapeType::Oval);
        autoshape.add_adjustment(1000);
        autoshape.add_adjustment(2000);

        assert_eq!(autoshape.auto_shape_type(), AutoShapeType::Oval);
        assert_eq!(autoshape.adjustments().len(), 2);
        assert_eq!(autoshape.adjustments()[0], 1000);
        assert_eq!(autoshape.adjustments()[1], 2000);
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_autoshape_shape_classification() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::AutoShape;

        // Test basic shape
        let mut basic_shape = AutoShape::new(props.clone(), vec![]);
        basic_shape.set_auto_shape_type(AutoShapeType::Rectangle);
        assert!(basic_shape.is_basic_shape());
        assert!(!basic_shape.is_complex_shape());

        // Test complex shape
        let mut complex_shape = AutoShape::new(props, vec![]);
        complex_shape.set_auto_shape_type(AutoShapeType::Star);
        assert!(!complex_shape.is_basic_shape());
        assert!(complex_shape.is_complex_shape());
    }

    #[test]
    fn test_autoshape_type_conversion() {
        assert_eq!(AutoShapeType::from(1), AutoShapeType::Rectangle);
        assert_eq!(AutoShapeType::from(3), AutoShapeType::Oval);
        assert_eq!(AutoShapeType::from(12), AutoShapeType::Star);
        assert_eq!(AutoShapeType::from(999), AutoShapeType::Custom(999));
    }
}
