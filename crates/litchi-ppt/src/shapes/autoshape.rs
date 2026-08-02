//! Auto shape implementation.
//!
//! Auto shapes are predefined shapes like rectangles, ovals, arrows, etc.
//! that can be used as building blocks for more complex graphics in PowerPoint.

use super::geometry::{
    GeometryRect, ShapePathType, extract_geometry_rect, extract_segment_info, extract_shape_path,
    extract_vertices,
};
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
///
/// Uses lifetime parameter `'a` to enable zero-copy parsing when the shape
/// data can be borrowed from a larger buffer.
#[derive(Debug, Clone)]
pub struct AutoShape<'a> {
    /// Shape container with properties and data
    container: ShapeContainer<'a>,
    /// The type of auto shape
    auto_shape_type: AutoShapeType,
    /// Adjustment values for shape parameters (for complex shapes)
    adjustments: Vec<i32>,
    /// Explicit custom/freeform geometry, when present.
    geometry: Option<AutoShapeGeometry>,
}

/// Owned OfficeArt geometry for a custom or freeform shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoShapeGeometry {
    coordinate_space: Option<GeometryRect>,
    path_type: Option<ShapePathType>,
    vertices: Vec<(i32, i32)>,
    segment_info: Vec<u16>,
}

impl AutoShapeGeometry {
    fn from_properties(props: &litchi_odraw::prop::Props<'_>) -> Option<Self> {
        let coordinate_space = extract_geometry_rect(props);
        let path_type = extract_shape_path(props);
        let vertices: Vec<(i32, i32)> = extract_vertices(props)
            .map(|vertices| vertices.iter().collect())
            .unwrap_or_default();
        let segment_info: Vec<u16> = extract_segment_info(props)
            .map(|segments| {
                segments
                    .chunks_exact(2)
                    .map(|segment| u16::from_le_bytes([segment[0], segment[1]]))
                    .collect()
            })
            .unwrap_or_default();

        (coordinate_space.is_some()
            || path_type.is_some()
            || !vertices.is_empty()
            || !segment_info.is_empty())
        .then_some(Self {
            coordinate_space,
            path_type,
            vertices,
            segment_info,
        })
    }

    /// Return the shape's internal geometry coordinate space.
    pub fn coordinate_space(&self) -> Option<GeometryRect> {
        self.coordinate_space
    }

    /// Return the high-level path connectivity mode.
    pub fn path_type(&self) -> Option<ShapePathType> {
        self.path_type
    }

    /// Return the freeform vertices in geometry-space units.
    pub fn vertices(&self) -> &[(i32, i32)] {
        &self.vertices
    }

    /// Return raw MSOPATHINFO words in document order.
    pub fn segment_info(&self) -> &[u16] {
        &self.segment_info
    }
}

impl<'a> AutoShape<'a> {
    /// Create a new auto shape with owned data.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            auto_shape_type: AutoShapeType::Rectangle, // Default
            adjustments: Vec::new(),
            geometry: None,
        }
    }

    /// Create an auto shape from already-parsed OfficeArt metadata.
    pub(crate) fn from_odraw(
        properties: ShapeProperties,
        native_shape_type: u16,
        odraw_properties: &litchi_odraw::prop::Props<'_>,
    ) -> Self {
        let mut shape = Self::new(properties, Vec::new());
        shape.auto_shape_type = AutoShapeType::from(native_shape_type);
        shape.adjustments = Self::extract_adjustments_from_properties(odraw_properties);
        shape.geometry = AutoShapeGeometry::from_properties(odraw_properties);
        shape
    }

    /// Create an auto shape from an existing container.
    pub fn from_container(container: ShapeContainer<'a>) -> super::super::package::Result<Self> {
        // Extract auto shape type from raw data
        let auto_shape_type = Self::extract_shape_type(&container.raw_data)?;

        // Extract adjustment values if available
        let adjustments = Self::extract_adjustments(&container.raw_data)?;
        let geometry = Self::extract_geometry(&container.raw_data)?;

        Ok(Self {
            container,
            auto_shape_type,
            adjustments,
            geometry,
        })
    }

    /// Extract auto shape type from raw shape data.
    ///
    /// In Escher records, the shape type is typically stored in the Sp (shape) record.
    /// The instance field of the Sp record contains the shape type ID.
    ///
    /// # Performance
    ///
    /// - Bounded depth-first record scanning
    /// - Early termination on match
    fn extract_shape_type(raw_data: &[u8]) -> super::super::package::Result<AutoShapeType> {
        Ok(
            Self::find_odraw_record(raw_data, litchi_odraw::RecordKind::Sp)?
                .map(|sp| AutoShapeType::from(sp.instance()))
                .unwrap_or(AutoShapeType::Rectangle),
        )
    }

    /// Extract adjustment values from raw shape data.
    ///
    /// Adjustment values control the geometry of complex shapes (e.g., arrow head size).
    /// Based on Apache POI's auto shape adjustment parsing.
    ///
    /// In Escher format, adjustments are stored in the Opt record as properties:
    /// - `AdjustValue` (0x0147): First adjustment value
    /// - `Adjust2Value` (0x0148) through `Adjust10Value` (0x0150): Additional adjustments
    ///
    /// # Performance
    ///
    /// - Returns empty vector for shapes without adjustments
    /// - Pre-allocated capacity for shapes with adjustments (max 10)
    /// - Bounded depth-first record scanning
    fn extract_adjustments(raw_data: &[u8]) -> super::super::package::Result<Vec<i32>> {
        let Some(opt) = Self::find_odraw_record(raw_data, litchi_odraw::RecordKind::Opt)? else {
            return Ok(Vec::new());
        };
        let properties = litchi_odraw::prop::Props::parse(&opt)?;
        Ok(Self::extract_adjustments_from_properties(&properties))
    }

    fn extract_geometry(
        raw_data: &[u8],
    ) -> super::super::package::Result<Option<AutoShapeGeometry>> {
        let Some(opt) = Self::find_odraw_record(raw_data, litchi_odraw::RecordKind::Opt)? else {
            return Ok(None);
        };
        let properties = litchi_odraw::prop::Props::parse(&opt)?;
        Ok(AutoShapeGeometry::from_properties(&properties))
    }

    fn find_odraw_record<'data>(
        data: &'data [u8],
        record_type: litchi_odraw::RecordKind,
    ) -> super::super::package::Result<Option<litchi_odraw::Record<'data>>> {
        const MAX_SCANNED_RECORDS: usize = 4_096;

        let mut streams = vec![data];
        let mut scanned = 0;
        while let Some(stream) = streams.pop() {
            let mut offset = 0;
            while stream.len().saturating_sub(offset) >= 8 {
                if scanned == MAX_SCANNED_RECORDS {
                    return Err(super::super::package::PptError::Corrupted(
                        "OfficeArt shape exceeds the PPT record limit".to_string(),
                    ));
                }
                let (record, consumed) = litchi_odraw::Record::parse(stream, offset)?;
                scanned += 1;
                if record.kind() == record_type {
                    return Ok(Some(record));
                }
                if record.is_container() {
                    streams.push(record.data());
                }
                if consumed == 0 {
                    break;
                }
                offset = offset.checked_add(consumed).ok_or_else(|| {
                    super::super::package::PptError::Corrupted(
                        "OfficeArt shape traversal offset overflow".to_string(),
                    )
                })?;
            }
        }
        Ok(None)
    }

    /// Extract adjustment values from Escher properties.
    ///
    /// This is the proper implementation using checked OfficeArt properties.
    /// Adjustments are stored as simple integer properties.
    ///
    /// # Arguments
    ///
    /// * `props` - Parsed Escher properties from Opt record
    ///
    /// # Returns
    ///
    /// Vector of adjustment values (up to 10 values)
    ///
    /// # Performance
    ///
    /// - Pre-allocated Vec with known max capacity (10)
    /// - O(1) property lookups via HashMap
    /// - Missing positions use the protocol-defined default of zero
    /// - Trailing absent positions are omitted
    pub fn extract_adjustments_from_properties(props: &litchi_odraw::prop::Props) -> Vec<i32> {
        use litchi_odraw::prop::Id;

        let adjustment_ids = [
            Id::AdjustValue,
            Id::Adjust2Value,
            Id::Adjust3Value,
            Id::Adjust4Value,
            Id::Adjust5Value,
            Id::Adjust6Value,
            Id::Adjust7Value,
            Id::Adjust8Value,
            Id::Adjust9Value,
            Id::Adjust10Value,
        ];

        let mut adjustments = vec![0; adjustment_ids.len()];
        let mut populated_len = 0;
        for (index, id) in adjustment_ids.into_iter().enumerate() {
            if let Some(value) = props.get_adjust(id) {
                adjustments[index] = value;
                populated_len = index + 1;
            }
        }
        adjustments.truncate(populated_len);

        adjustments
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

    /// Get the text contained by this shape.
    pub fn text(&self) -> &str {
        self.container.text_content.as_deref().unwrap_or_default()
    }

    /// Set the text contained by this shape.
    pub fn set_text(&mut self, text: String) {
        self.container.set_text(text);
    }

    /// Return explicit custom/freeform geometry, when the shape defines it.
    pub fn geometry(&self) -> Option<&AutoShapeGeometry> {
        self.geometry.as_ref()
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

impl<'a> Shape for AutoShape<'a>
where
    'a: 'static,
{
    fn properties(&self) -> &ShapeProperties {
        &self.container.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.container.properties
    }

    fn text(&self) -> super::super::package::Result<String> {
        Ok(self.text().to_owned())
    }

    fn has_text(&self) -> bool {
        !self.text().is_empty()
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

    fn autoshape_container() -> ShapeContainer<'static> {
        use crate::officeart_wire::{PropertyBuilder, ShapeBuilder, record_type, write_container};

        let mut children = Vec::new();
        ShapeBuilder::new(13, 7).write(&mut children).unwrap();
        let mut properties = PropertyBuilder::new();
        properties.add_simple(0x0147, 25);
        properties.add_simple(0x0149, 75);
        properties.write(&mut children).unwrap();

        let mut raw_data = Vec::new();
        write_container(&mut raw_data, 0, record_type::SP_CONTAINER, &children).unwrap();
        ShapeContainer::new(ShapeProperties::default(), raw_data)
    }

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

    #[test]
    fn test_autoshape_from_container_parses_officeart_metadata() {
        let autoshape = AutoShape::from_container(autoshape_container()).unwrap();

        assert_eq!(autoshape.auto_shape_type(), AutoShapeType::Arrow);
        assert_eq!(autoshape.adjustments(), &[25, 0, 75]);
    }
}
