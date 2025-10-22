//! Shape geometry extraction from Escher properties.
//!
//! This module provides functions to extract geometric information from Escher
//! shape properties, following Apache POI's approach.
//!
//! # Performance
//!
//! - Zero-copy property access via borrowing
//! - O(1) property lookups
//! - Minimal allocations

use super::super::escher::{EscherProperties, EscherPropertyId};

/// Geometric rectangle defined by coordinates.
///
/// In Escher format, geometry is defined by left, top, right, bottom coordinates
/// in master units (typically 1/576 inch).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GeometryRect {
    /// Left coordinate
    pub left: i32,
    /// Top coordinate
    pub top: i32,
    /// Right coordinate
    pub right: i32,
    /// Bottom coordinate
    pub bottom: i32,
}

impl GeometryRect {
    /// Create a new geometry rectangle.
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Get width in master units.
    #[inline]
    pub const fn width(&self) -> i32 {
        self.right - self.left
    }

    /// Get height in master units.
    #[inline]
    pub const fn height(&self) -> i32 {
        self.bottom - self.top
    }

    /// Get center point (x, y).
    #[inline]
    pub const fn center(&self) -> (i32, i32) {
        (
            self.left + (self.width() / 2),
            self.top + (self.height() / 2),
        )
    }
}

/// Extract geometry rectangle from Escher properties.
///
/// Geometry properties define the internal coordinate space of the shape.
/// These are different from the shape's position/size which is in the anchor.
///
/// # Arguments
///
/// * `props` - Parsed Escher properties from Opt record
///
/// # Returns
///
/// `Some(GeometryRect)` if all geometry properties are present, `None` otherwise
///
/// # Performance
///
/// - Direct property access (no iteration)
/// - Early return on missing properties
/// - No allocations
///
/// # Example
///
/// ```ignore
/// use litchi::ole::ppt::escher::EscherProperties;
/// use litchi::ole::ppt::shapes::geometry::extract_geometry_rect;
///
/// let props = EscherProperties::from_opt_record(&opt_record);
/// if let Some(geom) = extract_geometry_rect(&props) {
///     println!("Geometry: {}x{}", geom.width(), geom.height());
/// }
/// ```
#[inline]
pub fn extract_geometry_rect<'data>(props: &EscherProperties<'data>) -> Option<GeometryRect> {
    let left = props.get_coord(EscherPropertyId::GeomLeft)?;
    let top = props.get_coord(EscherPropertyId::GeomTop)?;
    let right = props.get_coord(EscherPropertyId::GeomRight)?;
    let bottom = props.get_coord(EscherPropertyId::GeomBottom)?;

    Some(GeometryRect::new(left, top, right, bottom))
}

/// Shape path types from MS-ODRAW specification.
///
/// The shape path defines how to interpret the vertices and segments
/// of a complex shape geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ShapePathType {
    /// Lines only, no curves
    LinesOnly = 0,
    /// Lines and curves
    LinesCurves = 1,
    /// Closed path
    Closed = 2,
    /// Unknown path type
    Unknown = 0xFFFFFFFF,
}

impl From<i32> for ShapePathType {
    fn from(value: i32) -> Self {
        match value as u32 {
            0 => Self::LinesOnly,
            1 => Self::LinesCurves,
            2 => Self::Closed,
            _ => Self::Unknown,
        }
    }
}

/// Extract shape path type from Escher properties.
///
/// The shape path determines how the shape's geometry is rendered.
///
/// # Arguments
///
/// * `props` - Parsed Escher properties
///
/// # Returns
///
/// The shape path type, or `None` if not specified
///
/// # Performance
///
/// - Single property lookup
/// - No allocations
#[inline]
pub fn extract_shape_path<'data>(props: &EscherProperties<'data>) -> Option<ShapePathType> {
    props
        .get_int(EscherPropertyId::ShapePath)
        .map(ShapePathType::from)
}

/// Vertex data for complex shapes.
///
/// Vertices define the points of a shape's geometry in the shape's
/// coordinate space (defined by the geometry rectangle).
///
/// # Format
///
/// Each vertex is typically 8 bytes: 4 bytes X, 4 bytes Y
#[derive(Debug, Clone)]
pub struct VertexData<'data> {
    /// Raw vertex data (borrowed)
    data: &'data [u8],
    /// Number of vertices
    count: usize,
}

impl<'data> VertexData<'data> {
    /// Create vertex data from raw bytes.
    ///
    /// # Arguments
    ///
    /// * `data` - Raw vertex data (8 bytes per vertex)
    ///
    /// # Returns
    ///
    /// `Some(VertexData)` if data is valid, `None` otherwise
    #[inline]
    pub fn new(data: &'data [u8]) -> Option<Self> {
        if data.is_empty() || !data.len().is_multiple_of(8) {
            return None;
        }

        Some(Self {
            data,
            count: data.len() / 8,
        })
    }

    /// Get number of vertices.
    #[inline]
    pub const fn count(&self) -> usize {
        self.count
    }

    /// Get vertex at index as (x, y) pair.
    ///
    /// # Arguments
    ///
    /// * `index` - Vertex index (0-based)
    ///
    /// # Returns
    ///
    /// `Some((x, y))` if index is valid, `None` otherwise
    ///
    /// # Performance
    ///
    /// - Direct offset calculation
    /// - No bounds checking in release builds (debug has assertions)
    #[inline]
    pub fn get(&self, index: usize) -> Option<(i32, i32)> {
        if index >= self.count {
            return None;
        }

        let offset = index * 8;
        let x = i32::from_le_bytes([
            self.data[offset],
            self.data[offset + 1],
            self.data[offset + 2],
            self.data[offset + 3],
        ]);
        let y = i32::from_le_bytes([
            self.data[offset + 4],
            self.data[offset + 5],
            self.data[offset + 6],
            self.data[offset + 7],
        ]);

        Some((x, y))
    }

    /// Iterate over all vertices as (x, y) pairs.
    ///
    /// # Performance
    ///
    /// - Zero-copy iteration
    /// - No allocations
    pub fn iter(&self) -> impl Iterator<Item = (i32, i32)> + '_ {
        (0..self.count).filter_map(move |i| self.get(i))
    }
}

/// Extract vertices from Escher properties.
///
/// Vertices define the points of complex shapes like polygons, freeform shapes, etc.
/// They are stored as an array property with 8 bytes per vertex (4 bytes X, 4 bytes Y).
///
/// # Arguments
///
/// * `props` - Parsed Escher properties
///
/// # Returns
///
/// `Some(VertexData)` if vertices are present and valid, `None` otherwise
///
/// # Performance
///
/// - Zero-copy: borrows data from properties
/// - No parsing overhead (vertices accessed lazily)
/// - O(1) lookup
///
/// # Example
///
/// ```ignore
/// if let Some(vertices) = extract_vertices(&props) {
///     for (x, y) in vertices.iter() {
///         println!("Vertex: ({}, {})", x, y);
///     }
/// }
/// ```
#[inline]
pub fn extract_vertices<'data>(props: &EscherProperties<'data>) -> Option<VertexData<'data>> {
    // Try to get vertices as array property
    if let Some(array) = props.get_array(EscherPropertyId::Vertices) {
        // Convert array elements to vertex data
        // Each element should be 8 bytes (x, y coordinates)
        let element_size = array.element_size();

        if element_size == 8 {
            // Get the raw data directly (zero-copy)
            let raw_data = array.raw_data();
            if raw_data.len() >= 6 {
                // Skip the 6-byte header, get element data
                return VertexData::new(&raw_data[6..]);
            }
        }
    }

    // Also try as complex property (some files may store it this way)
    props
        .get_binary(EscherPropertyId::Vertices)
        .and_then(VertexData::new)
}

/// Extract segment info from Escher properties.
///
/// Segment info defines how to connect vertices (lines, curves, etc.)
/// for complex shape paths.
///
/// # Arguments
///
/// * `props` - Parsed Escher properties
///
/// # Returns
///
/// Raw segment data if present
///
/// # Performance
///
/// - Zero-copy via borrow
/// - Single property lookup
#[inline]
pub fn extract_segment_info<'data>(props: &EscherProperties<'data>) -> Option<&'data [u8]> {
    props.get_binary(EscherPropertyId::SegmentInfo)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_geometry_rect_dimensions() {
        let geom = GeometryRect::new(0, 0, 1000, 500);
        assert_eq!(geom.width(), 1000);
        assert_eq!(geom.height(), 500);
        assert_eq!(geom.center(), (500, 250));
    }

    #[test]
    fn test_vertex_data_creation() {
        // Create vertex data: 2 vertices at (10, 20) and (30, 40)
        let data = [
            10, 0, 0, 0, // X1 = 10
            20, 0, 0, 0, // Y1 = 20
            30, 0, 0, 0, // X2 = 30
            40, 0, 0, 0, // Y2 = 40
        ];

        let vertices = VertexData::new(&data).unwrap();
        assert_eq!(vertices.count(), 2);

        let (x1, y1) = vertices.get(0).unwrap();
        assert_eq!(x1, 10);
        assert_eq!(y1, 20);

        let (x2, y2) = vertices.get(1).unwrap();
        assert_eq!(x2, 30);
        assert_eq!(y2, 40);
    }

    #[test]
    fn test_vertex_data_iteration() {
        let data = [
            10, 0, 0, 0, 20, 0, 0, 0, // Vertex 1: (10, 20)
            30, 0, 0, 0, 40, 0, 0, 0, // Vertex 2: (30, 40)
            50, 0, 0, 0, 60, 0, 0, 0, // Vertex 3: (50, 60)
        ];

        let vertices = VertexData::new(&data).unwrap();
        let points: Vec<_> = vertices.iter().collect();

        assert_eq!(points.len(), 3);
        assert_eq!(points[0], (10, 20));
        assert_eq!(points[1], (30, 40));
        assert_eq!(points[2], (50, 60));
    }

    #[test]
    fn test_shape_path_type_conversion() {
        assert_eq!(ShapePathType::from(0), ShapePathType::LinesOnly);
        assert_eq!(ShapePathType::from(1), ShapePathType::LinesCurves);
        assert_eq!(ShapePathType::from(2), ShapePathType::Closed);
        assert_eq!(ShapePathType::from(99), ShapePathType::Unknown);
    }
}
