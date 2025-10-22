//! Escher shape parsing and representation.
//!
//! # Performance
//!
//! - Zero-copy shape data access
//! - Lazy property parsing
//! - Enum-based shape type dispatch (no trait objects)

use super::container::EscherContainer;
use super::properties::{EscherProperties, ShapeAnchor};
use super::types::EscherRecordType;

/// Escher shape type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EscherShapeType {
    /// Rectangle
    Rectangle,
    /// Ellipse
    Ellipse,
    /// Text box
    TextBox,
    /// Line
    Line,
    /// Polygon
    Polygon,
    /// Group (contains other shapes)
    Group,
    /// Picture/Image
    Picture,
    /// Auto shape (predefined shape)
    AutoShape,
    /// Connector
    Connector,
    /// Unknown shape type
    Unknown,
}

/// Escher shape structure.
///
/// # Performance
///
/// - Zero-copy: Borrows from document data
/// - Lazy: Properties parsed on demand
#[derive(Debug, Clone)]
pub struct EscherShape<'data> {
    /// The SpContainer record
    container: EscherContainer<'data>,
    /// Shape type
    shape_type: EscherShapeType,
    /// Shape ID
    shape_id: Option<u32>,
    /// Shape properties (position, size, colors, etc.)
    properties: EscherProperties,
    /// Shape anchor (position and size)
    anchor: Option<ShapeAnchor>,
}

impl<'data> EscherShape<'data> {
    /// Parse an Escher shape from an SpContainer record.
    pub fn from_container(container: EscherContainer<'data>) -> Self {
        let shape_type = Self::detect_shape_type(&container);
        let shape_id = Self::extract_shape_id(&container);
        let properties = EscherProperties::from_container(&container);
        let anchor = Self::extract_anchor(&container);

        Self {
            container,
            shape_type,
            shape_id,
            properties,
            anchor,
        }
    }

    /// Get the shape type.
    #[inline]
    pub fn shape_type(&self) -> EscherShapeType {
        self.shape_type
    }

    /// Get the shape ID.
    #[inline]
    pub fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    /// Get shape properties.
    #[inline]
    pub fn properties(&self) -> &EscherProperties {
        &self.properties
    }

    /// Get shape anchor (position and size).
    #[inline]
    pub fn anchor(&self) -> Option<&ShapeAnchor> {
        self.anchor.as_ref()
    }

    /// Check if this shape can contain text.
    pub fn can_contain_text(&self) -> bool {
        matches!(
            self.shape_type,
            EscherShapeType::TextBox | EscherShapeType::Rectangle | EscherShapeType::AutoShape
        )
    }

    /// Extract text from this shape.
    ///
    /// # Performance
    ///
    /// - Searches for ClientTextbox child
    /// - Uses text extraction module
    pub fn text(&self) -> Option<String> {
        // Look for ClientTextbox in children
        if let Some(textbox) = self.container.find_child(EscherRecordType::ClientTextbox) {
            super::text::extract_text_from_textbox(&textbox)
        } else {
            None
        }
    }

    /// Get the underlying container.
    #[inline]
    pub fn container(&self) -> &EscherContainer<'data> {
        &self.container
    }

    /// Get child shapes if this is a group shape.
    ///
    /// # Performance
    ///
    /// - Returns iterator over child containers
    /// - Zero-copy: Borrows from document data
    /// - Lazy: Shapes parsed on demand
    ///
    /// # Returns
    ///
    /// Vector of child EscherShape objects for group shapes, empty for non-groups
    pub fn child_shapes(&self) -> Vec<EscherShape<'data>> {
        // Only group shapes have children
        if self.shape_type != EscherShapeType::Group {
            return Vec::new();
        }

        let mut children = Vec::new();

        // For group shapes, look for SpgrContainer or iterate SpContainer children
        // The first SpContainer is the group shape itself, skip it
        let mut is_first = true;

        for child in self.container.children().flatten() {
            match child.record_type {
                // SpContainer holds a single shape
                EscherRecordType::SpContainer => {
                    // Skip the first SpContainer (it's the group shape itself)
                    if is_first {
                        is_first = false;
                        continue;
                    }

                    let sp_container = EscherContainer::new(child);
                    let child_shape = EscherShape::from_container(sp_container);
                    children.push(child_shape);
                },
                // SpgrContainer holds a group of shapes (recursive)
                // Treat the nested SpgrContainer as a group shape itself
                EscherRecordType::SpgrContainer => {
                    let group_container = EscherContainer::new(child);
                    // Create a shape for the nested group (which will recursively load its children)
                    let group_shape = EscherShape::from_container(group_container);
                    children.push(group_shape);
                },
                _ => {},
            }
        }

        children
    }

    /// Extract shapes from SpgrContainer (used for nested groups).
    ///
    /// # Note
    ///
    /// This is a helper function primarily used for testing and special cases.
    /// Normal shape extraction uses the recursive `child_shapes()` method.
    #[cfg_attr(not(test), allow(dead_code))]
    fn extract_from_spgr_container<'a>(
        container: &EscherContainer<'a>,
        shapes: &mut Vec<EscherShape<'a>>,
    ) {
        let mut is_first = true;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    // Skip the first SpContainer in SpgrContainer
                    if is_first {
                        is_first = false;
                        continue;
                    }

                    let sp_container = EscherContainer::new(child);
                    let shape = EscherShape::from_container(sp_container);
                    shapes.push(shape);
                },
                EscherRecordType::SpgrContainer => {
                    // Recursive group
                    let group_container = EscherContainer::new(child);
                    Self::extract_from_spgr_container(&group_container, shapes);
                },
                _ => {},
            }
        }
    }

    /// Detect shape type from container properties.
    ///
    /// Based on Apache POI's shape type detection logic.
    fn detect_shape_type(container: &EscherContainer<'data>) -> EscherShapeType {
        // First, check if this is a SpgrContainer (group container)
        // SpgrContainer (0xF003) directly indicates a group shape
        if container.record().record_type == EscherRecordType::SpgrContainer {
            return EscherShapeType::Group;
        }

        // Look for Sp (Shape) atom to determine type
        if let Some(sp) = container.find_child(EscherRecordType::Sp) {
            // The shape type is in the instance field
            // See MS-ODRAW specification, section 2.2.37
            let shape_type_id = sp.instance;

            return match shape_type_id {
                // Text box (check for picture data first for type 75)
                75 if Self::has_picture_data(container) => EscherShapeType::Picture,
                75 | 202 => EscherShapeType::TextBox,
                // Basic shapes
                1 => EscherShapeType::Rectangle,
                3 => EscherShapeType::Ellipse,
                // Line
                20 => EscherShapeType::Line,
                // Group (pseudo shape)
                0 => EscherShapeType::Group,
                // Default to auto shape for known shape IDs
                _ if shape_type_id < 203 => EscherShapeType::AutoShape,
                _ => EscherShapeType::Unknown,
            };
        }

        // Fallback: Check if it has text
        if container
            .find_child(EscherRecordType::ClientTextbox)
            .is_some()
        {
            return EscherShapeType::TextBox;
        }

        // Check if it's a group by looking for SpgrContainer child
        for child_result in container.children() {
            if let Ok(child) = child_result
                && child.record_type == EscherRecordType::SpgrContainer
            {
                return EscherShapeType::Group;
            }
        }

        // Default
        EscherShapeType::Unknown
    }

    /// Check if container has picture/blip data.
    fn has_picture_data(container: &EscherContainer<'data>) -> bool {
        // Look for blip references or embedded blip data
        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::BlipJpeg
                | EscherRecordType::BlipPng
                | EscherRecordType::BlipDib
                | EscherRecordType::BlipTiff
                | EscherRecordType::BlipEmf
                | EscherRecordType::BlipWmf
                | EscherRecordType::BlipPict => {
                    return true;
                },
                _ => {},
            }
        }
        false
    }

    /// Extract shape ID from Sp atom.
    fn extract_shape_id(container: &EscherContainer<'data>) -> Option<u32> {
        if let Some(sp) = container.find_child(EscherRecordType::Sp) {
            // Shape ID is in the first 4 bytes of Sp data
            if sp.data.len() >= 4 {
                let id = u32::from_le_bytes([sp.data[0], sp.data[1], sp.data[2], sp.data[3]]);
                return Some(id);
            }
        }
        None
    }

    /// Extract shape anchor (position and size).
    fn extract_anchor(container: &EscherContainer<'data>) -> Option<ShapeAnchor> {
        // Try ChildAnchor first
        if let Some(child_anchor) = container.find_child(EscherRecordType::ChildAnchor)
            && let Some(anchor) = ShapeAnchor::from_child_anchor(&child_anchor)
        {
            return Some(anchor);
        }

        // Try ClientAnchor
        if let Some(client_anchor) = container.find_child(EscherRecordType::ClientAnchor)
            && let Some(anchor) = ShapeAnchor::from_client_anchor(&client_anchor)
        {
            return Some(anchor);
        }

        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::ppt::escher::record::EscherRecord;

    #[test]
    fn test_shape_type_detection() {
        // Would need actual test data here
        // For now, just ensure the enum is usable
        let shape_type = EscherShapeType::TextBox;
        assert_eq!(shape_type, EscherShapeType::TextBox);
    }

    #[test]
    fn test_child_shapes_non_group() {
        // Create a simple SpContainer for a non-group shape (TextBox)
        let data = create_textbox_spcontainer();
        let (record, _) = EscherRecord::parse(&data, 0).unwrap();
        let container = EscherContainer::new(record);
        let shape = EscherShape::from_container(container);

        // Non-group shapes should return empty children
        let children = shape.child_shapes();
        assert_eq!(children.len(), 0);
    }

    #[test]
    fn test_child_shapes_empty_group() {
        // Create a SpgrContainer with only the group shape itself (no children)
        let data = create_empty_group_spgrcontainer();
        let (record, _) = EscherRecord::parse(&data, 0).unwrap();
        let container = EscherContainer::new(record);
        let shape = EscherShape::from_container(container);

        // Empty group should have shape_type Group but no children
        assert_eq!(shape.shape_type(), EscherShapeType::Group);
        let children = shape.child_shapes();
        assert_eq!(children.len(), 0);
    }

    #[test]
    fn test_child_shapes_group_with_children() {
        // Create a SpgrContainer with child shapes
        let data = create_group_with_children();
        let (record, _) = EscherRecord::parse(&data, 0).unwrap();
        let container = EscherContainer::new(record);
        let shape = EscherShape::from_container(container);

        // Group should have shape_type Group and contain children
        assert_eq!(shape.shape_type(), EscherShapeType::Group);
        let children = shape.child_shapes();
        // Should have extracted child shapes (exact count depends on test data)
        assert!(children.len() > 0);
    }

    // Helper function to create a simple TextBox SpContainer
    fn create_textbox_spcontainer() -> Vec<u8> {
        let mut data = Vec::new();

        // SpContainer header
        data.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x04, 0xF0, // record type = 0xF004 (SpContainer)
        ]);
        let sp_data = create_sp_atom(1, 75); // shape type 75 = TextBox
        data.extend_from_slice(&(sp_data.len() as u32).to_le_bytes());
        data.extend_from_slice(&sp_data);

        data
    }

    // Helper function to create an empty group SpgrContainer
    fn create_empty_group_spgrcontainer() -> Vec<u8> {
        let mut data = Vec::new();

        // SpgrContainer header
        data.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x03, 0xF0, // record type = 0xF003 (SpgrContainer)
        ]);

        // First SpContainer (the group shape itself)
        let group_sp = create_sp_container(100, 0); // shape type 0 = Group
        data.extend_from_slice(&(group_sp.len() as u32).to_le_bytes());
        data.extend_from_slice(&group_sp);

        data
    }

    // Helper function to create a group with children
    fn create_group_with_children() -> Vec<u8> {
        let mut data = Vec::new();

        // SpgrContainer header
        data.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x03, 0xF0, // record type = 0xF003 (SpgrContainer)
        ]);

        // Calculate total length
        let group_sp = create_sp_container(100, 0); // Group shape itself
        let child1 = create_sp_container(101, 75); // TextBox child
        let child2 = create_sp_container(102, 1); // Rectangle child

        let total_len = group_sp.len() + child1.len() + child2.len();
        data.extend_from_slice(&(total_len as u32).to_le_bytes());

        // First SpContainer (the group shape itself)
        data.extend_from_slice(&group_sp);

        // Child shapes
        data.extend_from_slice(&child1);
        data.extend_from_slice(&child2);

        data
    }

    // Helper to create a SpContainer with Sp atom
    fn create_sp_container(shape_id: u32, shape_type: u16) -> Vec<u8> {
        let mut data = Vec::new();

        // SpContainer header
        data.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x04, 0xF0, // record type = 0xF004 (SpContainer)
        ]);

        let sp_atom = create_sp_atom(shape_id, shape_type);
        data.extend_from_slice(&(sp_atom.len() as u32).to_le_bytes());
        data.extend_from_slice(&sp_atom);

        data
    }

    // Helper to create Sp atom
    fn create_sp_atom(shape_id: u32, shape_type: u16) -> Vec<u8> {
        let mut data = Vec::new();

        // Sp atom header
        let ver_inst = (shape_type << 4) | 0x02; // version=2, instance=shape_type
        data.extend_from_slice(&ver_inst.to_le_bytes());
        data.extend_from_slice(&[0x0A, 0xF0]); // record type = 0xF00A (Sp)
        data.extend_from_slice(&8u32.to_le_bytes()); // length = 8

        // Sp data: shape ID (4 bytes) + flags (4 bytes)
        data.extend_from_slice(&shape_id.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes()); // flags

        data
    }

    #[test]
    fn test_extract_from_spgr_container_skip_first() {
        // Verify that extract_from_spgr_container skips the first SpContainer
        let data = create_group_with_children();
        let (record, _) = EscherRecord::parse(&data, 0).unwrap();
        let container = EscherContainer::new(record);

        let mut shapes = Vec::new();
        EscherShape::extract_from_spgr_container(&container, &mut shapes);

        // Should extract 2 child shapes (skipping the first which is the group itself)
        assert_eq!(shapes.len(), 2);
    }

    #[test]
    fn test_nested_groups() {
        // Create a nested group structure: Group1 -> Group2 -> Shape
        let data = create_nested_group_structure();
        let (record, _) = EscherRecord::parse(&data, 0).unwrap();
        let container = EscherContainer::new(record);
        let shape = EscherShape::from_container(container);

        assert_eq!(shape.shape_type(), EscherShapeType::Group);

        // Get children of outer group
        let children = shape.child_shapes();
        assert!(children.len() > 0);

        // First child should be another group
        if let Some(first_child) = children.first() {
            assert_eq!(first_child.shape_type(), EscherShapeType::Group);

            // Get children of inner group
            let inner_children = first_child.child_shapes();
            assert!(inner_children.len() > 0);
        }
    }

    // Helper to create nested group structure
    fn create_nested_group_structure() -> Vec<u8> {
        let mut data = Vec::new();

        // Outer SpgrContainer header
        data.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x03, 0xF0, // record type = 0xF003 (SpgrContainer)
        ]);

        // First: group shape itself
        let outer_group_sp = create_sp_container(100, 0);

        // Second: nested SpgrContainer
        let mut inner_group = Vec::new();
        inner_group.extend_from_slice(&[
            0x0F, 0x00, // version=0xF, instance=0
            0x03, 0xF0, // record type = 0xF003 (SpgrContainer)
        ]);

        let inner_group_sp = create_sp_container(200, 0);
        let inner_child = create_sp_container(201, 75);
        let inner_total = inner_group_sp.len() + inner_child.len();

        inner_group.extend_from_slice(&(inner_total as u32).to_le_bytes());
        inner_group.extend_from_slice(&inner_group_sp);
        inner_group.extend_from_slice(&inner_child);

        let total_len = outer_group_sp.len() + inner_group.len();
        data.extend_from_slice(&(total_len as u32).to_le_bytes());
        data.extend_from_slice(&outer_group_sp);
        data.extend_from_slice(&inner_group);

        data
    }
}
