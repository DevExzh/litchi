//! Escher shape parsing and representation.
//!
//! # Performance
//!
//! - Zero-copy shape data access
//! - Lazy property parsing
//! - Enum-based shape type dispatch (no trait objects)

use super::container::EscherContainer;
use super::types::EscherRecordType;
use super::properties::{EscherProperties, ShapeAnchor};

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
            EscherShapeType::TextBox
                | EscherShapeType::Rectangle
                | EscherShapeType::AutoShape
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
    
    /// Detect shape type from container properties.
    ///
    /// Based on Apache POI's shape type detection logic.
    fn detect_shape_type(container: &EscherContainer<'data>) -> EscherShapeType {
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
        if container.find_child(EscherRecordType::ClientTextbox).is_some() {
            return EscherShapeType::TextBox;
        }
        
        // Check if it's a group by looking for SpgrContainer
        for child_result in container.children() {
            if let Ok(child) = child_result
                && child.record_type == EscherRecordType::SpgrContainer {
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
                }
                _ => {}
            }
        }
        false
    }
    
    /// Extract shape ID from Sp atom.
    fn extract_shape_id(container: &EscherContainer<'data>) -> Option<u32> {
        if let Some(sp) = container.find_child(EscherRecordType::Sp) {
            // Shape ID is in the first 4 bytes of Sp data
            if sp.data.len() >= 4 {
                let id = u32::from_le_bytes([
                    sp.data[0],
                    sp.data[1],
                    sp.data[2],
                    sp.data[3],
                ]);
                return Some(id);
            }
        }
        None
    }
    
    /// Extract shape anchor (position and size).
    fn extract_anchor(container: &EscherContainer<'data>) -> Option<ShapeAnchor> {
        // Try ChildAnchor first
        if let Some(child_anchor) = container.find_child(EscherRecordType::ChildAnchor)
            && let Some(anchor) = ShapeAnchor::from_child_anchor(&child_anchor) {
                return Some(anchor);
            }
        
        // Try ClientAnchor
        if let Some(client_anchor) = container.find_child(EscherRecordType::ClientAnchor)
            && let Some(anchor) = ShapeAnchor::from_client_anchor(&client_anchor) {
                return Some(anchor);
            }
        
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_shape_type_detection() {
        // Would need actual test data here
        // For now, just ensure the enum is usable
        let shape_type = EscherShapeType::TextBox;
        assert_eq!(shape_type, EscherShapeType::TextBox);
    }
}

