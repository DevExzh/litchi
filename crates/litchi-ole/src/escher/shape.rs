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
    Rectangle,
    Ellipse,
    TextBox,
    Line,
    Polygon,
    Group,
    Picture,
    AutoShape,
    Connector,
    Unknown,
}

/// Escher shape structure.
///
/// # Performance
///
/// - Zero-copy: Borrows from document data
///   A parsed Escher shape with properties and children.
#[derive(Debug, Clone)]
pub struct EscherShape<'data> {
    pub shape_type: EscherShapeType,
    pub shape_id: Option<u32>,
    pub properties: EscherProperties<'data>,
    pub text: Option<String>,
    pub is_group: bool,
    pub children: Vec<EscherShape<'data>>,
    container: EscherContainer<'data>,
    anchor: Option<ShapeAnchor>,
}

impl<'data> EscherShape<'data> {
    /// Parse an Escher shape from an SpContainer record.
    pub fn from_container(container: EscherContainer<'data>) -> Self {
        let shape_type = Self::detect_shape_type(&container);
        let shape_id = Self::extract_shape_id(&container);
        let properties = EscherProperties::from_container(&container);
        let anchor = Self::extract_anchor(&container);

        let text = if let Some(textbox) = container.find_child(EscherRecordType::ClientTextbox) {
            super::text::extract_text_from_textbox(&textbox)
        } else {
            None
        };

        let is_group = shape_type == EscherShapeType::Group;

        let mut children = Vec::new();

        if is_group {
            let mut is_first = true;

            for child in container.children().flatten() {
                match child.record_type {
                    EscherRecordType::SpContainer => {
                        if is_first {
                            is_first = false;
                            continue;
                        }

                        let sp_container = EscherContainer::new(child);
                        let child_shape = EscherShape::from_container(sp_container);
                        children.push(child_shape);
                    },
                    EscherRecordType::SpgrContainer => {
                        let group_container = EscherContainer::new(child);
                        let group_shape = EscherShape::from_container(group_container);
                        children.push(group_shape);
                    },
                    _ => {},
                }
            }
        }

        Self {
            shape_type,
            shape_id,
            properties,
            text,
            is_group,
            children,
            container,
            anchor,
        }
    }

    #[inline]
    pub fn shape_type(&self) -> EscherShapeType {
        self.shape_type
    }

    #[inline]
    pub fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    #[inline]
    pub fn properties(&self) -> &EscherProperties<'data> {
        &self.properties
    }

    #[inline]
    pub fn anchor(&self) -> Option<&ShapeAnchor> {
        self.anchor.as_ref()
    }

    pub fn can_contain_text(&self) -> bool {
        matches!(
            self.shape_type,
            EscherShapeType::TextBox | EscherShapeType::Rectangle | EscherShapeType::AutoShape
        )
    }

    pub fn text(&self) -> Option<String> {
        if let Some(textbox) = self.container.find_child(EscherRecordType::ClientTextbox) {
            super::text::extract_text_from_textbox(&textbox)
        } else {
            None
        }
    }

    #[inline]
    pub fn container(&self) -> &EscherContainer<'data> {
        &self.container
    }

    pub fn child_shapes(&self) -> Vec<EscherShape<'data>> {
        if self.shape_type != EscherShapeType::Group {
            return Vec::new();
        }

        let mut children = Vec::new();
        let mut is_first = true;

        for child in self.container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    if is_first {
                        is_first = false;
                        continue;
                    }

                    let sp_container = EscherContainer::new(child);
                    let child_shape = EscherShape::from_container(sp_container);
                    children.push(child_shape);
                },
                EscherRecordType::SpgrContainer => {
                    let group_container = EscherContainer::new(child);
                    let group_shape = EscherShape::from_container(group_container);
                    children.push(group_shape);
                },
                _ => {},
            }
        }

        children
    }

    fn detect_shape_type(container: &EscherContainer<'data>) -> EscherShapeType {
        if container.record().record_type == EscherRecordType::SpgrContainer {
            return EscherShapeType::Group;
        }

        if let Some(sp) = container.find_child(EscherRecordType::Sp) {
            let shape_type_id = sp.instance;

            return match shape_type_id {
                75 if Self::has_picture_data(container) => EscherShapeType::Picture,
                75 | 202 => EscherShapeType::TextBox,
                1 => EscherShapeType::Rectangle,
                3 => EscherShapeType::Ellipse,
                20 => EscherShapeType::Line,
                0 => EscherShapeType::Group,
                _ if shape_type_id < 203 => EscherShapeType::AutoShape,
                _ => EscherShapeType::Unknown,
            };
        }

        if container
            .find_child(EscherRecordType::ClientTextbox)
            .is_some()
        {
            return EscherShapeType::TextBox;
        }

        for child_result in container.children() {
            if let Ok(child) = child_result
                && child.record_type == EscherRecordType::SpgrContainer
            {
                return EscherShapeType::Group;
            }
        }

        EscherShapeType::Unknown
    }

    fn has_picture_data(container: &EscherContainer<'data>) -> bool {
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

    fn extract_shape_id(container: &EscherContainer<'data>) -> Option<u32> {
        if let Some(sp) = container.find_child(EscherRecordType::Sp)
            && sp.data.len() >= 4
        {
            let id = u32::from_le_bytes([sp.data[0], sp.data[1], sp.data[2], sp.data[3]]);
            return Some(id);
        }
        None
    }

    fn extract_anchor(container: &EscherContainer<'data>) -> Option<ShapeAnchor> {
        if let Some(child_anchor) = container.find_child(EscherRecordType::ChildAnchor)
            && let Some(anchor) = ShapeAnchor::from_child_anchor(&child_anchor)
        {
            return Some(anchor);
        }

        if let Some(client_anchor) = container.find_child(EscherRecordType::ClientAnchor)
            && let Some(anchor) = ShapeAnchor::from_client_anchor(&client_anchor)
        {
            return Some(anchor);
        }

        None
    }
}
