//! Escher shape parsing and representation.
//!
//! # Performance
//!
//! - Zero-copy shape data access
//! - Lazy property parsing
//! - Enum-based shape type dispatch (no trait objects)

use super::container::EscherContainer;
use super::properties::{EscherProperties, EscherPropertyId, ShapeAnchor};
use super::types::EscherRecordType;

/// Escher shape type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EscherShapeType {
    Rectangle,
    Ellipse,
    TextBox,
    Placeholder,
    Line,
    Polygon,
    Group,
    Table,
    Picture,
    Object,
    Media,
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
    native_shape_type: Option<u16>,
    anchor: Option<ShapeAnchor>,
    placeholder: Option<EscherPlaceholder>,
    external_object_id: Option<u32>,
}

/// Placeholder metadata embedded in an OfficeArt client-data record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EscherPlaceholder {
    pub position: i32,
    pub placeholder_type: u8,
    pub size: u8,
}

impl<'data> EscherShape<'data> {
    /// Parse an Escher shape from an SpContainer or SpgrContainer record.
    pub fn from_container(container: EscherContainer<'data>) -> Self {
        let placeholder = Self::extract_placeholder(&container);
        let frame_info = Self::extract_frame_info(&container);
        let mut shape_type = Self::detect_shape_type(&container, frame_info.kind);
        if placeholder.is_some()
            && !matches!(shape_type, EscherShapeType::Group | EscherShapeType::Table)
        {
            shape_type = EscherShapeType::Placeholder;
        }
        // A group is represented by an SpgrContainer whose first SpContainer
        // carries the group's ID, properties, and anchor. Keep the outer
        // container for child traversal, but read metadata from that header.
        let group_header = if matches!(shape_type, EscherShapeType::Group | EscherShapeType::Table)
        {
            container
                .find_child(EscherRecordType::SpContainer)
                .map(EscherContainer::new)
        } else {
            None
        };
        let metadata_container = group_header.as_ref().unwrap_or(&container);
        let shape_id = Self::extract_shape_id(metadata_container);
        let native_shape_type = Self::extract_native_shape_type(metadata_container);
        let properties = EscherProperties::from_container(metadata_container);
        let anchor = Self::extract_anchor(metadata_container);

        let text =
            if let Some(textbox) = metadata_container.find_child(EscherRecordType::ClientTextbox) {
                super::text::extract_text_from_textbox(&textbox)
            } else {
                None
            };

        let is_group = matches!(shape_type, EscherShapeType::Group | EscherShapeType::Table);

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
            native_shape_type,
            anchor,
            placeholder,
            external_object_id: frame_info.external_object_id,
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

    /// Return the native MSOSPT value stored in the OfficeArt `Sp` atom.
    #[inline]
    pub fn native_shape_type(&self) -> Option<u16> {
        self.native_shape_type
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
            EscherShapeType::TextBox
                | EscherShapeType::Placeholder
                | EscherShapeType::Rectangle
                | EscherShapeType::Ellipse
                | EscherShapeType::Polygon
                | EscherShapeType::AutoShape
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

    /// Return parsed child shapes without reparsing or allocating.
    #[inline]
    pub fn children(&self) -> &[EscherShape<'data>] {
        &self.children
    }

    /// Return placeholder metadata when this is a placeholder shape.
    #[inline]
    pub fn placeholder(&self) -> Option<EscherPlaceholder> {
        self.placeholder
    }

    /// Return the external object reference used by an OLE or media frame.
    #[inline]
    pub fn external_object_id(&self) -> Option<u32> {
        self.external_object_id
    }

    /// Parse inert PowerPoint animation metadata from this shape's client data.
    pub fn animation_info(
        &self,
    ) -> crate::ppt::package::Result<Option<crate::ppt::animation::AnimationInfo>> {
        let group_header = if self.is_group {
            self.container
                .find_child(EscherRecordType::SpContainer)
                .map(EscherContainer::new)
        } else {
            None
        };
        let metadata = group_header.as_ref().unwrap_or(&self.container);
        let Some(client_data) = metadata.find_child(EscherRecordType::ClientData) else {
            return Ok(None);
        };
        let mut offset = 0usize;
        while offset + 8 <= client_data.data.len() {
            let (record, consumed) =
                crate::ppt::records::PptRecord::parse(client_data.data, offset)?;
            if record.record_type == crate::consts::PptRecordType::AnimationInfo {
                return crate::ppt::animation::parse_animation_info(&record).map(Some);
            }
            if consumed == 0 {
                return Err(crate::ppt::package::PptError::Corrupted(
                    "zero-length PPT record in OfficeArt client data".to_string(),
                ));
            }
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::ppt::package::PptError::Corrupted(
                    "OfficeArt client-data offset overflow".to_string(),
                )
            })?;
        }
        Ok(None)
    }

    /// Return owned copies of the child shapes.
    pub fn child_shapes(&self) -> Vec<EscherShape<'data>> {
        self.children.clone()
    }

    fn detect_shape_type(
        container: &EscherContainer<'data>,
        frame_kind: Option<EscherShapeType>,
    ) -> EscherShapeType {
        if container.record().record_type == EscherRecordType::SpgrContainer {
            return if Self::is_table_group(container) {
                EscherShapeType::Table
            } else {
                EscherShapeType::Group
            };
        }

        if let Some(sp) = container.find_child(EscherRecordType::Sp) {
            let shape_type_id = sp.instance;

            return match shape_type_id {
                // MSOSPT 75 is the frame used for pictures in binary Office files.
                // The image normally lives in the Pictures stream and is referenced
                // through the pib property rather than embedded in this container.
                75 => frame_kind.unwrap_or(EscherShapeType::Picture),
                202 => EscherShapeType::TextBox,
                1 => EscherShapeType::Rectangle,
                3 => EscherShapeType::Ellipse,
                20 => EscherShapeType::Line,
                // MSOSPT 0 is a non-primitive shape. POI treats one with
                // explicit vertices as a freeform and other instances as an
                // auto shape. Groups are identified by their SpgrContainer.
                0 if EscherProperties::from_container(container)
                    .has(EscherPropertyId::Vertices) =>
                {
                    EscherShapeType::Polygon
                },
                0 => EscherShapeType::AutoShape,
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

    fn is_table_group(container: &EscherContainer<'data>) -> bool {
        let Some(header) = container.find_child(EscherRecordType::SpContainer) else {
            return false;
        };
        let header = EscherContainer::new(header);
        let Some(user_defined) = header.find_child(EscherRecordType::TertiaryOpt) else {
            return false;
        };
        let properties = EscherProperties::from_opt_record(&user_defined);
        properties
            .get_int(super::properties::EscherPropertyId::GroupTableProperties)
            .is_some_and(|value| value & 1 != 0)
    }

    fn extract_placeholder(container: &EscherContainer<'data>) -> Option<EscherPlaceholder> {
        let client_data = container.find_child(EscherRecordType::ClientData)?;
        let mut offset = 0usize;
        while offset + 8 <= client_data.data.len() {
            let (record, consumed) =
                crate::ppt::records::PptRecord::parse(client_data.data, offset).ok()?;
            if record.record_type_raw == 3011 && record.data.len() >= 8 {
                return Some(EscherPlaceholder {
                    position: i32::from_le_bytes(record.data[0..4].try_into().ok()?),
                    placeholder_type: record.data[4],
                    size: record.data[5],
                });
            }
            if consumed == 0 {
                return None;
            }
            offset = offset.checked_add(consumed)?;
        }
        None
    }

    fn extract_frame_info(container: &EscherContainer<'data>) -> EscherFrameInfo {
        const EX_OBJ_REF_ATOM: u16 = 3009;
        const INTERACTIVE_INFO: u16 = 4082;
        const INTERACTIVE_INFO_ATOM: u16 = 4083;
        const ACTION_OLE: u8 = 5;
        const ACTION_MEDIA: u8 = 6;

        let Some(client_data) = container.find_child(EscherRecordType::ClientData) else {
            return EscherFrameInfo::default();
        };
        let mut info = EscherFrameInfo::default();
        let mut offset = 0usize;
        while offset + 8 <= client_data.data.len() {
            let Ok((record, consumed)) =
                crate::ppt::records::PptRecord::parse(client_data.data, offset)
            else {
                break;
            };
            match record.record_type_raw {
                EX_OBJ_REF_ATOM if record.data.len() >= 4 => {
                    info.external_object_id = Some(u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]));
                    info.kind.get_or_insert(EscherShapeType::Object);
                },
                INTERACTIVE_INFO => {
                    let action = record
                        .children
                        .iter()
                        .find(|child| {
                            child.record_type_raw == INTERACTIVE_INFO_ATOM && child.data.len() >= 9
                        })
                        .map(|atom| atom.data[8]);

                    info.kind = match action {
                        Some(ACTION_OLE) => Some(EscherShapeType::Object),
                        Some(ACTION_MEDIA) => Some(EscherShapeType::Media),
                        _ => info.kind,
                    };
                },
                _ => {},
            }
            if consumed == 0 {
                break;
            }
            let Some(next) = offset.checked_add(consumed) else {
                break;
            };
            offset = next;
        }

        info
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

    fn extract_native_shape_type(container: &EscherContainer<'data>) -> Option<u16> {
        container
            .find_child(EscherRecordType::Sp)
            .map(|sp| sp.instance)
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

#[derive(Debug, Clone, Copy, Default)]
struct EscherFrameInfo {
    kind: Option<EscherShapeType>,
    external_object_id: Option<u32>,
}
