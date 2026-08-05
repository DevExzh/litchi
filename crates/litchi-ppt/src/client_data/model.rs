//! Semantic OfficeArt client-data values.

use crate::embedded::reference::Reference;

pub(super) const MAX_DEFINED_CHILDREN: usize = 13;

/// Resource limits for a shape client-data container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ClientDataLimits {
    /// Maximum enclosing payload size.
    pub max_payload_bytes: usize,
    /// Maximum payload size of any individual child.
    pub max_child_payload_bytes: usize,
    /// Maximum number of child records.
    pub max_child_records: usize,
}

impl Default for ClientDataLimits {
    fn default() -> Self {
        Self {
            max_payload_bytes: 16 * 1024 * 1024,
            max_child_payload_bytes: 16 * 1024 * 1024,
            max_child_records: MAX_DEFINED_CHILDREN,
        }
    }
}

/// Every record alternative permitted by MS-PPT §§2.7.3-2.7.4.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ClientDataChildKind {
    ShapeFlags,
    ShapeFlags10,
    ExternalObjectReference,
    AnimationInfo,
    MouseClickInteractiveInfo,
    MouseOverInteractiveInfo,
    Placeholder,
    RecolorInfo,
    ProgrammableTags,
    RoundTripNewPlaceholderId12,
    RoundTripShapeId12,
    RoundTripHeaderFooterPlaceholder12,
    RoundTripShapeChecksumForCustomLayouts12,
}

/// One validated child record, retaining its exact payload and advisory instance.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ClientDataChild {
    pub(super) kind: ClientDataChildKind,
    pub(super) version: u16,
    pub(super) instance: u16,
    pub(super) payload: Vec<u8>,
}

impl ClientDataChild {
    /// Classified child kind.
    pub fn kind(&self) -> ClientDataChildKind {
        self.kind
    }

    /// Record version retained from the input.
    pub fn version(&self) -> u16 {
        self.version
    }

    /// Record instance retained from the input.
    pub fn instance(&self) -> u16 {
        self.instance
    }

    /// Exact child payload.
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// External object ID when this is an ExObjRefAtom.
    pub fn external_object_id(&self) -> Option<u32> {
        if self.kind != ClientDataChildKind::ExternalObjectReference {
            return None;
        }
        Reference::parse_payload(&self.payload)
            .ok()
            .map(|reference| reference.id)
    }

    /// PowerPoint 12 shape ID when this is the corresponding round-trip atom.
    pub fn round_trip_shape_id(&self) -> Option<u32> {
        (self.kind == ClientDataChildKind::RoundTripShapeId12).then(|| u32_at(&self.payload, 0))
    }

    /// Shape and text checksums when this is the checksum round-trip atom.
    pub fn round_trip_checksums(&self) -> Option<(u32, u32)> {
        (self.kind == ClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12)
            .then(|| (u32_at(&self.payload, 0), u32_at(&self.payload, 4)))
    }

    /// Placeholder ID carried by either one-byte placeholder round-trip atom.
    pub fn round_trip_placeholder_id(&self) -> Option<u8> {
        matches!(
            self.kind,
            ClientDataChildKind::RoundTripNewPlaceholderId12
                | ClientDataChildKind::RoundTripHeaderFooterPlaceholder12
        )
        .then(|| self.payload[0])
    }
}

/// A complete, ordered OfficeArtClientData container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ClientData {
    pub(super) children: Vec<ClientDataChild>,
}

impl ClientData {
    /// Ordered child records.
    pub fn children(&self) -> &[ClientDataChild] {
        &self.children
    }

    /// Return the unique record of a particular kind.
    pub fn child(&self, kind: ClientDataChildKind) -> Option<&ClientDataChild> {
        self.children.iter().find(|child| child.kind == kind)
    }

    pub fn shape_flags(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::ShapeFlags)
    }

    pub fn shape_flags10(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::ShapeFlags10)
    }

    pub fn external_object_reference(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::ExternalObjectReference)
    }

    pub fn animation_info(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::AnimationInfo)
    }

    pub fn mouse_click_interactive_info(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::MouseClickInteractiveInfo)
    }

    pub fn mouse_over_interactive_info(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::MouseOverInteractiveInfo)
    }

    pub fn placeholder(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::Placeholder)
    }

    pub fn recolor_info(&self) -> Option<&ClientDataChild> {
        self.child(ClientDataChildKind::RecolorInfo)
    }
}

pub(super) fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}
