//! Package-neutral custom-data properties model.

use std::sync::Arc;

/// Treatment of connection references when removing a Custom Data storage.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub enum RemovalDisposition {
    /// Refuse removal when connections still refer to this storage.
    #[default]
    RejectReferenced,
    /// Clear embeddedDataId values while preserving each connection.
    DetachConnections,
    /// Redirect references to the given remaining storage UID.
    RetargetConnections(String),
}

/// One self-contained `x14:extLst` subtree, retained without interpretation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub xml: Vec<u8>,
}

/// Typed properties for one embedded custom-data storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    pub id: String,
    pub extension_list: Option<ExtensionList>,
}

/// One complete custom-data storage.  The binary payload is inert bytes; the
/// ordinary API never interprets or executes it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomData {
    pub properties: Properties,
    pub data: Vec<u8>,
}

impl CustomData {
    /// Construct a storage with a typed UID and inert payload.
    #[must_use]
    pub fn new(id: impl Into<String>, data: Vec<u8>) -> Self {
        Self {
            properties: Properties {
                id: id.into(),
                extension_list: None,
            },
            data,
        }
    }
}

/// Compatibility name for a complete custom-data storage.
pub type Store = CustomData;

/// Immutable custom-data metadata and payload backed by shared allocations.
/// Cloning this view never copies the payload or extension XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomDataView {
    pub(crate) properties: Arc<Properties>,
    pub(crate) data: Arc<Vec<u8>>,
}

impl CustomDataView {
    /// Borrow the storage properties.
    #[must_use]
    pub fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Borrow the exact inert payload bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Explicitly copy the metadata and complete payload into an owned value.
    #[must_use]
    pub fn to_owned(&self) -> CustomData {
        CustomData {
            properties: self.properties.as_ref().clone(),
            data: self.data.as_ref().clone(),
        }
    }
}

impl From<CustomData> for CustomDataView {
    fn from(value: CustomData) -> Self {
        Self {
            properties: Arc::new(value.properties),
            data: Arc::new(value.data),
        }
    }
}
