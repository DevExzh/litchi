//! Package-neutral Custom Data values.
//!
//! The payload in this module is deliberately inert.  Package owners retain
//! relationship, source, and transaction state at their own boundaries;
//! these values contain only the X14 properties and the associated bytes.

use std::sync::Arc;

/// Treatment of host references when removing a Custom Data storage.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub enum RemovalDisposition {
    /// Refuse removal while a recognized host reference names the storage.
    #[default]
    RejectReferenced,
    /// Clear recognized references before removing the storage.
    DetachConnections,
    /// Rewrite recognized references to an existing remaining storage.
    RetargetConnections(String),
}

/// One complete, source-preserving X14 extension-list fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    /// The complete `<extLst>` element, including its original lexical bytes.
    pub xml: Vec<u8>,
}

/// Typed properties for one embedded Custom Data storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    /// The required `ST_Xstring` UID.  The empty value is legal.
    pub id: String,
    /// The optional direct X14 extension list.
    pub extension_list: Option<ExtensionList>,
}

/// One self-contained Custom Data storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomData {
    /// Typed properties for the associated payload.
    pub properties: Properties,
    /// Inert binary payload bytes.
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

/// Compatibility name for a complete Custom Data storage.
pub type Store = CustomData;

/// Immutable Custom Data metadata and payload backed by shared allocations.
///
/// The fields remain public for the existing XLSX adapter's source-bound
/// staging path.  Callers should prefer [`Self::properties`], [`Self::data`],
/// and [`Self::to_owned`] when working through the common facade.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomDataView {
    /// Shared typed properties.
    pub properties: Arc<Properties>,
    /// Shared inert payload.
    pub data: Arc<Vec<u8>>,
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

    /// Explicitly copy metadata and payload into an owned value.
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
