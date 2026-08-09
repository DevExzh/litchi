use std::fmt;
use std::str::FromStr;
use std::sync::Arc;

use litchi_cfb::OleError;

use super::{
    CAML_NAMESPACE, CONTENT_TYPE_NAMESPACE, COVER_PAGE_NAMESPACE, CUSTOM_PROPERTY_EDITOR_NAMESPACE,
    CUSTOM_XSN_NAMESPACE, LONG_PROPERTIES_NAMESPACE,
};

/// Resource limits applied while reading a legacy Custom XML data store.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_items: usize,
    pub max_item_bytes: usize,
    pub max_properties_bytes: usize,
    pub max_total_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_elements: usize,
    pub max_schema_references: usize,
    pub max_string_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_items: 4096,
            max_item_bytes: 16 * 1024 * 1024,
            max_properties_bytes: 1024 * 1024,
            max_total_bytes: 64 * 1024 * 1024,
            max_xml_depth: 256,
            max_xml_elements: 1_000_000,
            max_schema_references: 4096,
            max_string_bytes: 4 * 1024 * 1024,
        }
    }
}

#[derive(Debug)]
pub enum Error {
    Ole(OleError),
    Invalid(String),
    ResourceLimit(String),
    Xml(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => write!(formatter, "OLE Custom XML storage error: {error}"),
            Self::Invalid(message) => write!(formatter, "invalid Custom XML data store: {message}"),
            Self::ResourceLimit(message) => {
                write!(
                    formatter,
                    "Custom XML data-store resource limit exceeded: {message}"
                )
            },
            Self::Xml(message) => write!(formatter, "invalid Custom XML data-store XML: {message}"),
        }
    }
}

impl std::error::Error for Error {}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

pub type Result<T> = std::result::Result<T, Error>;

/// Meaning of the optional promotion marker storages.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Promotion {
    /// Neither marker exists; callers retain the IRM payload for interoperability.
    #[default]
    Unspecified,
    /// The public store is represented identically inside protected content.
    Redundant,
    /// The public store supersedes the older copy inside protected content.
    Modified,
}

/// Typed 128-bit item identifier serialized as an uppercase braced GUID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ItemId([u8; 16]);

impl ItemId {
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Produce a compact, case-insensitive-safe storage name for new items.
    #[must_use]
    pub fn storage_name(self) -> String {
        const ALPHABET: &[u8; 32] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
        let mut output = String::with_capacity(26);
        let mut accumulator = 0u32;
        let mut bits = 0u8;
        for byte in self.0 {
            accumulator = (accumulator << 8) | u32::from(byte);
            bits += 8;
            while bits >= 5 {
                bits -= 5;
                output.push(ALPHABET[((accumulator >> bits) & 0x1F) as usize] as char);
            }
        }
        if bits != 0 {
            output.push(ALPHABET[((accumulator << (5 - bits)) & 0x1F) as usize] as char);
        }
        output
    }
}

impl fmt::Display for ItemId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("{")?;
        for (index, byte) in self.0.iter().enumerate() {
            if matches!(index, 4 | 6 | 8 | 10) {
                formatter.write_str("-")?;
            }
            write!(formatter, "{byte:02X}")?;
        }
        formatter.write_str("}")
    }
}

impl FromStr for ItemId {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let Some(inner) = value
            .strip_prefix('{')
            .and_then(|candidate| candidate.strip_suffix('}'))
        else {
            return Err(invalid("itemID is not a braced GUID"));
        };
        if inner.len() != 36
            || inner.bytes().enumerate().any(|(index, byte)| match index {
                8 | 13 | 18 | 23 => byte != b'-',
                _ => !byte.is_ascii_hexdigit(),
            })
        {
            return Err(invalid("itemID is not a braced GUID"));
        }
        let mut bytes = [0; 16];
        for (nibble, byte) in inner.bytes().filter(|byte| *byte != b'-').enumerate() {
            let nibble_value = match byte {
                b'0'..=b'9' => byte - b'0',
                b'a'..=b'f' => byte - b'a' + 10,
                b'A'..=b'F' => byte - b'A' + 10,
                _ => return Err(invalid("itemID contains a non-hexadecimal digit")),
            };
            if nibble.is_multiple_of(2) {
                bytes[nibble / 2] = nibble_value << 4;
            } else {
                bytes[nibble / 2] |= nibble_value;
            }
        }
        Ok(Self(bytes))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    pub item_id: ItemId,
    /// Schema target namespaces. They are retained but never resolved.
    pub schema_references: Vec<String>,
}

/// Expanded root name of an inert Custom XML item.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RootName {
    pub namespace: Option<String>,
    pub local_name: String,
}

/// Known MS-OSHARED application-defined item family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ItemKind {
    CustomPropertyEditor,
    CustomXsn,
    ContentType,
    CoverPageProperties,
    LongProperties,
    CollaborativeApplicationMarkup,
    Other,
}

/// One exactly paired `Item` / `Properties` sub-storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Item {
    pub(crate) storage_name: String,
    pub(crate) xml: Arc<[u8]>,
    pub(crate) root_name: RootName,
    pub(crate) properties_xml: Arc<[u8]>,
    pub(crate) properties: Properties,
}

impl Item {
    /// Construct an item after validating both its payload and properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the storage name, item XML, or properties violate
    /// the default Custom XML limits.
    pub fn new(
        storage_name: impl Into<String>,
        xml: Vec<u8>,
        properties: Properties,
    ) -> Result<Self> {
        let limits = Limits::default();
        let name = storage_name.into();
        super::codec::validate_storage_name(&name)?;
        let root_name = super::xml::validate_payload(&xml, &limits)?;
        super::codec::validate_properties(&properties, &limits)?;
        let properties_xml = super::codec::write_properties(&properties)?;
        Ok(Self {
            storage_name: name,
            xml: Arc::from(xml),
            root_name,
            properties_xml: Arc::from(properties_xml),
            properties,
        })
    }

    #[must_use]
    pub fn storage_name(&self) -> &str {
        &self.storage_name
    }

    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    #[must_use]
    pub fn root_name(&self) -> &RootName {
        &self.root_name
    }

    #[must_use]
    pub fn kind(&self) -> ItemKind {
        match self.root_name.namespace.as_deref() {
            Some(CUSTOM_PROPERTY_EDITOR_NAMESPACE) => ItemKind::CustomPropertyEditor,
            Some(CUSTOM_XSN_NAMESPACE) => ItemKind::CustomXsn,
            Some(CONTENT_TYPE_NAMESPACE) => ItemKind::ContentType,
            Some(COVER_PAGE_NAMESPACE) => ItemKind::CoverPageProperties,
            Some(LONG_PROPERTIES_NAMESPACE) => ItemKind::LongProperties,
            Some(CAML_NAMESPACE) => ItemKind::CollaborativeApplicationMarkup,
            _ => ItemKind::Other,
        }
    }

    #[must_use]
    pub fn properties_xml(&self) -> &[u8] {
        &self.properties_xml
    }

    #[must_use]
    pub fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Replace the item XML after validating its bounded root projection.
    ///
    /// # Errors
    ///
    /// Returns an error if `xml` is malformed or exceeds the default limits.
    pub fn set_xml(&mut self, xml: Vec<u8>) -> Result<()> {
        let root_name = super::xml::validate_payload(&xml, &Limits::default())?;
        self.xml = Arc::from(xml);
        self.root_name = root_name;
        Ok(())
    }

    /// Replace the typed Properties projection and serialize it deterministically.
    ///
    /// # Errors
    ///
    /// Returns an error if `properties` violates the default format limits.
    pub fn set_properties(&mut self, properties: Properties) -> Result<()> {
        super::codec::validate_properties(&properties, &Limits::default())?;
        self.properties_xml = Arc::from(super::codec::write_properties(&properties)?);
        self.properties = properties;
        Ok(())
    }
}

/// Complete legacy Custom XML data store and its promotion state.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Store {
    pub promotion: Promotion,
    pub(crate) items: Vec<Item>,
}

impl Store {
    /// Construct a complete validated store.
    ///
    /// # Errors
    ///
    /// Returns an error if an item or its aggregate store state violates the
    /// default Custom XML limits.
    pub fn new(promotion: Promotion, items: Vec<Item>) -> Result<Self> {
        let value = Self { promotion, items };
        super::codec::validate_store(&value, &Limits::default())?;
        Ok(value)
    }

    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    pub fn items_mut(&mut self) -> &mut [Item] {
        &mut self.items
    }

    /// Append one item while preserving unique storage names and identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error if `item` duplicates an existing name or identifier,
    /// or would exceed the default item limit.
    pub fn push(&mut self, item: Item) -> Result<()> {
        if self.items.iter().any(|existing| {
            existing
                .storage_name
                .eq_ignore_ascii_case(&item.storage_name)
                || existing.properties.item_id == item.properties.item_id
        }) {
            return Err(invalid("custom XML storage name or itemID is duplicated"));
        }
        if self.items.len() >= Limits::default().max_items {
            return Err(limit("item count exceeds the default limit"));
        }
        self.items.push(item);
        Ok(())
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(message: impl Into<String>) -> Error {
    Error::ResourceLimit(message.into())
}

pub(crate) fn xml_error(message: impl Into<String>) -> Error {
    Error::Xml(message.into())
}
