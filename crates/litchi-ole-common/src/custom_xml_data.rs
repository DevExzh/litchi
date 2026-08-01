//! Legacy Office Custom XML Data Storage (`MsoDataStore`).
//!
//! Item XML is validated structurally and retained verbatim. Schema references
//! are metadata only: this module never resolves a URI, loads a schema, expands
//! an external entity, or executes application-specific XML.

use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt;
use std::io::{Read, Seek};
use std::str::FromStr;

use litchi_cfb::{OleError, OleFile, OleWriter};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub const MSO_DATA_STORE_STORAGE: &str = "MsoDataStore";
pub const REDUNDANT_PROMOTION_STORAGE: &str = "IsRedundantDataStorePromotion";
pub const MODIFIED_PROMOTION_STORAGE: &str = "IsModifiedDataStorePromotion";
pub const ITEM_STREAM: &str = "Item";
pub const PROPERTIES_STREAM: &str = "Properties";
pub const CUSTOM_XML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

const CUSTOM_PROPERTY_EDITOR_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/customDocumentInformationPanel";
const CUSTOM_XSN_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/metadata/customXsn";
const CONTENT_TYPE_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/metadata/contentType";
const COVER_PAGE_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/coverPageProps";
const LONG_PROPERTIES_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/2006/metadata/longProperties";
const CAML_NAMESPACE: &str = "office.server.policy";

/// Resource limits applied while reading a legacy Custom XML Data Storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MsoDataStoreLimits {
    pub max_items: usize,
    pub max_item_bytes: usize,
    pub max_properties_bytes: usize,
    pub max_total_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_elements: usize,
    pub max_schema_references: usize,
    pub max_string_bytes: usize,
}

impl Default for MsoDataStoreLimits {
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
pub enum MsoDataStoreError {
    Ole(OleError),
    Invalid(String),
    ResourceLimit(String),
    Xml(String),
}

impl fmt::Display for MsoDataStoreError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ole(error) => write!(formatter, "OLE Custom XML storage error: {error}"),
            Self::Invalid(message) => write!(formatter, "invalid MsoDataStore: {message}"),
            Self::ResourceLimit(message) => {
                write!(formatter, "MsoDataStore resource limit exceeded: {message}")
            },
            Self::Xml(message) => write!(formatter, "invalid MsoDataStore XML: {message}"),
        }
    }
}

impl std::error::Error for MsoDataStoreError {}

impl From<OleError> for MsoDataStoreError {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

pub type Result<T> = std::result::Result<T, MsoDataStoreError>;

/// Meaning of the optional MS-OFFCRYPTO promotion marker storages.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataStorePromotion {
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
pub struct CustomXmlItemId([u8; 16]);

impl CustomXmlItemId {
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Produce a compact, case-insensitive-safe storage name for new items.
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

impl fmt::Display for CustomXmlItemId {
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

impl FromStr for CustomXmlItemId {
    type Err = MsoDataStoreError;

    fn from_str(value: &str) -> Result<Self> {
        let Some(inner) = value
            .strip_prefix('{')
            .and_then(|value| value.strip_suffix('}'))
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
            let value = match byte {
                b'0'..=b'9' => byte - b'0',
                b'a'..=b'f' => byte - b'a' + 10,
                b'A'..=b'F' => byte - b'A' + 10,
                _ => return Err(invalid("itemID contains a non-hexadecimal digit")),
            };
            if nibble.is_multiple_of(2) {
                bytes[nibble / 2] = value << 4;
            } else {
                bytes[nibble / 2] |= value;
            }
        }
        Ok(Self(bytes))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlDataProperties {
    pub item_id: CustomXmlItemId,
    /// Schema target namespaces. They are retained but never resolved.
    pub schema_references: Vec<String>,
}

/// Expanded root name of an inert Custom XML item.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlRootName {
    pub namespace: Option<String>,
    pub local_name: String,
}

/// Known MS-OSHARED application-defined item family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomXmlItemKind {
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
pub struct CustomXmlDataItem {
    storage_name: String,
    xml: Vec<u8>,
    root_name: CustomXmlRootName,
    properties_xml: Vec<u8>,
    properties: CustomXmlDataProperties,
}

impl CustomXmlDataItem {
    pub fn new(
        storage_name: impl Into<String>,
        xml: Vec<u8>,
        properties: CustomXmlDataProperties,
    ) -> Result<Self> {
        let limits = MsoDataStoreLimits::default();
        let storage_name = storage_name.into();
        validate_storage_name(&storage_name)?;
        let root_name = validate_custom_xml_payload(&xml, &limits)?;
        validate_properties(&properties, &limits)?;
        let properties_xml = write_custom_xml_properties(&properties)?;
        Ok(Self {
            storage_name,
            xml,
            root_name,
            properties_xml,
            properties,
        })
    }

    pub fn storage_name(&self) -> &str {
        &self.storage_name
    }

    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    pub fn root_name(&self) -> &CustomXmlRootName {
        &self.root_name
    }

    pub fn kind(&self) -> CustomXmlItemKind {
        match self.root_name.namespace.as_deref() {
            Some(CUSTOM_PROPERTY_EDITOR_NAMESPACE) => CustomXmlItemKind::CustomPropertyEditor,
            Some(CUSTOM_XSN_NAMESPACE) => CustomXmlItemKind::CustomXsn,
            Some(CONTENT_TYPE_NAMESPACE) => CustomXmlItemKind::ContentType,
            Some(COVER_PAGE_NAMESPACE) => CustomXmlItemKind::CoverPageProperties,
            Some(LONG_PROPERTIES_NAMESPACE) => CustomXmlItemKind::LongProperties,
            Some(CAML_NAMESPACE) => CustomXmlItemKind::CollaborativeApplicationMarkup,
            _ => CustomXmlItemKind::Other,
        }
    }

    pub fn properties_xml(&self) -> &[u8] {
        &self.properties_xml
    }

    pub fn properties(&self) -> &CustomXmlDataProperties {
        &self.properties
    }

    pub fn set_xml(&mut self, xml: Vec<u8>) -> Result<()> {
        let root_name = validate_custom_xml_payload(&xml, &MsoDataStoreLimits::default())?;
        self.xml = xml;
        self.root_name = root_name;
        Ok(())
    }

    pub fn set_properties(&mut self, properties: CustomXmlDataProperties) -> Result<()> {
        validate_properties(&properties, &MsoDataStoreLimits::default())?;
        self.properties_xml = write_custom_xml_properties(&properties)?;
        self.properties = properties;
        Ok(())
    }
}

/// Complete legacy Custom XML Data Storage and its promotion state.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct MsoDataStore {
    pub promotion: DataStorePromotion,
    items: Vec<CustomXmlDataItem>,
}

impl MsoDataStore {
    pub fn new(promotion: DataStorePromotion, items: Vec<CustomXmlDataItem>) -> Result<Self> {
        let value = Self { promotion, items };
        validate_store(&value, &MsoDataStoreLimits::default())?;
        Ok(value)
    }

    pub fn items(&self) -> &[CustomXmlDataItem] {
        &self.items
    }

    pub fn items_mut(&mut self) -> &mut [CustomXmlDataItem] {
        &mut self.items
    }

    pub fn push(&mut self, item: CustomXmlDataItem) -> Result<()> {
        if self.items.iter().any(|existing| {
            existing
                .storage_name
                .eq_ignore_ascii_case(&item.storage_name)
                || existing.properties.item_id == item.properties.item_id
        }) {
            return Err(invalid("custom XML storage name or itemID is duplicated"));
        }
        if self.items.len() >= MsoDataStoreLimits::default().max_items {
            return Err(limit("item count exceeds the default limit"));
        }
        self.items.push(item);
        Ok(())
    }
}

/// Inspect a complete legacy Custom XML Data Storage with default limits.
pub fn inspect_mso_data_store<R: Read + Seek>(
    ole: &mut OleFile<R>,
) -> Result<Option<MsoDataStore>> {
    inspect_mso_data_store_with_limits(ole, MsoDataStoreLimits::default())
}

/// Inspect a complete legacy Custom XML Data Storage with caller-selected limits.
pub fn inspect_mso_data_store_with_limits<R: Read + Seek>(
    ole: &mut OleFile<R>,
    limits: MsoDataStoreLimits,
) -> Result<Option<MsoDataStore>> {
    validate_limits(&limits)?;
    let redundant = marker_exists(ole, REDUNDANT_PROMOTION_STORAGE)?;
    let modified = marker_exists(ole, MODIFIED_PROMOTION_STORAGE)?;
    if redundant && modified {
        return Err(invalid(
            "redundant and modified promotion storages are both present",
        ));
    }
    let promotion = if redundant {
        DataStorePromotion::Redundant
    } else if modified {
        DataStorePromotion::Modified
    } else {
        DataStorePromotion::Unspecified
    };

    if !ole.directory_exists(&[MSO_DATA_STORE_STORAGE]) {
        if ole.exists(&[MSO_DATA_STORE_STORAGE]) {
            return Err(invalid("MsoDataStore is not a storage"));
        }
        if promotion != DataStorePromotion::Unspecified {
            return Err(invalid("promotion marker exists without MsoDataStore"));
        }
        return Ok(None);
    }

    let entries = ole.list_directory_entries(&[MSO_DATA_STORE_STORAGE])?;
    if entries.len() > limits.max_items {
        return Err(limit(format!(
            "item count {} exceeds {}",
            entries.len(),
            limits.max_items
        )));
    }
    let mut names = Vec::with_capacity(entries.len());
    for entry in entries {
        if entry.entry_type != 1 {
            return Err(invalid(format!(
                "MsoDataStore child '{}' is not a storage",
                entry.name
            )));
        }
        validate_storage_name(&entry.name)?;
        names.push(entry.name.clone());
    }
    names.sort();

    let mut total_bytes = 0usize;
    let mut item_ids = HashSet::with_capacity(names.len());
    let mut items = Vec::with_capacity(names.len());
    for storage_name in names {
        let entries =
            ole.list_directory_entries(&[MSO_DATA_STORE_STORAGE, storage_name.as_str()])?;
        if entries.len() != 2 {
            return Err(invalid(format!(
                "custom XML sub-storage '{storage_name}' must contain exactly Item and Properties"
            )));
        }
        let mut item_size = None;
        let mut properties_size = None;
        for entry in entries {
            match (entry.name.as_str(), entry.entry_type) {
                (ITEM_STREAM, 2) => item_size = Some(stream_size(entry.size, "Item", &limits)?),
                (PROPERTIES_STREAM, 2) => {
                    properties_size = Some(stream_size(entry.size, "Properties", &limits)?)
                },
                _ => {
                    return Err(invalid(format!(
                        "custom XML sub-storage '{storage_name}' has unexpected entry '{}'",
                        entry.name
                    )));
                },
            }
        }
        let item_size = item_size.ok_or_else(|| invalid("custom XML Item stream is missing"))?;
        let properties_size =
            properties_size.ok_or_else(|| invalid("custom XML Properties stream is missing"))?;
        if item_size > limits.max_item_bytes {
            return Err(limit(format!(
                "Item stream has {item_size} bytes, limit is {}",
                limits.max_item_bytes
            )));
        }
        if properties_size > limits.max_properties_bytes {
            return Err(limit(format!(
                "Properties stream has {properties_size} bytes, limit is {}",
                limits.max_properties_bytes
            )));
        }
        total_bytes = total_bytes
            .checked_add(item_size)
            .and_then(|value| value.checked_add(properties_size))
            .ok_or_else(|| limit("aggregate stream size overflows usize"))?;
        if total_bytes > limits.max_total_bytes {
            return Err(limit(format!(
                "aggregate stream bytes exceed {}",
                limits.max_total_bytes
            )));
        }
        let xml = ole.open_stream(&[MSO_DATA_STORE_STORAGE, storage_name.as_str(), ITEM_STREAM])?;
        let properties_xml = ole.open_stream(&[
            MSO_DATA_STORE_STORAGE,
            storage_name.as_str(),
            PROPERTIES_STREAM,
        ])?;
        let root_name = validate_custom_xml_payload(&xml, &limits)?;
        let properties = parse_custom_xml_properties_with_limits(&properties_xml, &limits)?;
        if !item_ids.insert(properties.item_id) {
            return Err(invalid(format!(
                "itemID {} is used by multiple custom XML items",
                properties.item_id
            )));
        }
        items.push(CustomXmlDataItem {
            storage_name,
            xml,
            root_name,
            properties_xml,
            properties,
        });
    }
    Ok(Some(MsoDataStore { promotion, items }))
}

/// Materialize a validated store in a newly assembled OLE writer.
pub fn write_mso_data_store(writer: &mut OleWriter, store: &MsoDataStore) -> Result<()> {
    validate_store(store, &MsoDataStoreLimits::default())?;
    writer.create_storage(&[MSO_DATA_STORE_STORAGE])?;
    match store.promotion {
        DataStorePromotion::Unspecified => {},
        DataStorePromotion::Redundant => {
            writer.create_storage(&[REDUNDANT_PROMOTION_STORAGE])?;
        },
        DataStorePromotion::Modified => {
            writer.create_storage(&[MODIFIED_PROMOTION_STORAGE])?;
        },
    }
    for item in &store.items {
        writer.create_storage(&[MSO_DATA_STORE_STORAGE, item.storage_name.as_str()])?;
        writer.create_stream(
            &[
                MSO_DATA_STORE_STORAGE,
                item.storage_name.as_str(),
                ITEM_STREAM,
            ],
            &item.xml,
        )?;
        writer.create_stream(
            &[
                MSO_DATA_STORE_STORAGE,
                item.storage_name.as_str(),
                PROPERTIES_STREAM,
            ],
            &item.properties_xml,
        )?;
    }
    Ok(())
}

/// Parse the schema-defined Custom XML Data Storage Properties stream.
pub fn parse_custom_xml_properties(xml: &[u8]) -> Result<CustomXmlDataProperties> {
    parse_custom_xml_properties_with_limits(xml, &MsoDataStoreLimits::default())
}

/// Serialize Custom XML Data Storage Properties in stable schema order.
pub fn write_custom_xml_properties(properties: &CustomXmlDataProperties) -> Result<Vec<u8>> {
    validate_properties(properties, &MsoDataStoreLimits::default())?;
    let mut output = Vec::with_capacity(256);
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    output.extend_from_slice(CUSTOM_XML_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\" ds:itemID=\"");
    output.extend_from_slice(properties.item_id.to_string().as_bytes());
    output.extend_from_slice(b"\">");
    if !properties.schema_references.is_empty() {
        output.extend_from_slice(b"<ds:schemaRefs>");
        for uri in &properties.schema_references {
            output.extend_from_slice(b"<ds:schemaRef ds:uri=\"");
            push_escaped_attribute(&mut output, uri);
            output.extend_from_slice(b"\"/>");
        }
        output.extend_from_slice(b"</ds:schemaRefs>");
    }
    output.extend_from_slice(b"</ds:datastoreItem>");
    Ok(output)
}

fn parse_custom_xml_properties_with_limits(
    xml: &[u8],
    limits: &MsoDataStoreLimits,
) -> Result<CustomXmlDataProperties> {
    if xml.len() > limits.max_properties_bytes {
        return Err(limit("Properties XML exceeds its byte limit"));
    }
    let normalized = normalize_xml_encoding(xml)?;
    let mut reader = NsReader::from_reader(normalized.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut state = PropertiesParseState::default();

    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| xml_error(error.to_string()))?
        {
            Event::Start(element) => {
                process_properties_element(&reader, &element, limits, &mut state)?;
                state.depth += 1;
                if state.depth > limits.max_xml_depth {
                    return Err(limit("Properties XML depth exceeds its limit"));
                }
            },
            Event::Empty(element) => {
                process_properties_element(&reader, &element, limits, &mut state)?;
                if state.depth == 0 {
                    state.root_closed = true;
                }
            },
            Event::End(_) => {
                state.depth = state
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Properties XML has an unexpected closing tag"))?;
                if state.depth == 0 {
                    state.root_closed = true;
                }
            },
            Event::Text(text) if !is_xml_whitespace(text.as_ref()) => {
                return Err(invalid("Properties XML contains text content"));
            },
            Event::CData(text) if !is_xml_whitespace(text.as_ref()) => {
                return Err(invalid("Properties XML contains CDATA content"));
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Properties XML",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !state.root_seen || !state.root_closed || state.depth != 0 {
        return Err(invalid("Properties XML has no complete datastoreItem root"));
    }
    let properties = CustomXmlDataProperties {
        item_id: state
            .item_id
            .ok_or_else(|| invalid("datastoreItem lacks itemID"))?,
        schema_references: state.schema_references,
    };
    validate_properties(&properties, limits)?;
    Ok(properties)
}

#[derive(Default)]
struct PropertiesParseState {
    depth: usize,
    root_seen: bool,
    root_closed: bool,
    schema_refs_seen: bool,
    item_id: Option<CustomXmlItemId>,
    schema_references: Vec<String>,
    element_count: usize,
}

fn process_properties_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &MsoDataStoreLimits,
    state: &mut PropertiesParseState,
) -> Result<()> {
    let (namespace, local_name) = resolved_element(reader, element)?;
    state.element_count += 1;
    if state.element_count > limits.max_xml_elements {
        return Err(limit("Properties XML element count exceeds its limit"));
    }
    match (state.depth, local_name.as_slice(), namespace.as_deref()) {
        (0, b"datastoreItem", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes() && !state.root_seen =>
        {
            state.item_id = Some(
                required_attribute(reader, element, CUSTOM_XML_NAMESPACE, b"itemID")?.parse()?,
            );
            reject_other_attributes(reader, element, &[(CUSTOM_XML_NAMESPACE, b"itemID")])?;
            state.root_seen = true;
        },
        (1, b"schemaRefs", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes()
                && state.root_seen
                && !state.schema_refs_seen =>
        {
            reject_other_attributes(reader, element, &[])?;
            state.schema_refs_seen = true;
        },
        (2, b"schemaRef", Some(namespace))
            if namespace == CUSTOM_XML_NAMESPACE.as_bytes() && state.schema_refs_seen =>
        {
            if state.schema_references.len() >= limits.max_schema_references {
                return Err(limit("schema reference count exceeds its limit"));
            }
            state.schema_references.push(required_attribute(
                reader,
                element,
                CUSTOM_XML_NAMESPACE,
                b"uri",
            )?);
            reject_other_attributes(reader, element, &[(CUSTOM_XML_NAMESPACE, b"uri")])?;
        },
        _ => return Err(invalid("Properties XML violates datastoreItem grammar")),
    }
    Ok(())
}

fn validate_custom_xml_payload(
    xml: &[u8],
    limits: &MsoDataStoreLimits,
) -> Result<CustomXmlRootName> {
    if xml.is_empty() || xml.len() > limits.max_item_bytes {
        return Err(limit("Item XML is empty or exceeds its byte limit"));
    }
    let normalized = normalize_xml_encoding(xml)?;
    let mut reader = NsReader::from_reader(normalized.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut root_name = None;
    let mut elements = 0usize;
    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| xml_error(error.to_string()))?
        {
            Event::Start(element) => {
                elements += 1;
                if elements > limits.max_xml_elements {
                    return Err(limit("Item XML element count exceeds its limit"));
                }
                if depth == 0 {
                    roots += 1;
                    root_name = Some(expanded_root(&reader, &element)?);
                }
                depth += 1;
                if depth > limits.max_xml_depth {
                    return Err(limit("Item XML depth exceeds its limit"));
                }
            },
            Event::Empty(element) => {
                elements += 1;
                if elements > limits.max_xml_elements {
                    return Err(limit("Item XML element count exceeds its limit"));
                }
                if depth == 0 {
                    roots += 1;
                    root_name = Some(expanded_root(&reader, &element)?);
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Item XML has an unexpected closing tag"))?;
            },
            Event::Text(text) if depth == 0 && !is_xml_whitespace(text.as_ref()) => {
                return Err(invalid("Item XML has text outside its root"));
            },
            Event::CData(text) if depth == 0 && !is_xml_whitespace(text.as_ref()) => {
                return Err(invalid("Item XML has CDATA outside its root"));
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and general entity references are forbidden in Item XML",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid("Item XML must have exactly one complete root"));
    }
    root_name.ok_or_else(|| invalid("Item XML has no root"))
}

fn expanded_root(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<CustomXmlRootName> {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    let namespace = match resolved {
        ResolveResult::Bound(Namespace(value)) => Some(
            std::str::from_utf8(value)
                .map_err(|_| xml_error("root namespace is not UTF-8"))?
                .to_string(),
        ),
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => {
            return Err(xml_error(format!(
                "root uses unknown namespace prefix {:?}",
                prefix
            )));
        },
    };
    let local_name = std::str::from_utf8(local_name.as_ref())
        .map_err(|_| xml_error("root local name is not UTF-8"))?
        .to_string();
    Ok(CustomXmlRootName {
        namespace,
        local_name,
    })
}

fn validate_store(store: &MsoDataStore, limits: &MsoDataStoreLimits) -> Result<()> {
    validate_limits(limits)?;
    if store.items.len() > limits.max_items {
        return Err(limit("item count exceeds its limit"));
    }
    let mut names = HashSet::with_capacity(store.items.len());
    let mut ids = HashSet::with_capacity(store.items.len());
    let mut total = 0usize;
    for item in &store.items {
        validate_storage_name(&item.storage_name)?;
        if !names.insert(item.storage_name.to_uppercase()) {
            return Err(invalid("custom XML storage name is duplicated"));
        }
        if !ids.insert(item.properties.item_id) {
            return Err(invalid("custom XML itemID is duplicated"));
        }
        let root = validate_custom_xml_payload(&item.xml, limits)?;
        if root != item.root_name {
            return Err(invalid("cached Item root name disagrees with Item XML"));
        }
        let properties = parse_custom_xml_properties_with_limits(&item.properties_xml, limits)?;
        if properties != item.properties {
            return Err(invalid(
                "typed properties disagree with preserved Properties XML",
            ));
        }
        total = total
            .checked_add(item.xml.len())
            .and_then(|value| value.checked_add(item.properties_xml.len()))
            .ok_or_else(|| limit("aggregate stream size overflows usize"))?;
        if total > limits.max_total_bytes {
            return Err(limit("aggregate stream size exceeds its limit"));
        }
    }
    Ok(())
}

fn validate_properties(
    properties: &CustomXmlDataProperties,
    limits: &MsoDataStoreLimits,
) -> Result<()> {
    if properties.schema_references.len() > limits.max_schema_references {
        return Err(limit("schema reference count exceeds its limit"));
    }
    let total =
        properties
            .schema_references
            .iter()
            .try_fold(0usize, |total, value| -> Result<usize> {
                validate_xml_characters(value)?;
                total
                    .checked_add(value.len())
                    .ok_or_else(|| limit("schema reference strings overflow usize"))
            })?;
    if total > limits.max_string_bytes {
        return Err(limit("schema reference strings exceed their byte limit"));
    }
    Ok(())
}

fn validate_storage_name(value: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if value.is_empty() || units > 31 || value.contains('\0') {
        return Err(invalid(
            "custom XML sub-storage name is empty, too long, or contains NUL",
        ));
    }
    Ok(())
}

fn validate_limits(limits: &MsoDataStoreLimits) -> Result<()> {
    if limits.max_item_bytes == 0
        || limits.max_properties_bytes == 0
        || limits.max_total_bytes == 0
        || limits.max_xml_depth == 0
        || limits.max_xml_elements == 0
        || limits.max_string_bytes == 0
    {
        return Err(limit(
            "configured byte, depth, and element limits must be nonzero",
        ));
    }
    Ok(())
}

fn normalize_xml_encoding(xml: &[u8]) -> Result<Cow<'_, [u8]>> {
    let (encoding, bytes) = if let Some(bytes) = xml.strip_prefix(&[0xFF, 0xFE]) {
        (Some(Utf16Encoding::LittleEndian), bytes)
    } else if let Some(bytes) = xml.strip_prefix(&[0xFE, 0xFF]) {
        (Some(Utf16Encoding::BigEndian), bytes)
    } else if xml.starts_with(&[b'<', 0, b'?', 0])
        || xml.starts_with(&[b'<', 0, b'!', 0])
        || xml.starts_with(&[b'<', 0])
    {
        (Some(Utf16Encoding::LittleEndian), xml)
    } else if xml.starts_with(&[0, b'<', 0, b'?'])
        || xml.starts_with(&[0, b'<', 0, b'!'])
        || xml.starts_with(&[0, b'<'])
    {
        (Some(Utf16Encoding::BigEndian), xml)
    } else {
        (None, xml)
    };
    let Some(encoding) = encoding else {
        let text =
            std::str::from_utf8(xml).map_err(|_| xml_error("XML is not valid UTF-8 or UTF-16"))?;
        validate_xml_characters(text)?;
        return Ok(Cow::Borrowed(xml));
    };
    if bytes.len() % 2 != 0 {
        return Err(xml_error("UTF-16 XML has an odd byte length"));
    }
    let units = bytes
        .chunks_exact(2)
        .map(|pair| match encoding {
            Utf16Encoding::LittleEndian => u16::from_le_bytes([pair[0], pair[1]]),
            Utf16Encoding::BigEndian => u16::from_be_bytes([pair[0], pair[1]]),
        })
        .collect::<Vec<_>>();
    let text = String::from_utf16(&units)
        .map_err(|_| xml_error("UTF-16 XML is not well-formed Unicode"))?;
    validate_xml_characters(&text)?;
    Ok(Cow::Owned(text.into_bytes()))
}

fn validate_xml_characters(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(xml_error("XML contains a character forbidden by XML 1.0"));
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Utf16Encoding {
    LittleEndian,
    BigEndian,
}

fn stream_size(value: u64, label: &str, limits: &MsoDataStoreLimits) -> Result<usize> {
    let value =
        usize::try_from(value).map_err(|_| limit(format!("{label} size overflows usize")))?;
    if value > limits.max_total_bytes {
        return Err(limit(format!("{label} size exceeds aggregate byte limit")));
    }
    Ok(value)
}

fn marker_exists<R: Read + Seek>(ole: &OleFile<R>, name: &str) -> Result<bool> {
    if ole.directory_exists(&[name]) {
        return Ok(true);
    }
    if ole.exists(&[name]) {
        return Err(invalid(format!("{name} is not a storage")));
    }
    Ok(false)
}

fn resolved_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(Option<Vec<u8>>, Vec<u8>)> {
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = match resolved {
        ResolveResult::Bound(namespace) => Some(namespace.as_ref().to_vec()),
        ResolveResult::Unbound => None,
        ResolveResult::Unknown(prefix) => {
            return Err(xml_error(format!(
                "unknown element namespace prefix {:?}",
                prefix
            )));
        },
    };
    Ok((namespace, local.as_ref().to_vec()))
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local_name: &[u8],
) -> Result<String> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == local_name
            && matches!(resolved, ResolveResult::Bound(bound) if bound.as_ref() == namespace.as_bytes())
        {
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
            if value.replace(decoded).is_some() {
                return Err(invalid("XML element has a duplicate attribute"));
            }
        }
    }
    value.ok_or_else(|| {
        invalid(format!(
            "XML element lacks required attribute {}",
            String::from_utf8_lossy(local_name)
        ))
    })
}

fn reject_other_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    allowed: &[(&str, &[u8])],
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let is_allowed = match resolved {
            ResolveResult::Bound(namespace) => {
                allowed.iter().any(|(allowed_namespace, allowed_local)| {
                    namespace.as_ref() == allowed_namespace.as_bytes()
                        && local.as_ref() == *allowed_local
                })
            },
            ResolveResult::Unbound => allowed.iter().any(|(allowed_namespace, allowed_local)| {
                allowed_namespace.is_empty() && local.as_ref() == *allowed_local
            }),
            ResolveResult::Unknown(_) => {
                return Err(invalid("XML attribute uses an unknown namespace prefix"));
            },
        };
        if !is_allowed {
            return Err(invalid(format!(
                "XML element has unexpected attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn push_escaped_attribute(output: &mut Vec<u8>, value: &str) {
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            _ => output.push(byte),
        }
    }
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid(message: impl Into<String>) -> MsoDataStoreError {
    MsoDataStoreError::Invalid(message.into())
}

fn limit(message: impl Into<String>) -> MsoDataStoreError {
    MsoDataStoreError::ResourceLimit(message.into())
}

fn xml_error(message: impl Into<String>) -> MsoDataStoreError {
    MsoDataStoreError::Xml(message.into())
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;
    use std::path::PathBuf;

    use super::*;

    const ITEM_ID: &str = "{E0FCA697-D525-4175-A08B-DFD1F1FC7C9F}";

    fn properties() -> CustomXmlDataProperties {
        CustomXmlDataProperties {
            item_id: ITEM_ID.parse().unwrap(),
            schema_references: vec![
                "http://schemas.openxmlformats.org/officeDocument/2006/bibliography".to_string(),
            ],
        }
    }

    #[test]
    fn property_xml_and_compact_storage_name_round_trip() {
        let value = properties();
        let xml = write_custom_xml_properties(&value).unwrap();
        assert_eq!(parse_custom_xml_properties(&xml).unwrap(), value);
        assert_eq!(value.item_id.to_string(), ITEM_ID);
        assert_eq!(value.item_id.storage_name().len(), 26);
        assert!(
            value
                .item_id
                .storage_name()
                .bytes()
                .all(|byte| { byte.is_ascii_uppercase() || byte.is_ascii_digit() })
        );
    }

    #[test]
    fn synthetic_store_round_trips_with_modified_promotion() {
        let item = CustomXmlDataItem::new(
            properties().item_id.storage_name(),
            br#"<?xml version="1.0"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>"#
                .to_vec(),
            properties(),
        )
        .unwrap();
        let store = MsoDataStore::new(DataStorePromotion::Modified, vec![item]).unwrap();
        let mut writer = OleWriter::new();
        write_mso_data_store(&mut writer, &store).unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        assert_eq!(inspect_mso_data_store(&mut ole).unwrap().unwrap(), store);
    }

    #[test]
    fn reads_real_word_custom_xml_store() {
        let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/inline-endnote-and-footnote.doc");
        let mut ole = OleFile::open(std::fs::File::open(path).unwrap()).unwrap();
        let store = inspect_mso_data_store(&mut ole).unwrap().unwrap();
        assert_eq!(store.promotion, DataStorePromotion::Unspecified);
        assert_eq!(store.items().len(), 1);
        assert_eq!(store.items()[0].properties().item_id.to_string(), ITEM_ID);
        assert_eq!(store.items()[0].root_name().local_name, "Sources");
    }

    #[test]
    fn rejects_conflicting_markers_bad_shape_and_hostile_xml() {
        let mut writer = OleWriter::new();
        writer.create_storage(&[MSO_DATA_STORE_STORAGE]).unwrap();
        writer
            .create_storage(&[REDUNDANT_PROMOTION_STORAGE])
            .unwrap();
        writer
            .create_storage(&[MODIFIED_PROMOTION_STORAGE])
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        assert!(inspect_mso_data_store(&mut ole).is_err());

        assert!(
            CustomXmlDataItem::new("item", b"<!DOCTYPE x><x/>".to_vec(), properties()).is_err()
        );
        assert!(
            parse_custom_xml_properties(
                br#"<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{E0FCA697-D525-4175-A08B-DFD1F1FC7C9F}"><ds:bad/></ds:datastoreItem>"#
            )
            .is_err()
        );
        assert!(CustomXmlDataItem::new("item", b"<x>\0</x>".to_vec(), properties()).is_err());
        let mut invalid_properties = properties();
        invalid_properties.schema_references = vec!["urn:\0test".to_string()];
        assert!(write_custom_xml_properties(&invalid_properties).is_err());
    }

    #[test]
    fn zero_count_limits_can_disable_items_and_schema_references() {
        let mut writer = OleWriter::new();
        writer.create_storage(&[MSO_DATA_STORE_STORAGE]).unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        let limits = MsoDataStoreLimits {
            max_items: 0,
            max_schema_references: 0,
            ..MsoDataStoreLimits::default()
        };
        assert!(
            inspect_mso_data_store_with_limits(&mut ole, limits)
                .unwrap()
                .unwrap()
                .items()
                .is_empty()
        );

        let xml = write_custom_xml_properties(&properties()).unwrap();
        assert!(parse_custom_xml_properties_with_limits(&xml, &limits).is_err());
    }

    #[test]
    fn accepts_and_preserves_utf16_xml_streams() {
        fn utf16le(value: &str) -> Vec<u8> {
            let mut output = vec![0xFF, 0xFE];
            for unit in value.encode_utf16() {
                output.extend_from_slice(&unit.to_le_bytes());
            }
            output
        }

        let properties_xml = utf16le(
            &String::from_utf8(write_custom_xml_properties(&properties()).unwrap()).unwrap(),
        );
        assert_eq!(
            parse_custom_xml_properties(&properties_xml).unwrap(),
            properties()
        );

        let item_xml = utf16le(
            r#"<?xml version="1.0" encoding="UTF-16"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>"#,
        );
        let item = CustomXmlDataItem::new("UTF16ITEM", item_xml.clone(), properties()).unwrap();
        assert_eq!(item.xml(), item_xml);
        assert_eq!(item.root_name().local_name, "Sources");
    }
}
