//! ISO/IEC 29500 Custom XML Data Storage parts and their properties.
//!
//! Payload XML is validated and retained as inert bytes. This module never
//! retrieves schemas, performs validation against a schema, runs transforms,
//! resolves external entities, or interprets application-specific payloads.

use crate::common::{ExpandedName, MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::XmlPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::{HashMap, HashSet};

pub const TRANSITIONAL_CUSTOM_XML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
pub const STRICT_CUSTOM_XML_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/customXml";
pub const TRANSITIONAL_CUSTOM_XML_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
pub const STRICT_CUSTOM_XML_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
pub const TRANSITIONAL_CUSTOM_XML_PROPERTIES_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
pub const STRICT_CUSTOM_XML_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps";
pub const CUSTOM_XML_PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

pub const MAX_CUSTOM_XML_PART_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_CUSTOM_XML_PROPERTIES_BYTES: usize = 1024 * 1024;
pub const MAX_CUSTOM_XML_DEPTH: usize = 256;
pub const MAX_CUSTOM_XML_ITEMS: usize = 4096;
pub const MAX_CUSTOM_XML_SCHEMA_REFERENCES: usize = 4096;
pub const MAX_CUSTOM_XML_ELEMENTS: usize = 1_000_000;
pub const MAX_CUSTOM_XML_STRING_BYTES: usize = 4 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomXmlConformance {
    Transitional,
    Strict,
}

impl CustomXmlConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_CUSTOM_XML_NAMESPACE,
            Self::Strict => STRICT_CUSTOM_XML_NAMESPACE,
        }
    }

    fn data_relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_CUSTOM_XML_RELATIONSHIP,
            Self::Strict => STRICT_CUSTOM_XML_RELATIONSHIP,
        }
    }

    fn properties_relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_CUSTOM_XML_PROPERTIES_RELATIONSHIP,
            Self::Strict => STRICT_CUSTOM_XML_PROPERTIES_RELATIONSHIP,
        }
    }
}

/// Contents of a Custom XML Data Storage Properties part (§22.5).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlDataProperties {
    /// Globally unique `ST_Guid` for the associated data part.
    pub item_id: String,
    /// Target namespaces of associated schemas. These URIs are never resolved.
    pub schema_references: Vec<String>,
}

/// One relationship occurrence that targets a Custom XML Data Storage part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomXmlDataItem {
    pub source_part_name: PackURI,
    pub relationship_id: String,
    pub data_part_name: PackURI,
    pub content_type: String,
    pub root_name: ExpandedName,
    pub xml: Vec<u8>,
    pub properties_part_name: Option<PackURI>,
    pub properties: Option<CustomXmlDataProperties>,
}

/// Parameters for deterministic package creation.
#[derive(Debug, Clone)]
pub struct NewCustomXmlDataItem {
    pub source_part_name: PackURI,
    pub relationship_id: String,
    pub data_part_name: PackURI,
    pub content_type: String,
    pub xml: Vec<u8>,
    pub properties_part_name: Option<PackURI>,
    pub properties_relationship_id: String,
    pub properties: Option<CustomXmlDataProperties>,
    pub conformance: CustomXmlConformance,
}

/// Parse a Custom XML Data Storage Properties part with bounded MCE handling.
pub fn parse_custom_xml_properties(xml: &[u8]) -> Result<CustomXmlDataProperties> {
    if xml.len() > MAX_CUSTOM_XML_PROPERTIES_BYTES {
        return invalid(format!(
            "custom XML properties exceed {MAX_CUSTOM_XML_PROPERTIES_BYTES} bytes"
        ));
    }
    let mut capabilities = MceCapabilities::ooxml_baseline();
    capabilities
        .understand_namespace(TRANSITIONAL_CUSTOM_XML_NAMESPACE)
        .understand_namespace(STRICT_CUSTOM_XML_NAMESPACE);
    let limits = MceLimits {
        max_input_bytes: MAX_CUSTOM_XML_PROPERTIES_BYTES,
        max_output_bytes: MAX_CUSTOM_XML_PROPERTIES_BYTES * 2,
        max_depth: MAX_CUSTOM_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    parse_properties_xml(processed.xml.as_ref())
}

/// Serialize Custom XML Data Storage Properties in stable schema order.
pub fn write_custom_xml_properties(
    properties: &CustomXmlDataProperties,
    conformance: CustomXmlConformance,
) -> Result<Vec<u8>> {
    validate_properties(properties)?;
    let mut out = String::with_capacity(256 + properties.schema_references.len() * 64);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<ds:datastoreItem xmlns:ds=\"");
    escape_attribute(&mut out, conformance.namespace());
    out.push_str("\" ds:itemID=\"");
    escape_attribute(&mut out, &properties.item_id);
    out.push_str("\">");
    if !properties.schema_references.is_empty() {
        out.push_str("<ds:schemaRefs>");
        for uri in &properties.schema_references {
            out.push_str("<ds:schemaRef ds:uri=\"");
            escape_attribute(&mut out, uri);
            out.push_str("\"/>");
        }
        out.push_str("</ds:schemaRefs>");
    }
    out.push_str("</ds:datastoreItem>");
    Ok(out.into_bytes())
}

/// Discover and validate every explicit Custom XML Data Storage relationship.
pub fn discover_custom_xml_data(package: &OpcPackage) -> Result<Vec<CustomXmlDataItem>> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_data_relationship(relationship.reltype()))
    {
        return invalid("package root cannot source a Custom XML Data Storage relationship".into());
    }

    let mut occurrences = Vec::new();
    let mut property_ids: HashMap<String, PackURI> = HashMap::new();
    let mut validated_properties: HashMap<PackURI, CustomXmlDataProperties> = HashMap::new();
    let mut properties_owners: HashMap<PackURI, PackURI> = HashMap::new();
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| is_data_relationship(relationship.reltype()))
        {
            if occurrences.len() >= MAX_CUSTOM_XML_ITEMS {
                return invalid(format!(
                    "custom XML relationship count exceeds {MAX_CUSTOM_XML_ITEMS}"
                ));
            }
            if relationship.is_external() {
                return invalid(format!(
                    "custom XML relationship '{}' from '{}' must be internal",
                    relationship.r_id(),
                    source.partname().as_str()
                ));
            }
            let data_part_name = relationship.target_partname().map_err(|error| {
                OoxmlError::InvalidRelationship(format!(
                    "invalid custom XML target '{}': {error}",
                    relationship.r_id()
                ))
            })?;
            let data_part = package.get_part(&data_part_name).map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "custom XML part '{}': {error}",
                    data_part_name.as_str()
                ))
            })?;
            require_xml_content_type(data_part.content_type())?;
            let root_name = validate_custom_xml_payload(data_part.blob())?;
            let (properties_part_name, properties) = resolve_properties(
                package,
                data_part,
                &mut property_ids,
                &mut validated_properties,
                &mut properties_owners,
            )?;
            occurrences.push(CustomXmlDataItem {
                source_part_name: source.partname().clone(),
                relationship_id: relationship.r_id().into(),
                data_part_name,
                content_type: data_part.content_type().into(),
                root_name,
                xml: data_part.blob().to_vec(),
                properties_part_name,
                properties,
            });
        }
    }
    occurrences.sort_unstable_by(|left, right| {
        left.source_part_name
            .as_str()
            .cmp(right.source_part_name.as_str())
            .then_with(|| left.relationship_id.cmp(&right.relationship_id))
    });
    Ok(occurrences)
}

/// Add a validated data part, optional properties part, and both relationships.
pub fn add_custom_xml_data(
    package: &mut OpcPackage,
    item: NewCustomXmlDataItem,
) -> Result<()> {
    require_xml_content_type(&item.content_type)?;
    validate_custom_xml_payload(&item.xml)?;
    if item.relationship_id.is_empty() {
        return invalid("custom XML relationship ID must not be empty".into());
    }
    if item.properties.is_some() != item.properties_part_name.is_some() {
        return invalid("properties and properties_part_name must either both be present or absent".into());
    }
    if package
        .iter_parts()
        .any(|part| part.partname() == &item.data_part_name)
    {
        return invalid(format!(
            "custom XML target '{}' already exists",
            item.data_part_name.as_str()
        ));
    }
    let source = package.get_part(&item.source_part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "custom XML source '{}': {error}",
            item.source_part_name.as_str()
        ))
    })?;
    if source
        .rels()
        .iter()
        .any(|relationship| relationship.r_id() == item.relationship_id)
    {
        return invalid(format!(
            "relationship '{}' already exists on '{}'",
            item.relationship_id,
            item.source_part_name.as_str()
        ));
    }

    let mut data_part = XmlPart::new(
        item.data_part_name.clone(),
        item.content_type,
        item.xml,
    );
    if let (Some(properties), Some(properties_part_name)) =
        (item.properties, item.properties_part_name)
    {
        validate_properties(&properties)?;
        if item.properties_relationship_id.is_empty() {
            return invalid("custom XML properties relationship ID must not be empty".into());
        }
        if package
            .iter_parts()
            .any(|part| part.partname() == &properties_part_name)
        {
            return invalid(format!(
                "custom XML properties target '{}' already exists",
                properties_part_name.as_str()
            ));
        }
        for existing in discover_custom_xml_data(package)? {
            if existing.properties.as_ref().is_some_and(|candidate| {
                candidate.item_id.eq_ignore_ascii_case(&properties.item_id)
            }) {
                return invalid(format!(
                    "duplicate custom XML itemID '{}'",
                    properties.item_id
                ));
            }
        }
        let properties_xml = write_custom_xml_properties(&properties, item.conformance)?;
        let properties_target = properties_part_name.relative_ref(item.data_part_name.base_uri());
        data_part.rels_mut().add_relationship(
            item.conformance.properties_relationship().into(),
            properties_target,
            item.properties_relationship_id,
            false,
        );
        package.add_part(Box::new(XmlPart::new(
            properties_part_name,
            CUSTOM_XML_PROPERTIES_CONTENT_TYPE.into(),
            properties_xml,
        )));
    }
    package.add_part(Box::new(data_part));
    let target = item.data_part_name.relative_ref(item.source_part_name.base_uri());
    package
        .get_part_mut(&item.source_part_name)?
        .rels_mut()
        .add_relationship(
            item.conformance.data_relationship().into(),
            target,
            item.relationship_id,
            false,
        );
    Ok(())
}

fn resolve_properties(
    package: &OpcPackage,
    data_part: &dyn Part,
    property_ids: &mut HashMap<String, PackURI>,
    cache: &mut HashMap<PackURI, CustomXmlDataProperties>,
    owners: &mut HashMap<PackURI, PackURI>,
) -> Result<(Option<PackURI>, Option<CustomXmlDataProperties>)> {
    let relationships: Vec<_> = data_part.rels().iter().collect();
    if relationships.len() > 1 {
        return invalid(format!(
            "custom XML part '{}' has more than one outbound relationship",
            data_part.partname().as_str()
        ));
    }
    let Some(relationship) = relationships.first() else {
        return Ok((None, None));
    };
    if !is_properties_relationship(relationship.reltype()) {
        return invalid(format!(
            "custom XML part '{}' has forbidden relationship type '{}'",
            data_part.partname().as_str(),
            relationship.reltype()
        ));
    }
    if relationship.is_external() {
        return invalid("custom XML properties relationship must be internal".into());
    }
    let properties_part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("invalid custom XML properties target: {error}"))
    })?;
    if let Some(existing_owner) =
        owners.insert(properties_part_name.clone(), data_part.partname().clone())
    {
        if existing_owner != *data_part.partname() {
            return invalid(format!(
                "custom XML properties part '{}' is shared by '{}' and '{}'",
                properties_part_name.as_str(),
                existing_owner.as_str(),
                data_part.partname().as_str()
            ));
        }
    }
    let properties = if let Some(properties) = cache.get(&properties_part_name) {
        properties.clone()
    } else {
        let part = package.get_part(&properties_part_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "custom XML properties part '{}': {error}",
                properties_part_name.as_str()
            ))
        })?;
        if part.content_type() != CUSTOM_XML_PROPERTIES_CONTENT_TYPE {
            return Err(OoxmlError::InvalidContentType {
                expected: CUSTOM_XML_PROPERTIES_CONTENT_TYPE.into(),
                got: part.content_type().into(),
            });
        }
        if part.rels().iter().next().is_some() {
            return invalid(format!(
                "custom XML properties part '{}' must not have relationships",
                properties_part_name.as_str()
            ));
        }
        let properties = parse_custom_xml_properties(part.blob())?;
        let key = properties.item_id.to_ascii_lowercase();
        if let Some(existing) = property_ids.insert(key, properties_part_name.clone()) {
            if existing != properties_part_name {
                return invalid(format!(
                    "duplicate custom XML itemID '{}'",
                    properties.item_id
                ));
            }
        }
        cache.insert(properties_part_name.clone(), properties.clone());
        properties
    };
    Ok((Some(properties_part_name), Some(properties)))
}

fn validate_properties(properties: &CustomXmlDataProperties) -> Result<()> {
    if !is_st_guid(&properties.item_id) {
        return invalid(format!(
            "custom XML itemID '{}' is not ST_Guid",
            properties.item_id
        ));
    }
    if properties.schema_references.len() > MAX_CUSTOM_XML_SCHEMA_REFERENCES {
        return invalid(format!(
            "schema reference count exceeds {MAX_CUSTOM_XML_SCHEMA_REFERENCES}"
        ));
    }
    let string_bytes = properties.item_id.len()
        + properties
            .schema_references
            .iter()
            .map(String::len)
            .sum::<usize>();
    if string_bytes > MAX_CUSTOM_XML_STRING_BYTES {
        return invalid("custom XML properties strings exceed allocation cap".into());
    }
    if properties.schema_references.iter().any(String::is_empty) {
        return invalid("custom XML schema reference URI must not be empty".into());
    }
    Ok(())
}

fn is_st_guid(value: &str) -> bool {
    let Some(inner) = value.strip_prefix('{').and_then(|value| value.strip_suffix('}')) else {
        return false;
    };
    let groups = [8, 4, 4, 4, 12];
    let mut parts = inner.split('-');
    groups.into_iter().all(|length| {
        parts
            .next()
            .is_some_and(|part| part.len() == length && part.bytes().all(|byte| byte.is_ascii_hexdigit()))
    }) && parts.next().is_none()
}

fn validate_custom_xml_payload(xml: &[u8]) -> Result<ExpandedName> {
    if xml.len() > MAX_CUSTOM_XML_PART_BYTES {
        return invalid(format!(
            "custom XML payload exceeds {MAX_CUSTOM_XML_PART_BYTES} bytes"
        ));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut namespaces: Vec<HashMap<String, String>> = Vec::new();
    let mut root = None;
    let mut closed_root = false;
    let mut element_count = 0usize;
    let mut version = XmlVersion::Implicit1_0;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Decl(declaration) => version = declaration.xml_version()?,
            Event::Start(element) => {
                if namespaces.len() >= MAX_CUSTOM_XML_DEPTH {
                    return invalid(format!("custom XML depth exceeds {MAX_CUSTOM_XML_DEPTH}"));
                }
                let info = resolve_element(&reader, &element, namespaces.last(), version)?;
                element_count += 1;
                if element_count > MAX_CUSTOM_XML_ELEMENTS {
                    return invalid(format!("custom XML element count exceeds {MAX_CUSTOM_XML_ELEMENTS}"));
                }
                if namespaces.is_empty() {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots".into());
                    }
                    root = Some(ExpandedName {
                        namespace: info.namespace.clone(),
                        local_name: info.local_name.clone(),
                    });
                }
                namespaces.push(info.namespaces);
            }
            Event::Empty(element) => {
                let info = resolve_element(&reader, &element, namespaces.last(), version)?;
                element_count += 1;
                if element_count > MAX_CUSTOM_XML_ELEMENTS {
                    return invalid(format!("custom XML element count exceeds {MAX_CUSTOM_XML_ELEMENTS}"));
                }
                if namespaces.is_empty() {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots".into());
                    }
                    root = Some(ExpandedName {
                        namespace: info.namespace,
                        local_name: info.local_name,
                    });
                    closed_root = true;
                }
            }
            Event::End(_) => {
                if namespaces.pop().is_none() {
                    return invalid("custom XML payload has an unexpected end tag".into());
                }
                if namespaces.is_empty() {
                    closed_root = true;
                }
            }
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML payloads".into()),
            Event::GeneralRef(reference) => {
                let reference = reference.as_ref();
                let predefined = matches!(reference, b"lt" | b"gt" | b"amp" | b"apos" | b"quot");
                if !predefined && !reference.starts_with(b"#") {
                    return invalid("custom XML payload contains a non-predefined entity".into());
                }
            }
            Event::Text(text) if namespaces.is_empty() && !is_xml_whitespace(text.as_ref()) => {
                return invalid("custom XML payload has text outside its root".into());
            }
            Event::CData(text) if namespaces.is_empty() && !is_xml_whitespace(text.as_ref()) => {
                return invalid("custom XML payload has CDATA outside its root".into());
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !namespaces.is_empty() || !closed_root {
        return invalid("custom XML payload has no complete root element".into());
    }
    root.ok_or_else(|| OoxmlError::InvalidFormat("custom XML payload has no root".into()))
}

#[derive(Debug)]
struct ResolvedElement {
    namespace: String,
    local_name: String,
    attributes: Vec<(String, String, String)>,
    namespaces: HashMap<String, String>,
}

fn resolve_element(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    parent_namespaces: Option<&HashMap<String, String>>,
    version: XmlVersion,
) -> Result<ResolvedElement> {
    let mut namespaces = parent_namespaces.cloned().unwrap_or_default();
    namespaces.insert("xml".into(), "http://www.w3.org/XML/1998/namespace".into());
    let mut raw_attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        if name == "xmlns" {
            namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.into(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    let qname = element.name();
    let raw_name = std::str::from_utf8(qname.as_ref())
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    let (prefix, local_name) = split_qname(raw_name);
    let namespace = if prefix.is_empty() {
        namespaces.get(prefix).cloned().unwrap_or_default()
    } else {
        namespaces.get(prefix).cloned().ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("unbound XML namespace prefix '{prefix}'"))
        })?
    };
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (name, value) in raw_attributes {
        let (prefix, local_name) = split_qname(&name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces.get(prefix).cloned().ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("unbound XML attribute prefix '{prefix}'"))
            })?
        };
        if !seen.insert((namespace.clone(), local_name.to_owned())) {
            return invalid(format!("duplicate XML attribute {{{namespace}}}{local_name}"));
        }
        attributes.push((namespace, local_name.to_owned(), value));
    }
    Ok(ResolvedElement {
        namespace,
        local_name: local_name.into(),
        attributes,
        namespaces,
    })
}

fn parse_properties_xml(xml: &[u8]) -> Result<CustomXmlDataProperties> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut namespaces = Vec::new();
    let mut contexts = Vec::new();
    let mut item_id = None;
    let mut schema_references = Vec::new();
    let mut root_namespace = None;
    let mut seen_schema_references = false;
    let mut closed_root = false;
    let mut version = XmlVersion::Implicit1_0;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Decl(declaration) => version = declaration.xml_version()?,
            Event::Start(element) => {
                let info = resolve_element(&reader, &element, namespaces.last(), version)?;
                let context = properties_start(
                    &info,
                    contexts.last().copied(),
                    &mut item_id,
                    &mut schema_references,
                    &mut root_namespace,
                    &mut seen_schema_references,
                )?;
                contexts.push(context);
                namespaces.push(info.namespaces);
            }
            Event::Empty(element) => {
                let info = resolve_element(&reader, &element, namespaces.last(), version)?;
                properties_start(
                    &info,
                    contexts.last().copied(),
                    &mut item_id,
                    &mut schema_references,
                    &mut root_namespace,
                    &mut seen_schema_references,
                )?;
                if contexts.is_empty() {
                    closed_root = true;
                }
            }
            Event::End(_) => {
                if contexts.pop().is_none() || namespaces.pop().is_none() {
                    return invalid("unexpected custom XML properties end tag".into());
                }
                if contexts.is_empty() {
                    closed_root = true;
                }
            }
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML properties".into()),
            Event::Text(text) if !is_xml_whitespace(text.as_ref()) => {
                return invalid("text is not permitted in custom XML properties".into());
            }
            Event::CData(text) if !is_xml_whitespace(text.as_ref()) => {
                return invalid("CDATA is not permitted in custom XML properties".into());
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !closed_root || !contexts.is_empty() {
        return invalid("custom XML properties root is incomplete".into());
    }
    let properties = CustomXmlDataProperties {
        item_id: item_id.ok_or_else(|| {
            OoxmlError::InvalidFormat("datastoreItem requires ds:itemID".into())
        })?,
        schema_references,
    };
    validate_properties(&properties)?;
    Ok(properties)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PropertiesContext {
    DataStoreItem,
    SchemaReferences,
    SchemaReference,
}

fn properties_start(
    element: &ResolvedElement,
    parent: Option<PropertiesContext>,
    item_id: &mut Option<String>,
    schema_references: &mut Vec<String>,
    root_namespace: &mut Option<String>,
    seen_schema_references: &mut bool,
) -> Result<PropertiesContext> {
    let supported_namespace = is_custom_xml_namespace(&element.namespace);
    match parent {
        None if root_namespace.is_none()
            && supported_namespace
            && element.local_name == "datastoreItem" =>
        {
            reject_attributes_except(element, &[(element.namespace.as_str(), "itemID")])?;
            *item_id = Some(required_resolved_attr(element, &element.namespace, "itemID")?.into());
            *root_namespace = Some(element.namespace.clone());
            Ok(PropertiesContext::DataStoreItem)
        }
        Some(PropertiesContext::DataStoreItem)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRefs" =>
        {
            if *seen_schema_references {
                return invalid("datastoreItem has multiple schemaRefs elements".into());
            }
            *seen_schema_references = true;
            reject_attributes_except(element, &[])?;
            Ok(PropertiesContext::SchemaReferences)
        }
        Some(PropertiesContext::SchemaReferences)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRef" =>
        {
            reject_attributes_except(element, &[(element.namespace.as_str(), "uri")])?;
            if schema_references.len() >= MAX_CUSTOM_XML_SCHEMA_REFERENCES {
                return invalid(format!(
                    "schema reference count exceeds {MAX_CUSTOM_XML_SCHEMA_REFERENCES}"
                ));
            }
            schema_references.push(
                required_resolved_attr(element, &element.namespace, "uri")?.into(),
            );
            Ok(PropertiesContext::SchemaReference)
        }
        _ => invalid(format!(
            "unexpected custom XML properties element {{{}}}{}",
            element.namespace, element.local_name
        )),
    }
}

fn reject_attributes_except(element: &ResolvedElement, allowed: &[(&str, &str)]) -> Result<()> {
    for (namespace, local_name, _) in &element.attributes {
        if !allowed
            .iter()
            .any(|(allowed_namespace, allowed_name)| {
                namespace == allowed_namespace && local_name == allowed_name
            })
        {
            return invalid(format!(
                "unexpected attribute {{{namespace}}}{local_name} on {}",
                element.local_name
            ));
        }
    }
    Ok(())
}

fn required_resolved_attr<'a>(
    element: &'a ResolvedElement,
    namespace: &str,
    local_name: &str,
) -> Result<&'a str> {
    element
        .attributes
        .iter()
        .find(|(candidate_namespace, candidate_name, _)| {
            candidate_namespace == namespace && candidate_name == local_name
        })
        .map(|(_, _, value)| value.as_str())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "{} requires {{{namespace}}}{local_name}",
                element.local_name
            ))
        })
}

fn require_xml_content_type(content_type: &str) -> Result<()> {
    let mime = content_type
        .split_once(';')
        .map_or(content_type, |(mime, _)| mime)
        .trim();
    if mime.eq_ignore_ascii_case("application/xml")
        || mime.eq_ignore_ascii_case("text/xml")
        || mime.to_ascii_lowercase().ends_with("+xml")
    {
        Ok(())
    } else {
        invalid(format!(
            "Custom XML Data Storage content type '{content_type}' is not XML"
        ))
    }
}

fn is_data_relationship(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_CUSTOM_XML_RELATIONSHIP | STRICT_CUSTOM_XML_RELATIONSHIP
    )
}

fn is_properties_relationship(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_CUSTOM_XML_PROPERTIES_RELATIONSHIP
            | STRICT_CUSTOM_XML_PROPERTIES_RELATIONSHIP
    )
}

fn is_custom_xml_namespace(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_CUSTOM_XML_NAMESPACE | STRICT_CUSTOM_XML_NAMESPACE
    )
}

fn split_qname(value: &str) -> (&str, &str) {
    value.split_once(':').unwrap_or(("", value))
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value.iter().all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn escape_attribute(out: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(character),
        }
    }
}

fn invalid<T>(message: String) -> Result<T> {
    Err(OoxmlError::InvalidFormat(message))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    const POI_XLSX: &[u8] = include_bytes!(
        "../../../3rdparty/poi/test-data/spreadsheet/customIndexedColors.xlsx"
    );
    const LO_DOCX: &[u8] = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
    );

    #[test]
    fn loads_poi_and_libreoffice_reference_fixtures() {
        let poi = OpcPackage::from_bytes(POI_XLSX).unwrap();
        let items = discover_custom_xml_data(&poi).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].root_name.local_name, "easyPacket");
        assert_eq!(items[0].properties.as_ref().unwrap().schema_references.len(), 0);

        let libreoffice = OpcPackage::from_bytes(LO_DOCX).unwrap();
        let items = discover_custom_xml_data(&libreoffice).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].root_name.local_name, "Sources");
        assert_eq!(
            items[0].properties.as_ref().unwrap().schema_references,
            ["http://schemas.openxmlformats.org/officeDocument/2006/bibliography"]
        );
    }

    #[test]
    fn strict_properties_writer_is_deterministic_and_round_trips() {
        let properties = sample_properties();
        let first = write_custom_xml_properties(&properties, CustomXmlConformance::Strict).unwrap();
        let second = write_custom_xml_properties(&properties, CustomXmlConformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(std::str::from_utf8(&first)
            .unwrap()
            .contains(STRICT_CUSTOM_XML_NAMESPACE));
        assert_eq!(parse_custom_xml_properties(&first).unwrap(), properties);
    }

    #[test]
    fn mce_selects_fallback_schema_reference() {
        let xml = format!(
            r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_CUSTOM_XML_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs><mc:AlternateContent><mc:Choice Requires="x"><ds:schemaRef ds:uri="urn:wrong"/></mc:Choice><mc:Fallback><ds:schemaRef ds:uri="urn:right"/></mc:Fallback></mc:AlternateContent></ds:schemaRefs></ds:datastoreItem>"#
        );
        assert_eq!(
            parse_custom_xml_properties(xml.as_bytes())
                .unwrap()
                .schema_references,
            ["urn:right"]
        );
    }

    #[test]
    fn package_writer_round_trips_without_interpreting_payload() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .into(),
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>".to_vec(),
        )));
        add_custom_xml_data(
            &mut package,
            NewCustomXmlDataItem {
                source_part_name: PackURI::new("/word/document.xml").unwrap(),
                relationship_id: "rIdData".into(),
                data_part_name: PackURI::new("/customXml/item1.xml").unwrap(),
                content_type: "application/xml".into(),
                xml: b"<customer xmlns=\"urn:customer\" id=\"7\"/>".to_vec(),
                properties_part_name: Some(
                    PackURI::new("/customXml/itemProps1.xml").unwrap(),
                ),
                properties_relationship_id: "rIdProps".into(),
                properties: Some(sample_properties()),
                conformance: CustomXmlConformance::Transitional,
            },
        )
        .unwrap();
        let items = discover_custom_xml_data(&package).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].xml, b"<customer xmlns=\"urn:customer\" id=\"7\"/>");
        assert_eq!(items[0].properties.as_ref().unwrap(), &sample_properties());
    }

    #[test]
    fn rejects_malformed_properties_payloads_and_package_graphs() {
        assert!(parse_custom_xml_properties(br#"<!DOCTYPE x><x/>"#).is_err());
        let missing_id = format!(
            r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_CUSTOM_XML_NAMESPACE}"/>"#
        );
        assert!(parse_custom_xml_properties(missing_id.as_bytes()).is_err());
        let duplicate_refs = format!(
            r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_CUSTOM_XML_NAMESPACE}" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs/><ds:schemaRefs/></ds:datastoreItem>"#
        );
        assert!(parse_custom_xml_properties(duplicate_refs.as_bytes()).is_err());
        assert!(validate_custom_xml_payload(br#"<!DOCTYPE x><x/>"#).is_err());
        assert!(validate_custom_xml_payload(b"<a><b></a>").is_err());

        let mut package = OpcPackage::new();
        let mut source = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/xml".into(),
            b"<document/>".to_vec(),
        );
        source.rels_mut().add_relationship(
            TRANSITIONAL_CUSTOM_XML_RELATIONSHIP.into(),
            "https://example.invalid/data.xml".into(),
            "rId1".into(),
            true,
        );
        package.add_part(Box::new(source));
        assert!(discover_custom_xml_data(&package).is_err());
    }

    #[test]
    fn enforces_guid_schema_count_depth_and_size_caps() {
        let mut invalid_guid = sample_properties();
        invalid_guid.item_id = "not-a-guid".into();
        assert!(write_custom_xml_properties(
            &invalid_guid,
            CustomXmlConformance::Transitional
        )
        .is_err());

        let deep = format!(
            "{}{}",
            "<x>".repeat(MAX_CUSTOM_XML_DEPTH + 1),
            "</x>".repeat(MAX_CUSTOM_XML_DEPTH + 1)
        );
        assert!(validate_custom_xml_payload(deep.as_bytes()).is_err());
        assert!(validate_custom_xml_payload(&vec![b' '; MAX_CUSTOM_XML_PART_BYTES + 1]).is_err());
    }

    fn sample_properties() -> CustomXmlDataProperties {
        CustomXmlDataProperties {
            item_id: "{11111111-1111-1111-1111-111111111111}".into(),
            schema_references: vec!["urn:customer".into()],
        }
    }
}
