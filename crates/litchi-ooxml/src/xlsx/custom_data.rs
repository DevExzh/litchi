//! MS-XLSX Custom Data and Custom Data Properties parts.
//!
//! Custom Data payloads are deliberately opaque. This module inventories and
//! copies their bytes, but never parses, activates, dispatches, or executes them.

use crate::error::{OoxmlError, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

pub const CUSTOM_DATA_CONTENT_TYPE: &str = "application/binary";
pub const CUSTOM_DATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
pub const CUSTOM_DATA_PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
pub const CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";

const MAX_PROPERTIES_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_XML_BYTES: usize = 2 * 1024 * 1024;
const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_PAYLOAD_BYTES: usize = 256 * 1024 * 1024;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_STORES: usize = 4096;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomDataExtensionList {
    /// One self-contained `x14:extLst` subtree, retained without interpretation.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomDataProperties {
    pub id: String,
    pub extension_list: Option<CustomDataExtensionList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomDataPayload {
    pub part_name: String,
    /// Opaque add-in-owned data. No format sniffing or execution is performed.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookCustomData {
    pub properties_relationship_id: String,
    pub properties_part_name: String,
    pub properties: CustomDataProperties,
    pub data_relationship_id: Option<String>,
    pub payload: Option<CustomDataPayload>,
}

#[derive(Clone)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

/// Parse a Custom Data Properties part.
pub fn parse_custom_data_properties(xml: &[u8]) -> Result<CustomDataProperties> {
    let root = parse_document(xml)?;
    require(&root, X14, "datastoreItem")?;
    no_attributes(&root, &[("", "id")])?;
    whitespace(&root)?;
    if root.children.len() > 1 {
        return Err(invalid("datastoreItem permits at most one extLst"));
    }
    let extension_list = root
        .children
        .first()
        .map(|child| {
            require(child, X14, "extLst")?;
            let xml = serialize_node(child)?;
            if xml.len() > MAX_EXTENSION_XML_BYTES {
                return Err(limit("extension XML bytes"));
            }
            Ok(CustomDataExtensionList { xml })
        })
        .transpose()?;
    let value = CustomDataProperties {
        id: required(&root, "", "id")?.to_owned(),
        extension_list,
    };
    validate_properties(&value, true)?;
    Ok(value)
}

/// Deterministically serialize a Custom Data Properties part.
pub fn write_custom_data_properties(value: &CustomDataProperties) -> Result<Vec<u8>> {
    validate_properties(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x14:datastoreItem xmlns:x14=\"");
    escape_attr(&mut output, X14);
    output.push(b'\"');
    attr(&mut output, "id", &value.id);
    if let Some(extension) = &value.extension_list {
        output.push(b'>');
        output.extend_from_slice(&extension.xml);
        output.extend_from_slice(b"</x14:datastoreItem>");
    } else {
        output.extend_from_slice(b"/>");
    }
    if output.len() > MAX_PROPERTIES_XML_BYTES {
        return Err(limit("serialized properties XML bytes"));
    }
    Ok(output)
}

/// Load every Custom Data Properties part related from the workbook and its
/// optional opaque Custom Data payload.
pub fn load_custom_data(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Vec<WorkbookCustomData>> {
    reject_root_relationships(package)?;
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    require_workbook_root(&root)?;

    for part in package.iter_parts() {
        if part.partname().as_str() != workbook_name.as_str()
            && part.rels().iter().any(|relationship| {
                relationship.reltype() == CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE
            })
        {
            return Err(invalid(format!(
                "non-workbook part '{}' sources a Custom Data Properties relationship",
                part.partname()
            )));
        }
    }

    let property_relationships: Vec<_> = workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE)
        .collect();
    if property_relationships.len() > MAX_STORES {
        return Err(limit("store count"));
    }
    let mut property_targets = HashSet::new();
    let mut data_targets = HashSet::new();
    let mut ids = HashSet::new();
    let mut output = Vec::with_capacity(property_relationships.len());
    let mut total_payload = 0usize;
    for relationship in property_relationships {
        validate_relationship_id(relationship.r_id())?;
        if relationship.is_external() {
            return Err(invalid(format!(
                "Custom Data Properties relationship '{}' is external",
                relationship.r_id()
            )));
        }
        let target = relationship.target_partname()?;
        if !property_targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple workbook relationships target Custom Data Properties part '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != CUSTOM_DATA_PROPERTIES_CONTENT_TYPE {
            return Err(invalid(format!(
                "Custom Data Properties part '{target}' has content type '{}'",
                part.content_type()
            )));
        }
        let properties = parse_custom_data_properties(part.blob())?;
        if !ids.insert(properties.id.clone()) {
            return Err(invalid(format!(
                "duplicate Custom Data storage id '{}'",
                properties.id
            )));
        }
        let custom_relationships: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == CUSTOM_DATA_RELATIONSHIP_TYPE)
            .collect();
        if part.rels().len() != custom_relationships.len() {
            return Err(invalid(format!(
                "Custom Data Properties part '{target}' has a forbidden outbound relationship"
            )));
        }
        if custom_relationships.len() > 1 {
            return Err(invalid(format!(
                "Custom Data Properties part '{target}' has multiple Custom Data relationships"
            )));
        }
        let (data_relationship_id, payload) =
            if let Some(data_relationship) = custom_relationships.first() {
                validate_relationship_id(data_relationship.r_id())?;
                if data_relationship.is_external() {
                    return Err(invalid(format!(
                        "Custom Data relationship '{}' is external",
                        data_relationship.r_id()
                    )));
                }
                let data_target = data_relationship.target_partname()?;
                if !data_targets.insert(data_target.to_string()) {
                    return Err(invalid(format!(
                        "multiple properties parts target Custom Data part '{data_target}'"
                    )));
                }
                let data_part = package.get_part(&data_target)?;
                if data_part.content_type() != CUSTOM_DATA_CONTENT_TYPE {
                    return Err(invalid(format!(
                        "Custom Data part '{data_target}' has content type '{}'",
                        data_part.content_type()
                    )));
                }
                if !data_part.rels().is_empty() {
                    return Err(invalid(format!(
                        "Custom Data part '{data_target}' has forbidden outbound relationships"
                    )));
                }
                add_payload(&mut total_payload, data_part.blob().len())?;
                (
                    Some(data_relationship.r_id().to_owned()),
                    Some(CustomDataPayload {
                        part_name: data_target.to_string(),
                        data: data_part.blob().to_vec(),
                    }),
                )
            } else {
                (None, None)
            };
        output.push(WorkbookCustomData {
            properties_relationship_id: relationship.r_id().to_owned(),
            properties_part_name: target.to_string(),
            properties,
            data_relationship_id,
            payload,
        });
    }

    for part in package.iter_parts() {
        if part.content_type() == CUSTOM_DATA_PROPERTIES_CONTENT_TYPE
            && !property_targets.contains(part.partname().as_str())
        {
            return Err(invalid(format!(
                "orphan Custom Data Properties part '{}'",
                part.partname()
            )));
        }
        if part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == CUSTOM_DATA_RELATIONSHIP_TYPE)
            && !property_targets.contains(part.partname().as_str())
        {
            return Err(invalid(format!(
                "non-properties part '{}' sources a Custom Data relationship",
                part.partname()
            )));
        }
        if part.partname().as_str().starts_with("/xl/customData/")
            && part.content_type() == CUSTOM_DATA_CONTENT_TYPE
            && !data_targets.contains(part.partname().as_str())
        {
            return Err(invalid(format!(
                "orphan Custom Data part '{}'",
                part.partname()
            )));
        }
    }
    Ok(output)
}

/// Store a complete set of Custom Data Properties parts and optional opaque
/// payloads. The operation validates the entire plan before mutating the package.
pub fn store_custom_data(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    stores: &[WorkbookCustomData],
) -> Result<()> {
    if stores.is_empty() {
        return Err(invalid("at least one Custom Data store is required"));
    }
    validate_store_set(stores)?;
    if !load_custom_data(package, workbook_name)?.is_empty() {
        return Err(invalid("workbook already contains Custom Data stores"));
    }
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    require_workbook_root(&root)?;
    let mut property_plans = Vec::with_capacity(stores.len());
    let mut data_plans = Vec::new();
    for store in stores {
        if workbook
            .rels()
            .get(&store.properties_relationship_id)
            .is_some()
        {
            return Err(invalid(format!(
                "workbook relationship ID '{}' already exists",
                store.properties_relationship_id
            )));
        }
        let properties_uri =
            PackURI::new(&store.properties_part_name).map_err(OoxmlError::InvalidUri)?;
        if !properties_uri.as_str().starts_with("/xl/customData/")
            || !properties_uri.as_str().ends_with(".xml")
        {
            return Err(invalid(format!(
                "Custom Data Properties part '{properties_uri}' must be an XML part under /xl/customData"
            )));
        }
        if package
            .iter_parts()
            .any(|part| part.partname().as_str() == properties_uri.as_str())
        {
            return Err(invalid(format!("part '{properties_uri}' already exists")));
        }
        property_plans.push((
            store.properties_relationship_id.clone(),
            properties_uri.clone(),
            write_custom_data_properties(&store.properties)?,
        ));
        if let (Some(data_id), Some(payload)) = (&store.data_relationship_id, &store.payload) {
            let data_uri = PackURI::new(&payload.part_name).map_err(OoxmlError::InvalidUri)?;
            if !data_uri.as_str().starts_with("/xl/customData/") {
                return Err(invalid(format!(
                    "Custom Data part '{data_uri}' must be under /xl/customData"
                )));
            }
            if package
                .iter_parts()
                .any(|part| part.partname().as_str() == data_uri.as_str())
            {
                return Err(invalid(format!("part '{data_uri}' already exists")));
            }
            data_plans.push((
                properties_uri,
                data_id.clone(),
                data_uri,
                payload.data.clone(),
            ));
        }
    }
    for (_, uri, xml) in &property_plans {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            CUSTOM_DATA_PROPERTIES_CONTENT_TYPE.into(),
            xml.clone(),
        )));
    }
    for (_, _, uri, data) in &data_plans {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            CUSTOM_DATA_CONTENT_TYPE.into(),
            data.clone(),
        )));
    }
    for (id, uri, _) in &property_plans {
        package
            .get_part_mut(workbook_name)?
            .rels_mut()
            .add_relationship(
                CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE.into(),
                uri.relative_ref(workbook_name.base_uri()),
                id.clone(),
                false,
            );
    }
    for (source, id, target, _) in &data_plans {
        package.get_part_mut(source)?.rels_mut().add_relationship(
            CUSTOM_DATA_RELATIONSHIP_TYPE.into(),
            target.relative_ref(source.base_uri()),
            id.clone(),
            false,
        );
    }
    Ok(())
}

fn validate_properties(value: &CustomDataProperties, extension_already_parsed: bool) -> Result<()> {
    if value.id.is_empty() {
        return Err(invalid("Custom Data storage id cannot be empty"));
    }
    if value.id.chars().count() >= 65_536 {
        return Err(invalid(
            "Custom Data storage id must contain fewer than 65536 characters",
        ));
    }
    bounded(&value.id, "storage id bytes")?;
    if let Some(extension) = &value.extension_list {
        if extension.xml.len() > MAX_EXTENSION_XML_BYTES {
            return Err(limit("extension XML bytes"));
        }
        if !extension_already_parsed {
            let root = parse_document(&extension.xml)?;
            require(&root, X14, "extLst")?;
        }
    }
    Ok(())
}

fn validate_store_set(stores: &[WorkbookCustomData]) -> Result<()> {
    if stores.len() > MAX_STORES {
        return Err(limit("store count"));
    }
    let mut ids = HashSet::new();
    let mut property_ids = HashSet::new();
    let mut property_targets = HashSet::new();
    let mut data_targets = HashSet::new();
    let mut total_payload = 0usize;
    for store in stores {
        validate_properties(&store.properties, false)?;
        validate_relationship_id(&store.properties_relationship_id)?;
        if !ids.insert(store.properties.id.clone()) {
            return Err(invalid(format!(
                "duplicate Custom Data storage id '{}'",
                store.properties.id
            )));
        }
        if !property_ids.insert(store.properties_relationship_id.clone()) {
            return Err(invalid(
                "duplicate workbook Custom Data Properties relationship ID",
            ));
        }
        if !property_targets.insert(store.properties_part_name.clone()) {
            return Err(invalid("duplicate Custom Data Properties part name"));
        }
        match (&store.data_relationship_id, &store.payload) {
            (None, None) => {},
            (Some(id), Some(payload)) => {
                validate_relationship_id(id)?;
                if !data_targets.insert(payload.part_name.clone()) {
                    return Err(invalid("duplicate Custom Data payload part name"));
                }
                add_payload(&mut total_payload, payload.data.len())?;
            },
            _ => {
                return Err(invalid(
                    "Custom Data relationship ID and payload must either both be present or both be absent",
                ));
            },
        }
    }
    Ok(())
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_PROPERTIES_XML_BYTES {
        return Err(limit("properties XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let is_empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if is_empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(_) => {
                return Err(invalid("CDATA is rejected in Custom Data Properties XML"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated Custom Data Properties XML"));
    }
    root.ok_or_else(|| invalid("missing Custom Data Properties root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::<String, ()>::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            X14 => "x14".into(),
            SML | STRICT_SML => "x".into(),
            _ => {
                let prefix = format!("n{next}");
                next += 1;
                prefix
            },
        };
        prefixes.insert(namespace.clone(), prefix);
    }
    let mut output = Vec::new();
    write_node(node, &prefixes, true, &mut output);
    Ok(output)
}
fn collect_namespaces(node: &Node, namespaces: &mut BTreeMap<String, ()>) {
    if !node.namespace.is_empty() {
        namespaces.insert(node.namespace.clone(), ());
    }
    for attribute in &node.attributes {
        if !attribute.namespace.is_empty() {
            namespaces.insert(attribute.namespace.clone(), ());
        }
    }
    for child in &node.children {
        collect_namespaces(child, namespaces);
    }
}
fn write_node(node: &Node, prefixes: &HashMap<String, String>, root: bool, output: &mut Vec<u8>) {
    output.push(b'<');
    qname(output, &node.namespace, &node.name, prefixes);
    if root {
        let mut values: Vec<_> = prefixes.iter().collect();
        values.sort_by(|a, b| a.1.cmp(b.1));
        for (namespace, prefix) in values {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
            escape_attr(output, namespace);
            output.push(b'\"');
        }
    }
    for attribute in &node.attributes {
        output.push(b' ');
        qname(output, &attribute.namespace, &attribute.name, prefixes);
        output.extend_from_slice(b"=\"");
        escape_attr(output, &attribute.value);
        output.push(b'\"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_node(child, prefixes, false, output);
    }
    output.extend_from_slice(b"</");
    qname(output, &node.namespace, &node.name, prefixes);
    output.push(b'>');
}
fn qname(output: &mut Vec<u8>, namespace: &str, name: &str, prefixes: &HashMap<String, String>) {
    if !namespace.is_empty() {
        output.extend_from_slice(prefixes[namespace].as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}

fn reject_root_relationships(package: &OpcPackage) -> Result<()> {
    if package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE | CUSTOM_DATA_RELATIONSHIP_TYPE
        )
    }) {
        Err(invalid(
            "package root cannot source Custom Data relationships",
        ))
    } else {
        Ok(())
    }
}
fn require_workbook_root(root: &Node) -> Result<()> {
    if root.name == "workbook" && matches!(root.namespace.as_str(), SML | STRICT_SML) {
        Ok(())
    } else {
        Err(invalid(
            "Custom Data Properties source must be a workbook part",
        ))
    }
}
fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected {{{namespace}}}{name}, got {{{}}}{}",
            node.namespace, node.name
        )))
    }
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn validate_relationship_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
fn add_payload(total: &mut usize, size: usize) -> Result<()> {
    if size > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total payload bytes"))?;
    if *total > MAX_TOTAL_PAYLOAD_BYTES {
        Err(limit("total payload bytes"))
    } else {
        Ok(())
    }
}
fn bounded(value: &str, name: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit(name))
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape_attr(output, value);
    output.push(b'\"');
}
fn escape_attr(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("Custom Data {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn properties() -> CustomDataProperties {
        CustomDataProperties { id: "Storage-1".into(), extension_list: Some(CustomDataExtensionList { xml: format!(r#"<x14:extLst xmlns:x14="{X14}"><x14:ext uri="urn:test"><v:opaque xmlns:v="urn:vendor" value="kept"/></x14:ext></x14:extLst>"#).into_bytes() }) }
    }
    fn store() -> WorkbookCustomData {
        WorkbookCustomData {
            properties_relationship_id: "rIdProps1".into(),
            properties_part_name: "/xl/customData/itemProps1.xml".into(),
            properties: properties(),
            data_relationship_id: Some("rIdData1".into()),
            payload: Some(CustomDataPayload {
                part_name: "/xl/customData/item1.bin".into(),
                data: b"MZ\0macro-looking bytes are inert\xff".to_vec(),
            }),
        }
    }
    fn package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        )));
        (package, workbook)
    }

    #[test]
    fn typed_properties_and_extensions_round_trip() {
        let expected = properties();
        let xml = write_custom_data_properties(&expected).unwrap();
        let parsed = parse_custom_data_properties(&xml).unwrap();
        assert_eq!(parsed.id, expected.id);
        assert!(String::from_utf8_lossy(&parsed.extension_list.unwrap().xml).contains("opaque"));
    }

    #[test]
    fn package_round_trip_keeps_binary_inert() {
        let (mut package, workbook) = package();
        let expected = store();
        store_custom_data(&mut package, &workbook, std::slice::from_ref(&expected)).unwrap();
        let loaded = load_custom_data(&package, &workbook).unwrap();
        assert_eq!(loaded.len(), 1);
        assert_eq!(loaded[0].properties.id, expected.properties.id);
        assert_eq!(
            loaded[0].properties_relationship_id,
            expected.properties_relationship_id
        );
        assert_eq!(
            loaded[0].data_relationship_id,
            expected.data_relationship_id
        );
        assert_eq!(loaded[0].payload, expected.payload);
        assert!(
            String::from_utf8_lossy(&loaded[0].properties.extension_list.as_ref().unwrap().xml)
                .contains("value=\"kept\"")
        );
    }

    #[test]
    fn rejects_hostile_xml_identity_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x14:datastoreItem xmlns:x14="{X14}" id="x"/>"#),
            format!(r#"<x14:datastoreItem xmlns:x14="{X14}"/>"#),
            format!(
                r#"<x14:datastoreItem xmlns:x14="{X14}" id="x"><x14:extLst/><x14:extLst/></x14:datastoreItem>"#
            ),
        ] {
            assert!(parse_custom_data_properties(xml.as_bytes()).is_err());
        }
        assert!(parse_custom_data_properties(&vec![b' '; MAX_PROPERTIES_XML_BYTES + 1]).is_err());
        let mut stores = vec![store(), store()];
        stores[1].properties_relationship_id = "rIdProps2".into();
        stores[1].properties_part_name = "/xl/customData/itemProps2.xml".into();
        stores[1].data_relationship_id = None;
        stores[1].payload = None;
        assert!(validate_store_set(&stores).is_err());
    }

    #[test]
    fn rejects_external_wrong_type_or_outbound_graphs() {
        let (mut external_package, workbook) = package();
        external_package
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                CUSTOM_DATA_PROPERTIES_RELATIONSHIP_TYPE.into(),
                "https://example.invalid/itemProps.xml".into(),
                "rIdExternal".into(),
                true,
            );
        assert!(load_custom_data(&external_package, &workbook).is_err());
        let (mut package, workbook) = package();
        store_custom_data(&mut package, &workbook, &[store()]).unwrap();
        let payload = PackURI::new("/xl/customData/item1.bin").unwrap();
        package
            .get_part_mut(&payload)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.bin".into(),
                "rIdBad".into(),
                false,
            );
        assert!(load_custom_data(&package, &workbook).is_err());
    }
}
