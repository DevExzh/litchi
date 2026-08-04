//! Bounded SpreadsheetML custom-data properties XML codec.

use super::model::{ExtensionList, Properties};
use crate::error::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

const MAX_PROPERTIES_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_XML_BYTES: usize = 2 * 1024 * 1024;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;

/// Parse a Custom Data Properties part.
pub fn parse_properties(xml: &[u8]) -> Result<Properties> {
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
            Ok(ExtensionList { xml })
        })
        .transpose()?;
    let value = Properties {
        id: required(&root, "", "id")?.to_owned(),
        extension_list,
    };
    validate_properties(&value, true)?;
    Ok(value)
}

/// Deterministically serialize a Custom Data Properties part.
pub fn write_properties(value: &Properties) -> Result<Vec<u8>> {
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

/// Validate that XML belongs to a SpreadsheetML workbook part.
//
// The host calls this while retaining ownership of package graph traversal.
// Keeping the bounded XML root check beside the shared parser avoids a second
// permissive parser in the compatibility layer.
pub fn validate_workbook_root(xml: &[u8]) -> Result<()> {
    let root = parse_document(xml)?;
    require_workbook_root(&root)
}

fn validate_properties(value: &Properties, extension_already_parsed: bool) -> Result<()> {
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
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
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

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(name: &str) -> Error {
    invalid(format!("Custom Data {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn properties() -> Properties {
        Properties {
            id: "Storage-1".into(),
            extension_list: Some(ExtensionList {
                xml: format!(
                    r#"<x14:extLst xmlns:x14="{X14}"><x14:ext uri="urn:test"><v:opaque xmlns:v="urn:vendor" value="kept"/></x14:ext></x14:extLst>"#
                )
                .into_bytes(),
            }),
        }
    }

    #[test]
    fn typed_properties_and_extensions_round_trip() {
        let expected = properties();
        let xml = write_properties(&expected).unwrap();
        let parsed = parse_properties(&xml).unwrap();
        assert_eq!(parsed.id, expected.id);
        assert!(String::from_utf8_lossy(&parsed.extension_list.unwrap().xml).contains("opaque"));
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
            assert!(parse_properties(xml.as_bytes()).is_err());
        }
        assert!(parse_properties(&vec![b' '; MAX_PROPERTIES_XML_BYTES + 1]).is_err());
    }

    #[test]
    fn validates_workbook_namespace() {
        assert!(validate_workbook_root(format!(r#"<workbook xmlns="{SML}"/>"#).as_bytes()).is_ok());
        assert!(validate_workbook_root(b"<workbook xmlns=\"urn:not-spreadsheetml\"/>").is_err());
    }
}
