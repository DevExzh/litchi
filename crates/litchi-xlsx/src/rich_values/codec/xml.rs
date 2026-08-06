//! Small namespace-aware XML DOM used only at the rich-values boundary.

use crate::error::Result;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap};

use super::super::{
    MAX_DEPTH, MAX_NODES, MAX_STRING_BYTES, MAX_XML_BYTES, invalid, limit, xml_error,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Attribute {
    pub(crate) namespace: String,
    pub(crate) name: String,
    pub(crate) value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Node {
    pub(crate) namespace: String,
    pub(crate) name: String,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) children: Vec<Node>,
    pub(crate) text: String,
}

pub(crate) fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let is_empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, &element, reader.decoder(), &mut strings)?;
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
                add_string(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::CData(text) => {
                let decoded = std::str::from_utf8(text.as_ref()).map_err(xml_error)?;
                add_string(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("CDATA outside XML root"));
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
                add_string(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated rich-values XML"));
    }
    root.ok_or_else(|| invalid("missing rich-values XML root"))
}

pub(crate) fn validate_fragment(xml: &[u8]) -> Result<()> {
    parse_document(xml).map(|_| ())
}

pub(crate) fn opaque(node: &Node) -> Result<super::super::model::Opaque> {
    let xml = serialize_node(node)?;
    if xml.len() > super::super::MAX_OPAQUE_BYTES {
        return Err(limit("opaque XML"));
    }
    Ok(super::super::model::Opaque::from_serialized(xml))
}

pub(crate) fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::<String, ()>::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            super::super::RICH_DATA => "rd".to_owned(),
            super::super::RICH_DATA_2 => "rd2".to_owned(),
            super::super::FEATURE_BAG => "fpb".to_owned(),
            super::super::RICH_VALUE_REL => "rvr".to_owned(),
            super::super::RELATIONSHIPS | super::super::STRICT_RELATIONSHIPS => "r".to_owned(),
            super::super::SPREADSHEETML => "x".to_owned(),
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

pub(crate) fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    attribute(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}

pub(crate) fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    attribute(node, namespace, name)
}

pub(crate) fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        return Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )));
    }
    Ok(())
}

pub(crate) fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}

pub(crate) fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected {{{namespace}}}{name}, got {{{}}}{}",
            node.namespace, node.name
        )))
    }
}

pub(crate) fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape_attr(output, value);
    output.push(b'\"');
}

pub(crate) fn escape_attr(output: &mut Vec<u8>, value: &str) {
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

pub(crate) fn escape_text(output: &mut Vec<u8>, value: &str) {
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
    add_string(strings, namespace.len() + name.len())?;
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
        add_string(strings, namespace.len() + name.len() + value.len())?;
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

fn attribute<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn add_string(total: &mut usize, size: usize) -> Result<()> {
    *total = total.checked_add(size).ok_or_else(|| limit("XML string"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string"))
    } else {
        Ok(())
    }
}
