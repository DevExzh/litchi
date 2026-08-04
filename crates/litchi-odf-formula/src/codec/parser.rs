//! Bounded MathML parsing for OpenDocument Formula packages.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::model::{Attribute, Content, Element, MATHML_NAMESPACE};

const MAX_MATH_DEPTH: usize = 128;
const MAX_MATH_NODES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 32 * 1_048_576;

pub fn parse(xml: &str) -> Result<Element> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut root_closed = false;
    let mut node_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid formula MathML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "formula contains multiple root elements".to_string(),
                    ));
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                if stack.is_empty()
                    && (node.namespace_uri() != Some(MATHML_NAMESPACE)
                        || node.local_name() != "math")
                {
                    return Err(Error::InvalidFormat(
                        "formula content must have a MathML math root".to_string(),
                    ));
                }
                stack.push(node);
                if stack.len() > MAX_MATH_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "MathML nesting exceeds {MAX_MATH_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                if stack.is_empty() {
                    if root_closed {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple root elements".to_string(),
                        ));
                    }
                    let resolved_namespace_uri = namespace_uri(&namespace)?;
                    let node =
                        make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                    if node.namespace_uri() != Some(MATHML_NAMESPACE) || node.local_name() != "math"
                    {
                        return Err(Error::InvalidFormat(
                            "formula content must have a MathML math root".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                    buffer.clear();
                    continue;
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                stack
                    .last_mut()
                    .expect("parent exists")
                    .content_mut()
                    .push(Content::Element(node));
            },
            Event::End(_) => {
                let node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("MathML element stack underflow".to_string())
                })?;
                if let Some(parent) = stack.last_mut() {
                    parent.content_mut().push(Content::Element(node));
                } else {
                    if root.is_some() {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple MathML roots".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML text: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML CDATA: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                push_text(
                    stack.last_mut().expect("element exists"),
                    decode_reference(reference)?,
                    &mut text_bytes,
                )?;
            },
            Event::Text(ref text) if !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !root_closed {
        return Err(Error::InvalidFormat(
            "formula contains incomplete MathML".to_string(),
        ));
    }
    root.ok_or_else(|| Error::InvalidFormat("formula has no MathML root".to_string()))
}

fn make_element(
    reader: &NsReader<&[u8]>,
    resolved_namespace_uri: Option<String>,
    element: &BytesStart<'_>,
    node_count: &mut usize,
) -> Result<Element> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("MathML node count overflow".to_string()))?;
    if *node_count > MAX_MATH_NODES {
        return Err(Error::InvalidFormat(format!(
            "formula exceeds {MAX_MATH_NODES} MathML elements"
        )));
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "MathML element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let local_name = decode_utf8(element.local_name().as_ref(), "element name")?;
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid MathML attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let local_name = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace_uri() == namespace_uri.as_deref()
                && existing.local_name() == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded MathML attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid MathML attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "MathML attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(Attribute::from_parts(namespace_uri, local_name, value));
    }
    Ok(Element::from_parts(
        resolved_namespace_uri,
        local_name,
        attributes,
        Vec::new(),
    ))
}

fn push_text(element: &mut Element, value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("MathML text size overflow".to_string()))?;
    if *total > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(
            "formula exceeds 32 MiB of MathML text".to_string(),
        ));
    }
    if let Some(Content::Text(existing)) = element.content_mut().last_mut() {
        existing.push_str(&value);
    } else {
        element.content_mut().push(Content::Text(value));
    }
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown MathML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 MathML {kind}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid MathML character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid MathML entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Ok(format!("&{name};")),
    }
}
