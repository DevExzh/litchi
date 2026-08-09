//! Bounded `MathML` parsing for `OpenDocument` Formula packages.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::Limits;
use crate::model::{Attribute, Content, Element, MATHML_NAMESPACE};

/// Parse bounded `MathML` markup into an inert element tree.
///
/// # Errors
///
/// Returns an error when the markup is not well-formed, when the root is not
/// a `math` element in the `MathML` namespace, or when a safety limit on
/// depth, node count, attribute count or size, or text size is exceeded.
///
pub fn parse(xml: &str) -> Result<Element> {
    parse_with_limits(xml, Limits::default())
}

/// Parse inert `MathML` using caller-selected finite limits.
///
/// # Errors
///
/// Returns an error for malformed markup, an invalid root, or any exceeded
/// byte, depth, element, attribute, or text ceiling.
pub fn parse_with_limits(xml: &str, limits: Limits) -> Result<Element> {
    if xml.len() > limits.xml_bytes() {
        return Err(Error::InvalidFormat(format!(
            "formula content.xml has {} bytes, exceeding the {} byte limit",
            xml.len(),
            limits.xml_bytes()
        )));
    }
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
                let node = make_element(
                    &reader,
                    resolved_namespace_uri,
                    element,
                    &mut node_count,
                    limits,
                )?;
                if stack.is_empty()
                    && (node.namespace_uri() != Some(MATHML_NAMESPACE)
                        || node.local_name() != "math")
                {
                    return Err(Error::InvalidFormat(
                        "formula content must have a MathML math root".to_string(),
                    ));
                }
                stack.push(node);
                if stack.len() > limits.depth() {
                    return Err(Error::InvalidFormat(format!(
                        "MathML nesting exceeds {} levels",
                        limits.depth()
                    )));
                }
            },
            Event::Empty(ref element) => {
                let element_depth = stack.len().checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("MathML nesting depth overflow".to_string())
                })?;
                if element_depth > limits.depth() {
                    return Err(Error::InvalidFormat(format!(
                        "MathML nesting exceeds {} levels",
                        limits.depth()
                    )));
                }
                if stack.is_empty() {
                    if root_closed {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple root elements".to_string(),
                        ));
                    }
                    let resolved_namespace_uri = namespace_uri(&namespace)?;
                    let node = make_element(
                        &reader,
                        resolved_namespace_uri,
                        element,
                        &mut node_count,
                        limits,
                    )?;
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
                let node = make_element(
                    &reader,
                    resolved_namespace_uri,
                    element,
                    &mut node_count,
                    limits,
                )?;
                let parent = stack.last_mut().ok_or_else(|| {
                    Error::InvalidFormat("MathML parent stack is empty".to_string())
                })?;
                parent.content_mut().push(Content::Element(node));
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
                let top = stack.last_mut().ok_or_else(|| {
                    Error::InvalidFormat("MathML text stack is empty".to_string())
                })?;
                push_text(top, value.into_owned(), &mut text_bytes, limits)?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML CDATA: {error}"))
                })?;
                let top = stack.last_mut().ok_or_else(|| {
                    Error::InvalidFormat("MathML CDATA stack is empty".to_string())
                })?;
                push_text(top, value.into_owned(), &mut text_bytes, limits)?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                let top = stack.last_mut().ok_or_else(|| {
                    Error::InvalidFormat("MathML entity stack is empty".to_string())
                })?;
                push_text(top, decode_reference(reference)?, &mut text_bytes, limits)?;
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
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
    limits: Limits,
) -> Result<Element> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("MathML node count overflow".to_string()))?;
    if *node_count > limits.nodes() {
        return Err(Error::InvalidFormat(format!(
            "formula exceeds {} MathML elements",
            limits.nodes()
        )));
    }
    if element.attributes().count() > limits.attributes() {
        return Err(Error::InvalidFormat(format!(
            "MathML element exceeds {} attributes",
            limits.attributes()
        )));
    }
    let local_name = decode_utf8(element.local_name().as_ref(), "element name")?;
    let mut attributes = Vec::new();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid MathML attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let attribute_local_name = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace_uri() == namespace_uri.as_deref()
                && existing.local_name() == attribute_local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded MathML attribute '{attribute_local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid MathML attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > limits.attribute_bytes() {
            return Err(Error::InvalidFormat(format!(
                "MathML attribute exceeds {} bytes",
                limits.attribute_bytes()
            )));
        }
        attributes.push(Attribute::from_parts(
            namespace_uri,
            attribute_local_name,
            value,
        ));
    }
    Ok(Element::from_parts(
        resolved_namespace_uri,
        local_name,
        attributes,
        Vec::new(),
    ))
}

fn push_text(
    element: &mut Element,
    value: String,
    total: &mut usize,
    limits: Limits,
) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("MathML text size overflow".to_string()))?;
    if *total > limits.text_bytes() {
        return Err(Error::InvalidFormat(format!(
            "formula exceeds {} bytes of MathML text",
            limits.text_bytes()
        )));
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
    match std::str::from_utf8(bytes) {
        Ok(text) => Ok(text.to_string()),
        Err(_) => Err(Error::InvalidFormat(format!("non-UTF-8 MathML {kind}"))),
    }
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
