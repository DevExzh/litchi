//! Bounded, namespace-aware codecs for Custom XML Data Storage.

use crate::mce::{Capabilities, Limits, Name, process_markup_compatibility};
use crate::xml::decode_xml_reference;
use crate::xml_name;
use crate::{Error, Result};
use litchi_opc::ContentType;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

use super::model::{
    Conformance, MAX_DEPTH, MAX_ELEMENTS, MAX_PART_BYTES, MAX_PROPS_BYTES, MAX_SCHEMA_REFS,
    MAX_STRING_BYTES, Props, STRICT_NAMESPACE, STRICT_PROPS_RELATIONSHIP, STRICT_RELATIONSHIP,
    TRANSITIONAL_NAMESPACE, TRANSITIONAL_PROPS_RELATIONSHIP, TRANSITIONAL_RELATIONSHIP,
};

/// Parse a Custom XML Data Storage Properties part with bounded MCE handling.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn read_props(xml: &[u8]) -> Result<Props> {
    require_at_most("custom XML properties bytes", xml.len(), MAX_PROPS_BYTES)?;
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities
        .understand_namespace(TRANSITIONAL_NAMESPACE)
        .understand_namespace(STRICT_NAMESPACE);
    let limits = Limits {
        max_input_bytes: MAX_PROPS_BYTES,
        max_output_bytes: MAX_PROPS_BYTES.saturating_mul(2),
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    parse_props_xml(processed.xml.as_ref())
}

/// Serialize properties in stable schema order.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn write_props(props: &Props, conformance: Conformance) -> Result<Vec<u8>> {
    validate_props(props)?;
    let output_len = props_output_len(props, conformance)?;
    let mut out = Vec::with_capacity(output_len);
    out.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    push_escaped_attr(&mut out, conformance.namespace());
    out.extend_from_slice(b"\" ds:itemID=\"");
    push_escaped_attr(&mut out, &props.id);
    out.extend_from_slice(b"\">");
    if !props.schemas.is_empty() {
        out.extend_from_slice(b"<ds:schemaRefs>");
        for uri in &props.schemas {
            out.extend_from_slice(b"<ds:schemaRef ds:uri=\"");
            push_escaped_attr(&mut out, uri);
            out.extend_from_slice(b"\"/>");
        }
        out.extend_from_slice(b"</ds:schemaRefs>");
    }
    out.extend_from_slice(b"</ds:datastoreItem>");
    if out.len() != output_len {
        return invalid("custom XML properties size calculation disagrees with serialization");
    }
    Ok(out)
}

/// Rewrite the typed properties projection while retaining source markup when
/// the schema-reference topology is unchanged. The bounded lexical patch is
/// deliberately limited to the known `itemID` attribute; structural schema
/// edits use the deterministic writer and never interpret extension XML.
pub(super) fn rewrite_props(
    source: &[u8],
    before: &Props,
    after: &Props,
    conformance: Conformance,
) -> Result<Vec<u8>> {
    if before == after {
        return Ok(source.to_vec());
    }
    if read_props(source)? != *before {
        return invalid("custom XML properties source projection is stale");
    }
    if before.schemas == after.schemas
        && let Some((value_start, value_end)) = item_id_span(source)?
    {
        let mut output = Vec::with_capacity(
            source
                .len()
                .saturating_sub(value_end.saturating_sub(value_start))
                .saturating_add(escaped_attr_len(&after.id)?),
        );
        output.extend_from_slice(&source[..value_start]);
        push_escaped_attr(&mut output, &after.id);
        output.extend_from_slice(&source[value_end..]);
        if read_props(&output).is_ok_and(|props| props == *after) {
            return Ok(output);
        }
    }
    write_props(after, conformance)
}

fn item_id_span(source: &[u8]) -> Result<Option<(usize, usize)>> {
    let mut start = None;
    let mut position = 0usize;
    while let Some(offset) = source
        .get(position..)
        .and_then(|tail| tail.iter().position(|b| *b == b'<'))
    {
        let tag_start = position + offset;
        if source.get(tag_start + 1) == Some(&b'?')
            || source.get(tag_start + 1) == Some(&b'!')
            || source.get(tag_start + 1) == Some(&b'/')
        {
            position = tag_start.saturating_add(2);
            continue;
        }
        start = Some(tag_start);
        break;
    }
    let Some(tag_start) = start else {
        return Ok(None);
    };
    let mut quote = None;
    let mut tag_end = None;
    for (offset, byte) in source[tag_start..].iter().enumerate() {
        match (quote, *byte) {
            (None, b'\'' | b'"') => quote = Some(*byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => {
                tag_end = Some(tag_start + offset);
                break;
            },
            _ => {},
        }
    }
    let Some(tag_end) = tag_end else {
        return invalid("custom XML properties root tag is incomplete");
    };
    let mut cursor = tag_start + 1;
    while cursor < tag_end && !matches!(source[cursor], b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>')
    {
        cursor += 1;
    }
    while cursor < tag_end {
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || source[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < tag_end
            && !source[cursor].is_ascii_whitespace()
            && !matches!(source[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name = &source[name_start..cursor];
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || source[cursor] != b'=' {
            return invalid("custom XML properties root attribute is malformed");
        }
        cursor += 1;
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let Some(&delimiter) = source.get(cursor) else {
            return invalid("custom XML properties root attribute has no value");
        };
        if !matches!(delimiter, b'\'' | b'"') {
            return invalid("custom XML properties root attribute is unquoted");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag_end && source[cursor] != delimiter {
            cursor += 1;
        }
        if cursor >= tag_end {
            return invalid("custom XML properties root attribute is unterminated");
        }
        if name
            .rsplit(|byte| *byte == b':')
            .next()
            .is_some_and(|local| local == b"itemID")
        {
            return Ok(Some((value_start, cursor)));
        }
        cursor += 1;
    }
    Ok(None)
}

/// Validate a typed properties value without serializing it.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_props(props: &Props) -> Result<()> {
    if !valid_guid(&props.id) {
        return invalid(format!("custom XML itemID '{}' is not ST_Guid", props.id));
    }
    require_at_most(
        "custom XML schema references",
        props.schemas.len(),
        MAX_SCHEMA_REFS,
    )?;
    let mut string_bytes = props.id.len();
    for schema in &props.schemas {
        if schema.is_empty() {
            return invalid("custom XML schema reference URI must not be empty");
        }
        validate_xml_chars(schema)?;
        string_bytes = string_bytes.checked_add(schema.len()).ok_or_else(|| {
            limit(
                "custom XML property string bytes",
                MAX_STRING_BYTES,
                usize::MAX,
            )
        })?;
    }
    require_at_most(
        "custom XML property string bytes",
        string_bytes,
        MAX_STRING_BYTES,
    )
}

/// Return whether `value` is the braced hexadecimal `ST_Guid` lexical form.
#[must_use]
pub fn valid_guid(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
    else {
        return false;
    };
    let groups = [8, 4, 4, 4, 12];
    let mut parts = inner.split('-');
    groups.into_iter().all(|length| {
        parts.next().is_some_and(|part| {
            part.len() == length && part.bytes().all(|byte| byte.is_ascii_hexdigit())
        })
    }) && parts.next().is_none()
}

/// Validate a bounded XML payload and return its expanded document-element name.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_payload(xml: &[u8]) -> Result<Name> {
    require_at_most("custom XML payload bytes", xml.len(), MAX_PART_BYTES)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    let mut closed_root = false;
    let mut elements = 0usize;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML payload has a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                require_nested_depth(depth)?;
                elements = bump_elements(elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("custom XML depth", MAX_DEPTH, usize::MAX))?;
            },
            Event::Empty(element) => {
                require_nested_depth(depth)?;
                elements = bump_elements(elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                    closed_root = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("custom XML payload has an unexpected end tag".into())
                })?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML payloads"),
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return invalid("custom XML payload has a reference outside its root");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
            },
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 && !is_xml_whitespace(decoded.as_bytes()) {
                    return invalid("custom XML payload has text outside its root");
                }
            },
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 {
                    return invalid("custom XML payload has CDATA outside its root");
                }
            },
            Event::Comment(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
            },
            Event::PI(instruction) => validate_instruction(&reader, &instruction)?,
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if depth != 0 || !closed_root {
        return invalid("custom XML payload has no complete root element");
    }
    root.ok_or_else(|| Error::Invalid("custom XML payload has no root".into()))
}

/// Validate that a Custom XML data-part content type is a well-formed XML media type.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_content_type(content_type: &str) -> Result<()> {
    let parsed = ContentType::new(content_type).map_err(|_error| Error::ContentType {
        expected: "a well-formed XML media type".into(),
        actual: content_type.into(),
    })?;
    let media_type = parsed.as_str().split(';').next().unwrap_or_default();
    let Some((type_name, subtype)) = media_type.split_once('/') else {
        return Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        });
    };
    let xml = (type_name.eq_ignore_ascii_case("application")
        || type_name.eq_ignore_ascii_case("text"))
        && subtype.eq_ignore_ascii_case("xml");
    let suffix = subtype.len() > 4
        && subtype
            .get(subtype.len() - 4..)
            .is_some_and(|value| value.eq_ignore_ascii_case("+xml"));
    if xml || suffix {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        })
    }
}

#[derive(Debug)]
struct ResolvedElement {
    namespace: String,
    local_name: String,
    attributes: Vec<(String, String, String)>,
}

fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
    capture: bool,
) -> Result<Option<Name>> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = resolved_namespace(resolved, "element")?;
    let namespace =
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
    let local =
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let captured = capture.then(|| Name {
        namespace: namespace.into(),
        local_name: local.into(),
    });

    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(resolved, "attribute")?;
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
        if !seen.insert((namespace.to_vec(), local.as_ref().to_vec())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
    }
    Ok(captured)
}

fn resolve_props_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
) -> Result<ResolvedElement> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = std::str::from_utf8(resolved_namespace(resolved, "element")?)
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let local_name = std::str::from_utf8(local.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = std::str::from_utf8(resolved_namespace(resolved, "attribute")?)
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let attribute_name = std::str::from_utf8(local.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert((attribute_namespace.clone(), attribute_name.clone())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        attributes.push((attribute_namespace, attribute_name, value));
    }
    Ok(ResolvedElement {
        namespace,
        local_name,
        attributes,
    })
}

fn resolved_namespace<'a>(result: ResolveResult<'a>, kind: &str) -> Result<&'a [u8]> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(value),
        ResolveResult::Unbound => Ok(b""),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn parse_props_xml(xml: &[u8]) -> Result<Props> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut contexts = Vec::new();
    let mut id = None;
    let mut schemas = Vec::new();
    let mut root_namespace = None;
    let mut seen_schemas = false;
    let mut closed_root = false;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML properties have a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                require_nested_depth(contexts.len())?;
                let element = resolve_props_element(&reader, &element, version)?;
                let context = props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                contexts.push(context);
            },
            Event::Empty(element) => {
                require_nested_depth(contexts.len())?;
                let element = resolve_props_element(&reader, &element, version)?;
                props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::End(_) => {
                if contexts.pop().is_none() {
                    return invalid("unexpected custom XML properties end tag");
                }
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML properties"),
            Event::Text(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(text.as_bytes())
                {
                    return invalid("text is not permitted in custom XML properties");
                }
            },
            Event::CData(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if contexts.is_empty() {
                    return invalid("CDATA is not permitted outside custom XML properties");
                }
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(text.as_bytes())
                {
                    return invalid("CDATA is not permitted in custom XML properties");
                }
            },
            Event::GeneralRef(reference) => {
                if contexts.is_empty() {
                    return invalid("references are not permitted outside custom XML properties");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
                if !matches!(contexts.last(), Some(PropsContext::Opaque))
                    && !is_xml_whitespace(value.as_bytes())
                {
                    return invalid("references are not permitted in custom XML properties");
                }
            },
            Event::Comment(text) => {
                text.decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::PI(_) if !matches!(contexts.last(), Some(PropsContext::Opaque)) => {
                return invalid("processing instructions are forbidden in custom XML properties");
            },
            Event::PI(_) => {},
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if !closed_root || !contexts.is_empty() {
        return invalid("custom XML properties root is incomplete");
    }
    let props = Props {
        id: id.ok_or_else(|| Error::Invalid("datastoreItem requires ds:itemID".into()))?,
        schemas,
    };
    validate_props(&props)?;
    Ok(props)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PropsContext {
    Item,
    Schemas,
    Schema,
    Opaque,
}

fn props_start(
    element: &ResolvedElement,
    parent: Option<PropsContext>,
    id: &mut Option<String>,
    schemas: &mut Vec<String>,
    root_namespace: &mut Option<String>,
    seen_schemas: &mut bool,
) -> Result<PropsContext> {
    let supported_namespace = is_custom_namespace(&element.namespace);
    match parent {
        None if root_namespace.is_none()
            && supported_namespace
            && element.local_name == "datastoreItem" =>
        {
            reject_attributes_except(element, &[(element.namespace.as_str(), "itemID")]);
            *id = Some(required_attr(element, &element.namespace, "itemID")?.into());
            *root_namespace = Some(element.namespace.clone());
            Ok(PropsContext::Item)
        },
        Some(PropsContext::Item)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRefs" =>
        {
            if *seen_schemas {
                return invalid("datastoreItem has multiple schemaRefs elements");
            }
            *seen_schemas = true;
            reject_attributes_except(element, &[]);
            Ok(PropsContext::Schemas)
        },
        Some(PropsContext::Schemas)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRef" =>
        {
            let next = schemas.len().checked_add(1).ok_or_else(|| {
                limit("custom XML schema references", MAX_SCHEMA_REFS, usize::MAX)
            })?;
            require_at_most("custom XML schema references", next, MAX_SCHEMA_REFS)?;
            reject_attributes_except(element, &[(element.namespace.as_str(), "uri")]);
            schemas.push(required_attr(element, &element.namespace, "uri")?.into());
            Ok(PropsContext::Schema)
        },
        Some(PropsContext::Opaque) => Ok(PropsContext::Opaque),
        Some(PropsContext::Item | PropsContext::Schemas | PropsContext::Schema) => {
            // Future properties markup is retained in the source part and is
            // intentionally not interpreted by this host-neutral layer.
            Ok(PropsContext::Opaque)
        },
        _ => invalid(format!(
            "unexpected custom XML properties element {{{}}}{}",
            element.namespace, element.local_name
        )),
    }
}

fn reject_attributes_except(element: &ResolvedElement, allowed: &[(&str, &str)]) {
    // Unknown attributes are opaque extension data. Known attributes are
    // still checked by `required_attr`; retaining this helper keeps the
    // semantic call sites explicit without discarding future markup.
    let _ = (element, allowed);
}

fn required_attr<'a>(
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
            Error::Invalid(format!(
                "{} requires {{{namespace}}}{local_name}",
                element.local_name
            ))
        })
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<XmlVersion> {
    let version = declaration.xml_version()?;
    let declaration_text =
        std::str::from_utf8(declaration.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut declaration_state = 0u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            return invalid(format!(
                "unexpected XML declaration attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        declaration_state = match (declaration_state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return invalid(format!(
                    "unexpected or out-of-order XML declaration attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                ));
            },
        };
        std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        let encoding =
            std::str::from_utf8(&encoding).map_err(|error| Error::Xml(error.to_string()))?;
        if !valid_encoding_name(encoding) {
            return invalid(format!(
                "XML declaration encoding '{encoding}' is not an EncName"
            ));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return invalid("XML declaration standalone must be 'yes' or 'no'");
        }
    }
    Ok(version)
}

fn validate_instruction(
    reader: &NsReader<&[u8]>,
    instruction: &quick_xml::events::BytesPI<'_>,
) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if !xml_name::is_xml_name(&target) {
        return Err(Error::Xml(format!(
            "invalid processing-instruction target '{target}'"
        )));
    }
    if target.eq_ignore_ascii_case("xml") {
        return invalid("processing-instruction target cannot be 'xml'");
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| Error::Xml(error.to_string()))?;
    validate_xml_chars(&content)?;
    Ok(())
}

fn props_output_len(props: &Props, conformance: Conformance) -> Result<usize> {
    const DECL: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    const ROOT_PREFIX: &[u8] = b"<ds:datastoreItem xmlns:ds=\"";
    const ID_PREFIX: &[u8] = b"\" ds:itemID=\"";
    const ROOT_OPEN_END: &[u8] = b"\">";
    const SCHEMAS_OPEN: &[u8] = b"<ds:schemaRefs>";
    const SCHEMA_PREFIX: &[u8] = b"<ds:schemaRef ds:uri=\"";
    const SCHEMA_END: &[u8] = b"\"/>";
    const SCHEMAS_END: &[u8] = b"</ds:schemaRefs>";
    const ROOT_END: &[u8] = b"</ds:datastoreItem>";

    let mut total = 0usize;
    add_len(&mut total, DECL.len())?;
    add_len(&mut total, ROOT_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(conformance.namespace())?)?;
    add_len(&mut total, ID_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(&props.id)?)?;
    add_len(&mut total, ROOT_OPEN_END.len())?;
    if !props.schemas.is_empty() {
        add_len(&mut total, SCHEMAS_OPEN.len())?;
        for schema in &props.schemas {
            add_len(&mut total, SCHEMA_PREFIX.len())?;
            add_len(&mut total, escaped_attr_len(schema)?)?;
            add_len(&mut total, SCHEMA_END.len())?;
        }
        add_len(&mut total, SCHEMAS_END.len())?;
    }
    add_len(&mut total, ROOT_END.len())?;
    require_at_most(
        "custom XML serialized properties bytes",
        total,
        MAX_PROPS_BYTES,
    )?;
    Ok(total)
}

fn escaped_attr_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let width = match character {
            '&' | '\t' | '\n' | '\r' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        total.checked_add(width).ok_or_else(|| {
            limit(
                "custom XML serialized properties bytes",
                MAX_PROPS_BYTES,
                usize::MAX,
            )
        })
    })
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total.checked_add(value).ok_or_else(|| {
        limit(
            "custom XML serialized properties bytes",
            MAX_PROPS_BYTES,
            usize::MAX,
        )
    })?;
    Ok(())
}

fn push_escaped_attr(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut encoded = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::Xml(
            "custom XML contains a character forbidden by XML 1.0".into(),
        ));
    }
    Ok(())
}

pub(super) fn require_rel_id(value: &str, label: &str) -> Result<()> {
    if xml_name::is_ncname(value) {
        Ok(())
    } else {
        Err(Error::Relationship(format!(
            "{label} ID '{value}' is not an XML NCName"
        )))
    }
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    if xml_name::is_qualified_name(value) {
        Ok(())
    } else {
        Err(Error::Xml(format!("invalid XML {kind} QName '{value}'")))
    }
}

fn validate_namespace_declaration(value: &[u8]) -> Result<()> {
    if value == b"xmlns" {
        return Ok(());
    }
    let Some(prefix) = value.strip_prefix(b"xmlns:") else {
        return Err(Error::Xml("invalid XML namespace declaration".into()));
    };
    let prefix = std::str::from_utf8(prefix).map_err(|error| Error::Xml(error.to_string()))?;
    if xml_name::is_ncname(prefix) {
        Ok(())
    } else {
        Err(Error::Xml(format!(
            "invalid XML namespace prefix '{prefix}'"
        )))
    }
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn require_nested_depth(depth: usize) -> Result<()> {
    if depth >= MAX_DEPTH {
        Err(limit(
            "custom XML depth",
            MAX_DEPTH,
            depth.saturating_add(1),
        ))
    } else {
        Ok(())
    }
}

fn bump_elements(elements: usize) -> Result<usize> {
    let next = elements
        .checked_add(1)
        .ok_or_else(|| limit("custom XML elements", MAX_ELEMENTS, usize::MAX))?;
    require_at_most("custom XML elements", next, MAX_ELEMENTS)?;
    Ok(next)
}

pub(super) fn require_at_most(resource: &'static str, actual: usize, max: usize) -> Result<()> {
    if actual > max {
        Err(limit(resource, max, actual))
    } else {
        Ok(())
    }
}

pub(super) fn limit(resource: &'static str, max: usize, actual: usize) -> Error {
    Error::Limit {
        resource,
        max,
        actual,
    }
}

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Invalid(message.into()))
}

pub(super) fn is_data_relationship(value: &str) -> bool {
    matches!(value, TRANSITIONAL_RELATIONSHIP | STRICT_RELATIONSHIP)
}

pub(super) fn is_props_relationship(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_PROPS_RELATIONSHIP | STRICT_PROPS_RELATIONSHIP
    )
}

fn is_custom_namespace(value: &str) -> bool {
    matches!(value, TRANSITIONAL_NAMESPACE | STRICT_NAMESPACE)
}

fn is_namespace_declaration(value: &[u8]) -> bool {
    value == b"xmlns" || value.starts_with(b"xmlns:")
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}
