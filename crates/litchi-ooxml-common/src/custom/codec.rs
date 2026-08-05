//! Namespace-aware XML codec for the custom-properties part.

use chrono::{DateTime, SecondsFormat, Utc};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;
use std::borrow::Cow;
use std::collections::{BTreeSet, HashSet};

use super::model::{
    Props, Value, WireKind, fold_name, validate_name, validate_value, validate_xml_text,
    value_text_bytes,
};
use super::schema::*;
use crate::{Error, Result};

struct PendingProperty {
    name: String,
    pid: i32,
    format_id: String,
    value: Option<(WireKind, Value)>,
}

pub(crate) fn decode(xml: &[u8]) -> Result<Props> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut props = Props::new();
    let mut pids = HashSet::new();
    let mut pending: Option<PendingProperty> = None;
    let mut value_kind: Option<WireKind> = None;
    let mut value_text = String::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Decl(_) => {
                count_node(&mut nodes)?;
                if declaration_seen || root_seen {
                    return Err(invalid("XML declaration must occur once before the root"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD declarations are forbidden in custom-properties XML",
                ));
            },
            Event::PI(_) => {
                return Err(invalid(
                    "processing instructions are forbidden in custom-properties XML",
                ));
            },
            Event::Start(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                    },
                    2 => {
                        if !is_name(&namespace, &element, CUSTOM_NS, b"property") {
                            return Err(invalid(format!(
                                "unexpected element '{}' below custom-properties root",
                                display_name(element.name().as_ref())
                            )));
                        }
                        if pending.is_some() {
                            return Err(invalid("custom properties cannot be nested"));
                        }
                        let parsed = parse_property_attributes(&element, decoder)?;
                        if !pids.insert(parsed.pid) {
                            return Err(invalid(format!(
                                "duplicate custom property PID {}",
                                parsed.pid
                            )));
                        }
                        pending = Some(parsed);
                    },
                    3 => {
                        let property = pending.as_ref().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        if property.value.is_some() || value_kind.is_some() {
                            return Err(invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            )));
                        }
                        value_kind = Some(parse_value_element(&namespace, &element)?);
                        value_text.clear();
                        validate_value_attributes(&element)?;
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                        root_closed = true;
                    },
                    2 => {
                        return Err(invalid(
                            "custom property must contain exactly one typed value",
                        ));
                    },
                    3 => {
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        if property.value.is_some() || value_kind.is_some() {
                            return Err(invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            )));
                        }
                        let kind = parse_value_element(&namespace, &element)?;
                        validate_value_attributes(&element)?;
                        property.value = Some((kind, parse_value(kind, "")?));
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
            },
            Event::End(_) => {
                count_node(&mut nodes)?;
                match depth {
                    3 => {
                        let kind = value_kind
                            .take()
                            .ok_or_else(|| invalid("custom-property value state is incomplete"))?;
                        let value = parse_value(kind, &value_text)?;
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        property.value = Some((kind, value));
                        value_text.clear();
                    },
                    2 => {
                        let property = pending
                            .take()
                            .ok_or_else(|| invalid("custom-property record is incomplete"))?;
                        let (wire, value) = property.value.ok_or_else(|| {
                            invalid(format!(
                                "custom property '{}' must contain exactly one value",
                                property.name
                            ))
                        })?;
                        props.insert_parsed(
                            property.name,
                            property.pid,
                            property.format_id,
                            wire,
                            value,
                        )?;
                    },
                    1 => {
                        root_closed = true;
                    },
                    _ => return Err(invalid("unexpected XML end element")),
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("custom-properties XML depth underflow"))?;
            },
            Event::Text(text) => {
                count_node(&mut nodes)?;
                let text = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(format!("invalid XML text: {error}")))?;
                if depth == 3 {
                    append_value_text(&mut value_text, &text)?;
                } else if !text.trim().is_empty() {
                    return Err(invalid(
                        "non-whitespace text is not allowed outside a property value",
                    ));
                }
            },
            Event::CData(text) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    return Err(invalid("CDATA is only allowed inside a property value"));
                }
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid CDATA text: {error}")))?;
                append_value_text(&mut value_text, &text)?;
            },
            Event::GeneralRef(reference) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    return Err(invalid(
                        "entity references are only allowed inside a property value",
                    ));
                }
                let decoded = decode_reference(&reference)?;
                append_value_text(&mut value_text, &decoded)?;
            },
            Event::Comment(_) => count_node(&mut nodes)?,
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || depth != 0 || pending.is_some() || value_kind.is_some() {
        return Err(invalid(
            "custom-properties XML must contain one complete Properties root",
        ));
    }
    Ok(props)
}

pub(crate) fn encode(props: &Props) -> Result<Vec<u8>> {
    validate_collection(props)?;
    let estimated = estimated_xml_size(props)?;
    let mut writer = Writer::new(Vec::with_capacity(estimated));
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut root = BytesStart::new("Properties");
    root.push_attribute((
        "xmlns",
        std::str::from_utf8(CUSTOM_NS).map_err(|error| {
            Error::Xml(format!(
                "invalid built-in custom-properties namespace: {error}"
            ))
        })?,
    ));
    root.push_attribute((
        "xmlns:vt",
        std::str::from_utf8(VT_NS).map_err(|error| {
            Error::Xml(format!("invalid built-in variant-types namespace: {error}"))
        })?,
    ));
    writer.write_event(Event::Start(root))?;

    let mut ordered: Vec<_> = props.properties.iter().collect();
    ordered.sort_unstable_by_key(|(_, property)| property.pid);
    for (name, property) in ordered {
        let pid = property.pid.to_string();
        let mut element = BytesStart::new("property");
        element.push_attribute(("fmtid", property.format_id.as_str()));
        element.push_attribute(("pid", pid.as_str()));
        element.push_attribute(("name", name.as_str()));
        writer.write_event(Event::Start(element))?;

        let value_name = property.wire.qualified_name();
        if property.wire == WireKind::Empty {
            writer.write_event(Event::Empty(BytesStart::new(value_name)))?;
        } else {
            writer.write_event(Event::Start(BytesStart::new(value_name)))?;
            let value = value_lexical(&property.value)?;
            writer.write_event(Event::Text(BytesText::new(&value)))?;
            writer.write_event(Event::End(BytesEnd::new(value_name)))?;
        }
        writer.write_event(Event::End(BytesEnd::new("property")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("Properties")))?;
    let xml = writer.into_inner();
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    Ok(xml)
}

fn validate_collection(props: &Props) -> Result<()> {
    if props.properties.len() > MAX_PROPERTIES {
        return Err(limit(
            "custom properties",
            MAX_PROPERTIES,
            props.properties.len(),
        ));
    }
    let mut pids = HashSet::with_capacity(props.properties.len());
    let mut folded = BTreeSet::new();
    let mut name_bytes = 0usize;
    let mut text_bytes = 0usize;
    for (name, property) in &props.properties {
        validate_name(name)?;
        validate_value(&property.value)?;
        validate_wire_value(property.wire, &property.value)?;
        if property.pid < 2 || !pids.insert(property.pid) {
            return Err(invalid(format!(
                "invalid or duplicate custom property PID {}",
                property.pid
            )));
        }
        validate_format_id(&property.format_id)?;
        if !folded.insert(fold_name(name)) {
            return Err(invalid(format!(
                "duplicate custom property name '{name}' (names are case-insensitive)"
            )));
        }
        name_bytes = checked_total(
            name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        text_bytes = checked_total(
            text_bytes,
            value_text_bytes(&property.value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
    }
    Ok(())
}

fn estimated_xml_size(props: &Props) -> Result<usize> {
    let mut size = 256usize;
    for (name, property) in &props.properties {
        size = size
            .checked_add(256)
            .and_then(|size| {
                name.len()
                    .checked_mul(6)
                    .and_then(|name| size.checked_add(name))
            })
            .and_then(|size| {
                value_text_bytes(&property.value)
                    .checked_mul(6)
                    .and_then(|text| size.checked_add(text))
            })
            .ok_or_else(|| limit("custom-properties XML bytes", MAX_XML_BYTES, usize::MAX))?;
        if size > MAX_XML_BYTES {
            return Err(limit("custom-properties XML bytes", MAX_XML_BYTES, size));
        }
    }
    Ok(size)
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    if !is_name(namespace, element, CUSTOM_NS, b"Properties") {
        return Err(invalid(format!(
            "custom-properties root must be Properties in namespace '{}'",
            String::from_utf8_lossy(CUSTOM_NS)
        )));
    }
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            let name = attribute.key.local_name().as_ref().to_vec();
            let _ = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(format!("invalid root attribute: {error}")))?;
            return Err(invalid(format!(
                "unexpected custom-properties root attribute '{}'",
                display_name(&name)
            )));
        }
    }
    Ok(())
}

fn parse_property_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<PendingProperty> {
    let mut name = None;
    let mut pid = None;
    let mut format_id = None;
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        if attribute.key.prefix().is_some() {
            return Err(invalid(format!(
                "custom property has unexpected qualified attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(format!("invalid property attribute: {error}")))?
            .into_owned();
        match attribute.key.local_name().as_ref() {
            b"name" if name.is_none() => name = Some(value),
            b"pid" if pid.is_none() => {
                let parsed = value.parse::<i32>().map_err(|error| {
                    invalid(format!("invalid custom property PID '{value}': {error}"))
                })?;
                pid = Some(parsed);
            },
            b"fmtid" if format_id.is_none() => format_id = Some(normalize_format_id(&value)?),
            local => {
                return Err(invalid(format!(
                    "duplicate or unexpected custom property attribute '{}'",
                    display_name(local)
                )));
            },
        }
    }
    let name = name.ok_or_else(|| invalid("custom property is missing its name attribute"))?;
    validate_name(&name)?;
    let pid = pid.ok_or_else(|| invalid(format!("custom property '{name}' is missing its PID")))?;
    if pid < 2 {
        return Err(invalid(format!(
            "custom property '{name}' has PID {pid}; PIDs must be at least 2"
        )));
    }
    let format_id = format_id
        .ok_or_else(|| invalid(format!("custom property '{name}' is missing its format ID")))?;
    Ok(PendingProperty {
        name,
        pid,
        format_id,
        value: None,
    })
}

fn validate_value_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "custom-property value has unexpected attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn parse_value_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<WireKind> {
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == VT_NS) {
        return Err(invalid(format!(
            "custom-property value '{}' is not in the variant-types namespace",
            display_name(element.name().as_ref())
        )));
    }
    WireKind::from_local_name(element.local_name().as_ref()).ok_or_else(|| {
        invalid(format!(
            "unsupported custom-property value type '{}'",
            display_name(element.local_name().as_ref())
        ))
    })
}

fn parse_value(kind: WireKind, text: &str) -> Result<Value> {
    match kind {
        WireKind::Empty => {
            if text.is_empty() {
                Ok(Value::Empty)
            } else {
                Err(invalid("vt:empty cannot contain text"))
            }
        },
        WireKind::Lpstr | WireKind::Lpwstr => {
            validate_xml_text(text, "custom-property text")?;
            Ok(Value::Text(text.to_owned()))
        },
        WireKind::I4 => text
            .trim()
            .parse::<i32>()
            .map(Value::I32)
            .map_err(|error| invalid(format!("invalid vt:i4 value '{text}': {error}"))),
        WireKind::I8 => text
            .trim()
            .parse::<i64>()
            .map(Value::I64)
            .map_err(|error| invalid(format!("invalid vt:i8 value '{text}': {error}"))),
        WireKind::R4 => {
            let value = text
                .trim()
                .parse::<f32>()
                .map_err(|error| invalid(format!("invalid vt:r4 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r4 custom property must be finite"));
            }
            Ok(Value::F32(value))
        },
        WireKind::R8 => {
            let value = text
                .trim()
                .parse::<f64>()
                .map_err(|error| invalid(format!("invalid vt:r8 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r8 custom property must be finite"));
            }
            Ok(Value::F64(value))
        },
        WireKind::Bool => match text.trim() {
            "true" | "1" => Ok(Value::Bool(true)),
            "false" | "0" => Ok(Value::Bool(false)),
            value => Err(invalid(format!("invalid vt:bool value '{value}'"))),
        },
        WireKind::Filetime => {
            let value = DateTime::parse_from_rfc3339(text.trim()).map_err(|error| {
                invalid(format!(
                    "invalid vt:filetime RFC3339 date-time '{text}': {error}"
                ))
            })?;
            Ok(Value::Time(value.with_timezone(&Utc)))
        },
    }
}

fn value_lexical(value: &Value) -> Result<Cow<'_, str>> {
    validate_value(value)?;
    Ok(match value {
        Value::Empty => Cow::Borrowed(""),
        Value::Text(text) => Cow::Borrowed(text),
        Value::I32(value) => Cow::Owned(value.to_string()),
        Value::I64(value) => Cow::Owned(value.to_string()),
        Value::F32(value) => Cow::Owned(value.to_string()),
        Value::F64(value) => Cow::Owned(value.to_string()),
        Value::Bool(true) => Cow::Borrowed("true"),
        Value::Bool(false) => Cow::Borrowed("false"),
        Value::Time(value) => Cow::Owned(value.to_rfc3339_opts(SecondsFormat::AutoSi, true)),
    })
}

fn validate_wire_value(wire: WireKind, value: &Value) -> Result<()> {
    let matches = matches!(
        (wire, value),
        (WireKind::Empty, Value::Empty)
            | (WireKind::Lpstr | WireKind::Lpwstr, Value::Text(_))
            | (WireKind::I4, Value::I32(_))
            | (WireKind::I8, Value::I64(_))
            | (WireKind::R4, Value::F32(_))
            | (WireKind::R8, Value::F64(_))
            | (WireKind::Bool, Value::Bool(_))
            | (WireKind::Filetime, Value::Time(_))
    );
    if matches {
        Ok(())
    } else {
        Err(invalid(
            "custom-property wire type does not match its value",
        ))
    }
}

fn validate_format_id(format_id: &str) -> Result<()> {
    let normalized = normalize_format_id(format_id)?;
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(())
}

fn normalize_format_id(format_id: &str) -> Result<String> {
    let bytes = format_id.as_bytes();
    let valid = bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && [9, 14, 19, 24]
            .iter()
            .all(|position| bytes.get(*position) == Some(&b'-'))
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || byte.is_ascii_hexdigit()
        });
    if !valid {
        return Err(invalid(format!(
            "invalid custom property format ID '{format_id}'"
        )));
    }
    let normalized = format_id.to_ascii_uppercase();
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(normalized)
}

fn append_value_text(buffer: &mut String, text: &str) -> Result<()> {
    let actual = buffer
        .len()
        .checked_add(text.len())
        .ok_or_else(|| limit("custom-property text bytes", MAX_TEXT_BYTES, usize::MAX))?;
    if actual > MAX_TEXT_BYTES {
        return Err(limit("custom-property text bytes", MAX_TEXT_BYTES, actual));
    }
    buffer.push_str(text);
    Ok(())
}

fn decode_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| Error::Xml(format!("invalid character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(format!("invalid entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_owned()),
        "lt" => Ok("<".to_owned()),
        "gt" => Ok(">".to_owned()),
        "quot" => Ok("\"".to_owned()),
        "apos" => Ok("'".to_owned()),
        _ => Err(invalid(format!("unsupported entity reference '&{name};'"))),
    }
}

fn is_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    element.local_name().as_ref() == expected_local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected_namespace)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn display_name(name: &[u8]) -> String {
    String::from_utf8_lossy(name).into_owned()
}
