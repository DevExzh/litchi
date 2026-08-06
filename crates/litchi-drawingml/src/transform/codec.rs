//! XML codec for `a:CT_Transform2D`.

use std::{fmt::Write as _, io::Write};

use quick_xml::{
    events::{BytesStart, Event},
    reader::Reader,
};

use litchi_ooxml_common::xml::unqualified_attribute_value;

use crate::{
    Error, Result,
    coordinate::{Coordinate, Extent},
};

use super::{Angle, Point, Size, Transform, validation};

/// Maximum accepted serialized transform fragment.
pub(crate) const MAX_XML_BYTES: usize = 1 << 20;
const MAX_NODES: usize = 128;

/// Read one complete `a:xfrm` element.
///
/// The reader is intentionally fragment-oriented: a host may provide a
/// prefix and namespace declaration on an ancestor, so element names are
/// matched by local name while all scalar values retain their strict schema
/// domains. Unsupported attributes and children fail rather than being lost
/// by a later edit.
pub fn read(xml: &[u8]) -> Result<Transform> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("DrawingML transform XML", MAX_XML_BYTES));
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut transform = Transform::new();
    let mut root_seen = false;
    let mut root_open = false;
    let mut root_closed = false;
    let mut last_child_order = 0u8;
    let mut nodes = 0usize;

    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("DrawingML transform nodes", MAX_NODES))?;
        if nodes > MAX_NODES {
            return Err(limit("DrawingML transform nodes", MAX_NODES));
        }

        match event {
            Event::Start(element) if !root_seen => {
                require_root(&element)?;
                parse_root_attributes(&element, reader.decoder(), &mut transform)?;
                root_seen = true;
                root_open = true;
            },
            Event::Empty(element) if !root_seen => {
                require_root(&element)?;
                parse_root_attributes(&element, reader.decoder(), &mut transform)?;
                root_seen = true;
                root_closed = true;
            },
            Event::Start(element) if root_open => {
                let order = child_order(element.local_name().as_ref())
                    .ok_or_else(|| unsupported_child(element.local_name().as_ref()))?;
                if order <= last_child_order {
                    return Err(invalid(
                        "DrawingML transform children are duplicated or out of order",
                    ));
                }
                last_child_order = order;
                parse_child_start(
                    &element,
                    reader.decoder(),
                    &mut transform,
                    &mut reader,
                    &mut buffer,
                    &mut nodes,
                )?;
            },
            Event::Empty(element) if root_open => {
                let order = child_order(element.local_name().as_ref())
                    .ok_or_else(|| unsupported_child(element.local_name().as_ref()))?;
                if order <= last_child_order {
                    return Err(invalid(
                        "DrawingML transform children are duplicated or out of order",
                    ));
                }
                last_child_order = order;
                parse_child_empty(&element, reader.decoder(), &mut transform)?;
            },
            Event::End(element) if root_open => {
                if element.local_name().as_ref() != b"xfrm" {
                    return Err(invalid("DrawingML transform root close tag does not match"));
                }
                root_open = false;
                root_closed = true;
            },
            Event::Start(_) | Event::Empty(_) if root_closed => {
                return Err(invalid("DrawingML transform contains more than one root"));
            },
            Event::Start(_) | Event::Empty(_) => {
                return Err(invalid(
                    "DrawingML transform contains an unexpected element",
                ));
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(invalid("DrawingML transform contains unexpected text"));
            },
            Event::CData(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(invalid("DrawingML transform contains unexpected text"));
            },
            Event::Comment(_) | Event::PI(_) | Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "DrawingML transform contains unsupported document markup",
                ));
            },
            Event::End(_) => {
                return Err(invalid(
                    "DrawingML transform contains an unexpected close tag",
                ));
            },
            Event::Text(_) | Event::CData(_) => {},
            Event::GeneralRef(_) => {
                return Err(invalid(
                    "DrawingML transform contains an unsupported entity",
                ));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }

    if !root_seen {
        return Err(invalid("DrawingML transform has no root element"));
    }
    if root_open || !root_closed {
        return Err(invalid("DrawingML transform root is not closed"));
    }
    validation::validate(&transform)?;
    Ok(transform)
}

/// Serialize a transform with the canonical DrawingML `a:` prefix.
pub fn write(transform: &Transform) -> Result<Vec<u8>> {
    validation::validate(transform)?;

    let mut xml = String::with_capacity(192);
    xml.push_str("<a:xfrm");
    if let Some(value) = transform.authored_rotation() {
        write!(xml, r#" rot="{}""#, value.value()).map_err(format_error)?;
    }
    if let Some(value) = transform.authored_flip_horizontal() {
        write!(xml, r#" flipH="{}""#, bool_token(value)).map_err(format_error)?;
    }
    if let Some(value) = transform.authored_flip_vertical() {
        write!(xml, r#" flipV="{}""#, bool_token(value)).map_err(format_error)?;
    }

    let has_children = transform.offset().is_some()
        || transform.extent().is_some()
        || transform.child_offset().is_some()
        || transform.child_extent().is_some();
    if !has_children {
        xml.push_str("/>");
        return checked_output(xml.into_bytes());
    }

    xml.push('>');
    if let Some(value) = transform.offset() {
        write_point(&mut xml, "off", value).map_err(format_error)?;
    }
    if let Some(value) = transform.extent() {
        write_size(&mut xml, "ext", value).map_err(format_error)?;
    }
    if let Some(value) = transform.child_offset() {
        write_point(&mut xml, "chOff", value).map_err(format_error)?;
    }
    if let Some(value) = transform.child_extent() {
        write_size(&mut xml, "chExt", value).map_err(format_error)?;
    }
    xml.push_str("</a:xfrm>");
    checked_output(xml.into_bytes())
}

/// Serialize a transform to a caller-owned sink.
pub fn write_to<W: Write>(writer: &mut W, transform: &Transform) -> Result<()> {
    writer.write_all(&write(transform)?)?;
    Ok(())
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    transform: &mut Transform,
) -> Result<()> {
    validate_attributes(element, &[b"rot", b"flipH", b"flipV"], "transform root")?;
    if let Some(value) = unqualified_attribute_value(element, b"rot", decoder)? {
        transform.set_rotation(Some(Angle::parse(&value)?));
    }
    if let Some(value) = unqualified_attribute_value(element, b"flipH", decoder)? {
        transform.set_flip_horizontal(Some(parse_bool(&value, "flipH")?));
    }
    if let Some(value) = unqualified_attribute_value(element, b"flipV", decoder)? {
        transform.set_flip_vertical(Some(parse_bool(&value, "flipV")?));
    }
    Ok(())
}

fn parse_child_start(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    transform: &mut Transform,
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    nodes: &mut usize,
) -> Result<()> {
    let local = element.local_name();
    let value = parse_child_value(local.as_ref(), element, decoder)?;
    consume_leaf(reader, buffer, local.as_ref(), nodes)?;
    set_child(transform, local.as_ref(), value)
}

fn parse_child_empty(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    transform: &mut Transform,
) -> Result<()> {
    let local = element.local_name();
    let value = parse_child_value(local.as_ref(), element, decoder)?;
    set_child(transform, local.as_ref(), value)
}

fn parse_child_value(
    local: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<ChildValue> {
    match local {
        b"off" | b"chOff" => {
            validate_attributes(element, &[b"x", b"y"], "DrawingML transform point")?;
            let x = required_coordinate(element, b"x", decoder, "x")?;
            let y = required_coordinate(element, b"y", decoder, "y")?;
            Ok(ChildValue::Point(Point::new(x, y)))
        },
        b"ext" | b"chExt" => {
            validate_attributes(element, &[b"cx", b"cy"], "DrawingML transform size")?;
            let width = required_extent(element, b"cx", decoder, "cx")?;
            let height = required_extent(element, b"cy", decoder, "cy")?;
            Ok(ChildValue::Size(Size::new(width, height)))
        },
        _ => Err(unsupported_child(local)),
    }
}

fn set_child(transform: &mut Transform, local: &[u8], value: ChildValue) -> Result<()> {
    match (local, value) {
        (b"off", ChildValue::Point(value)) => {
            transform.set_offset(Some(value));
        },
        (b"ext", ChildValue::Size(value)) => {
            transform.set_extent(Some(value));
        },
        (b"chOff", ChildValue::Point(value)) => {
            transform.set_child_offset(Some(value));
        },
        (b"chExt", ChildValue::Size(value)) => {
            transform.set_child_extent(Some(value));
        },
        _ => {
            return Err(invalid(
                "DrawingML transform child value has the wrong type",
            ));
        },
    }
    Ok(())
}

enum ChildValue {
    Point(Point),
    Size(Size),
}

fn consume_leaf(
    reader: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    expected: &[u8],
    nodes: &mut usize,
) -> Result<()> {
    loop {
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        *nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("DrawingML transform nodes", MAX_NODES))?;
        if *nodes > MAX_NODES {
            return Err(limit("DrawingML transform nodes", MAX_NODES));
        }
        match event {
            Event::End(element) if element.local_name().as_ref() == expected => return Ok(()),
            Event::End(_) => {
                return Err(invalid(
                    "DrawingML transform child close tag does not match",
                ));
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::CData(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Start(_) | Event::Empty(_) => {
                return Err(invalid("DrawingML transform point or size is not a leaf"));
            },
            Event::Comment(_) | Event::PI(_) | Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "DrawingML transform child contains unsupported document markup",
                ));
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) | Event::Eof => {
                return Err(invalid("DrawingML transform child is unterminated"));
            },
        }
        buffer.clear();
    }
}

fn require_root(element: &BytesStart<'_>) -> Result<()> {
    if element.local_name().as_ref() != b"xfrm" {
        return Err(invalid("DrawingML transform root must be xfrm"));
    }
    Ok(())
}

fn validate_attributes(
    element: &BytesStart<'_>,
    allowed: &[&[u8]],
    description: &str,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if !allowed.iter().any(|name| *name == raw) {
            return Err(invalid(format!(
                "unsupported {description} attribute '{}'",
                String::from_utf8_lossy(raw)
            )));
        }
    }
    Ok(())
}

fn required_coordinate(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<Coordinate> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("DrawingML transform {description} is missing")))?;
    Coordinate::parse(&value)
        .map_err(|error| invalid(format!("invalid transform {description}: {error}")))
}

fn required_extent(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<Extent> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("DrawingML transform {description} is missing")))?;
    Extent::parse(&value)
        .map_err(|error| invalid(format!("invalid transform {description}: {error}")))
}

fn child_order(local: &[u8]) -> Option<u8> {
    match local {
        b"off" => Some(1),
        b"ext" => Some(2),
        b"chOff" => Some(3),
        b"chExt" => Some(4),
        _ => None,
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value.trim() {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid DrawingML transform {name} boolean '{value}'"
        ))),
    }
}

fn bool_token(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

fn write_point(xml: &mut String, name: &str, point: &Point) -> std::fmt::Result {
    write!(xml, r#"<a:{name} x="{}" y="{}"/>"#, point.x(), point.y())
}

fn write_size(xml: &mut String, name: &str, size: Size) -> std::fmt::Result {
    write!(
        xml,
        r#"<a:{name} cx="{}" cy="{}"/>"#,
        size.width(),
        size.height()
    )
}

fn checked_output(output: Vec<u8>) -> Result<Vec<u8>> {
    if output.len() > MAX_XML_BYTES {
        return Err(limit("DrawingML transform output", MAX_XML_BYTES));
    }
    Ok(output)
}

fn format_error(error: std::fmt::Error) -> Error {
    Error::Xml(error.to_string())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn unsupported_child(local: &[u8]) -> Error {
    invalid(format!(
        "unsupported DrawingML transform child '{}'",
        String::from_utf8_lossy(local)
    ))
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
