//! XML codec for `a:CT_Transform2D`.

use std::{fmt, io::Write};

use quick_xml::{
    events::{BytesStart, Event},
    reader::Reader,
};

use crate::{
    Error, Result,
    coordinate::{Coordinate, Extent},
};

use super::{Angle, Point, Size, Transform, validation};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// Maximum accepted serialized transform fragment.
pub(crate) const MAX_XML_BYTES: usize = 1 << 20;
const MAX_NODES: usize = 128;

/// Caller-selected bounds for one raw transform fragment.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[must_use]
pub struct Limits {
    pub max_xml_bytes: usize,
    pub max_nodes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_nodes: MAX_NODES,
    };

    #[must_use]
    pub const fn new(max_xml_bytes: usize, max_nodes: usize) -> Self {
        Self {
            max_xml_bytes,
            max_nodes,
        }
    }

    pub(crate) fn validate(self) -> Result<Self> {
        for (kind, value, maximum) in [
            (
                crate::LimitKind::XmlBytes,
                self.max_xml_bytes,
                MAX_XML_BYTES,
            ),
            (crate::LimitKind::XmlNodes, self.max_nodes, MAX_NODES),
        ] {
            if value == 0 || value > maximum {
                return Err(Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Read one complete `a:xfrm` element.
///
/// The reader is intentionally fragment-oriented: a host may provide a
/// prefix and namespace declaration on an ancestor, so element names are
/// matched by local name while all scalar values retain their strict schema
/// domains. Unsupported attributes and children fail rather than being lost
/// by a later edit.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8], limits: &Limits) -> Result<Transform> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(limit(
            crate::LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
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
            .ok_or_else(|| limit(crate::LimitKind::XmlNodes, nodes, limits.max_nodes))?;
        if nodes > limits.max_nodes {
            return Err(limit(crate::LimitKind::XmlNodes, nodes, limits.max_nodes));
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
                    limits.max_nodes,
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

/// Serialize a transform with the canonical `DrawingML` `a:` prefix.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write(transform: &Transform, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    validation::validate(transform)?;

    let mut length = LengthWriter::default();
    write_transform_xml(&mut length, transform).map_err(format_error)?;
    if length.overflow {
        return Err(limit(
            crate::LimitKind::XmlBytes,
            usize::MAX,
            limits.max_xml_bytes,
        ));
    }
    if length.len > limits.max_xml_bytes {
        return Err(limit(
            crate::LimitKind::XmlBytes,
            length.len,
            limits.max_xml_bytes,
        ));
    }
    let mut xml = String::with_capacity(length.len);
    write_transform_xml(&mut xml, transform).map_err(format_error)?;
    Ok(xml.into_bytes())
}

fn write_transform_xml<W: fmt::Write>(xml: &mut W, transform: &Transform) -> fmt::Result {
    xml.write_str("<a:xfrm")?;
    if let Some(value) = transform.authored_rotation() {
        write!(xml, r#" rot="{}""#, value.value())?;
    }
    if let Some(value) = transform.authored_flip_horizontal() {
        write!(xml, r#" flipH="{}""#, bool_token(value))?;
    }
    if let Some(value) = transform.authored_flip_vertical() {
        write!(xml, r#" flipV="{}""#, bool_token(value))?;
    }

    let has_children = transform.offset().is_some()
        || transform.extent().is_some()
        || transform.child_offset().is_some()
        || transform.child_extent().is_some();
    if !has_children {
        return xml.write_str("/>");
    }

    xml.write_char('>')?;
    if let Some(value) = transform.offset() {
        write_point(xml, "off", value)?;
    }
    if let Some(value) = transform.extent() {
        write_size(xml, "ext", value)?;
    }
    if let Some(value) = transform.child_offset() {
        write_point(xml, "chOff", value)?;
    }
    if let Some(value) = transform.child_extent() {
        write_size(xml, "chExt", value)?;
    }
    xml.write_str("</a:xfrm>")
}

/// Serialize a transform to a caller-owned sink.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_to<W: Write>(writer: &mut W, transform: &Transform, limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    validation::validate(transform)?;
    let mut length = LengthWriter::default();
    write_transform_xml(&mut length, transform).map_err(format_error)?;
    if length.overflow {
        return Err(limit(
            crate::LimitKind::XmlBytes,
            usize::MAX,
            limits.max_xml_bytes,
        ));
    }
    if length.len > limits.max_xml_bytes {
        return Err(limit(
            crate::LimitKind::XmlBytes,
            length.len,
            limits.max_xml_bytes,
        ));
    }
    write_transform_to(writer, transform)?;
    Ok(())
}

fn write_transform_to<W: Write>(writer: &mut W, transform: &Transform) -> Result<()> {
    let mut sink = IoFmtWriter {
        writer,
        error: None,
    };
    let result = write_transform_xml(&mut sink, transform);
    if let Some(error) = sink.error {
        return Err(Error::Io(error));
    }
    result.map_err(format_error)
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
    max_nodes: usize,
) -> Result<()> {
    let local = element.local_name();
    let value = parse_child_value(local.as_ref(), element, decoder)?;
    consume_leaf(reader, buffer, local.as_ref(), nodes, max_nodes)?;
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
    max_nodes: usize,
) -> Result<()> {
    loop {
        let event = reader
            .read_event_into(buffer)
            .map_err(xml_error)?
            .into_owned();
        *nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit(crate::LimitKind::XmlNodes, *nodes, max_nodes))?;
        if *nodes > max_nodes {
            return Err(limit(crate::LimitKind::XmlNodes, *nodes, max_nodes));
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
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if !allowed.contains(&raw) {
            return Err(invalid(format!("unsupported {description} attribute")));
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
    Coordinate::parse(&value).map_err(|error| match error {
        crate::coordinate::ParseError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind,
            observed,
            maximum,
        },
        _ => invalid(format!("invalid transform {description}")),
    })
}

fn required_extent(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<Extent> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("DrawingML transform {description} is missing")))?;
    Extent::parse(&value).map_err(|error| match error {
        crate::coordinate::ParseError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind,
            observed,
            maximum,
        },
        _ => invalid(format!("invalid transform {description}")),
    })
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
            "invalid DrawingML transform {name} boolean"
        ))),
    }
}

fn bool_token(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

fn write_point<W: fmt::Write>(xml: &mut W, name: &str, point: &Point) -> fmt::Result {
    write!(xml, r#"<a:{name} x="{}" y="{}"/>"#, point.x(), point.y())
}

fn write_size<W: fmt::Write>(xml: &mut W, name: &str, size: Size) -> fmt::Result {
    write!(
        xml,
        r#"<a:{name} cx="{}" cy="{}"/>"#,
        size.width(),
        size.height()
    )
}

struct IoFmtWriter<'a, W> {
    writer: &'a mut W,
    error: Option<std::io::Error>,
}

impl<W: Write> fmt::Write for IoFmtWriter<'_, W> {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        match self.writer.write_all(value.as_bytes()) {
            Ok(()) => Ok(()),
            Err(error) => {
                self.error = Some(error);
                Err(fmt::Error)
            },
        }
    }
}

#[derive(Default)]
struct LengthWriter {
    len: usize,
    overflow: bool,
}

impl fmt::Write for LengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        match self.len.checked_add(value.len()) {
            Some(length) => self.len = length,
            None => {
                self.len = usize::MAX;
                self.overflow = true;
            },
        }
        Ok(())
    }
}

fn format_error(error: fmt::Error) -> Error {
    Error::Xml(error.to_string())
}

fn xml_error(_error: impl fmt::Display) -> Error {
    Error::Xml("malformed DrawingML transform XML".into())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn unsupported_child(_local: &[u8]) -> Error {
    invalid("unsupported DrawingML transform child")
}

fn unqualified_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    litchi_ooxml_common::xml::unqualified_attribute_value(element, name, decoder).map_err(xml_error)
}

fn limit(kind: crate::LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

#[cfg(test)]
mod envelope_tests {
    //! What the transform reader accepts before, around and after `xfrm`, and the exact error reported otherwise. Each
    //! expected line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
            ("reference before root", format!("&amp;{d}")),
            ("reference after root", format!("{d}&amp;")),
            ("CDATA after root", format!("{d}<![CDATA[x]]>")),
            ("end tag before root", format!("</foreign>{d}")),
            (
                "declaration inside root",
                format!("{open}{DECLARATION}{body}{close}"),
            ),
            (
                "text, CDATA and reference inside root",
                format!("{open}{body}text<![CDATA[x]]>&amp;{close}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&str) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(&xml)))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The transform reader.
    #[test]
    fn transform_read_envelope() {
        let root = Root {
            open: r#"<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            body: r#"<a:off x="1" y="2"/>"#,
            close: r"</a:xfrm>",
            empty: r#"<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"DrawingML transform has no root element\")",
                "whitespace only: Invalid(\"DrawingML transform has no root element\")",
                "declaration only: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "empty-element root: Ok(None)",
                "start and end root: Ok(None)",
                "root then whitespace: Ok(Some(Point { x: Coordinate(1), y: Coordinate(2) }))",
                "root then comment: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "root then PI: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "root then text: Invalid(\"DrawingML transform contains unexpected text\")",
                "root then second root: Invalid(\"DrawingML transform contains more than one root\")",
                "root then foreign start: Invalid(\"DrawingML transform contains more than one root\")",
                "root then foreign empty: Invalid(\"DrawingML transform contains more than one root\")",
                "declaration after root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "two declarations: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "comment before declaration: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "text before root: Invalid(\"DrawingML transform contains unexpected text\")",
                "unterminated root: Invalid(\"DrawingML transform root is not closed\")",
                "stray end tag: Xml(\"malformed DrawingML transform XML\")",
                "DOCTYPE: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "top-level CDATA: Invalid(\"DrawingML transform contains unexpected text\")",
                "declaration then root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "whitespace before declaration: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "PI before declaration: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "comment then root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "PI then root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "root then text then second root: Invalid(\"DrawingML transform contains unexpected text\")",
                "empty-element root then root: Invalid(\"DrawingML transform contains more than one root\")",
                "declaration after empty-element root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "reference before root: Invalid(\"DrawingML transform contains an unsupported entity\")",
                "reference after root: Invalid(\"DrawingML transform contains an unsupported entity\")",
                "CDATA after root: Invalid(\"DrawingML transform contains unexpected text\")",
                "end tag before root: Xml(\"malformed DrawingML transform XML\")",
                "declaration inside root: Invalid(\"DrawingML transform contains unsupported document markup\")",
                "text, CDATA and reference inside root: Invalid(\"DrawingML transform contains unexpected text\")",
            ],
            |xml| {
                outcome(
                    super::read(xml.as_bytes(), &super::Limits::DEFAULT),
                    |transform| format!("{:?}", transform.offset()),
                )
            },
        );
    }
}
