//! Typed, inert ODF drawing marker resources.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_MARKERS: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_PATH_BYTES: usize = 4 * 1_048_576;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// The four integer components of an SVG marker view box.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ViewBox {
    pub min_x: i64,
    pub min_y: i64,
    pub width: i64,
    pub height: i64,
}

impl ViewBox {
    pub const fn new(min_x: i64, min_y: i64, width: i64, height: i64) -> Self {
        Self {
            min_x,
            min_y,
            width,
            height,
        }
    }
}

impl FromStr for ViewBox {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let mut values = value.split_ascii_whitespace();
        let mut next = || {
            let component = values
                .next()
                .ok_or_else(|| make_error("svg:viewBox must contain four integers"))?;
            if component.contains('.') || component.contains('e') || component.contains('E') {
                return invalid(format!("invalid svg:viewBox integer '{component}'"));
            }
            component
                .parse::<i64>()
                .map_err(|_| make_error(format!("invalid svg:viewBox integer '{component}'")))
        };
        let result = Self::new(next()?, next()?, next()?, next()?);
        if values.next().is_some() {
            return invalid("svg:viewBox must contain exactly four integers");
        }
        Ok(result)
    }
}

impl fmt::Display for ViewBox {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} {} {} {}",
            self.min_x, self.min_y, self.width, self.height
        )
    }
}

/// Bounded, inert ODF marker path data.
///
/// ODF 1.2 intentionally declares `pathData` as an unrestricted string, so the
/// lexical value is retained rather than narrowed to one SVG command dialect.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct PathData(String);

impl PathData {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, "svg:d", true, MAX_PATH_BYTES)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for PathData {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl FromStr for PathData {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// One named `draw:marker` resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    pub name: String,
    pub display_name: Option<String>,
    pub view_box: ViewBox,
    pub path_data: PathData,
}

impl Definition {
    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "draw:name", false, MAX_VALUE_BYTES)?;
        if let Some(display_name) = &self.display_name {
            validate_text(display_name, "draw:display-name", true, MAX_VALUE_BYTES)?;
        }
        validate_text(self.path_data.as_str(), "svg:d", true, MAX_PATH_BYTES)
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192 + self.path_data.as_str().len());
        write_marker(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered marker resources from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Collection {
    pub markers: Vec<Definition>,
}

impl Collection {
    pub fn get(&self, name: &str) -> Option<&Definition> {
        self.markers.iter().find(|marker| marker.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.markers.len() > MAX_MARKERS {
            return invalid(format!("drawing styles exceed {MAX_MARKERS} markers"));
        }
        let mut names = HashSet::with_capacity(self.markers.len());
        let mut aggregate = 0usize;
        for marker in &self.markers {
            marker.validate()?;
            if !names.insert(marker.name.as_str()) {
                return invalid(format!("duplicate drawing marker name '{}'", marker.name));
            }
            aggregate = aggregate
                .checked_add(marker.name.len())
                .and_then(|size| {
                    size.checked_add(marker.display_name.as_ref().map_or(0, String::len))
                })
                .and_then(|size| size.checked_add(marker.path_data.as_str().len()))
                .ok_or_else(|| make_error("drawing marker size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("drawing marker values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize a standalone schema-positioned `office:styles` fragment.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let capacity = self.markers.iter().fold(192usize, |size, marker| {
            size.saturating_add(192 + marker.path_data.as_str().len())
        });
        let mut output = String::with_capacity(capacity);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">"#,
        );
        for marker in &self.markers {
            write_marker(&mut output, marker, false);
        }
        output.push_str("</office:styles>");
        Ok(output)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Draw,
    Svg,
    Other,
}

#[derive(Clone)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}

struct ActiveDefinition {
    parent_depth: usize,
    value: Definition,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parse marker resources from an ODF styles or flat-document XML part.
pub fn parse_drawing_markers(xml: &str) -> Result<Collection> {
    if !xml.contains("marker") {
        return Ok(Collection::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing marker XML exceeds 64 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveDefinition> = None;
    let mut result = Collection::default();
    let mut aggregate = 0usize;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing marker XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    return invalid("draw:marker cannot contain child elements");
                }
                if namespace == NamespaceKind::Draw && local == "marker" {
                    ensure_location(&stack)?;
                    ensure_count(result.markers.len())?;
                    active = Some(ActiveDefinition {
                        parent_depth: stack.len(),
                        value: parse_marker(&reader, element, &mut aggregate)?,
                    });
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("drawing marker XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    return invalid("draw:marker cannot contain child elements");
                }
                if namespace == NamespaceKind::Draw && local == "marker" {
                    ensure_location(&stack)?;
                    ensure_count(result.markers.len())?;
                    result
                        .markers
                        .push(parse_marker(&reader, element, &mut aggregate)?);
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("drawing marker XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|marker| marker.parent_depth == stack.len())
                {
                    if frame.namespace != NamespaceKind::Draw || frame.local != "marker" {
                        return invalid("unexpected drawing marker end element");
                    }
                    result
                        .markers
                        .push(active.take().expect("active marker checked").value);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid marker text: {error}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("draw:marker must be empty");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("draw:marker cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in marker XML");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated drawing marker XML");
    }
    result.validate()?;
    Ok(result)
}

fn ensure_count(count: usize) -> Result<()> {
    if count >= MAX_MARKERS {
        invalid(format!("drawing styles exceed {MAX_MARKERS} markers"))
    } else {
        Ok(())
    }
}

fn parse_marker(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Definition> {
    let mut values = attributes(reader, element, aggregate)?;
    let name = required(&mut values, NamespaceKind::Draw, "name", "draw:name")?;
    let display_name = take(&mut values, NamespaceKind::Draw, "display-name");
    let view_box = required(&mut values, NamespaceKind::Svg, "viewBox", "svg:viewBox")?.parse()?;
    let path_data = PathData::new(required(&mut values, NamespaceKind::Svg, "d", "svg:d")?)?;
    reject_attributes(&values)?;
    let marker = Definition {
        name,
        display_name,
        view_box,
        path_data,
    };
    marker.validate()?;
    Ok(marker)
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| make_error(format!("invalid marker attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid marker attribute: {error}")))?
            .into_owned();
        let limit = if namespace == NamespaceKind::Svg && local == "d" {
            MAX_PATH_BYTES
        } else {
            MAX_VALUE_BYTES
        };
        if value.len() > limit {
            return invalid(format!("marker attribute '{local}' exceeds its size limit"));
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("marker aggregate size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("marker values exceed 16 MiB");
        }
        if values.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded marker attribute");
        }
    }
    Ok(values)
}

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if stack
        .last()
        .is_some_and(|frame| frame.namespace == NamespaceKind::Office && frame.local == "styles")
    {
        Ok(())
    } else {
        invalid("draw:marker must be a direct child of office:styles")
    }
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) => {
            let bytes: &[u8] = namespace.as_ref();
            Ok(if bytes == OFFICE_NS {
                NamespaceKind::Office
            } else if bytes == DRAW_NS {
                NamespaceKind::Draw
            } else if bytes == SVG_NS {
                NamespaceKind::Svg
            } else {
                NamespaceKind::Other
            })
        },
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "marker" && namespace != NamespaceKind::Draw {
        return invalid("marker element uses the wrong namespace");
    }
    Ok(())
}

fn take(values: &mut Attributes, namespace: NamespaceKind, local: &str) -> Option<String> {
    values.remove(&(namespace, local.to_owned()))
}

fn required(
    values: &mut Attributes,
    namespace: NamespaceKind,
    local: &str,
    qualified: &str,
) -> Result<String> {
    take(values, namespace, local)
        .ok_or_else(|| make_error(format!("missing required {qualified} attribute")))
}

fn reject_attributes(values: &Attributes) -> Result<()> {
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported marker attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn validate_text(value: &str, name: &str, allow_empty: bool, limit: usize) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > limit {
        return invalid(format!("invalid {name} length"));
    }
    if value
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return invalid(format!("{name} contains prohibited control characters"));
    }
    Ok(())
}

fn write_marker(output: &mut String, marker: &Definition, standalone: bool) {
    output.push_str("<draw:marker");
    if standalone {
        output.push_str(
            r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0""#,
        );
    }
    write_attribute(output, "draw:name", &marker.name);
    if let Some(display_name) = &marker.display_name {
        write_attribute(output, "draw:display-name", display_name);
    }
    write_attribute(output, "svg:viewBox", &marker.view_box.to_string());
    write_attribute(output, "svg:d", marker.path_data.as_str());
    output.push_str("/>");
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_xml(output, value);
    output.push('"');
}

fn escape_xml(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn decode(value: &[u8], what: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| make_error(format!("invalid UTF-8 in marker {what}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
