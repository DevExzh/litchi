//! Typed ODF drawing opacity-gradient resources.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::fmt;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const LOEXT_NS: &[u8] = b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_OPACITIES: usize = 65_536;
const MAX_STOPS: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// One of the six ODF opacity-gradient geometries.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Style {
    Linear,
    Axial,
    Radial,
    Ellipsoid,
    Square,
    Rectangular,
}

impl Style {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "linear" => Ok(Self::Linear),
            "axial" => Ok(Self::Axial),
            "radial" => Ok(Self::Radial),
            "ellipsoid" => Ok(Self::Ellipsoid),
            "square" => Ok(Self::Square),
            "rectangular" => Ok(Self::Rectangular),
            _ => invalid(format!("unsupported draw:opacity style '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Linear => "linear",
            Self::Axial => "axial",
            Self::Radial => "radial",
            Self::Ellipsoid => "ellipsoid",
            Self::Square => "square",
            Self::Rectangular => "rectangular",
        }
    }
}

/// A finite percentage constrained to the ODF `0%..=100%` datatype.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Percent(f64);

impl Percent {
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() || !(0.0..=100.0).contains(&value) {
            return invalid("opacity percentage must be between 0% and 100%");
        }
        Ok(Self(value))
    }

    pub const fn value(self) -> f64 {
        self.0
    }
}

impl fmt::Display for Percent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}%", canonical_number(self.0))
    }
}

/// A finite signed percentage used by ODF gradient geometry.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct GeometryPercent(f64);

impl GeometryPercent {
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return invalid("opacity geometry percentage must be finite");
        }
        Ok(Self(value))
    }

    pub const fn value(self) -> f64 {
        self.0
    }
}

impl fmt::Display for GeometryPercent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}%", canonical_number(self.0))
    }
}

/// A validated lexical ODF angle, retained because ODF leaves its unit grammar open.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Angle(String);

impl Angle {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, "draw:angle", false)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A LibreOffice extension stop coordinate or opacity.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum StopValue {
    Fraction(f64),
    Percent(Percent),
}

impl StopValue {
    pub fn fraction(value: f64) -> Result<Self> {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return invalid("opacity stop fraction must be between 0 and 1");
        }
        Ok(Self::Fraction(value))
    }

    fn parse(value: &str, name: &str) -> Result<Self> {
        if let Some(number) = value.strip_suffix('%') {
            return Ok(Self::Percent(parse_bounded_percent(number, name)?));
        }
        let value = parse_decimal(value, false, name)?;
        Self::fraction(value)
    }
}

impl fmt::Display for StopValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Fraction(value) => formatter.write_str(&canonical_number(*value)),
            Self::Percent(value) => value.fmt(formatter),
        }
    }
}

/// One ordered LibreOffice opacity stop.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Stop {
    pub offset: StopValue,
    pub opacity: StopValue,
}

/// One `draw:opacity` resource.
#[derive(Clone, Debug, PartialEq)]
pub struct Definition {
    pub name: Option<String>,
    pub display_name: Option<String>,
    pub style: Style,
    pub center_x: Option<GeometryPercent>,
    pub center_y: Option<GeometryPercent>,
    pub start: Option<Percent>,
    pub end: Option<Percent>,
    pub angle: Option<Angle>,
    pub border: Option<GeometryPercent>,
    pub extension_stops: Vec<Stop>,
}

impl Definition {
    pub fn validate(&self) -> Result<()> {
        if let Some(name) = &self.name {
            validate_text(name, "draw:name", false)?;
        }
        if let Some(display_name) = &self.display_name {
            validate_text(display_name, "draw:display-name", true)?;
        }
        if let Some(value) = self.center_x {
            GeometryPercent::new(value.0)?;
        }
        if let Some(value) = self.center_y {
            GeometryPercent::new(value.0)?;
        }
        if let Some(value) = self.start {
            Percent::new(value.0)?;
        }
        if let Some(value) = self.end {
            Percent::new(value.0)?;
        }
        if let Some(angle) = &self.angle {
            validate_text(angle.as_str(), "draw:angle", false)?;
        }
        if let Some(value) = self.border {
            GeometryPercent::new(value.0)?;
        }
        if self.extension_stops.len() > MAX_STOPS {
            return invalid(format!("opacity gradient exceeds {MAX_STOPS} stops"));
        }
        for stop in &self.extension_stops {
            validate_stop_value(stop.offset)?;
            validate_stop_value(stop.opacity)?;
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + self.extension_stops.len() * 96);
        write_opacity(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered opacity resources from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Collection {
    pub opacities: Vec<Definition>,
}

impl Collection {
    pub fn get(&self, name: &str) -> Option<&Definition> {
        self.opacities
            .iter()
            .find(|opacity| opacity.name.as_deref() == Some(name))
    }

    pub fn validate(&self) -> Result<()> {
        if self.opacities.len() > MAX_OPACITIES {
            return invalid(format!("drawing styles exceed {MAX_OPACITIES} opacities"));
        }
        let mut names = HashSet::with_capacity(self.opacities.len());
        let mut aggregate = 0usize;
        for opacity in &self.opacities {
            opacity.validate()?;
            if let Some(name) = &opacity.name {
                if !names.insert(name.as_str()) {
                    return invalid(format!("duplicate drawing opacity name '{name}'"));
                }
                aggregate = aggregate
                    .checked_add(name.len())
                    .ok_or_else(|| make_error("opacity aggregate size overflow"))?;
            }
            aggregate = aggregate
                .checked_add(opacity.display_name.as_ref().map_or(0, String::len))
                .and_then(|size| {
                    size.checked_add(
                        opacity
                            .angle
                            .as_ref()
                            .map_or(0, |angle| angle.as_str().len()),
                    )
                })
                .ok_or_else(|| make_error("opacity aggregate size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("opacity resources exceed 16 MiB");
            }
        }
        Ok(())
    }

    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + self.opacities.len() * 256);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">"#,
        );
        for opacity in &self.opacities {
            write_opacity(&mut output, opacity, false);
        }
        output.push_str("</office:styles>");
        Ok(output)
    }
}

impl crate::OpenDocumentPackage {
    pub fn drawing_opacities(&self) -> Result<Collection> {
        let styles = self.styles_xml()?;
        parse_drawing_opacities(styles.as_deref().unwrap_or_default())
    }
}

impl crate::FlatOpenDocument {
    pub fn drawing_opacities(&self) -> Result<Collection> {
        parse_drawing_opacities(self.xml())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Draw,
    Svg,
    Loext,
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

pub fn parse_drawing_opacities(xml: &str) -> Result<Collection> {
    if !xml.contains("opacity") {
        return Ok(Collection::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing opacity XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveDefinition> = None;
    let mut result = Collection::default();
    let mut aggregate = 0usize;
    let mut stop_count = 0usize;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing opacity XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    if namespace == NamespaceKind::Loext && local == "opacity-stop" {
                        return invalid("loext:opacity-stop must be empty");
                    }
                    return invalid("draw:opacity contains an unsupported child element");
                }
                if namespace == NamespaceKind::Draw && local == "opacity" {
                    ensure_location(&stack)?;
                    ensure_count(result.opacities.len(), MAX_OPACITIES, "opacity resources")?;
                    active = Some(ActiveDefinition {
                        parent_depth: stack.len(),
                        value: parse_opacity(&reader, element, &mut aggregate)?,
                    });
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("opacity XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if let Some(opacity) = active.as_mut() {
                    if stack.len() != opacity.parent_depth + 1
                        || namespace != NamespaceKind::Loext
                        || local != "opacity-stop"
                    {
                        return invalid("draw:opacity contains an unsupported child element");
                    }
                    ensure_count(stop_count, MAX_STOPS, "opacity stops")?;
                    opacity.value.extension_stops.push(parse_stop(
                        &reader,
                        element,
                        &mut aggregate,
                    )?);
                    stop_count += 1;
                } else if namespace == NamespaceKind::Draw && local == "opacity" {
                    ensure_location(&stack)?;
                    ensure_count(result.opacities.len(), MAX_OPACITIES, "opacity resources")?;
                    result
                        .opacities
                        .push(parse_opacity(&reader, element, &mut aggregate)?);
                } else if namespace == NamespaceKind::Loext && local == "opacity-stop" {
                    return invalid("loext:opacity-stop must be inside draw:opacity");
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("opacity XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|opacity| opacity.parent_depth == stack.len())
                {
                    if frame.namespace != NamespaceKind::Draw || frame.local != "opacity" {
                        return invalid("unexpected draw:opacity end element");
                    }
                    result
                        .opacities
                        .push(active.take().expect("active opacity checked").value);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid opacity text: {error}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("draw:opacity cannot contain text");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("draw:opacity cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in opacities");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated drawing opacity XML");
    }
    result.validate()?;
    Ok(result)
}

fn parse_opacity(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Definition> {
    let mut values = attributes(reader, element, aggregate)?;
    let name = take(&mut values, NamespaceKind::Draw, "name");
    let display_name = take(&mut values, NamespaceKind::Draw, "display-name");
    let style = Style::parse(&required(
        &mut values,
        NamespaceKind::Draw,
        "style",
        "draw:style",
    )?)?;
    let center_x = take_geometry_percent(&mut values, "cx")?;
    let center_y = take_geometry_percent(&mut values, "cy")?;
    let start = take_bounded_percent(&mut values, "start")?;
    let end = take_bounded_percent(&mut values, "end")?;
    let angle = take(&mut values, NamespaceKind::Draw, "angle")
        .map(Angle::new)
        .transpose()?;
    let border = take_geometry_percent(&mut values, "border")?;
    reject_attributes(&values, "draw:opacity")?;
    Ok(Definition {
        name,
        display_name,
        style,
        center_x,
        center_y,
        start,
        end,
        angle,
        border,
        extension_stops: Vec::new(),
    })
}

fn parse_stop(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Stop> {
    let mut values = attributes(reader, element, aggregate)?;
    let offset = StopValue::parse(
        &required(&mut values, NamespaceKind::Svg, "offset", "svg:offset")?,
        "svg:offset",
    )?;
    let opacity = StopValue::parse(
        &required(
            &mut values,
            NamespaceKind::Svg,
            "stop-opacity",
            "svg:stop-opacity",
        )?,
        "svg:stop-opacity",
    )?;
    reject_attributes(&values, "loext:opacity-stop")?;
    Ok(Stop { offset, opacity })
}

fn take_bounded_percent(values: &mut Attributes, local: &str) -> Result<Option<Percent>> {
    take(values, NamespaceKind::Draw, local)
        .map(|value| {
            let number = value
                .strip_suffix('%')
                .ok_or_else(|| make_error(format!("draw:{local} must be a percentage")))?;
            parse_bounded_percent(number, &format!("draw:{local}"))
        })
        .transpose()
}

fn take_geometry_percent(values: &mut Attributes, local: &str) -> Result<Option<GeometryPercent>> {
    take(values, NamespaceKind::Draw, local)
        .map(|value| {
            let number = value
                .strip_suffix('%')
                .ok_or_else(|| make_error(format!("draw:{local} must be a percentage")))?;
            GeometryPercent::new(parse_decimal(number, true, &format!("draw:{local}"))?)
        })
        .transpose()
}

fn parse_bounded_percent(number: &str, name: &str) -> Result<Percent> {
    Percent::new(parse_decimal(number, false, name)?)
}

fn parse_decimal(value: &str, signed: bool, name: &str) -> Result<f64> {
    if value.is_empty()
        || value.starts_with('+')
        || (!signed && value.starts_with('-'))
        || value.chars().any(char::is_whitespace)
        || value.contains(['e', 'E'])
    {
        return invalid(format!("invalid {name} value '{value}'"));
    }
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let (integer, fraction) = unsigned.split_once('.').unwrap_or((unsigned, ""));
    if (integer.is_empty() && fraction.is_empty())
        || (!integer.is_empty() && !integer.bytes().all(|byte| byte.is_ascii_digit()))
        || (!fraction.is_empty() && !fraction.bytes().all(|byte| byte.is_ascii_digit()))
        || unsigned.matches('.').count() > 1
    {
        return invalid(format!("invalid {name} value '{value}'"));
    }
    value
        .parse::<f64>()
        .map_err(|_| make_error(format!("invalid {name} value '{value}'")))
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| make_error(format!("invalid opacity attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid opacity attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("opacity attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("opacity aggregate size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("opacity values exceed 16 MiB");
        }
        if values.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded opacity attribute");
        }
    }
    Ok(values)
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
            } else if bytes == LOEXT_NS {
                NamespaceKind::Loext
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

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if stack
        .last()
        .is_some_and(|frame| frame.namespace == NamespaceKind::Office && frame.local == "styles")
    {
        Ok(())
    } else {
        invalid("draw:opacity must be a direct child of office:styles")
    }
}

fn ensure_count(count: usize, maximum: usize, name: &str) -> Result<()> {
    if count >= maximum {
        invalid(format!("{name} exceed {maximum}"))
    } else {
        Ok(())
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "opacity" && namespace != NamespaceKind::Draw {
        return invalid("opacity element uses the wrong namespace");
    }
    if local == "opacity-stop" && namespace != NamespaceKind::Loext {
        return invalid("opacity-stop element uses the wrong namespace");
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

fn reject_attributes(values: &Attributes, element: &str) -> Result<()> {
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported {element} attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn validate_stop_value(value: StopValue) -> Result<()> {
    match value {
        StopValue::Fraction(value) => StopValue::fraction(value).map(|_| ()),
        StopValue::Percent(value) => Percent::new(value.0).map(|_| ()),
    }
}

fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > MAX_VALUE_BYTES {
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

fn write_opacity(output: &mut String, opacity: &Definition, standalone: bool) {
    output.push_str("<draw:opacity");
    if standalone {
        output.push_str(
            r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#,
        );
    }
    if let Some(name) = &opacity.name {
        write_attribute(output, "draw:name", name);
    }
    if let Some(display_name) = &opacity.display_name {
        write_attribute(output, "draw:display-name", display_name);
    }
    write_attribute(output, "draw:style", opacity.style.as_str());
    if let Some(value) = opacity.center_x {
        write_attribute(output, "draw:cx", &value.to_string());
    }
    if let Some(value) = opacity.center_y {
        write_attribute(output, "draw:cy", &value.to_string());
    }
    if let Some(value) = opacity.start {
        write_attribute(output, "draw:start", &value.to_string());
    }
    if let Some(value) = opacity.end {
        write_attribute(output, "draw:end", &value.to_string());
    }
    if let Some(angle) = &opacity.angle {
        write_attribute(output, "draw:angle", angle.as_str());
    }
    if let Some(value) = opacity.border {
        write_attribute(output, "draw:border", &value.to_string());
    }
    if opacity.extension_stops.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        for stop in &opacity.extension_stops {
            output.push_str("<loext:opacity-stop");
            write_attribute(output, "svg:offset", &stop.offset.to_string());
            write_attribute(output, "svg:stop-opacity", &stop.opacity.to_string());
            output.push_str("/>");
        }
        output.push_str("</draw:opacity>");
    }
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

fn canonical_number(value: f64) -> String {
    if value == 0.0 {
        "0".to_owned()
    } else {
        value.to_string()
    }
}

fn decode(value: &[u8], what: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| make_error(format!("invalid UTF-8 in opacity {what}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

    fn wrap(body: &str) -> String {
        format!(
            r#"<office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:svg="{SVG}" xmlns:loext="{LOEXT}">{body}</office:styles>"#
        )
    }

    #[test]
    fn parses_and_round_trips_all_geometries_and_stops() {
        let styles = [
            "linear",
            "axial",
            "radial",
            "ellipsoid",
            "square",
            "rectangular",
        ];
        let body = styles.iter().enumerate().map(|(index, style)| format!(r#"<draw:opacity draw:name="o{index}" draw:style="{style}" draw:cx="-5.5%" draw:cy=".5%" draw:start="0%" draw:end="100.0%" draw:angle="1rad" draw:border="25.%"><loext:opacity-stop svg:offset="0" svg:stop-opacity="20%"/><loext:opacity-stop svg:offset="1" svg:stop-opacity=".5"/></draw:opacity>"#)).collect::<String>();
        let parsed = parse_drawing_opacities(&wrap(&body)).unwrap();
        assert_eq!(parsed.opacities.len(), 6);
        assert_eq!(parsed.get("o3").unwrap().style, Style::Ellipsoid);
        assert_eq!(parsed.opacities[0].extension_stops.len(), 2);
        let serialized = parsed.to_xml().unwrap();
        assert_eq!(parse_drawing_opacities(&serialized).unwrap(), parsed);
    }

    #[test]
    fn rejects_malformed_or_misplaced_opacities() {
        for xml in [
            wrap(r#"<draw:opacity draw:name="x"/>"#),
            wrap(r#"<draw:opacity draw:name="x" draw:style="cone"/>"#),
            wrap(r#"<draw:opacity draw:name="x" draw:style="linear" draw:start="101%"/>"#),
            wrap(
                r#"<draw:opacity draw:name="x" draw:style="linear"><loext:opacity-stop svg:offset="2" svg:stop-opacity="1"/></draw:opacity>"#,
            ),
            wrap(
                r#"<draw:opacity draw:name="x" draw:style="linear"><loext:opacity-stop svg:offset="0"/></draw:opacity>"#,
            ),
            wrap(r#"<draw:opacity draw:name="x" draw:style="linear"><draw:g/></draw:opacity>"#),
            wrap(r#"<loext:opacity-stop svg:offset="0" svg:stop-opacity="1"/>"#),
            wrap(
                r#"<draw:opacity draw:name="x" draw:style="linear"/><draw:opacity draw:name="x" draw:style="axial"/>"#,
            ),
            format!(
                r#"<office:document xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><draw:opacity draw:style="linear"/></office:document>"#
            ),
            format!(
                r#"<!DOCTYPE x><office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><draw:opacity draw:style="linear"/></office:styles>"#
            ),
        ] {
            assert!(parse_drawing_opacities(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn parses_local_angles_and_extension_stops() {
        let angles_xml = include_str!("../../../../../test-data/odf/drawing/opacity-angles.fodg");
        let angles = crate::FlatOpenDocument::from_bytes(angles_xml.as_bytes().to_vec()).unwrap();
        let values = angles.drawing_opacities().unwrap();
        assert_eq!(values.opacities.len(), 6);
        assert_eq!(
            values.opacities[0].angle.as_ref().unwrap().as_str(),
            "90deg"
        );
        assert_eq!(
            values.opacities[2].angle.as_ref().unwrap().as_str(),
            "1.0rad"
        );
        assert_eq!(
            values.opacities[3].angle.as_ref().unwrap().as_str(),
            "1000grad"
        );

        let stops_xml =
            include_str!("../../../../../test-data/odf/drawing/opacity-extension-stops.fodt");
        let stops = crate::FlatOpenDocument::from_bytes(stops_xml.as_bytes().to_vec()).unwrap();
        let values = stops.drawing_opacities().unwrap();
        let value = values.get("Transparency_20_1").unwrap();
        assert_eq!(value.style, Style::Ellipsoid);
        assert_eq!(value.extension_stops.len(), 2);
    }
}
