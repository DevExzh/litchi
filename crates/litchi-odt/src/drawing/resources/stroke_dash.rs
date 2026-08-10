//! Typed, inert ODF drawing stroke-dash resources.

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
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_DASHES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;
const MAX_DOT_COUNT: u32 = 1_000_000;

/// The cap shape used for each segment in a stroke-dash pattern.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Style {
    Rect,
    Round,
}

impl Style {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "rect" => Ok(Self::Rect),
            "round" => Ok(Self::Round),
            _ => invalid(format!("unsupported draw:stroke-dash style '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Rect => "rect",
            Self::Round => "round",
        }
    }
}

/// Unit for a stroke-dash length or percentage.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MeasureUnit {
    Centimeter,
    Millimeter,
    Inch,
    Point,
    Pica,
    Pixel,
    Percent,
}

impl MeasureUnit {
    const fn suffix(self) -> &'static str {
        match self {
            Self::Centimeter => "cm",
            Self::Millimeter => "mm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Pixel => "px",
            Self::Percent => "%",
        }
    }
}

/// A finite, nonnegative ODF length or percentage.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Measure {
    value: f64,
    unit: MeasureUnit,
}

impl Measure {
    pub fn new(value: f64, unit: MeasureUnit) -> Result<Self> {
        if !value.is_finite() || value < 0.0 {
            return invalid("stroke-dash measure must be finite and nonnegative");
        }
        Ok(Self { value, unit })
    }

    pub const fn value(self) -> f64 {
        self.value
    }

    pub const fn unit(self) -> MeasureUnit {
        self.unit
    }
}

impl FromStr for Measure {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let (number, unit) = split_measure(value)?;
        validate_decimal(number, value)?;
        let number = number
            .parse::<f64>()
            .map_err(|_error| make_error(format!("invalid stroke-dash measure '{value}'")))?;
        Self::new(number, unit)
    }
}

impl fmt::Display for Measure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{}{}",
            canonical_number(self.value),
            self.unit.suffix()
        )
    }
}

/// One named `draw:stroke-dash` resource.
#[derive(Clone, Debug, PartialEq)]
pub struct Definition {
    pub name: String,
    pub display_name: Option<String>,
    pub style: Option<Style>,
    pub dots1: Option<u32>,
    pub dots1_length: Option<Measure>,
    pub dots2: Option<u32>,
    pub dots2_length: Option<Measure>,
    pub distance: Option<Measure>,
}

impl Definition {
    /// ODF defaults an omitted cap style to `rect`.
    pub fn effective_style(&self) -> Style {
        self.style.unwrap_or(Style::Rect)
    }

    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "draw:name", false)?;
        if let Some(display_name) = &self.display_name {
            validate_text(display_name, "draw:display-name", true)?;
        }
        for (name, count) in [("draw:dots1", self.dots1), ("draw:dots2", self.dots2)] {
            if count.is_some_and(|count| count > MAX_DOT_COUNT) {
                return invalid(format!("{name} exceeds {MAX_DOT_COUNT}"));
            }
        }
        for measure in [self.dots1_length, self.dots2_length, self.distance]
            .into_iter()
            .flatten()
        {
            Measure::new(measure.value, measure.unit)?;
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192);
        write_dash(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered stroke-dash resources from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Collection {
    pub dashes: Vec<Definition>,
}

impl Collection {
    pub fn get(&self, name: &str) -> Option<&Definition> {
        self.dashes.iter().find(|dash| dash.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.dashes.len() > MAX_DASHES {
            return invalid(format!("drawing styles exceed {MAX_DASHES} stroke dashes"));
        }
        let mut names = HashSet::with_capacity(self.dashes.len());
        let mut aggregate = 0usize;
        for dash in &self.dashes {
            dash.validate()?;
            if !names.insert(dash.name.as_str()) {
                return invalid(format!(
                    "duplicate drawing stroke-dash name '{}'",
                    dash.name
                ));
            }
            aggregate = aggregate
                .checked_add(dash.name.len())
                .and_then(|size| {
                    size.checked_add(dash.display_name.as_ref().map_or(0, String::len))
                })
                .ok_or_else(|| make_error("drawing stroke-dash size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("drawing stroke-dash values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize a standalone schema-positioned `office:styles` fragment.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192 + self.dashes.len() * 192);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">"#,
        );
        for dash in &self.dashes {
            write_dash(&mut output, dash, false);
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
    Other,
}

#[derive(Clone)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}

struct ActiveDash {
    parent_depth: usize,
    value: Definition,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parse stroke-dash resources from an ODF styles or flat-document XML part.
pub fn parse_drawing_stroke_dashes(xml: &str) -> Result<Collection> {
    if !xml.contains("stroke-dash") {
        return Ok(Collection::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing stroke-dash XML exceeds 64 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveDash> = None;
    let mut result = Collection::default();
    let mut aggregate = 0usize;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing stroke-dash XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    return invalid("draw:stroke-dash cannot contain child elements");
                }
                if namespace == NamespaceKind::Draw && local == "stroke-dash" {
                    ensure_location(&stack)?;
                    if result.dashes.len() >= MAX_DASHES {
                        return invalid(format!(
                            "drawing styles exceed {MAX_DASHES} stroke dashes"
                        ));
                    }
                    active = Some(ActiveDash {
                        parent_depth: stack.len(),
                        value: parse_dash(&reader, element, &mut aggregate)?,
                    });
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!(
                        "drawing stroke-dash XML exceeds {MAX_DEPTH} levels"
                    ));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    return invalid("draw:stroke-dash cannot contain child elements");
                }
                if namespace == NamespaceKind::Draw && local == "stroke-dash" {
                    ensure_location(&stack)?;
                    if result.dashes.len() >= MAX_DASHES {
                        return invalid(format!(
                            "drawing styles exceed {MAX_DASHES} stroke dashes"
                        ));
                    }
                    result
                        .dashes
                        .push(parse_dash(&reader, element, &mut aggregate)?);
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("drawing stroke-dash XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|dash| dash.parent_depth == stack.len())
                {
                    if frame.namespace != NamespaceKind::Draw || frame.local != "stroke-dash" {
                        return invalid("unexpected drawing stroke-dash end element");
                    }
                    let dash = active
                        .take()
                        .ok_or_else(|| make_error("missing completed stroke dash"))?;
                    result.dashes.push(dash.value);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid stroke-dash text: {error}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("draw:stroke-dash must be empty");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("draw:stroke-dash cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "DTDs and processing instructions are prohibited in stroke-dash XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated drawing stroke-dash XML");
    }
    result.validate()?;
    Ok(result)
}

fn parse_dash(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Definition> {
    let mut values = attributes(reader, element, aggregate)?;
    let name = required(&mut values, NamespaceKind::Draw, "name", "draw:name")?;
    let display_name = take(&mut values, NamespaceKind::Draw, "display-name");
    let style = take(&mut values, NamespaceKind::Draw, "style")
        .map(|value| Style::parse(&value))
        .transpose()?;
    let dots1 = take_count(&mut values, "dots1")?;
    let dots1_length = take_measure(&mut values, "dots1-length")?;
    let dots2 = take_count(&mut values, "dots2")?;
    let dots2_length = take_measure(&mut values, "dots2-length")?;
    let distance = take_measure(&mut values, "distance")?;
    reject_attributes(&values)?;
    let value = Definition {
        name,
        display_name,
        style,
        dots1,
        dots1_length,
        dots2,
        dots2_length,
        distance,
    };
    value.validate()?;
    Ok(value)
}

fn take_count(values: &mut Attributes, local: &str) -> Result<Option<u32>> {
    take(values, NamespaceKind::Draw, local)
        .map(|value| {
            if value.starts_with('+') || value.chars().any(char::is_whitespace) {
                return invalid(format!("invalid draw:{local} integer '{value}'"));
            }
            let value = value
                .parse::<u32>()
                .map_err(|_error| make_error(format!("invalid draw:{local} integer '{value}'")))?;
            if value > MAX_DOT_COUNT {
                return invalid(format!("draw:{local} exceeds {MAX_DOT_COUNT}"));
            }
            Ok(value)
        })
        .transpose()
}

fn take_measure(values: &mut Attributes, local: &str) -> Result<Option<Measure>> {
    take(values, NamespaceKind::Draw, local)
        .map(|value| value.parse())
        .transpose()
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid stroke-dash attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid stroke-dash attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("stroke-dash attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("stroke-dash aggregate size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("stroke-dash values exceed 16 MiB");
        }
        if values.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded stroke-dash attribute");
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
        invalid("draw:stroke-dash must be a direct child of office:styles")
    }
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) if namespace.as_ref() == OFFICE_NS => {
            Ok(NamespaceKind::Office)
        },
        ResolveResult::Bound(namespace) if namespace.as_ref() == DRAW_NS => Ok(NamespaceKind::Draw),
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "stroke-dash" && namespace != NamespaceKind::Draw {
        return invalid("stroke-dash element uses the wrong namespace");
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
            "unsupported stroke-dash attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn split_measure(value: &str) -> Result<(&str, MeasureUnit)> {
    for (suffix, unit) in [
        ("cm", MeasureUnit::Centimeter),
        ("mm", MeasureUnit::Millimeter),
        ("in", MeasureUnit::Inch),
        ("pt", MeasureUnit::Point),
        ("pc", MeasureUnit::Pica),
        ("px", MeasureUnit::Pixel),
        ("%", MeasureUnit::Percent),
    ] {
        if let Some(number) = value.strip_suffix(suffix) {
            return Ok((number, unit));
        }
    }
    invalid(format!(
        "stroke-dash measure lacks a supported unit: '{value}'"
    ))
}

fn validate_decimal(number: &str, original: &str) -> Result<()> {
    if number.is_empty()
        || number.starts_with('+')
        || number.chars().any(char::is_whitespace)
        || number.contains('e')
        || number.contains('E')
    {
        return invalid(format!("invalid stroke-dash measure '{original}'"));
    }
    let unsigned = number.strip_prefix('-').unwrap_or(number);
    let mut parts = unsigned.split('.');
    let integer = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some()
        || integer.is_empty()
        || !integer.bytes().all(|byte| byte.is_ascii_digit())
        || fraction
            .is_some_and(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return invalid(format!("invalid stroke-dash measure '{original}'"));
    }
    Ok(())
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

fn write_dash(output: &mut String, value: &Definition, standalone: bool) {
    output.push_str("<draw:stroke-dash");
    if standalone {
        output.push_str(r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0""#);
    }
    write_attribute(output, "draw:name", &value.name);
    if let Some(display_name) = &value.display_name {
        write_attribute(output, "draw:display-name", display_name);
    }
    if let Some(style) = value.style {
        write_attribute(output, "draw:style", style.as_str());
    }
    if let Some(count) = value.dots1 {
        write_attribute(output, "draw:dots1", &count.to_string());
    }
    if let Some(length) = value.dots1_length {
        write_attribute(output, "draw:dots1-length", &length.to_string());
    }
    if let Some(count) = value.dots2 {
        write_attribute(output, "draw:dots2", &count.to_string());
    }
    if let Some(length) = value.dots2_length {
        write_attribute(output, "draw:dots2-length", &length.to_string());
    }
    if let Some(distance) = value.distance {
        write_attribute(output, "draw:distance", &distance.to_string());
    }
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
        .map_err(|_error| make_error(format!("invalid UTF-8 in stroke-dash {what}")))
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

    #[test]
    fn parses_and_round_trips_all_styles_and_measure_units() {
        let xml = format!(
            r#"<office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><draw:stroke-dash draw:name="mixed&amp;units" draw:display-name="Mixed" draw:style="round" draw:dots1="1" draw:dots1-length="2cm" draw:dots2="3" draw:dots2-length="4.5px" draw:distance="125%"/><draw:stroke-dash draw:name="rect" draw:style="rect" draw:dots1-length="1mm" draw:dots2-length="2in" draw:distance="3pt"/></office:styles>"#,
        );
        let parsed = parse_drawing_stroke_dashes(&xml).unwrap();
        assert_eq!(parsed.dashes.len(), 2);
        assert_eq!(
            parsed.get("mixed&units").unwrap().effective_style(),
            Style::Round
        );
        assert_eq!(
            parsed.dashes[0].distance.unwrap().unit(),
            MeasureUnit::Percent
        );
        assert_eq!(
            parsed.dashes[1].dots2_length.unwrap().unit(),
            MeasureUnit::Inch
        );
        let serialized = parsed.to_xml().unwrap();
        assert_eq!(parse_drawing_stroke_dashes(&serialized).unwrap(), parsed);
    }

    #[test]
    fn rejects_malformed_or_misplaced_resources() {
        let wrap = |body: &str| {
            format!(
                r#"<office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}">{body}</office:styles>"#
            )
        };
        for xml in [
            wrap(r#"<draw:stroke-dash draw:name="x" draw:distance="-1cm"/>"#),
            wrap(r#"<draw:stroke-dash draw:name="x" draw:dots1="1000001"/>"#),
            wrap(r#"<draw:stroke-dash draw:name="x" draw:style="square"/>"#),
            wrap(
                r#"<draw:stroke-dash draw:name="x"><draw:stroke-dash draw:name="y"/></draw:stroke-dash>"#,
            ),
            wrap(r#"<draw:stroke-dash draw:name="x"/><draw:stroke-dash draw:name="x"/>"#),
            format!(
                r#"<office:document xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><draw:stroke-dash draw:name="x"/></office:document>"#
            ),
            format!(
                r#"<office:styles xmlns:office="{OFFICE}" xmlns:evil="urn:evil"><evil:stroke-dash evil:name="x"/></office:styles>"#
            ),
            format!(
                r#"<!DOCTYPE x><office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><draw:stroke-dash draw:name="x"/></office:styles>"#
            ),
        ] {
            assert!(parse_drawing_stroke_dashes(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn parses_local_dashed_line_fixture() {
        let xml = include_str!("../../../../../test-data/odf/drawing/dashed-line.fodg");
        let parsed = parse_drawing_stroke_dashes(xml).unwrap();
        let dash = parsed.get("DoubleDashDotDot").unwrap();
        assert_eq!(dash.dots1, Some(1));
        assert_eq!(dash.dots2, Some(2));
        assert_eq!(dash.dots1_length.unwrap().value(), 800.0);
        assert_eq!(dash.distance.unwrap().value(), 300.0);
    }
}
