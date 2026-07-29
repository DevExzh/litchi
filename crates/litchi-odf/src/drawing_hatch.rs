//! Typed ODF drawing hatch resources.

use crate::drawing_gradient::OdfRgbColor;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_HATCHES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// Number of line directions in an ODF hatch.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfHatchStyle {
    Single,
    Double,
    Triple,
}

impl OdfHatchStyle {
    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "single" => Self::Single,
            "double" => Self::Double,
            "triple" => Self::Triple,
            _ => return invalid(format!("unsupported draw:hatch style '{value}'")),
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Single => "single",
            Self::Double => "double",
            Self::Triple => "triple",
        }
    }
}

/// Physical unit accepted by ODF hatch spacing.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfHatchLengthUnit {
    Centimeter,
    Millimeter,
    Inch,
    Point,
    Pica,
    Pixel,
}

impl OdfHatchLengthUnit {
    const fn suffix(self) -> &'static str {
        match self {
            Self::Centimeter => "cm",
            Self::Millimeter => "mm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Pixel => "px",
        }
    }
}

/// A finite signed physical hatch spacing.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct OdfHatchLength {
    value: f64,
    unit: OdfHatchLengthUnit,
}

impl OdfHatchLength {
    pub fn new(value: f64, unit: OdfHatchLengthUnit) -> Result<Self> {
        if !value.is_finite() {
            return invalid("hatch distance must be finite");
        }
        Ok(Self { value, unit })
    }

    pub const fn value(self) -> f64 {
        self.value
    }

    pub const fn unit(self) -> OdfHatchLengthUnit {
        self.unit
    }
}

impl FromStr for OdfHatchLength {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        if value.len() < 2 {
            return invalid(format!("invalid hatch distance '{value}'"));
        }
        let (number, suffix) = value.split_at(value.len() - 2);
        let unit = match suffix {
            "cm" => OdfHatchLengthUnit::Centimeter,
            "mm" => OdfHatchLengthUnit::Millimeter,
            "in" => OdfHatchLengthUnit::Inch,
            "pt" => OdfHatchLengthUnit::Point,
            "pc" => OdfHatchLengthUnit::Pica,
            "px" => OdfHatchLengthUnit::Pixel,
            _ => return invalid(format!("invalid hatch distance '{value}'")),
        };
        validate_decimal(number, value)?;
        let number = number
            .parse::<f64>()
            .map_err(|_| make_error(format!("invalid hatch distance '{value}'")))?;
        Self::new(number, unit)
    }
}

impl fmt::Display for OdfHatchLength {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let value = if self.value == 0.0 { 0.0 } else { self.value };
        write!(formatter, "{}{}", value, self.unit.suffix())
    }
}

/// Inert lexical hatch rotation.
///
/// ODF 1.2 deliberately leaves the angle datatype open. Keeping a validated
/// newtype preserves degrees, grads, radians, and unitless producer values.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfHatchRotation(String);

impl OdfHatchRotation {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, "hatch rotation", false)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// One named `draw:hatch` resource.
#[derive(Clone, Debug, PartialEq)]
pub struct OdfDrawingHatch {
    pub name: String,
    pub display_name: Option<String>,
    pub style: OdfHatchStyle,
    pub color: Option<OdfRgbColor>,
    pub distance: Option<OdfHatchLength>,
    pub rotation: Option<OdfHatchRotation>,
}

impl OdfDrawingHatch {
    pub fn new(name: impl Into<String>, style: OdfHatchStyle) -> Result<Self> {
        let value = Self {
            name: name.into(),
            display_name: None,
            style,
            color: None,
            distance: None,
            rotation: None,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "hatch name", false)?;
        if let Some(value) = &self.display_name {
            validate_text(value, "hatch display name", true)?;
        }
        if let Some(value) = &self.rotation {
            validate_text(value.as_str(), "hatch rotation", false)?;
        }
        if let Some(value) = self.distance {
            if !value.value().is_finite() {
                return invalid("hatch distance must be finite");
            }
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192);
        write_hatch(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered hatch resources from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct OdfDrawingHatches {
    pub hatches: Vec<OdfDrawingHatch>,
}

impl OdfDrawingHatches {
    pub fn get(&self, name: &str) -> Option<&OdfDrawingHatch> {
        self.hatches.iter().find(|hatch| hatch.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.hatches.len() > MAX_HATCHES {
            return invalid(format!("drawing styles exceed {MAX_HATCHES} hatches"));
        }
        let mut names = HashSet::with_capacity(self.hatches.len());
        let mut aggregate = 0usize;
        for hatch in &self.hatches {
            hatch.validate()?;
            if !names.insert(hatch.name.as_str()) {
                return invalid(format!("duplicate drawing hatch name '{}'", hatch.name));
            }
            aggregate = aggregate
                .checked_add(hatch.name.len())
                .and_then(|value| {
                    value.checked_add(hatch.display_name.as_deref().map_or(0, str::len))
                })
                .and_then(|value| {
                    value.checked_add(
                        hatch
                            .rotation
                            .as_ref()
                            .map_or(0, |item| item.as_str().len()),
                    )
                })
                .ok_or_else(|| make_error("drawing hatch size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("drawing hatch values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize a standalone schema-positioned `office:styles` fragment.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192 + self.hatches.len() * 160);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">"#,
        );
        for hatch in &self.hatches {
            write_hatch(&mut output, hatch, false);
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

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parse named hatch resources from an ODF styles or flat-document XML part.
pub fn parse_drawing_hatches(xml: &str) -> Result<OdfDrawingHatches> {
    if !xml.contains("hatch") {
        return Ok(OdfDrawingHatches::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing hatch XML exceeds 64 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut result = OdfDrawingHatches::default();

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing hatch XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if namespace == NamespaceKind::Draw && local == "hatch" {
                    ensure_location(&stack)?;
                    return invalid("draw:hatch must be empty");
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("drawing hatch XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if namespace == NamespaceKind::Draw && local == "hatch" {
                    ensure_location(&stack)?;
                    if result.hatches.len() >= MAX_HATCHES {
                        return invalid(format!("drawing styles exceed {MAX_HATCHES} hatches"));
                    }
                    result.hatches.push(parse_hatch(&reader, element)?);
                }
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| make_error("drawing hatch XML depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in hatches");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unterminated drawing hatch XML");
    }
    result.validate()?;
    Ok(result)
}

fn parse_hatch(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<OdfDrawingHatch> {
    let mut values = attributes(reader, element)?;
    let name = required(&mut values, "name", "draw:name")?;
    let display_name = take(&mut values, "display-name");
    let style = OdfHatchStyle::parse(&required(&mut values, "style", "draw:style")?)?;
    let color = take(&mut values, "color")
        .map(|value| value.parse())
        .transpose()?;
    let distance = take(&mut values, "distance")
        .map(|value| value.parse())
        .transpose()?;
    let rotation = take(&mut values, "rotation")
        .map(OdfHatchRotation::new)
        .transpose()?;
    reject_attributes(&values)?;
    let value = OdfDrawingHatch {
        name,
        display_name,
        style,
        color,
        distance,
        rotation,
    };
    value.validate()?;
    Ok(value)
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| make_error(format!("invalid hatch attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref(), "attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid hatch attribute value: {error}")))?
            .into_owned();
        validate_text(&value, "hatch attribute", true)?;
        if result.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded hatch attribute");
        }
    }
    Ok(result)
}

fn take(values: &mut Attributes, local: &str) -> Option<String> {
    values.remove(&(NamespaceKind::Draw, local.to_string()))
}

fn required(values: &mut Attributes, local: &str, context: &str) -> Result<String> {
    take(values, local).ok_or_else(|| make_error(format!("missing required {context}")))
}

fn reject_attributes(values: &Attributes) -> Result<()> {
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported draw:hatch attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if !matches!(stack.last(), Some(Frame { namespace: NamespaceKind::Office, local }) if local == "styles")
    {
        return invalid("draw:hatch must be a direct office:styles child");
    }
    Ok(())
}

fn namespace_kind(value: &ResolveResult<'_>) -> Result<NamespaceKind> {
    Ok(match value {
        ResolveResult::Unbound => NamespaceKind::None,
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        ResolveResult::Bound(_) => NamespaceKind::Other,
        ResolveResult::Unknown(prefix) => {
            return invalid(format!(
                "undeclared drawing hatch prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ));
        },
    })
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "hatch" && namespace != NamespaceKind::Draw {
        return invalid("hatch element uses an invalid namespace");
    }
    Ok(())
}

fn write_hatch(output: &mut String, hatch: &OdfDrawingHatch, standalone: bool) {
    output.push_str("<draw:hatch");
    if standalone {
        output.push_str(r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0""#);
    }
    push_attribute(output, "draw:name", &hatch.name);
    if let Some(value) = &hatch.display_name {
        push_attribute(output, "draw:display-name", value);
    }
    push_attribute(output, "draw:style", hatch.style.as_str());
    if let Some(value) = hatch.color {
        push_attribute(output, "draw:color", &value.to_string());
    }
    if let Some(value) = hatch.distance {
        push_attribute(output, "draw:distance", &value.to_string());
    }
    if let Some(value) = &hatch.rotation {
        push_attribute(output, "draw:rotation", value.as_str());
    }
    output.push_str("/>");
}

fn push_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
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
    output.push('"');
}

fn validate_decimal(value: &str, complete: &str) -> Result<()> {
    let value = value.strip_prefix('-').unwrap_or(value);
    if value.is_empty() {
        return invalid(format!("invalid hatch distance '{complete}'"));
    }
    let mut parts = value.split('.');
    let integer = parts.next().expect("split always yields one value");
    let fraction = parts.next();
    if parts.next().is_some()
        || !integer.bytes().all(|byte| byte.is_ascii_digit())
        || fraction
            .is_some_and(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
        || integer.is_empty() && fraction.is_none()
    {
        return invalid(format!("invalid hatch distance '{complete}'"));
    }
    Ok(())
}

fn validate_text(value: &str, context: &str, empty_allowed: bool) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return invalid(format!("{context} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{context} exceeds 64 KiB"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{context} contains an XML-prohibited character"));
    }
    Ok(())
}

fn decode_name(value: &[u8], context: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| make_error(format!("invalid UTF-8 in hatch {context} name")))
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

    const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:styles>"#;
    const SUFFIX: &str = "</office:styles></office:document-styles>";

    #[test]
    fn parses_and_round_trips_all_hatch_styles() {
        let xml = format!(
            r##"{PREFIX}<draw:hatch draw:name="single" draw:display-name="Single &amp; Blue" draw:style="single" draw:color="#0000ff" draw:distance="0.5cm" draw:rotation="58.5deg"/><draw:hatch draw:name="double" draw:style="double" draw:distance="2pt" draw:rotation="65grad"/><draw:hatch draw:name="triple" draw:style="triple" draw:rotation="1.02101761241558rad"/>{SUFFIX}"##
        );
        let hatches = parse_drawing_hatches(&xml).unwrap();
        assert_eq!(hatches.hatches.len(), 3);
        assert_eq!(
            hatches.hatches[0].display_name.as_deref(),
            Some("Single & Blue")
        );
        assert_eq!(hatches.hatches[0].distance.unwrap().value(), 0.5);
        assert_eq!(
            hatches.hatches[2].rotation.as_ref().unwrap().as_str(),
            "1.02101761241558rad"
        );

        let serialized = hatches.to_xml().unwrap();
        assert_eq!(parse_drawing_hatches(&serialized).unwrap(), hatches);
    }

    #[test]
    fn rejects_malformed_hatches() {
        for body in [
            r#"<draw:hatch draw:style="single"/>"#,
            r#"<draw:hatch draw:name="x"/>"#,
            r#"<draw:hatch draw:name="x" draw:style="quadruple"/>"#,
            r##"<draw:hatch draw:name="x" draw:style="single" draw:color="#gg0000"/>"##,
            r#"<draw:hatch draw:name="x" draw:style="single" draw:distance="1%"/>"#,
            r#"<draw:hatch draw:name="x" draw:style="single"></draw:hatch>"#,
            r#"<draw:hatch draw:name="x" draw:style="single"/><draw:hatch draw:name="x" draw:style="double"/>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(parse_drawing_hatches(&xml).is_err(), "accepted {body}");
        }
        let misplaced = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><draw:hatch draw:name="x" draw:style="single"/></office:automatic-styles></office:document-styles>"#.to_string();
        assert!(parse_drawing_hatches(&misplaced).is_err());
    }

    #[test]
    fn parses_local_angle_fixture() {
        let xml = include_str!("../../../test-data/odf/drawing/hatch-angles.fodg");
        let hatches = parse_drawing_hatches(xml).unwrap();
        assert_eq!(hatches.hatches.len(), 4);
        assert_eq!(
            hatches.hatches[0].rotation.as_ref().unwrap().as_str(),
            "58.5deg"
        );
        assert_eq!(
            hatches.hatches[1].rotation.as_ref().unwrap().as_str(),
            "65grad"
        );
        assert_eq!(
            hatches.hatches[2].rotation.as_ref().unwrap().as_str(),
            "1.02101761241558rad"
        );
        assert_eq!(
            hatches.hatches[3].rotation.as_ref().unwrap().as_str(),
            "585"
        );
        assert_eq!(
            parse_drawing_hatches(&hatches.to_xml().unwrap())
                .unwrap()
                .hatches
                .len(),
            4
        );
    }
}
