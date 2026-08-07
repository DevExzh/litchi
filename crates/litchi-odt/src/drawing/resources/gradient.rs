//! Typed, inert ODF drawing gradient resources.

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
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const LOEXT_NS: &[u8] = b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_GRADIENTS: usize = 65_536;
const MAX_STOPS: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// Six legacy ODF gradient geometries.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LegacyStyle {
    Linear,
    Axial,
    Radial,
    Ellipsoid,
    Square,
    Rectangular,
}

impl LegacyStyle {
    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "linear" => Self::Linear,
            "axial" => Self::Axial,
            "radial" => Self::Radial,
            "ellipsoid" => Self::Ellipsoid,
            "square" => Self::Square,
            "rectangular" => Self::Rectangular,
            _ => return invalid(format!("unsupported draw:gradient style '{value}'")),
        })
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

/// A finite signed percentage.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Percent(f64);

impl Percent {
    pub fn new(value: f64) -> Result<Self> {
        finite(value, "gradient percentage")?;
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

/// A gradient intensity constrained to 0 through 100 percent.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Intensity(f64);

impl Intensity {
    pub fn new(value: f64) -> Result<Self> {
        finite(value, "gradient intensity")?;
        if !(0.0..=100.0).contains(&value) {
            return invalid("gradient intensity must be between 0% and 100%");
        }
        Ok(Self(value))
    }

    pub const fn value(self) -> f64 {
        self.0
    }
}

impl fmt::Display for Intensity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}%", canonical_number(self.0))
    }
}

/// An RGB color represented by its three channels.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RgbColor {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
}

impl RgbColor {
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }
}

impl FromStr for RgbColor {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        if value.len() != 7 || !value.starts_with('#') {
            return invalid(format!("invalid ODF color '{value}'"));
        }
        let red = u8::from_str_radix(&value[1..3], 16)
            .map_err(|_| make_error(format!("invalid ODF color '{value}'")))?;
        let green = u8::from_str_radix(&value[3..5], 16)
            .map_err(|_| make_error(format!("invalid ODF color '{value}'")))?;
        let blue = u8::from_str_radix(&value[5..7], 16)
            .map_err(|_| make_error(format!("invalid ODF color '{value}'")))?;
        Ok(Self::new(red, green, blue))
    }
}

impl fmt::Display for RgbColor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "#{:02x}{:02x}{:02x}",
            self.red, self.green, self.blue
        )
    }
}

/// An inert lexical ODF angle, retained because ODF 1.2 deliberately leaves its grammar open.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Angle(String);

impl Angle {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, "gradient angle", false)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Unit for an SVG gradient coordinate.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CoordinateUnit {
    Centimeter,
    Millimeter,
    Inch,
    Point,
    Pica,
    Pixel,
    Percent,
}

impl CoordinateUnit {
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

/// A finite SVG coordinate expressed as an ODF length or percentage.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Coordinate {
    value: f64,
    unit: CoordinateUnit,
}

impl Coordinate {
    pub fn new(value: f64, unit: CoordinateUnit) -> Result<Self> {
        finite(value, "gradient coordinate")?;
        Ok(Self { value, unit })
    }

    pub const fn value(self) -> f64 {
        self.value
    }

    pub const fn unit(self) -> CoordinateUnit {
        self.unit
    }
}

impl FromStr for Coordinate {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let (number, unit) = split_measure(value)?;
        validate_decimal(number, value)?;
        let number = number
            .parse::<f64>()
            .map_err(|_| make_error(format!("invalid gradient coordinate '{value}'")))?;
        Self::new(number, unit)
    }
}

impl fmt::Display for Coordinate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{}{}",
            canonical_number(self.value),
            self.unit.suffix()
        )
    }
}

/// SVG gradient spread behavior.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SpreadMethod {
    Pad,
    Reflect,
    Repeat,
}

impl SpreadMethod {
    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "pad" => Self::Pad,
            "reflect" => Self::Reflect,
            "repeat" => Self::Repeat,
            _ => return invalid(format!("unsupported SVG gradient spread method '{value}'")),
        })
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Pad => "pad",
            Self::Reflect => "reflect",
            Self::Repeat => "repeat",
        }
    }
}

/// A stop position represented either as a number or percentage.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum StopOffset {
    Number(f64),
    Percent(Percent),
}

impl StopOffset {
    fn parse(value: &str) -> Result<Self> {
        if let Some(number) = value.strip_suffix('%') {
            validate_decimal(number, value)?;
            return Ok(Self::Percent(Percent::new(parse_number(
                number,
                "gradient stop offset",
            )?)?));
        }
        Ok(Self::Number(parse_number(value, "gradient stop offset")?))
    }

    fn validate(self) -> Result<()> {
        match self {
            Self::Number(value) => finite(value, "gradient stop offset"),
            Self::Percent(value) => finite(value.value(), "gradient stop percentage"),
        }
    }
}

impl fmt::Display for StopOffset {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Number(value) => formatter.write_str(&canonical_number(*value)),
            Self::Percent(value) => value.fmt(formatter),
        }
    }
}

/// One standard SVG gradient stop.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgStop {
    pub offset: StopOffset,
    pub color: Option<RgbColor>,
    pub opacity: Option<f64>,
}

impl SvgStop {
    fn validate(&self) -> Result<()> {
        self.offset.validate()?;
        if let Some(value) = self.opacity {
            finite(value, "gradient stop opacity")?;
        }
        Ok(())
    }
}

/// Color representation used by `LibreOffice` multi-color gradient stops.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum LibreOfficeColorType {
    Rgb,
    Theme,
    Other(String),
}

impl LibreOfficeColorType {
    fn parse(value: String) -> Self {
        match value.as_str() {
            "rgb" => Self::Rgb,
            "theme" => Self::Theme,
            _ => Self::Other(value),
        }
    }

    fn as_str(&self) -> &str {
        match self {
            Self::Rgb => "rgb",
            Self::Theme => "theme",
            Self::Other(value) => value,
        }
    }
}

/// Inert `LibreOffice` multi-color stop metadata retained on a legacy gradient.
#[derive(Clone, Debug, PartialEq)]
pub struct LibreOfficeStop {
    pub offset: StopOffset,
    pub color_type: LibreOfficeColorType,
    pub color_value: String,
}

impl LibreOfficeStop {
    fn validate(&self) -> Result<()> {
        self.offset.validate()?;
        validate_text(
            self.color_type.as_str(),
            "LibreOffice gradient color type",
            false,
        )?;
        validate_text(&self.color_value, "LibreOffice gradient color value", false)
    }
}

/// Legacy ODF `draw:gradient` resource.
#[derive(Clone, Debug, PartialEq)]
pub struct Legacy {
    pub name: Option<String>,
    pub display_name: Option<String>,
    pub style: LegacyStyle,
    pub center_x: Option<Percent>,
    pub center_y: Option<Percent>,
    pub start_color: Option<RgbColor>,
    pub end_color: Option<RgbColor>,
    pub start_intensity: Option<Intensity>,
    pub end_intensity: Option<Intensity>,
    pub angle: Option<Angle>,
    pub border: Option<Percent>,
    pub extension_stops: Vec<LibreOfficeStop>,
}

impl Legacy {
    fn validate(&self) -> Result<()> {
        if let Some(value) = &self.name {
            validate_text(value, "gradient name", false)?;
        }
        if let Some(value) = &self.display_name {
            validate_text(value, "gradient display name", true)?;
        }
        if let Some(value) = &self.angle {
            validate_text(value.as_str(), "gradient angle", false)?;
        }
        if self.extension_stops.len() > MAX_STOPS {
            return invalid(format!("gradient exceeds {MAX_STOPS} extension stops"));
        }
        for stop in &self.extension_stops {
            stop.validate()?;
        }
        Ok(())
    }
}

/// Attributes shared by SVG linear and radial gradients.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgCommon {
    pub name: String,
    pub display_name: Option<String>,
    pub object_bounding_box_units: Option<bool>,
    pub transform: Option<String>,
    pub spread_method: Option<SpreadMethod>,
    pub stops: Vec<SvgStop>,
}

impl SvgCommon {
    fn validate(&self) -> Result<()> {
        validate_text(&self.name, "SVG gradient name", false)?;
        if let Some(value) = &self.display_name {
            validate_text(value, "SVG gradient display name", true)?;
        }
        if self.object_bounding_box_units == Some(false) {
            return invalid("ODF SVG gradient units can only be objectBoundingBox");
        }
        if let Some(value) = &self.transform {
            validate_text(value, "SVG gradient transform", true)?;
        }
        if self.stops.len() > MAX_STOPS {
            return invalid(format!("SVG gradient exceeds {MAX_STOPS} stops"));
        }
        for stop in &self.stops {
            stop.validate()?;
        }
        Ok(())
    }
}

/// Standard SVG linear gradient resource.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgLinear {
    pub common: SvgCommon,
    pub x1: Option<Coordinate>,
    pub y1: Option<Coordinate>,
    pub x2: Option<Coordinate>,
    pub y2: Option<Coordinate>,
}

/// Standard SVG radial gradient resource.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgRadial {
    pub common: SvgCommon,
    pub center_x: Option<Coordinate>,
    pub center_y: Option<Coordinate>,
    pub radius: Option<Coordinate>,
    pub focus_x: Option<Coordinate>,
    pub focus_y: Option<Coordinate>,
}

/// One gradient resource in document order.
#[derive(Clone, Debug, PartialEq)]
pub enum Definition {
    Legacy(Legacy),
    Linear(SvgLinear),
    Radial(SvgRadial),
}

impl Definition {
    pub fn name(&self) -> Option<&str> {
        match self {
            Self::Legacy(value) => value.name.as_deref(),
            Self::Linear(value) => Some(&value.common.name),
            Self::Radial(value) => Some(&value.common.name),
        }
    }

    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Legacy(value) => value.validate(),
            Self::Linear(value) => value.common.validate(),
            Self::Radial(value) => value.common.validate(),
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::new();
        write_gradient(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered named gradients from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Collection {
    pub gradients: Vec<Definition>,
}

impl Collection {
    pub fn get(&self, name: &str) -> Option<&Definition> {
        self.gradients
            .iter()
            .find(|value| value.name() == Some(name))
    }

    pub fn validate(&self) -> Result<()> {
        if self.gradients.len() > MAX_GRADIENTS {
            return invalid(format!("drawing styles exceed {MAX_GRADIENTS} gradients"));
        }
        let mut names = HashSet::with_capacity(self.gradients.len());
        let mut aggregate = 0usize;
        for gradient in &self.gradients {
            gradient.validate()?;
            if let Some(name) = gradient.name() {
                if !names.insert(name) {
                    return invalid(format!("duplicate drawing gradient name '{name}'"));
                }
                aggregate = aggregate
                    .checked_add(name.len())
                    .ok_or_else(|| make_error("drawing gradient size overflow"))?;
            }
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("drawing gradient values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize a standalone schema-positioned `office:styles` fragment.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + self.gradients.len() * 256);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">"#,
        );
        for gradient in &self.gradients {
            write_gradient(&mut output, gradient, false);
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
    Loext,
    Other,
}

#[derive(Clone)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}

struct ActiveDefinition {
    depth: usize,
    value: Definition,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parse legacy and SVG gradient resources from an ODF styles or flat-document XML part.
pub fn parse_drawing_gradients(xml: &str) -> Result<Collection> {
    if !xml.contains("gradient") && !xml.contains("Gradient") {
        return Ok(Collection::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing gradient XML exceeds 64 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveDefinition> = None;
    let mut result = Collection::default();

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing gradient XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if active.is_some() {
                    if matches!(
                        (namespace, local.as_str()),
                        (NamespaceKind::Svg, "stop") | (NamespaceKind::Loext, "gradient-stop")
                    ) {
                        return invalid("gradient stop elements must be empty");
                    }
                    return invalid("gradient resource contains an unsupported child element");
                }
                if let Some(value) = parse_gradient_start(&reader, namespace, &local, element)? {
                    ensure_location(&stack)?;
                    if result.gradients.len() >= MAX_GRADIENTS {
                        return invalid(format!("drawing styles exceed {MAX_GRADIENTS} gradients"));
                    }
                    active = Some(ActiveDefinition {
                        depth: stack.len(),
                        value,
                    });
                } else if is_stop(namespace, &local) {
                    return invalid("gradient stop must be inside a gradient resource");
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("drawing gradient XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode_name(element.local_name().as_ref(), "element")?;
                reject_spoofed_name(namespace, &local)?;
                if let Some(active) = active.as_mut() {
                    if stack.len() != active.depth + 1 {
                        return invalid("gradient stop must be a direct gradient child");
                    }
                    add_stop(&reader, namespace, &local, element, &mut active.value)?;
                } else if let Some(value) =
                    parse_gradient_start(&reader, namespace, &local, element)?
                {
                    ensure_location(&stack)?;
                    result.gradients.push(value);
                } else if is_stop(namespace, &local) {
                    return invalid("gradient stop must be inside a gradient resource");
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("drawing gradient XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|gradient| gradient.depth == stack.len())
                {
                    if !is_gradient(frame.namespace, &frame.local) {
                        return invalid("unexpected drawing gradient end element");
                    }
                    result
                        .gradients
                        .push(active.take().expect("active gradient checked").value);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid gradient text: {error}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("gradient resources cannot contain text");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("gradient resources cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in gradients");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated drawing gradient XML");
    }
    result.validate()?;
    Ok(result)
}

fn parse_gradient_start(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    local: &str,
    element: &BytesStart<'_>,
) -> Result<Option<Definition>> {
    Ok(Some(match (namespace, local) {
        (NamespaceKind::Draw, "gradient") => {
            Definition::Legacy(parse_legacy_gradient(reader, element)?)
        },
        (NamespaceKind::Svg, "linearGradient") => {
            Definition::Linear(parse_linear_gradient(reader, element)?)
        },
        (NamespaceKind::Svg, "radialGradient") => {
            Definition::Radial(parse_radial_gradient(reader, element)?)
        },
        _ => return Ok(None),
    }))
}

fn parse_legacy_gradient(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Legacy> {
    let mut values = attributes(reader, element)?;
    let name = take(&mut values, NamespaceKind::Draw, "name");
    let display_name = take(&mut values, NamespaceKind::Draw, "display-name");
    let style = LegacyStyle::parse(&required(
        &mut values,
        NamespaceKind::Draw,
        "style",
        "draw:style",
    )?)?;
    let center_x = take(&mut values, NamespaceKind::Draw, "cx")
        .map(|value| parse_percent(&value, "draw:cx"))
        .transpose()?;
    let center_y = take(&mut values, NamespaceKind::Draw, "cy")
        .map(|value| parse_percent(&value, "draw:cy"))
        .transpose()?;
    let start_color = take(&mut values, NamespaceKind::Draw, "start-color")
        .map(|value| value.parse())
        .transpose()?;
    let end_color = take(&mut values, NamespaceKind::Draw, "end-color")
        .map(|value| value.parse())
        .transpose()?;
    let start_intensity = take(&mut values, NamespaceKind::Draw, "start-intensity")
        .map(|value| parse_intensity(&value, "draw:start-intensity"))
        .transpose()?;
    let end_intensity = take(&mut values, NamespaceKind::Draw, "end-intensity")
        .map(|value| parse_intensity(&value, "draw:end-intensity"))
        .transpose()?;
    let angle = take(&mut values, NamespaceKind::Draw, "angle")
        .map(Angle::new)
        .transpose()?;
    let border = take(&mut values, NamespaceKind::Draw, "border")
        .map(|value| parse_percent(&value, "draw:border"))
        .transpose()?;
    reject_attributes(&values, "draw:gradient")?;
    Ok(Legacy {
        name,
        display_name,
        style,
        center_x,
        center_y,
        start_color,
        end_color,
        start_intensity,
        end_intensity,
        angle,
        border,
        extension_stops: Vec::new(),
    })
}

fn parse_linear_gradient(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SvgLinear> {
    let mut values = attributes(reader, element)?;
    let common = parse_svg_common(&mut values)?;
    let x1 = take_coordinate(&mut values, "x1")?;
    let y1 = take_coordinate(&mut values, "y1")?;
    let x2 = take_coordinate(&mut values, "x2")?;
    let y2 = take_coordinate(&mut values, "y2")?;
    reject_attributes(&values, "svg:linearGradient")?;
    Ok(SvgLinear {
        common,
        x1,
        y1,
        x2,
        y2,
    })
}

fn parse_radial_gradient(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SvgRadial> {
    let mut values = attributes(reader, element)?;
    let common = parse_svg_common(&mut values)?;
    let center_x = take_coordinate(&mut values, "cx")?;
    let center_y = take_coordinate(&mut values, "cy")?;
    let radius = take_coordinate(&mut values, "r")?;
    let focus_x = take_coordinate(&mut values, "fx")?;
    let focus_y = take_coordinate(&mut values, "fy")?;
    reject_attributes(&values, "svg:radialGradient")?;
    Ok(SvgRadial {
        common,
        center_x,
        center_y,
        radius,
        focus_x,
        focus_y,
    })
}

fn parse_svg_common(values: &mut Attributes) -> Result<SvgCommon> {
    let name = required(values, NamespaceKind::Draw, "name", "draw:name")?;
    let display_name = take(values, NamespaceKind::Draw, "display-name");
    let object_bounding_box_units = take(values, NamespaceKind::Svg, "gradientUnits")
        .map(|value| {
            if value == "objectBoundingBox" {
                Ok(true)
            } else {
                invalid(format!("unsupported svg:gradientUnits '{value}'"))
            }
        })
        .transpose()?;
    let transform = take(values, NamespaceKind::Svg, "gradientTransform");
    let spread_method = take(values, NamespaceKind::Svg, "spreadMethod")
        .map(|value| SpreadMethod::parse(&value))
        .transpose()?;
    Ok(SvgCommon {
        name,
        display_name,
        object_bounding_box_units,
        transform,
        spread_method,
        stops: Vec::new(),
    })
}

fn add_stop(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    local: &str,
    element: &BytesStart<'_>,
    gradient: &mut Definition,
) -> Result<()> {
    match gradient {
        Definition::Legacy(value)
            if namespace == NamespaceKind::Loext && local == "gradient-stop" =>
        {
            if value.extension_stops.len() >= MAX_STOPS {
                return invalid(format!("gradient exceeds {MAX_STOPS} extension stops"));
            }
            value
                .extension_stops
                .push(parse_loext_stop(reader, element)?);
        },
        Definition::Linear(value) if namespace == NamespaceKind::Svg && local == "stop" => {
            if value.common.stops.len() >= MAX_STOPS {
                return invalid(format!("SVG gradient exceeds {MAX_STOPS} stops"));
            }
            value.common.stops.push(parse_svg_stop(reader, element)?);
        },
        Definition::Radial(value) if namespace == NamespaceKind::Svg && local == "stop" => {
            if value.common.stops.len() >= MAX_STOPS {
                return invalid(format!("SVG gradient exceeds {MAX_STOPS} stops"));
            }
            value.common.stops.push(parse_svg_stop(reader, element)?);
        },
        _ => return invalid("gradient contains an unsupported stop element"),
    }
    Ok(())
}

fn parse_svg_stop(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SvgStop> {
    let mut values = attributes(reader, element)?;
    let offset = StopOffset::parse(&required(
        &mut values,
        NamespaceKind::Svg,
        "offset",
        "svg:offset",
    )?)?;
    let color = take(&mut values, NamespaceKind::Svg, "stop-color")
        .map(|value| value.parse())
        .transpose()?;
    let opacity = take(&mut values, NamespaceKind::Svg, "stop-opacity")
        .map(|value| parse_number(&value, "svg:stop-opacity"))
        .transpose()?;
    reject_attributes(&values, "svg:stop")?;
    let value = SvgStop {
        offset,
        color,
        opacity,
    };
    value.validate()?;
    Ok(value)
}

fn parse_loext_stop(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<LibreOfficeStop> {
    let mut values = attributes(reader, element)?;
    let offset = StopOffset::parse(&required(
        &mut values,
        NamespaceKind::Svg,
        "offset",
        "svg:offset",
    )?)?;
    let color_type = LibreOfficeColorType::parse(required(
        &mut values,
        NamespaceKind::Loext,
        "color-type",
        "loext:color-type",
    )?);
    let color_value = required(
        &mut values,
        NamespaceKind::Loext,
        "color-value",
        "loext:color-value",
    )?;
    reject_attributes(&values, "loext:gradient-stop")?;
    let value = LibreOfficeStop {
        offset,
        color_type,
        color_value,
    };
    value.validate()?;
    Ok(value)
}

fn take_coordinate(values: &mut Attributes, local: &str) -> Result<Option<Coordinate>> {
    take(values, NamespaceKind::Svg, local)
        .map(|value| value.parse())
        .transpose()
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid gradient attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref(), "attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid gradient attribute value: {error}")))?
            .into_owned();
        validate_text(&value, "gradient attribute", true)?;
        if result.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded gradient attribute");
        }
    }
    Ok(result)
}

fn take(values: &mut Attributes, namespace: NamespaceKind, local: &str) -> Option<String> {
    values.remove(&(namespace, local.to_string()))
}

fn required(
    values: &mut Attributes,
    namespace: NamespaceKind,
    local: &str,
    context: &str,
) -> Result<String> {
    take(values, namespace, local).ok_or_else(|| make_error(format!("missing required {context}")))
}

fn reject_attributes(values: &Attributes, context: &str) -> Result<()> {
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported {context} attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if !matches!(stack.last(), Some(Frame { namespace: NamespaceKind::Office, local }) if local == "styles")
    {
        return invalid("gradient resources must be direct office:styles children");
    }
    Ok(())
}

fn is_gradient(namespace: NamespaceKind, local: &str) -> bool {
    matches!(
        (namespace, local),
        (NamespaceKind::Draw, "gradient")
            | (NamespaceKind::Svg, "linearGradient" | "radialGradient")
    )
}

fn is_stop(namespace: NamespaceKind, local: &str) -> bool {
    matches!(
        (namespace, local),
        (NamespaceKind::Svg, "stop") | (NamespaceKind::Loext, "gradient-stop")
    )
}

fn namespace_kind(value: &ResolveResult<'_>) -> Result<NamespaceKind> {
    Ok(match value {
        ResolveResult::Unbound => NamespaceKind::None,
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NS => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NS => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(value)) if *value == SVG_NS => NamespaceKind::Svg,
        ResolveResult::Bound(Namespace(value)) if *value == LOEXT_NS => NamespaceKind::Loext,
        ResolveResult::Bound(_) => NamespaceKind::Other,
        ResolveResult::Unknown(prefix) => {
            return invalid(format!(
                "undeclared drawing gradient prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ));
        },
    })
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "gradient" && namespace != NamespaceKind::Draw {
        return invalid("gradient element uses an invalid namespace");
    }
    if matches!(local, "linearGradient" | "radialGradient" | "stop")
        && namespace != NamespaceKind::Svg
    {
        return invalid("SVG gradient element uses an invalid namespace");
    }
    if local == "gradient-stop" && namespace != NamespaceKind::Loext {
        return invalid("gradient-stop element uses an invalid namespace");
    }
    Ok(())
}

fn write_gradient(output: &mut String, gradient: &Definition, standalone: bool) {
    match gradient {
        Definition::Legacy(value) => write_legacy(output, value, standalone),
        Definition::Linear(value) => write_linear(output, value, standalone),
        Definition::Radial(value) => write_radial(output, value, standalone),
    }
}

fn write_legacy(output: &mut String, value: &Legacy, standalone: bool) {
    output.push_str("<draw:gradient");
    write_namespaces(output, standalone);
    if let Some(name) = &value.name {
        push_attribute(output, "draw:name", name);
    }
    if let Some(name) = &value.display_name {
        push_attribute(output, "draw:display-name", name);
    }
    push_attribute(output, "draw:style", value.style.as_str());
    push_optional_display(output, "draw:cx", value.center_x);
    push_optional_display(output, "draw:cy", value.center_y);
    push_optional_display(output, "draw:start-color", value.start_color);
    push_optional_display(output, "draw:end-color", value.end_color);
    push_optional_display(output, "draw:start-intensity", value.start_intensity);
    push_optional_display(output, "draw:end-intensity", value.end_intensity);
    if let Some(angle) = &value.angle {
        push_attribute(output, "draw:angle", angle.as_str());
    }
    push_optional_display(output, "draw:border", value.border);
    if value.extension_stops.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for stop in &value.extension_stops {
        output.push_str("<loext:gradient-stop");
        push_attribute(output, "svg:offset", &stop.offset.to_string());
        push_attribute(output, "loext:color-type", stop.color_type.as_str());
        push_attribute(output, "loext:color-value", &stop.color_value);
        output.push_str("/>");
    }
    output.push_str("</draw:gradient>");
}

fn write_linear(output: &mut String, value: &SvgLinear, standalone: bool) {
    output.push_str("<svg:linearGradient");
    write_namespaces(output, standalone);
    write_svg_common(output, &value.common);
    push_optional_display(output, "svg:x1", value.x1);
    push_optional_display(output, "svg:y1", value.y1);
    push_optional_display(output, "svg:x2", value.x2);
    push_optional_display(output, "svg:y2", value.y2);
    write_svg_stops(output, "linearGradient", &value.common.stops);
}

fn write_radial(output: &mut String, value: &SvgRadial, standalone: bool) {
    output.push_str("<svg:radialGradient");
    write_namespaces(output, standalone);
    write_svg_common(output, &value.common);
    push_optional_display(output, "svg:cx", value.center_x);
    push_optional_display(output, "svg:cy", value.center_y);
    push_optional_display(output, "svg:r", value.radius);
    push_optional_display(output, "svg:fx", value.focus_x);
    push_optional_display(output, "svg:fy", value.focus_y);
    write_svg_stops(output, "radialGradient", &value.common.stops);
}

fn write_svg_common(output: &mut String, value: &SvgCommon) {
    push_attribute(output, "draw:name", &value.name);
    if let Some(name) = &value.display_name {
        push_attribute(output, "draw:display-name", name);
    }
    if value.object_bounding_box_units.is_some() {
        push_attribute(output, "svg:gradientUnits", "objectBoundingBox");
    }
    if let Some(transform) = &value.transform {
        push_attribute(output, "svg:gradientTransform", transform);
    }
    if let Some(method) = value.spread_method {
        push_attribute(output, "svg:spreadMethod", method.as_str());
    }
}

fn write_svg_stops(output: &mut String, tag: &str, stops: &[SvgStop]) {
    if stops.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for stop in stops {
        output.push_str("<svg:stop");
        push_attribute(output, "svg:offset", &stop.offset.to_string());
        push_optional_display(output, "svg:stop-color", stop.color);
        if let Some(opacity) = stop.opacity {
            push_attribute(output, "svg:stop-opacity", &canonical_number(opacity));
        }
        output.push_str("/>");
    }
    output.push_str("</svg:");
    output.push_str(tag);
    output.push('>');
}

fn write_namespaces(output: &mut String, standalone: bool) {
    if standalone {
        output.push_str(
            r#" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""#,
        );
    }
}

fn push_optional_display<T: fmt::Display>(output: &mut String, name: &str, value: Option<T>) {
    if let Some(value) = value {
        push_attribute(output, name, &value.to_string());
    }
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

fn split_measure(value: &str) -> Result<(&str, CoordinateUnit)> {
    if let Some(number) = value.strip_suffix('%') {
        return Ok((number, CoordinateUnit::Percent));
    }
    if value.len() < 2 {
        return invalid(format!("invalid gradient coordinate '{value}'"));
    }
    let (number, suffix) = value.split_at(value.len() - 2);
    let unit = match suffix {
        "cm" => CoordinateUnit::Centimeter,
        "mm" => CoordinateUnit::Millimeter,
        "in" => CoordinateUnit::Inch,
        "pt" => CoordinateUnit::Point,
        "pc" => CoordinateUnit::Pica,
        "px" => CoordinateUnit::Pixel,
        _ => return invalid(format!("invalid gradient coordinate '{value}'")),
    };
    Ok((number, unit))
}

fn parse_percent(value: &str, context: &str) -> Result<Percent> {
    let number = value
        .strip_suffix('%')
        .ok_or_else(|| make_error(format!("{context} must be a percentage")))?;
    validate_decimal(number, value)?;
    Percent::new(parse_number(number, context)?)
}

fn parse_intensity(value: &str, context: &str) -> Result<Intensity> {
    let number = value
        .strip_suffix('%')
        .ok_or_else(|| make_error(format!("{context} must be a percentage")))?;
    validate_decimal(number, value)?;
    Intensity::new(parse_number(number, context)?)
}

fn parse_number(value: &str, context: &str) -> Result<f64> {
    let value = value
        .parse::<f64>()
        .map_err(|_| make_error(format!("invalid {context} value '{value}'")))?;
    finite(value, context)?;
    Ok(value)
}

fn finite(value: f64, context: &str) -> Result<()> {
    if !value.is_finite() {
        return invalid(format!("{context} must be finite"));
    }
    Ok(())
}

fn canonical_number(value: f64) -> String {
    if value == 0.0 {
        "0".to_string()
    } else {
        value.to_string()
    }
}

fn validate_decimal(value: &str, complete: &str) -> Result<()> {
    let value = value.strip_prefix('-').unwrap_or(value);
    if value.is_empty() {
        return invalid(format!("invalid decimal '{complete}'"));
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
        return invalid(format!("invalid decimal '{complete}'"));
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
        .map_err(|_| make_error(format!("invalid UTF-8 in gradient {context} name")))
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

    const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><office:styles>"#;
    const SUFFIX: &str = "</office:styles></office:document-styles>";

    #[test]
    fn parses_and_round_trips_all_gradient_kinds() {
        let xml = format!(
            r##"{PREFIX}<draw:gradient draw:name="legacy" draw:style="rectangular" draw:cx="50%" draw:cy="25%" draw:start-color="#ff0000" draw:end-color="#0000ff" draw:start-intensity="100%" draw:end-intensity="75%" draw:angle="30deg" draw:border="0%"><loext:gradient-stop svg:offset="0" loext:color-type="rgb" loext:color-value="#ff0000"/><loext:gradient-stop svg:offset="100%" loext:color-type="theme" loext:color-value="accent1"/></draw:gradient><svg:linearGradient draw:name="linear" svg:gradientUnits="objectBoundingBox" svg:gradientTransform="rotate(30)" svg:spreadMethod="reflect" svg:x1="0%" svg:y1="1cm" svg:x2="100%" svg:y2="2cm"><svg:stop svg:offset="0" svg:stop-color="#ffffff" svg:stop-opacity="1"/><svg:stop svg:offset="100%" svg:stop-color="#000000"/></svg:linearGradient><svg:radialGradient draw:name="radial" svg:cx="50%" svg:cy="50%" svg:r="5cm" svg:fx="45%" svg:fy="45%"/>{SUFFIX}"##
        );
        let gradients = parse_drawing_gradients(&xml).unwrap();
        assert_eq!(gradients.gradients.len(), 3);
        assert_eq!(gradients.get("legacy").unwrap().name(), Some("legacy"));
        let Definition::Legacy(legacy) = &gradients.gradients[0] else {
            panic!("expected legacy gradient");
        };
        assert_eq!(legacy.extension_stops.len(), 2);
        assert_eq!(legacy.start_color, Some(RgbColor::new(255, 0, 0)));

        let serialized = gradients.to_xml().unwrap();
        assert_eq!(parse_drawing_gradients(&serialized).unwrap(), gradients);
    }

    #[test]
    fn rejects_malformed_gradients() {
        for body in [
            r#"<draw:gradient draw:name="x"/>"#,
            r#"<draw:gradient draw:name="x" draw:style="unknown"/>"#,
            r##"<draw:gradient draw:name="x" draw:style="linear" draw:start-color="#zzz000"/>"##,
            r#"<draw:gradient draw:name="x" draw:style="linear" draw:start-intensity="101%"/>"#,
            r#"<svg:linearGradient draw:name="x"><svg:stop/></svg:linearGradient>"#,
            r#"<svg:linearGradient draw:name="x"><svg:stop svg:offset="0"></svg:stop></svg:linearGradient>"#,
            r#"<draw:gradient draw:name="x" draw:style="linear"/><svg:radialGradient draw:name="x"/>"#,
        ] {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(parse_drawing_gradients(&xml).is_err(), "accepted {body}");
        }
        let misplaced = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><draw:gradient draw:name="x" draw:style="linear"/></office:automatic-styles></office:document-styles>"#.to_string();
        assert!(parse_drawing_gradients(&misplaced).is_err());
    }

    #[test]
    fn parses_local_multicolor_gradient_fixture() {
        let xml = include_str!("../../../../../test-data/odf/drawing/multicolor-gradient.fodp");
        let gradients = parse_drawing_gradients(xml).unwrap();
        assert!(gradients.gradients.len() >= 6);
        let Definition::Legacy(first) = &gradients.gradients[0] else {
            panic!("local fixture should begin with a legacy gradient");
        };
        assert_eq!(first.extension_stops.len(), 2);
        assert!(
            !parse_drawing_gradients(&gradients.to_xml().unwrap())
                .unwrap()
                .gradients
                .is_empty()
        );
    }
}
