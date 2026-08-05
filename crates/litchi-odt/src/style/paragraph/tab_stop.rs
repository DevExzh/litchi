//! Typed, inert ODF paragraph-style tab stops.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::{fmt, str::FromStr};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 4_096;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;
pub const MAX_STOPS: usize = 64;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Position(String);

impl Position {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_length(&value, false, "style:position")?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for Position {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for Position {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Type {
    #[default]
    Left,
    Center,
    Right,
    Character(char),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum LeaderType {
    None,
    Single,
    Double,
}

impl LeaderType {
    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Single => "single",
            Self::Double => "double",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum LeaderStyle {
    None,
    Solid,
    Dotted,
    Dash,
    LongDash,
    DotDash,
    DotDotDash,
    Wave,
}

impl LeaderStyle {
    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::Dotted => "dotted",
            Self::Dash => "dash",
            Self::LongDash => "long-dash",
            Self::DotDash => "dot-dash",
            Self::DotDotDash => "dot-dot-dash",
            Self::Wave => "wave",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct LeaderWidth(String);

impl LeaderWidth {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        let keyword = matches!(
            value.as_str(),
            "auto" | "normal" | "bold" | "thin" | "medium" | "thick"
        );
        let integer = value.parse::<u64>().is_ok_and(|number| number > 0);
        let percent = value.strip_suffix('%').is_some_and(valid_positive_decimal);
        if !keyword
            && !integer
            && !percent
            && validate_length(&value, true, "style:leader-width").is_err()
        {
            return invalid(format!("invalid style:leader-width '{value}'"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl FromStr for LeaderWidth {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum LeaderColor {
    FontColor,
    Rgb(u8, u8, u8),
}

impl LeaderColor {
    fn parse(value: &str) -> Result<Self> {
        if value == "font-color" {
            return Ok(Self::FontColor);
        }
        let hex = value
            .strip_prefix('#')
            .filter(|hex| hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()))
            .ok_or_else(|| error(format!("invalid style:leader-color '{value}'")))?;
        Ok(Self::Rgb(
            u8::from_str_radix(&hex[0..2], 16).expect("validated hex"),
            u8::from_str_radix(&hex[2..4], 16).expect("validated hex"),
            u8::from_str_radix(&hex[4..6], 16).expect("validated hex"),
        ))
    }
    fn lexical(self) -> String {
        match self {
            Self::FontColor => "font-color".into(),
            Self::Rgb(r, g, b) => format!("#{r:02X}{g:02X}{b:02X}"),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Stop {
    pub position: Position,
    pub tab_type: Type,
    pub leader_type: Option<LeaderType>,
    pub leader_style: Option<LeaderStyle>,
    pub leader_width: Option<LeaderWidth>,
    pub leader_color: Option<LeaderColor>,
    pub leader_text: Option<char>,
    pub leader_text_style: Option<String>,
}

impl Stop {
    pub fn new(position: Position) -> Self {
        Self {
            position,
            tab_type: Type::Left,
            leader_type: None,
            leader_style: None,
            leader_width: None,
            leader_color: None,
            leader_text: None,
            leader_text_style: None,
        }
    }
    pub fn validate(&self) -> Result<()> {
        validate_length(self.position.as_str(), false, "style:position")?;
        if let Some(value) = &self.leader_text_style {
            validate_text(value, "style:leader-text-style")?;
        }
        Ok(())
    }
}

/// Explicitly present `style:tab-stops`; an empty collection clears inheritance.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct Stops(Vec<Stop>);

impl Stops {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn try_from_vec(stops: Vec<Stop>) -> Result<Self> {
        let result = Self(stops);
        result.validate()?;
        Ok(result)
    }
    pub fn push(&mut self, stop: Stop) -> Result<()> {
        if self.0.len() >= MAX_STOPS {
            return invalid(format!("paragraph style exceeds {MAX_STOPS} tab stops"));
        }
        stop.validate()?;
        self.0.push(stop);
        Ok(())
    }
    pub fn as_slice(&self) -> &[Stop] {
        &self.0
    }
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Stop> {
        self.0.iter()
    }
    pub fn len(&self) -> usize {
        self.0.len()
    }
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
    pub fn validate(&self) -> Result<()> {
        if self.0.len() > MAX_STOPS {
            return invalid(format!("paragraph style exceeds {MAX_STOPS} tab stops"));
        }
        self.0.iter().try_for_each(Stop::validate)
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from("<style:tab-stops>");
        for stop in &self.0 {
            write_stop(&mut xml, stop);
        }
        xml.push_str("</style:tab-stops>");
        Ok(xml)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    /// `None` means inherit; `Some(empty)` explicitly clears inherited stops.
    pub tab_stops: Option<Stops>,
}

impl Style {
    pub fn named(name: impl Into<String>, tab_stops: Option<Stops>) -> Result<Self> {
        let name = name.into();
        validate_text(&name, "style:name")?;
        Ok(Self {
            name: Some(name),
            parent_style_name: None,
            is_default_style: false,
            tab_stops,
        })
    }
    pub fn default_style(tab_stops: Option<Stops>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            tab_stops,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(name), false) => validate_text(name, "style:name")?,
            (None, true) => {},
            _ => return invalid("named/default paragraph style identity is inconsistent"),
        }
        if let Some(parent) = &self.parent_style_name {
            validate_text(parent, "style:parent-style-name")?;
        }
        if let Some(stops) = &self.tab_stops {
            stops.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(
            "<style:{tag} xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" style:family=\"paragraph\""
        );
        if let Some(name) = &self.name {
            attr(&mut xml, "style:name", name);
        }
        if let Some(parent) = &self.parent_style_name {
            attr(&mut xml, "style:parent-style-name", parent);
        }
        if let Some(stops) = &self.tab_stops {
            xml.push_str("><style:paragraph-properties>");
            xml.push_str(&stops.to_xml_fragment()?);
            xml.push_str("</style:paragraph-properties></style:");
            xml.push_str(tag);
            xml.push('>');
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}

impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&Style> {
        self.styles.iter().find(|style| style.is_default_style)
    }
    pub fn resolved_tab_stops(&self, name: &str) -> Result<Option<&Stops>> {
        let mut current = self.get(name);
        let mut seen = HashSet::new();
        while let Some(style) = current {
            let identity = style.name.as_deref().unwrap_or("");
            if !seen.insert(identity) {
                return invalid("paragraph style parent cycle");
            }
            if let Some(stops) = &style.tab_stops {
                return Ok(Some(stops));
            }
            current = style
                .parent_style_name
                .as_deref()
                .and_then(|parent| self.get(parent));
        }
        Ok(self
            .default_style()
            .and_then(|style| style.tab_stops.as_ref()))
    }
}

impl crate::OpenDocumentPackage {
    pub fn paragraph_style_tab_stops(&self) -> Result<Styles> {
        parse(self.styles_xml()?.as_deref().unwrap_or_default())
    }
}

impl crate::FlatOpenDocument {
    pub fn paragraph_style_tab_stops(&self) -> Result<Styles> {
        parse(self.xml())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Ns {
    None,
    Office,
    Style,
    Other,
}
#[derive(Clone)]
struct Frame {
    ns: Ns,
    local: String,
}
struct Active {
    depth: usize,
    value: Style,
    props_depth: Option<usize>,
    saw_props: bool,
    stops_depth: Option<usize>,
    saw_stops: bool,
    stop_depth: Option<usize>,
}
type Attributes = HashMap<(Ns, String), String>;

pub fn parse(xml: &str) -> Result<Styles> {
    if !xml.contains("tab-stop") {
        return Ok(Styles::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("paragraph style XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let (mut buffer, mut stack, mut active, mut result, mut aggregate) = (
        Vec::new(),
        Vec::<Frame>::new(),
        None,
        Styles::default(),
        0usize,
    );
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|err| error(format!("invalid paragraph tab-stop XML: {err}")))?;
        let namespace = ns(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                spoof(namespace, &local)?;
                start(
                    &reader,
                    element,
                    namespace,
                    &local,
                    &stack,
                    &mut active,
                    &mut result,
                    &mut aggregate,
                    false,
                )?;
                stack.push(Frame {
                    ns: namespace,
                    local,
                });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("paragraph style XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                spoof(namespace, &local)?;
                start(
                    &reader,
                    element,
                    namespace,
                    &local,
                    &stack,
                    &mut active,
                    &mut result,
                    &mut aggregate,
                    true,
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| error("paragraph style XML depth underflow"))?;
                if let Some(style) = &mut active {
                    if style.stop_depth == Some(stack.len()) {
                        if frame.ns != Ns::Style || frame.local != "tab-stop" {
                            return invalid("unexpected tab-stop end");
                        }
                        style.stop_depth = None;
                    }
                    if style.stops_depth == Some(stack.len()) {
                        style.stops_depth = None;
                    }
                    if style.props_depth == Some(stack.len()) {
                        style.props_depth = None;
                    }
                    if style.depth == stack.len() {
                        let style = active.take().expect("active checked").value;
                        push_style(&mut result, style)?;
                    }
                }
            },
            Event::Text(ref text)
                if active.as_ref().is_some_and(|s: &Active| {
                    s.stop_depth.is_some() || s.stops_depth.is_some()
                }) =>
            {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|err| error(format!("invalid tab-stop text: {err}")))?;
                if !value.chars().all(char::is_whitespace) {
                    return invalid("tab-stop elements cannot contain text");
                }
            },
            Event::CData(_) | Event::GeneralRef(_)
                if active.as_ref().is_some_and(|s: &Active| {
                    s.stop_depth.is_some() || s.stops_depth.is_some()
                }) =>
            {
                return invalid("tab-stop elements cannot contain character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "DTDs and processing instructions are prohibited in paragraph style XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated paragraph style XML");
    }
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Ns,
    local: &str,
    stack: &[Frame],
    active: &mut Option<Active>,
    result: &mut Styles,
    aggregate: &mut usize,
    empty: bool,
) -> Result<()> {
    if active.is_none() && namespace == Ns::Style && matches!(local, "style" | "default-style") {
        if !stack.last().is_some_and(|f| {
            f.ns == Ns::Office && matches!(f.local.as_str(), "styles" | "automatic-styles")
        }) {
            return Ok(());
        }
        let mut attrs = attributes(reader, element, aggregate)?;
        if take(&mut attrs, "family").as_deref() != Some("paragraph") {
            return Ok(());
        }
        let is_default = local == "default-style";
        let name = take(&mut attrs, "name");
        if is_default == name.is_some() {
            return invalid("invalid named/default paragraph style identity");
        }
        let value = Style {
            name,
            parent_style_name: take(&mut attrs, "parent-style-name"),
            is_default_style: is_default,
            tab_stops: None,
        };
        value.validate()?;
        if empty {
            push_style(result, value)?;
        } else {
            *active = Some(Active {
                depth: stack.len(),
                value,
                props_depth: None,
                saw_props: false,
                stops_depth: None,
                saw_stops: false,
                stop_depth: None,
            });
        }
        return Ok(());
    }
    let Some(style) = active else {
        return Ok(());
    };
    if style.stop_depth.is_some() {
        return invalid("style:tab-stop must be empty");
    }
    if namespace == Ns::Style && local == "paragraph-properties" && stack.len() == style.depth + 1 {
        if style.saw_props {
            return invalid("multiple style:paragraph-properties elements");
        }
        style.saw_props = true;
        if !empty {
            style.props_depth = Some(stack.len());
        }
    } else if namespace == Ns::Style && local == "tab-stops" {
        if style.props_depth != stack.len().checked_sub(1) {
            return invalid("style:tab-stops has the wrong parent");
        }
        if style.saw_stops {
            return invalid("multiple style:tab-stops elements");
        }
        reject(&attributes(reader, element, aggregate)?, "style:tab-stops")?;
        style.saw_stops = true;
        style.value.tab_stops = Some(Stops::new());
        if !empty {
            style.stops_depth = Some(stack.len());
        }
    } else if namespace == Ns::Style && local == "tab-stop" {
        if style.stops_depth != stack.len().checked_sub(1) {
            return invalid("style:tab-stop has the wrong parent");
        }
        let stop = parse_stop(reader, element, aggregate)?;
        style
            .value
            .tab_stops
            .as_mut()
            .expect("parent checked")
            .push(stop)?;
        if !empty {
            style.stop_depth = Some(stack.len());
        }
    }
    Ok(())
}

fn parse_stop(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Stop> {
    let mut attrs = attributes(reader, element, aggregate)?;
    let position = Position::new(
        take(&mut attrs, "position").ok_or_else(|| error("missing style:position"))?,
    )?;
    let kind = take(&mut attrs, "type").unwrap_or_else(|| "left".into());
    let character = take(&mut attrs, "char");
    let tab_type = match kind.as_str() {
        "left" if character.is_none() => Type::Left,
        "center" if character.is_none() => Type::Center,
        "right" if character.is_none() => Type::Right,
        "char" => Type::Character(one(character, "style:char")?),
        _ => return invalid(format!("invalid style:type '{kind}'")),
    };
    let leader_type = take(&mut attrs, "leader-type")
        .map(|v| match v.as_str() {
            "none" => Ok(LeaderType::None),
            "single" => Ok(LeaderType::Single),
            "double" => Ok(LeaderType::Double),
            _ => invalid(format!("invalid style:leader-type '{v}'")),
        })
        .transpose()?;
    let leader_style = take(&mut attrs, "leader-style")
        .map(|v| match v.as_str() {
            "none" => Ok(LeaderStyle::None),
            "solid" => Ok(LeaderStyle::Solid),
            "dotted" => Ok(LeaderStyle::Dotted),
            "dash" => Ok(LeaderStyle::Dash),
            "long-dash" => Ok(LeaderStyle::LongDash),
            "dot-dash" => Ok(LeaderStyle::DotDash),
            "dot-dot-dash" => Ok(LeaderStyle::DotDotDash),
            "wave" => Ok(LeaderStyle::Wave),
            _ => invalid(format!("invalid style:leader-style '{v}'")),
        })
        .transpose()?;
    let leader_width = take(&mut attrs, "leader-width")
        .map(LeaderWidth::new)
        .transpose()?;
    let leader_color = take(&mut attrs, "leader-color")
        .map(|v| LeaderColor::parse(&v))
        .transpose()?;
    let leader_text = take(&mut attrs, "leader-text")
        .map(|v| one(Some(v), "style:leader-text"))
        .transpose()?;
    let leader_text_style = take(&mut attrs, "leader-text-style");
    reject(&attrs, "style:tab-stop")?;
    Ok(Stop {
        position,
        tab_type,
        leader_type,
        leader_style,
        leader_width,
        leader_color,
        leader_text,
        leader_text_style,
    })
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|err| error(format!("invalid tab-stop attribute: {err}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(&resolved)?, decode(local.as_ref(), "attribute name")?);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|err| error(format!("invalid tab-stop attribute: {err}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("tab-stop attribute exceeds 4096 bytes");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| error("tab-stop size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("paragraph tab-stop values exceed 16 MiB");
        }
        if result.insert(key, value).is_some() {
            return invalid("duplicate expanded tab-stop attribute");
        }
    }
    Ok(result)
}

fn push_style(result: &mut Styles, style: Style) -> Result<()> {
    if result.styles.len() >= MAX_STYLES {
        return invalid(format!("paragraph styles exceed {MAX_STYLES} entries"));
    }
    if result
        .styles
        .iter()
        .any(|item| item.is_default_style == style.is_default_style && item.name == style.name)
    {
        return invalid("duplicate paragraph style identity");
    }
    result.styles.push(style);
    Ok(())
}

fn ns(resolved: &ResolveResult<'_>) -> Result<Ns> {
    match resolved {
        ResolveResult::Unbound => Ok(Ns::None),
        ResolveResult::Bound(value) => {
            let value: &[u8] = value.as_ref();
            Ok(if value == OFFICE_NS {
                Ns::Office
            } else if value == STYLE_NS {
                Ns::Style
            } else {
                Ns::Other
            })
        },
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn spoof(namespace: Ns, local: &str) -> Result<()> {
    if matches!(local, "paragraph-properties" | "tab-stops" | "tab-stop") && namespace != Ns::Style
    {
        return invalid(format!("{local} uses the wrong namespace"));
    }
    Ok(())
}

fn take(attrs: &mut Attributes, local: &str) -> Option<String> {
    attrs.remove(&(Ns::Style, local.into()))
}
fn reject(attrs: &Attributes, element: &str) -> Result<()> {
    if let Some(((namespace, local), _)) = attrs.iter().next() {
        return invalid(format!(
            "unsupported {element} attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}
fn one(value: Option<String>, name: &str) -> Result<char> {
    let value = value.ok_or_else(|| error(format!("missing required {name}")))?;
    let mut chars = value.chars();
    let first = chars
        .next()
        .ok_or_else(|| error(format!("{name} must contain one character")))?;
    if chars.next().is_some() {
        return invalid(format!("{name} must contain one character"));
    }
    Ok(first)
}

fn validate_length(value: &str, positive: bool, name: &str) -> Result<()> {
    let unit = ["cm", "mm", "in", "pt", "pc", "px"]
        .into_iter()
        .find(|unit| value.ends_with(unit))
        .ok_or_else(|| error(format!("invalid {name} '{value}'")))?;
    let number = &value[..value.len() - unit.len()];
    let unsigned = number
        .strip_prefix('+')
        .or_else(|| number.strip_prefix('-'))
        .unwrap_or(number);
    if value.len() > MAX_VALUE_BYTES
        || !decimal(unsigned)
        || (positive && (number.starts_with('-') || zero(unsigned)))
    {
        return invalid(format!("invalid {name} '{value}'"));
    }
    Ok(())
}
fn decimal(value: &str) -> bool {
    let mut parts = value.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    parts.next().is_none()
        && whole.bytes().all(|b| b.is_ascii_digit())
        && fraction.map_or(!whole.is_empty(), |f| {
            !f.is_empty() && f.bytes().all(|b| b.is_ascii_digit())
        })
}
fn valid_positive_decimal(value: &str) -> bool {
    decimal(value) && !zero(value)
}
fn zero(value: &str) -> bool {
    value.bytes().all(|b| matches!(b, b'0' | b'.'))
}
fn validate_text(value: &str, name: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES || value.chars().any(|c| c.is_control()) {
        return invalid(format!("invalid {name}"));
    }
    Ok(())
}

fn write_stop(xml: &mut String, stop: &Stop) {
    xml.push_str("<style:tab-stop");
    if let Some(value) = stop.leader_color {
        attr(xml, "style:leader-color", &value.lexical());
    }
    if let Some(value) = stop.leader_style {
        attr(xml, "style:leader-style", value.as_str());
    }
    if let Some(value) = stop.leader_text {
        attr(xml, "style:leader-text", &value.to_string());
    }
    if let Some(value) = &stop.leader_text_style {
        attr(xml, "style:leader-text-style", value);
    }
    if let Some(value) = stop.leader_type {
        attr(xml, "style:leader-type", value.as_str());
    }
    if let Some(value) = &stop.leader_width {
        attr(xml, "style:leader-width", value.as_str());
    }
    attr(xml, "style:position", stop.position.as_str());
    match stop.tab_type {
        Type::Left => {},
        Type::Center => attr(xml, "style:type", "center"),
        Type::Right => attr(xml, "style:type", "right"),
        Type::Character(c) => {
            attr(xml, "style:char", &c.to_string());
            attr(xml, "style:type", "char");
        },
    }
    xml.push_str("/>");
}
fn attr(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    escape(xml, value);
    xml.push('"');
}
fn escape(xml: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(c),
        }
    }
}
fn decode(value: &[u8], what: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| error(format!("invalid UTF-8 in {what}")))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(error(message))
}
fn error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
