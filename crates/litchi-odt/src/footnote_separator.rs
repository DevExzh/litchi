//! Typed ODF `style:footnote-sep` page-layout properties.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};
use std::{collections::HashMap, fmt, str::FromStr};

const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_SEPARATORS: usize = 65_536;
const MAX_VALUE_BYTES: usize = 4_096;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// ODF `length` lexical value used by a footnote separator.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FootnoteSeparatorLength(String);

impl FootnoteSeparatorLength {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_measure(&value, false)?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
impl fmt::Display for FootnoteSeparatorLength {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}
impl FromStr for FootnoteSeparatorLength {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// ODF `percent` lexical value used by `style:rel-width`.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FootnoteSeparatorPercent(String);

impl FootnoteSeparatorPercent {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_measure(&value, true)?;
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
impl fmt::Display for FootnoteSeparatorPercent {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}
impl FromStr for FootnoteSeparatorPercent {
    type Err = Error;
    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FootnoteSeparatorLineStyle {
    None,
    Solid,
    Dotted,
    Dash,
    LongDash,
    DotDash,
    DotDotDash,
    Wave,
}
impl FootnoteSeparatorLineStyle {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "solid" => Ok(Self::Solid),
            "dotted" => Ok(Self::Dotted),
            "dash" => Ok(Self::Dash),
            "long-dash" => Ok(Self::LongDash),
            "dot-dash" => Ok(Self::DotDash),
            "dot-dot-dash" => Ok(Self::DotDotDash),
            "wave" => Ok(Self::Wave),
            _ => invalid(format!("invalid style:line-style '{value}'")),
        }
    }
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

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FootnoteSeparatorAdjustment {
    Left,
    Center,
    Right,
}
impl FootnoteSeparatorAdjustment {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            _ => invalid(format!("invalid style:adjustment '{value}'")),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
        }
    }
}

/// One optional footnote-separator rule in `style:page-layout-properties`.
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct StyleFootnoteSeparator {
    pub width: Option<FootnoteSeparatorLength>,
    pub relative_width: Option<FootnoteSeparatorPercent>,
    pub color: Option<(u8, u8, u8)>,
    pub line_style: Option<FootnoteSeparatorLineStyle>,
    pub adjustment: Option<FootnoteSeparatorAdjustment>,
    pub distance_before: Option<FootnoteSeparatorLength>,
    pub distance_after: Option<FootnoteSeparatorLength>,
}

impl StyleFootnoteSeparator {
    pub fn validate(&self) -> Result<()> {
        for value in [&self.width, &self.distance_before, &self.distance_after]
            .into_iter()
            .flatten()
        {
            validate_measure(value.as_str(), false)?;
        }
        if let Some(value) = &self.relative_width {
            validate_measure(value.as_str(), true)?;
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from(
            r#"<style:footnote-sep xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#,
        );
        if let Some(value) = &self.width {
            attr(&mut xml, "style:width", value.as_str());
        }
        if let Some(value) = &self.relative_width {
            attr(&mut xml, "style:rel-width", value.as_str());
        }
        if let Some((red, green, blue)) = self.color {
            attr(
                &mut xml,
                "style:color",
                &format!("#{red:02X}{green:02X}{blue:02X}"),
            );
        }
        if let Some(value) = self.line_style {
            attr(&mut xml, "style:line-style", value.as_str());
        }
        if let Some(value) = self.adjustment {
            attr(&mut xml, "style:adjustment", value.as_str());
        }
        if let Some(value) = &self.distance_before {
            attr(&mut xml, "style:distance-before-sep", value.as_str());
        }
        if let Some(value) = &self.distance_after {
            attr(&mut xml, "style:distance-after-sep", value.as_str());
        }
        xml.push_str("/>");
        Ok(xml)
    }

    pub(crate) fn to_page_layout_fragment(&self, name: &str) -> Result<String> {
        validate_style_name(name)?;
        Ok(format!(
            r#"<style:page-layout style:name="{}"><style:page-layout-properties>{}</style:page-layout-properties></style:page-layout>"#,
            escaped(name),
            self.to_xml_fragment()?
        ))
    }
}

impl crate::OpenDocumentPackage {
    pub fn style_footnote_separators(&self) -> Result<Vec<StyleFootnoteSeparator>> {
        let mut values =
            parse_style_footnote_separators(self.styles_xml()?.as_deref().unwrap_or_default())?;
        values.extend(parse_style_footnote_separators(&self.content_xml()?)?);
        if values.len() > MAX_SEPARATORS {
            return invalid("package exceeds 65536 style:footnote-sep values");
        }
        Ok(values)
    }
}

impl crate::FlatOpenDocument {
    pub fn style_footnote_separators(&self) -> Result<Vec<StyleFootnoteSeparator>> {
        parse_style_footnote_separators(self.xml())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum Ns {
    None,
    Style,
    Other,
}
#[derive(Clone)]
struct Frame {
    namespace: Ns,
    local: String,
    saw_separator: bool,
}
struct Active {
    depth: usize,
    value: StyleFootnoteSeparator,
}
type Attributes = HashMap<(Ns, String), String>;

/// Parse all typed page-layout footnote separators in one ODF XML part.
pub fn parse_style_footnote_separators(xml: &str) -> Result<Vec<StyleFootnoteSeparator>> {
    if !xml.contains("footnote-sep") {
        return Ok(Vec::new());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("footnote-separator XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<Active> = None;
    let mut values = Vec::new();
    let mut aggregate = 0usize;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid footnote-separator XML: {error}")))?;
        let namespace = ns(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref())?;
                spoof(namespace, &local)?;
                start(
                    &reader,
                    element,
                    namespace,
                    &local,
                    &mut stack,
                    &mut active,
                    &mut values,
                    &mut aggregate,
                    false,
                )?;
                stack.push(Frame {
                    namespace,
                    local,
                    saw_separator: false,
                });
                if stack.len() > MAX_DEPTH {
                    return invalid("footnote-separator XML exceeds 256 levels");
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref())?;
                spoof(namespace, &local)?;
                start(
                    &reader,
                    element,
                    namespace,
                    &local,
                    &mut stack,
                    &mut active,
                    &mut values,
                    &mut aggregate,
                    true,
                )?;
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| make_error("footnote-separator XML depth underflow"))?;
                if active
                    .as_ref()
                    .is_some_and(|current| current.depth == stack.len())
                {
                    let value = active.take().expect("active separator checked").value;
                    value.validate()?;
                    push_value(&mut values, value)?;
                }
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("style:footnote-sep must be empty");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "DTDs and processing instructions are prohibited in footnote-separator XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated footnote-separator XML");
    }
    Ok(values)
}

pub(crate) fn parse_page_layout_property_footnote_separators(
    xml: &str,
) -> Result<Vec<StyleFootnoteSeparator>> {
    let (wrapped, _, _) = crate::style_columns::scoped_property_xml(xml)?;
    parse_style_footnote_separators(&wrapped)
}

#[allow(clippy::too_many_arguments)]
fn start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Ns,
    local: &str,
    stack: &mut [Frame],
    active: &mut Option<Active>,
    values: &mut Vec<StyleFootnoteSeparator>,
    aggregate: &mut usize,
    empty: bool,
) -> Result<()> {
    if active.is_some() {
        return invalid("style:footnote-sep cannot contain child elements");
    }
    if namespace != Ns::Style || local != "footnote-sep" {
        return Ok(());
    }
    let parent = stack
        .last_mut()
        .ok_or_else(|| make_error("style:footnote-sep has no parent"))?;
    if parent.namespace != Ns::Style || parent.local != "page-layout-properties" {
        return invalid("style:footnote-sep must be a direct style:page-layout-properties child");
    }
    if parent.saw_separator {
        return invalid("page-layout-properties has multiple style:footnote-sep children");
    }
    parent.saw_separator = true;
    let value = parse_separator(reader, element, aggregate)?;
    if empty {
        push_value(values, value)?;
    } else {
        *active = Some(Active {
            depth: stack.len(),
            value,
        });
    }
    Ok(())
}

fn parse_separator(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<StyleFootnoteSeparator> {
    let mut values = attributes(reader, element, aggregate)?;
    let result = StyleFootnoteSeparator {
        width: take(&mut values, "width")
            .map(FootnoteSeparatorLength::new)
            .transpose()?,
        relative_width: take(&mut values, "rel-width")
            .map(FootnoteSeparatorPercent::new)
            .transpose()?,
        color: take(&mut values, "color")
            .map(|value| parse_color(&value))
            .transpose()?,
        line_style: take(&mut values, "line-style")
            .map(|value| FootnoteSeparatorLineStyle::parse(&value))
            .transpose()?,
        adjustment: take(&mut values, "adjustment")
            .map(|value| FootnoteSeparatorAdjustment::parse(&value))
            .transpose()?,
        distance_before: take(&mut values, "distance-before-sep")
            .map(FootnoteSeparatorLength::new)
            .transpose()?,
        distance_after: take(&mut values, "distance-after-sep")
            .map(FootnoteSeparatorLength::new)
            .transpose()?,
    };
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported style:footnote-sep attribute {namespace:?}:{local}"
        ));
    }
    result.validate()?;
    Ok(result)
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    if element.attributes().count() > 32 {
        return invalid("style:footnote-sep exceeds 32 attributes");
    }
    let mut result = Attributes::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            make_error(format!("invalid footnote-separator attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = ns(&resolved)?;
        let local = decode(local.as_ref())?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid footnote-separator attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("footnote-separator attribute exceeds 4096 bytes");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("footnote-separator size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("footnote-separator values exceed 16 MiB");
        }
        if result.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded footnote-separator attribute");
        }
    }
    Ok(result)
}

pub(crate) fn replace_page_layout_footnote_separator(
    layout: &crate::PageLayout,
    separator: &StyleFootnoteSeparator,
) -> Result<String> {
    separator.validate()?;
    let fragment = separator.to_xml_fragment()?;
    if let Some(properties) = &layout.properties {
        let existing = parse_page_layout_property_footnote_separators(&properties.xml)?;
        let new_properties = if existing.is_empty() {
            crate::style_columns::insert_before_end(
                &properties.xml,
                &fragment,
                "style:page-layout-properties",
            )?
        } else {
            replace_first(&properties.xml, &fragment)?
        };
        return crate::style_columns::self_contained_layout(&layout.xml.replacen(
            &properties.xml,
            &new_properties,
            1,
        ));
    }
    let properties = format!(
        r#"<style:page-layout-properties xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">{fragment}</style:page-layout-properties>"#,
    );
    crate::style_columns::self_contained_layout(&crate::style_columns::insert_before_end(
        &layout.xml,
        &properties,
        "style:page-layout",
    )?)
}

fn replace_first(xml: &str, replacement: &str) -> Result<String> {
    let (wrapped, prefix, suffix) = crate::style_columns::scoped_property_xml(xml)?;
    let mut reader = NsReader::from_str(&wrapped);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize)> = None;
    loop {
        let start = reader.buffer_position() as usize;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                make_error(format!(
                    "invalid footnote-separator replacement XML: {error}"
                ))
            })?;
        let selected = ns(&resolved)? == Ns::Style;
        let event = event.into_owned();
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element)
                if active.is_none()
                    && selected
                    && element.local_name().as_ref() == b"footnote-sep" =>
            {
                active = Some((start, 1))
            },
            Event::Empty(ref element)
                if active.is_none()
                    && selected
                    && element.local_name().as_ref() == b"footnote-sep" =>
            {
                return splice(&wrapped, start, end, replacement, prefix, suffix);
            },
            Event::Start(_) if active.is_some() => active.as_mut().expect("active").1 += 1,
            Event::End(_) if active.is_some() => {
                let current = active.as_mut().expect("active");
                current.1 -= 1;
                if current.1 == 0 {
                    return splice(&wrapped, current.0, end, replacement, prefix, suffix);
                }
            },
            Event::Eof => {
                return invalid("page-layout properties have no style:footnote-sep to replace");
            },
            _ => {},
        }
        buffer.clear();
    }
}

fn push_value(
    values: &mut Vec<StyleFootnoteSeparator>,
    value: StyleFootnoteSeparator,
) -> Result<()> {
    if values.len() >= MAX_SEPARATORS {
        return invalid("XML exceeds 65536 style:footnote-sep values");
    }
    values.push(value);
    Ok(())
}
fn ns(value: &ResolveResult<'_>) -> Result<Ns> {
    match value {
        ResolveResult::Unbound => Ok(Ns::None),
        ResolveResult::Bound(value) => Ok(if value.as_ref() == STYLE_NS {
            Ns::Style
        } else {
            Ns::Other
        }),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}
fn spoof(namespace: Ns, local: &str) -> Result<()> {
    if local == "footnote-sep" && namespace != Ns::Style {
        return invalid("footnote-sep uses the wrong namespace");
    }
    Ok(())
}
fn take(values: &mut Attributes, local: &str) -> Option<String> {
    values.remove(&(Ns::Style, local.to_owned()))
}
fn parse_color(value: &str) -> Result<(u8, u8, u8)> {
    let hex = value
        .strip_prefix('#')
        .filter(|hex| hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .ok_or_else(|| make_error("invalid style:color"))?;
    Ok((
        u8::from_str_radix(&hex[0..2], 16).expect("hex"),
        u8::from_str_radix(&hex[2..4], 16).expect("hex"),
        u8::from_str_radix(&hex[4..6], 16).expect("hex"),
    ))
}
fn validate_measure(value: &str, percent: bool) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES {
        return invalid("invalid ODF measure");
    }
    let number = if percent {
        value
            .strip_suffix('%')
            .ok_or_else(|| make_error(format!("invalid ODF percent '{value}'")))?
    } else {
        let unit = ["cm", "mm", "in", "pt", "pc", "px"]
            .into_iter()
            .find(|unit| value.ends_with(unit))
            .ok_or_else(|| make_error(format!("invalid ODF length '{value}'")))?;
        &value[..value.len() - unit.len()]
    };
    let number = number.strip_prefix('-').unwrap_or(number);
    let mut parts = number.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    let valid = parts.next().is_none()
        && (whole.bytes().all(|byte| byte.is_ascii_digit()))
        && match fraction {
            Some(fraction) => !whole.is_empty() || !fraction.is_empty(),
            None => !whole.is_empty(),
        }
        && fraction.is_none_or(|fraction| fraction.bytes().all(|byte| byte.is_ascii_digit()));
    if !valid {
        return invalid(if percent {
            format!("invalid ODF percent '{value}'")
        } else {
            format!("invalid ODF length '{value}'")
        });
    }
    Ok(())
}
fn validate_style_name(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES || value.chars().any(char::is_control) {
        return invalid("invalid style name");
    }
    Ok(())
}
fn attr(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    escape(xml, value);
    xml.push('"');
}
fn escaped(value: &str) -> String {
    let mut output = String::new();
    escape(&mut output, value);
    output
}
fn escape(xml: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}
fn splice(
    xml: &str,
    start: usize,
    end: usize,
    replacement: &str,
    prefix: usize,
    suffix: usize,
) -> Result<String> {
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output[prefix..output.len() - suffix].to_owned())
}
fn decode(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| make_error("non-UTF-8 footnote-separator name"))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}
fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
