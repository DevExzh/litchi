//! Bounded XML codec and snapshot mutation for footnote separators.

use std::collections::HashMap;

use litchi_core::Result;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use super::model::{Adjustment, Length, LineStyle, Percent, Separator};
use super::{
    MAX_AGGREGATE_BYTES, MAX_DEPTH, MAX_SEPARATORS, MAX_VALUE_BYTES, MAX_XML_BYTES, STYLE_NS,
    invalid, make_error,
};

impl Separator {
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
    value: Separator,
}

type Attributes = HashMap<(Ns, String), String>;

/// Parse all typed page-layout footnote separators in one ODF XML part.
pub fn parse(xml: &str) -> Result<Vec<Separator>> {
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
                    let value = active
                        .take()
                        .ok_or_else(|| make_error("missing completed footnote separator"))?
                        .value;
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

pub(crate) fn parse_page_layout_property_footnote_separators(xml: &str) -> Result<Vec<Separator>> {
    let (wrapped, _, _) = crate::style::columns::scoped_property_xml(xml)?;
    parse(&wrapped)
}

#[allow(clippy::too_many_arguments)]
fn start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Ns,
    local: &str,
    stack: &mut [Frame],
    active: &mut Option<Active>,
    values: &mut Vec<Separator>,
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
) -> Result<Separator> {
    let mut values = attributes(reader, element, aggregate)?;
    let result = Separator {
        width: take(&mut values, "width").map(Length::new).transpose()?,
        relative_width: take(&mut values, "rel-width")
            .map(Percent::new)
            .transpose()?,
        color: take(&mut values, "color")
            .map(|value| parse_color(&value))
            .transpose()?,
        line_style: take(&mut values, "line-style")
            .map(|value| LineStyle::parse(&value))
            .transpose()?,
        adjustment: take(&mut values, "adjustment")
            .map(|value| Adjustment::parse(&value))
            .transpose()?,
        distance_before: take(&mut values, "distance-before-sep")
            .map(Length::new)
            .transpose()?,
        distance_after: take(&mut values, "distance-after-sep")
            .map(Length::new)
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
    separator: &Separator,
) -> Result<String> {
    separator.validate()?;
    let fragment = separator.to_xml_fragment()?;
    if let Some(properties) = &layout.properties {
        let existing = parse_page_layout_property_footnote_separators(&properties.xml)?;
        let new_properties = if existing.is_empty() {
            crate::style::columns::insert_before_end(
                &properties.xml,
                &fragment,
                "style:page-layout-properties",
            )?
        } else {
            replace_first(&properties.xml, &fragment)?
        };
        return crate::style::columns::self_contained_layout(&layout.xml.replacen(
            &properties.xml,
            &new_properties,
            1,
        ));
    }
    let properties = format!(
        r#"<style:page-layout-properties xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">{fragment}</style:page-layout-properties>"#,
    );
    crate::style::columns::self_contained_layout(&crate::style::columns::insert_before_end(
        &layout.xml,
        &properties,
        "style:page-layout",
    )?)
}

fn replace_first(xml: &str, replacement: &str) -> Result<String> {
    let (wrapped, prefix, suffix) = crate::style::columns::scoped_property_xml(xml)?;
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
                active = Some((start, 1));
            },
            Event::Empty(ref element)
                if active.is_none()
                    && selected
                    && element.local_name().as_ref() == b"footnote-sep" =>
            {
                return splice(&wrapped, start, end, replacement, prefix, suffix);
            },
            Event::Start(_) if active.is_some() => {
                active
                    .as_mut()
                    .ok_or_else(|| make_error("missing active footnote separator range"))?
                    .1 += 1;
            },
            Event::End(_) if active.is_some() => {
                let current = active
                    .as_mut()
                    .ok_or_else(|| make_error("missing active footnote separator range"))?;
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

fn push_value(values: &mut Vec<Separator>, value: Separator) -> Result<()> {
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
        u8::from_str_radix(&hex[0..2], 16)
            .map_err(|error| make_error(format!("invalid red color component: {error}")))?,
        u8::from_str_radix(&hex[2..4], 16)
            .map_err(|error| make_error(format!("invalid green color component: {error}")))?,
        u8::from_str_radix(&hex[4..6], 16)
            .map_err(|error| make_error(format!("invalid blue color component: {error}")))?,
    ))
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
