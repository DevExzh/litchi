//! ODF text line-numbering configuration.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use super::model::{Configuration, Format, NonNegativeLength, Position, Separator, validate_value};
use super::{
    MAX_DOCUMENT_XML_BYTES, MAX_VALUE_BYTES, MAX_XML_DEPTH, OFFICE_NAMESPACE, STYLE_NAMESPACE,
    TEXT_NAMESPACE, invalid, xml_error,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NamespaceKind {
    Office,
    Text,
    Style,
    Other,
}

impl Configuration {
    /// Serialize a namespace-complete line-numbering configuration element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(384);
        output.push_str("<text:linenumbering-configuration xmlns:text=\"");
        output.push_str(std::str::from_utf8(TEXT_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
        output.push('"');
        write_bool_attr(&mut output, "text:number-lines", self.number_lines);
        write_attr(
            &mut output,
            "style:num-format",
            self.number_format.as_ref().map(Format::as_str),
        );
        write_bool_attr(&mut output, "style:num-letter-sync", self.letter_sync);
        write_attr(&mut output, "text:style-name", self.style_name.as_deref());
        write_u64_attr(&mut output, "text:increment", self.increment);
        write_attr(
            &mut output,
            "text:number-position",
            self.number_position.map(Position::as_str),
        );
        write_attr(
            &mut output,
            "text:offset",
            self.offset.as_ref().map(NonNegativeLength::as_str),
        );
        write_bool_attr(
            &mut output,
            "text:count-empty-lines",
            self.count_empty_lines,
        );
        write_bool_attr(
            &mut output,
            "text:count-in-text-boxes",
            self.count_in_text_boxes,
        );
        write_bool_attr(&mut output, "text:restart-on-page", self.restart_on_page);
        let Some(separator) = &self.separator else {
            output.push_str("/>");
            return Ok(output);
        };
        output.push_str("><text:linenumbering-separator");
        write_u64_attr(&mut output, "text:increment", separator.increment);
        if separator.text.is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            escape_text(&mut output, &separator.text);
            output.push_str("</text:linenumbering-separator>");
        }
        output.push_str("</text:linenumbering-configuration>");
        Ok(output)
    }
}

/// Parse the optional line-numbering declaration from an ODF styles document.
pub fn parse(xml: &str) -> Result<Option<Configuration>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte line-numbering limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut styles_content_depth = None;
    let mut result = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                if styles_content_depth != Some(depth) {
                    return invalid(
                        "text:linenumbering-configuration must be a direct office:styles child",
                    );
                }
                if result.is_some() {
                    return invalid("ODF styles contain duplicate line-numbering configurations");
                }
                let mut configuration = parse_attributes(&reader, &element)?;
                configuration.separator = parse_configuration_body(&mut reader)?;
                configuration.validate()?;
                result = Some(configuration);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                if styles_content_depth != Some(depth) {
                    return invalid(
                        "text:linenumbering-configuration must be a direct office:styles child",
                    );
                }
                if result.is_some() {
                    return invalid("ODF styles contain duplicate line-numbering configurations");
                }
                let configuration = parse_attributes(&reader, &element)?;
                configuration.validate()?;
                result = Some(configuration);
            },
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth.replace(depth).is_some()
                {
                    return invalid("ODF XML contains nested office:styles elements");
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth == Some(depth)
                {
                    styles_content_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF line metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

enum StylesSite {
    Content(usize),
    Empty(XmlSpan, String),
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end].rfind('<').ok_or_else(|| {
        Error::InvalidFormat("invalid line-numbering XML event boundary".to_string())
    })
}

fn locate_configuration(xml: &str) -> Result<(Option<XmlSpan>, StylesSite)> {
    parse(xml)?;

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(NamespaceKind, Vec<u8>)> = Vec::new();
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut styles_site = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved);
        match event {
            Event::Start(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len() + 1;
                if namespace == NamespaceKind::Text
                    && local == b"linenumbering-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                {
                    open_target = Some((depth, start));
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                if namespace == NamespaceKind::Text
                    && local == b"linenumbering-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                {
                    target = Some(XmlSpan { start, end });
                }
                if namespace == NamespaceKind::Office && local == b"styles" {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Empty(
                        XmlSpan { start, end },
                        std::str::from_utf8(element.name().as_ref())
                            .map_err(|_| {
                                Error::InvalidFormat("invalid office:styles QName".to_string())
                            })?
                            .to_string(),
                    ));
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let depth = stack.len();
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, target_start) = open_target.take().expect("target depth checked");
                    target = Some(XmlSpan {
                        start: target_start,
                        end,
                    });
                }
                if matches!(stack.last(), Some((NamespaceKind::Office, local)) if local == b"styles")
                {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Content(start));
                }
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    Ok((
        target,
        styles_site
            .ok_or_else(|| Error::InvalidFormat("document has no office:styles".to_string()))?,
    ))
}

/// Insert or replace the document line-numbering declaration without rewriting
/// unrelated style XML.
pub(crate) fn set_xml(xml: &str, configuration: &Configuration) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_configuration(xml)?;
    let fragment = configuration.to_xml()?;
    if let Some(span) = target {
        return Ok(format!(
            "{}{}{}",
            &xml[..span.start],
            fragment,
            &xml[span.end..]
        ));
    }
    match site {
        StylesSite::Content(insertion) => Ok(format!(
            "{}{}{}",
            &xml[..insertion],
            fragment,
            &xml[insertion..]
        )),
        StylesSite::Empty(span, qname) => {
            let raw = &xml[span.start..span.end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| Error::InvalidFormat("invalid empty office:styles".to_string()))?;
            Ok(format!(
                "{}{}>{}</{}>{}",
                &xml[..span.start],
                &raw[..slash],
                fragment,
                qname,
                &xml[span.end..]
            ))
        },
    }
}

/// Remove the document line-numbering declaration without rewriting unrelated
/// style XML.
pub(crate) fn remove_xml(xml: &str) -> Result<String> {
    let (target, _) = locate_configuration(xml)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

fn parse_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Configuration> {
    let mut configuration = Configuration::default();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        if !seen.insert((namespace, local.as_ref().to_vec())) {
            return invalid("duplicate line-numbering configuration attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_value(&value, "line-numbering attribute", true)?;
        match (namespace, local.as_ref()) {
            (NamespaceKind::Text, b"number-lines") => {
                configuration.number_lines = Some(parse_bool(&value, "text:number-lines")?)
            },
            (NamespaceKind::Style, b"num-format") => {
                configuration.number_format = Some(Format::parse(value)?)
            },
            (NamespaceKind::Style, b"num-letter-sync") => {
                configuration.letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?)
            },
            (NamespaceKind::Text, b"style-name") => {
                validate_value(&value, "text:style-name", false)?;
                configuration.style_name = Some(value);
            },
            (NamespaceKind::Text, b"increment") => {
                configuration.increment = Some(parse_nonnegative_integer(&value, "text:increment")?)
            },
            (NamespaceKind::Text, b"number-position") => {
                configuration.number_position = Some(Position::parse(&value)?)
            },
            (NamespaceKind::Text, b"offset") => {
                configuration.offset = Some(NonNegativeLength::new(value)?)
            },
            (NamespaceKind::Text, b"count-empty-lines") => {
                configuration.count_empty_lines =
                    Some(parse_bool(&value, "text:count-empty-lines")?)
            },
            (NamespaceKind::Text, b"count-in-text-boxes") => {
                configuration.count_in_text_boxes =
                    Some(parse_bool(&value, "text:count-in-text-boxes")?)
            },
            (NamespaceKind::Text, b"restart-on-page") => {
                configuration.restart_on_page = Some(parse_bool(&value, "text:restart-on-page")?)
            },
            _ => return invalid("unsupported line-numbering configuration attribute"),
        }
    }
    Ok(configuration)
}

fn parse_configuration_body(reader: &mut NsReader<&[u8]>) -> Result<Option<Separator>> {
    let mut separator = None;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                if separator.is_some() {
                    return invalid("duplicate text:linenumbering-separator");
                }
                separator = Some(parse_separator(reader, &element)?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                if separator.is_some() {
                    return invalid("duplicate text:linenumbering-separator");
                }
                separator = Some(Separator {
                    increment: parse_separator_increment(reader, &element)?,
                    text: String::new(),
                });
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-configuration" =>
            {
                break;
            },
            Event::Text(text) => require_whitespace(
                &text.decode().map_err(xml_error)?,
                "linenumbering-configuration",
            )?,
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in line numbering"),
            Event::Eof => return invalid("unterminated text:linenumbering-configuration"),
            _ => return invalid("unsupported child in linenumbering-configuration"),
        }
        buffer.clear();
    }
    Ok(separator)
}

fn parse_separator(reader: &mut NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Separator> {
    let increment = parse_separator_increment(reader, element)?;
    let mut text = String::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Text(value) => {
                let value = value.decode().map_err(xml_error)?;
                append_separator_text(&mut text, &value)?;
            },
            Event::CData(value) => {
                let value = value.decode().map_err(xml_error)?;
                append_separator_text(&mut text, &value)?;
            },
            Event::GeneralRef(reference) => {
                if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
                    let mut encoded = [0u8; 4];
                    append_separator_text(&mut text, character.encode_utf8(&mut encoded))?;
                } else {
                    let name = reference.decode().map_err(xml_error)?;
                    let value = match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "apos" => "'",
                        "quot" => "\"",
                        _ => return invalid("unsupported entity in line-numbering separator"),
                    };
                    append_separator_text(&mut text, value)?;
                }
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"linenumbering-separator" =>
            {
                break;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated text:linenumbering-separator"),
            _ => return invalid("text:linenumbering-separator may contain text only"),
        }
        buffer.clear();
    }
    Ok(Separator { increment, text })
}

fn parse_separator_increment(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<u64>> {
    let mut increment = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(&namespace) != NamespaceKind::Text || local.as_ref() != b"increment" {
            return invalid("unsupported linenumbering-separator attribute");
        }
        if increment.is_some() {
            return invalid("duplicate linenumbering-separator increment");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?;
        increment = Some(parse_nonnegative_integer(
            &value,
            "linenumbering-separator increment",
        )?);
    }
    Ok(increment)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        _ => NamespaceKind::Other,
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("{name} must be true, false, 1, or 0")),
    }
}

fn parse_nonnegative_integer(value: &str, name: &str) -> Result<u64> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid(format!("{name} must be a nonnegative integer"));
    }
    value
        .parse::<u64>()
        .map_err(|_| Error::InvalidFormat(format!("{name} exceeds the u64 range")))
}

fn append_separator_text(output: &mut String, value: &str) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("separator text size overflow".to_string()))?;
    if size > MAX_VALUE_BYTES {
        return invalid(format!(
            "line-numbering separator exceeds the {MAX_VALUE_BYTES} byte limit"
        ));
    }
    output.push_str(value);
    Ok(())
}

fn require_whitespace(value: &str, context: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        invalid(format!("text:{context} cannot contain text"))
    }
}

fn write_attr(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_attribute(output, value);
    output.push('"');
}

fn write_bool_attr(output: &mut String, name: &str, value: Option<bool>) {
    write_attr(
        output,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn write_u64_attr(output: &mut String, name: &str, value: Option<u64>) {
    let Some(value) = value else { return };
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&value.to_string());
    output.push('"');
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}

fn escape_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}
