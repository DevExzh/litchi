//! Bounded XML codec and lossless mutation for ODF note configurations.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

use super::model::{
    Class, Configuration, Configurations, NumberingScope, Position, validate_value,
};
use super::{
    MAX_DOCUMENT_XML_BYTES, MAX_VALUE_BYTES, MAX_XML_DEPTH, OFFICE_NAMESPACE, STYLE_NAMESPACE,
    TEXT_NAMESPACE, invalid, xml_error,
};
use crate::line_numbering::Format;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NamespaceKind {
    Office,
    Text,
    Style,
    Other,
}

impl Configuration {
    /// Serialize a namespace-complete notes configuration element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(512);
        output.push_str("<text:notes-configuration xmlns:text=\"");
        output.push_str(std::str::from_utf8(TEXT_NAMESPACE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE_NAMESPACE).expect("namespace is UTF-8"));
        output.push('"');
        write_attr(
            &mut output,
            "text:note-class",
            Some(self.note_class.as_str()),
        );
        write_attr(
            &mut output,
            "text:citation-style-name",
            self.citation_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:citation-body-style-name",
            self.citation_body_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:default-style-name",
            self.default_style_name.as_deref(),
        );
        write_attr(
            &mut output,
            "text:master-page-name",
            self.master_page_name.as_deref(),
        );
        write_u64_attr(&mut output, "text:start-value", self.start_value);
        write_attr(
            &mut output,
            "style:num-prefix",
            self.number_prefix.as_deref(),
        );
        write_attr(
            &mut output,
            "style:num-suffix",
            self.number_suffix.as_deref(),
        );
        write_attr(
            &mut output,
            "style:num-format",
            self.number_format.as_ref().map(Format::as_str),
        );
        write_bool_attr(&mut output, "style:num-letter-sync", self.letter_sync);
        write_attr(
            &mut output,
            "text:start-numbering-at",
            self.start_numbering_at.map(NumberingScope::as_str),
        );
        write_attr(
            &mut output,
            "text:footnotes-position",
            self.footnotes_position.map(Position::as_str),
        );
        if self.continuation_notice_forward.is_none() && self.continuation_notice_backward.is_none()
        {
            output.push_str("/>");
            return Ok(output);
        }
        output.push('>');
        write_notice(
            &mut output,
            "text:note-continuation-notice-forward",
            self.continuation_notice_forward.as_deref(),
        );
        write_notice(
            &mut output,
            "text:note-continuation-notice-backward",
            self.continuation_notice_backward.as_deref(),
        );
        output.push_str("</text:notes-configuration>");
        Ok(output)
    }
}

impl Configurations {
    /// Serialize both configurations in canonical footnote/endnote order.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::new();
        if let Some(configuration) = &self.footnote {
            output.push_str(&configuration.to_xml()?);
        }
        if let Some(configuration) = &self.endnote {
            output.push_str(&configuration.to_xml()?);
        }
        Ok(output)
    }
}

/// Parse footnote/endnote configurations from an ODF styles document.
pub fn parse(xml: &str) -> Result<Configurations> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte notes limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut styles_content_depth = None;
    let mut section_properties_depth = None;
    let mut result = Configurations::default();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                let collect = if styles_content_depth == Some(depth) {
                    true
                } else if section_properties_depth == Some(depth) {
                    false
                } else {
                    require_styles_scope(styles_content_depth, depth)?;
                    unreachable!("scope validation always fails outside supported parents")
                };
                let mut configuration = parse_attributes(&reader, &element)?;
                parse_notices(&mut reader, &mut configuration)?;
                configuration.validate()?;
                if collect {
                    insert_configuration(&mut result, configuration)?;
                }
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                let collect = if styles_content_depth == Some(depth) {
                    true
                } else if section_properties_depth == Some(depth) {
                    false
                } else {
                    require_styles_scope(styles_content_depth, depth)?;
                    unreachable!("scope validation always fails outside supported parents")
                };
                let configuration = parse_attributes(&reader, &element)?;
                configuration.validate()?;
                if collect {
                    insert_configuration(&mut result, configuration)?;
                }
            },
            Event::Start(element) if element.local_name().as_ref() == b"notes-configuration" => {
                return invalid("notes-configuration uses the wrong namespace");
            },
            Event::Empty(element) if element.local_name().as_ref() == b"notes-configuration" => {
                return invalid("notes-configuration uses the wrong namespace");
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
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"section-properties"
                    && section_properties_depth.replace(depth).is_some()
                {
                    return invalid("ODF XML contains nested style:section-properties elements");
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"styles"
                    && styles_content_depth == Some(depth)
                {
                    styles_content_depth = None;
                }
                if namespace == NamespaceKind::Style
                    && element.local_name().as_ref() == b"section-properties"
                    && section_properties_depth == Some(depth)
                {
                    section_properties_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF note metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

fn parse_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Configuration> {
    let mut note_class = None;
    let mut citation_style_name = None;
    let mut citation_body_style_name = None;
    let mut default_style_name = None;
    let mut master_page_name = None;
    let mut start_value = None;
    let mut number_prefix = None;
    let mut number_suffix = None;
    let mut number_format = None;
    let mut letter_sync = None;
    let mut start_numbering_at = None;
    let mut footnotes_position = None;
    let mut seen = HashSet::new();

    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        if !seen.insert((namespace, local.as_ref().to_vec())) {
            return invalid("duplicate text:notes-configuration attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_value(&value, "notes configuration attribute", true)?;
        match (namespace, local.as_ref()) {
            (NamespaceKind::Text, b"note-class") => note_class = Some(Class::parse(&value)?),
            (NamespaceKind::Text, b"citation-style-name") => citation_style_name = Some(value),
            (NamespaceKind::Text, b"citation-body-style-name") => {
                citation_body_style_name = Some(value)
            },
            (NamespaceKind::Text, b"default-style-name") => default_style_name = Some(value),
            (NamespaceKind::Text, b"master-page-name") => master_page_name = Some(value),
            (NamespaceKind::Text, b"start-value") => {
                start_value = Some(parse_nonnegative_integer(&value, "text:start-value")?)
            },
            (NamespaceKind::Style, b"num-prefix") => number_prefix = Some(value),
            (NamespaceKind::Style, b"num-suffix") => number_suffix = Some(value),
            (NamespaceKind::Style, b"num-format") => number_format = Some(Format::parse(value)?),
            (NamespaceKind::Style, b"num-letter-sync") => {
                letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?)
            },
            (NamespaceKind::Text, b"start-numbering-at") => {
                start_numbering_at = Some(NumberingScope::parse(&value)?)
            },
            (NamespaceKind::Text, b"footnotes-position") => {
                footnotes_position = Some(Position::parse(&value)?)
            },
            _ => return invalid("unsupported text:notes-configuration attribute"),
        }
    }
    let note_class = note_class.ok_or_else(|| {
        Error::InvalidFormat("notes configuration requires text:note-class".to_string())
    })?;
    Ok(Configuration {
        note_class,
        citation_style_name,
        citation_body_style_name,
        default_style_name,
        master_page_name,
        start_value,
        number_prefix,
        number_suffix,
        number_format,
        letter_sync,
        start_numbering_at,
        footnotes_position,
        continuation_notice_forward: None,
        continuation_notice_backward: None,
    })
}

fn parse_notices(reader: &mut NsReader<&[u8]>, configuration: &mut Configuration) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-forward" =>
            {
                reject_attributes(&element)?;
                if configuration.continuation_notice_forward.is_some() {
                    return invalid("duplicate forward note continuation notice");
                }
                configuration.continuation_notice_forward =
                    Some(parse_notice(reader, b"note-continuation-notice-forward")?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-forward" =>
            {
                if configuration
                    .continuation_notice_forward
                    .replace(String::new())
                    .is_some()
                {
                    return invalid("duplicate forward note continuation notice");
                }
                reject_attributes(&element)?;
            },
            Event::Start(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-backward" =>
            {
                reject_attributes(&element)?;
                if configuration.continuation_notice_backward.is_some() {
                    return invalid("duplicate backward note continuation notice");
                }
                configuration.continuation_notice_backward =
                    Some(parse_notice(reader, b"note-continuation-notice-backward")?);
            },
            Event::Empty(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"note-continuation-notice-backward" =>
            {
                if configuration
                    .continuation_notice_backward
                    .replace(String::new())
                    .is_some()
                {
                    return invalid("duplicate backward note continuation notice");
                }
                reject_attributes(&element)?;
            },
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"notes-configuration" =>
            {
                break;
            },
            Event::Text(text) => {
                require_whitespace(&text.decode().map_err(xml_error)?, "notes-configuration")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in notes configuration"),
            Event::Eof => return invalid("unterminated text:notes-configuration"),
            _ => return invalid("unsupported child in text:notes-configuration"),
        }
        buffer.clear();
    }
    Ok(())
}

fn parse_notice(reader: &mut NsReader<&[u8]>, expected_local: &[u8]) -> Result<String> {
    let mut output = String::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Text(value) => append_text(&mut output, &value.decode().map_err(xml_error)?)?,
            Event::CData(value) => append_text(&mut output, &value.decode().map_err(xml_error)?)?,
            Event::GeneralRef(reference) => append_reference(&mut output, &reference)?,
            Event::End(element)
                if namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == expected_local =>
            {
                break;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated note continuation notice"),
            _ => return invalid("note continuation notices may contain text only"),
        }
        buffer.clear();
    }
    Ok(output)
}

fn reject_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() != b"xmlns" && !attribute.key.as_ref().starts_with(b"xmlns:") {
            return invalid("note continuation notice cannot contain attributes");
        }
    }
    Ok(())
}

fn insert_configuration(result: &mut Configurations, configuration: Configuration) -> Result<()> {
    let slot = match configuration.note_class {
        Class::Footnote => &mut result.footnote,
        Class::Endnote => &mut result.endnote,
    };
    if slot.replace(configuration).is_some() {
        return invalid("duplicate notes configuration for the same note class");
    }
    Ok(())
}

fn require_styles_scope(styles_depth: Option<usize>, depth: usize) -> Result<()> {
    if styles_depth == Some(depth) {
        Ok(())
    } else {
        invalid("text:notes-configuration must be a direct office:styles child")
    }
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
    xml[..end]
        .rfind('<')
        .ok_or_else(|| Error::InvalidFormat("invalid notes XML event boundary".to_string()))
}

fn locate_configuration(xml: &str, note_class: Class) -> Result<(Option<XmlSpan>, StylesSite)> {
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
                    && local == b"notes-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                    && parse_attributes(&reader, element)?.note_class == note_class
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
                    && local == b"notes-configuration"
                    && matches!(stack.last(), Some((NamespaceKind::Office, parent)) if parent == b"styles")
                    && parse_attributes(&reader, element)?.note_class == note_class
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
                            .map_err(|_| Error::InvalidFormat("invalid styles QName".to_string()))?
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

/// Insert or replace one note-class configuration without rewriting unrelated XML.
pub fn set_xml(xml: &str, configuration: &Configuration) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_configuration(xml, configuration.note_class)?;
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

/// Remove one note-class configuration without rewriting unrelated XML.
pub fn remove_xml(xml: &str, note_class: Class) -> Result<String> {
    let (target, _) = locate_configuration(xml, note_class)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(value) if value.as_ref() == STYLE_NAMESPACE => NamespaceKind::Style,
        _ => NamespaceKind::Other,
    }
}

fn append_reference(
    output: &mut String,
    reference: &quick_xml::events::BytesRef<'_>,
) -> Result<()> {
    if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
        let mut encoded = [0u8; 4];
        append_text(output, character.encode_utf8(&mut encoded))
    } else {
        let name = reference.decode().map_err(xml_error)?;
        append_text(
            output,
            match name.as_ref() {
                "amp" => "&",
                "lt" => "<",
                "gt" => ">",
                "apos" => "'",
                "quot" => "\"",
                _ => return invalid("unsupported entity in note continuation notice"),
            },
        )
    }
}

fn append_text(output: &mut String, value: &str) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("note notice size overflow".to_string()))?;
    if size > MAX_VALUE_BYTES {
        return invalid(format!(
            "note continuation notice exceeds the {MAX_VALUE_BYTES} byte limit"
        ));
    }
    output.push_str(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => invalid(format!("{name} must be true or false")),
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

fn require_whitespace(value: &str, context: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        invalid(format!("text:{context} cannot contain text"))
    }
}

fn write_notice(output: &mut String, name: &str, value: Option<&str>) {
    let Some(value) = value else { return };
    output.push('<');
    output.push_str(name);
    if value.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        escape_text(output, value);
        output.push_str("</");
        output.push_str(name);
        output.push('>');
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
