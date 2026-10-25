//! Bounded XML codec and lossless mutation for ODF note configurations.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use litchi_odf_common::{ResolvedReader, XmlStreamLimits};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
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

fn namespace_text(namespace: &[u8]) -> Result<&str> {
    std::str::from_utf8(namespace).map_err(|error| {
        Error::InvalidFormat(format!("invalid static namespace encoding: {error}"))
    })
}

impl Configuration {
    /// Serialize a namespace-complete notes configuration element.
    pub fn to_xml(&self, limits: &XmlStreamLimits) -> Result<String> {
        self.validate()?;
        let capacity = configuration_size(self)?;
        ensure_output_size(capacity, limits)?;
        let mut output = String::with_capacity(capacity);
        output.push_str("<text:notes-configuration xmlns:text=\"");
        output.push_str(namespace_text(TEXT_NAMESPACE)?);
        output.push_str("\" xmlns:style=\"");
        output.push_str(namespace_text(STYLE_NAMESPACE)?);
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
            return ensure_output(output, limits);
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
        ensure_output(output, limits)
    }
}

impl Configurations {
    /// Serialize both configurations in canonical footnote/endnote order.
    pub fn to_xml_fragment(&self, limits: &XmlStreamLimits) -> Result<String> {
        self.validate()?;
        let mut capacity = 0usize;
        if let Some(configuration) = &self.footnote {
            add_size(&mut capacity, configuration_size(configuration)?)?;
        }
        if let Some(configuration) = &self.endnote {
            add_size(&mut capacity, configuration_size(configuration)?)?;
        }
        ensure_output_size(capacity, limits)?;
        let mut output = String::with_capacity(capacity);
        if let Some(configuration) = &self.footnote {
            output.push_str(&configuration.to_xml(limits)?);
        }
        if let Some(configuration) = &self.endnote {
            output.push_str(&configuration.to_xml(limits)?);
        }
        ensure_output(output, limits)
    }
}

/// Parse footnote/endnote configurations from an ODF styles document.
pub fn parse(xml: &str, limits: &XmlStreamLimits) -> Result<Configurations> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES
        || u64::try_from(xml.len()).unwrap_or(u64::MAX) > limits.max_bytes()
    {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte notes limit"
        ));
    }
    let mut reader = ResolvedReader::from_xml_document(xml, *limits).map_err(|error| {
        litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
    })?;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut styles_content_depth = None;
    let mut section_properties_depth = None;
    let mut result = Configurations::default();
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
        })?;
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
                if depth > MAX_XML_DEPTH.min(limits.max_depth()) {
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
    }
    Ok(result)
}

fn parse_attributes(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
) -> Result<Configuration> {
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
                citation_body_style_name = Some(value);
            },
            (NamespaceKind::Text, b"default-style-name") => default_style_name = Some(value),
            (NamespaceKind::Text, b"master-page-name") => master_page_name = Some(value),
            (NamespaceKind::Text, b"start-value") => {
                start_value = Some(parse_nonnegative_integer(&value, "text:start-value")?);
            },
            (NamespaceKind::Style, b"num-prefix") => number_prefix = Some(value),
            (NamespaceKind::Style, b"num-suffix") => number_suffix = Some(value),
            (NamespaceKind::Style, b"num-format") => number_format = Some(Format::parse(value)?),
            (NamespaceKind::Style, b"num-letter-sync") => {
                letter_sync = Some(parse_bool(&value, "style:num-letter-sync")?);
            },
            (NamespaceKind::Text, b"start-numbering-at") => {
                start_numbering_at = Some(NumberingScope::parse(&value)?);
            },
            (NamespaceKind::Text, b"footnotes-position") => {
                footnotes_position = Some(Position::parse(&value)?);
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

fn parse_notices(reader: &mut ResolvedReader<'_>, configuration: &mut Configuration) -> Result<()> {
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
        })?;
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
                require_whitespace(&text.decode().map_err(xml_error)?, "notes-configuration")?;
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in notes configuration"),
            Event::Eof => return invalid("unterminated text:notes-configuration"),
            _ => return invalid("unsupported child in text:notes-configuration"),
        }
    }
    Ok(())
}

fn parse_notice(reader: &mut ResolvedReader<'_>, expected_local: &[u8]) -> Result<String> {
    let mut output = String::new();
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
        })?;
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

fn locate_configuration(
    xml: &str,
    note_class: Class,
    limits: &XmlStreamLimits,
) -> Result<(Option<XmlSpan>, StylesSite)> {
    parse(xml, limits)?;
    let mut reader = ResolvedReader::from_xml_document(xml, *limits).map_err(|error| {
        litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
    })?;
    let mut stack: Vec<(NamespaceKind, Vec<u8>)> = Vec::new();
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut styles_site = None;
    loop {
        let (resolved, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid ODT notes XML")
        })?;
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
                            .map_err(|_error| {
                                Error::InvalidFormat("invalid styles QName".to_string())
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
                    let (_, target_start) = open_target.take().ok_or_else(|| {
                        Error::InvalidFormat("missing notes-configuration target start".to_string())
                    })?;
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
    }
    Ok((
        target,
        styles_site
            .ok_or_else(|| Error::InvalidFormat("document has no office:styles".to_string()))?,
    ))
}

/// Insert or replace one note-class configuration without rewriting unrelated XML.
pub fn set_xml(
    xml: &str,
    configuration: &Configuration,
    limits: &XmlStreamLimits,
) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_configuration(xml, configuration.note_class, limits)?;
    let fragment = configuration.to_xml(limits)?;
    if let Some(span) = target {
        let output_len = replaced_size(xml.len(), span.end - span.start, fragment.len(), limits)?;
        let mut output = String::with_capacity(output_len);
        output.push_str(&xml[..span.start]);
        output.push_str(&fragment);
        output.push_str(&xml[span.end..]);
        return ensure_output(output, limits);
    }
    match site {
        StylesSite::Content(insertion) => {
            let output_len = replaced_size(xml.len(), 0, fragment.len(), limits)?;
            let mut output = String::with_capacity(output_len);
            output.push_str(&xml[..insertion]);
            output.push_str(&fragment);
            output.push_str(&xml[insertion..]);
            ensure_output(output, limits)
        },
        StylesSite::Empty(span, qname) => {
            let raw = &xml[span.start..span.end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| Error::InvalidFormat("invalid empty office:styles".to_string()))?;
            let replacement_len = raw[..slash]
                .len()
                .checked_add(1)
                .and_then(|length| length.checked_add(fragment.len()))
                .and_then(|length| length.checked_add(2 + qname.len() + 1))
                .ok_or_else(|| Error::InvalidFormat("notes output size overflow".to_string()))?;
            let output_len = replaced_size(xml.len(), raw.len(), replacement_len, limits)?;
            let mut output = String::with_capacity(output_len);
            output.push_str(&xml[..span.start]);
            output.push_str(&raw[..slash]);
            output.push('>');
            output.push_str(&fragment);
            output.push_str("</");
            output.push_str(&qname);
            output.push('>');
            output.push_str(&xml[span.end..]);
            ensure_output(output, limits)
        },
    }
}

/// Remove one note-class configuration without rewriting unrelated XML.
pub fn remove_xml(xml: &str, note_class: Class, limits: &XmlStreamLimits) -> Result<String> {
    let (target, _) = locate_configuration(xml, note_class, limits)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    let output_len = replaced_size(xml.len(), span.end - span.start, 0, limits)?;
    let mut output = String::with_capacity(output_len);
    output.push_str(&xml[..span.start]);
    output.push_str(&xml[span.end..]);
    ensure_output(output, limits)
}

fn ensure_output(output: String, limits: &XmlStreamLimits) -> Result<String> {
    if u64::try_from(output.len()).unwrap_or(u64::MAX) > limits.max_bytes() {
        return invalid("serialized notes configuration exceeds the configured XML limit");
    }
    Ok(output)
}

fn replaced_size(
    source_len: usize,
    removed_len: usize,
    replacement_len: usize,
    limits: &XmlStreamLimits,
) -> Result<usize> {
    let size = source_len
        .checked_sub(removed_len)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| Error::InvalidFormat("notes output size overflow".to_string()))?;
    ensure_output_size(size, limits)?;
    Ok(size)
}

fn configuration_size(configuration: &Configuration) -> Result<usize> {
    let mut size = 0usize;
    add_size(&mut size, b"<text:notes-configuration xmlns:text=\"".len())?;
    add_size(&mut size, namespace_text(TEXT_NAMESPACE)?.len())?;
    add_size(&mut size, b"\" xmlns:style=\"".len())?;
    add_size(&mut size, namespace_text(STYLE_NAMESPACE)?.len())?;
    add_size(&mut size, 1)?;
    add_attr_size(
        &mut size,
        "text:note-class",
        Some(configuration.note_class.as_str()),
    )?;
    add_attr_size(
        &mut size,
        "text:citation-style-name",
        configuration.citation_style_name.as_deref(),
    )?;
    add_attr_size(
        &mut size,
        "text:citation-body-style-name",
        configuration.citation_body_style_name.as_deref(),
    )?;
    add_attr_size(
        &mut size,
        "text:default-style-name",
        configuration.default_style_name.as_deref(),
    )?;
    add_attr_size(
        &mut size,
        "text:master-page-name",
        configuration.master_page_name.as_deref(),
    )?;
    add_u64_attr_size(&mut size, "text:start-value", configuration.start_value)?;
    add_attr_size(
        &mut size,
        "style:num-prefix",
        configuration.number_prefix.as_deref(),
    )?;
    add_attr_size(
        &mut size,
        "style:num-suffix",
        configuration.number_suffix.as_deref(),
    )?;
    add_attr_size(
        &mut size,
        "style:num-format",
        configuration.number_format.as_ref().map(Format::as_str),
    )?;
    add_attr_size(
        &mut size,
        "style:num-letter-sync",
        configuration
            .letter_sync
            .map(|value| if value { "true" } else { "false" }),
    )?;
    add_attr_size(
        &mut size,
        "text:start-numbering-at",
        configuration.start_numbering_at.map(NumberingScope::as_str),
    )?;
    add_attr_size(
        &mut size,
        "text:footnotes-position",
        configuration.footnotes_position.map(Position::as_str),
    )?;
    let notices = configuration.continuation_notice_forward.is_some()
        || configuration.continuation_notice_backward.is_some();
    if !notices {
        add_size(&mut size, 2)?;
        return Ok(size);
    }
    add_size(&mut size, 1)?;
    add_notice_size(
        &mut size,
        "text:note-continuation-notice-forward",
        configuration.continuation_notice_forward.as_deref(),
    )?;
    add_notice_size(
        &mut size,
        "text:note-continuation-notice-backward",
        configuration.continuation_notice_backward.as_deref(),
    )?;
    add_size(&mut size, b"</text:notes-configuration>".len())?;
    Ok(size)
}

fn add_notice_size(size: &mut usize, name: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    add_size(size, 1)?;
    add_size(size, name.len())?;
    if value.is_empty() {
        add_size(size, 2)
    } else {
        add_size(size, 1)?;
        add_size(size, escaped_text_len(value)?)?;
        add_size(size, 2 + name.len() + 1)
    }
}

fn add_attr_size(size: &mut usize, name: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    add_size(size, 1)?;
    add_size(size, name.len())?;
    add_size(size, 2)?;
    add_size(size, escaped_attribute_len(value)?)?;
    add_size(size, 1)
}

fn add_u64_attr_size(size: &mut usize, name: &str, value: Option<u64>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    add_size(size, 1 + name.len() + 2 + value.to_string().len() + 1)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |size, character| {
        let encoded = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\r' | '\n' => 5,
            '\t' => 4,
            _ => character.len_utf8(),
        };
        size.checked_add(encoded)
            .ok_or_else(|| Error::InvalidFormat("notes output size overflow".to_string()))
    })
}

fn escaped_text_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |size, character| {
        let encoded = match character {
            '&' => 5,
            '<' | '>' => 4,
            _ => character.len_utf8(),
        };
        size.checked_add(encoded)
            .ok_or_else(|| Error::InvalidFormat("notes output size overflow".to_string()))
    })
}

fn add_size(size: &mut usize, amount: usize) -> Result<()> {
    *size = size
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("notes output size overflow".to_string()))?;
    Ok(())
}

fn ensure_output_size(size: usize, limits: &XmlStreamLimits) -> Result<()> {
    if u64::try_from(size).unwrap_or(u64::MAX) > limits.max_bytes() {
        return invalid("serialized notes configuration exceeds the configured XML limit");
    }
    Ok(())
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
        .map_err(|_error| Error::InvalidFormat(format!("{name} exceeds the u64 range")))
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
