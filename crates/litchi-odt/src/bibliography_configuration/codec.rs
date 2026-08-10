//! Bounded XML codec and lossless styles.xml mutation.

use super::model::{Configuration, Field, SortKey};
use super::{
    FO, MAX_AGGREGATE_BYTES, MAX_DEPTH, MAX_SORT_KEYS, MAX_VALUE_BYTES, MAX_XML_BYTES, OFFICE,
    STYLE, TEXT,
};
use crate::variable_declaration::Part;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

impl Configuration {
    /// Serialize the document-level bibliography policy and ordered sort keys.
    ///
    /// This fragment belongs in an ODF styles declaration context; it is not a
    /// child of `text:bibliography-source`.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut attributes = Vec::<(&str, &str, String)>::new();
        if let Some(value) = &self.prefix {
            attributes.push((TEXT, "prefix", value.clone()));
        }
        if let Some(value) = &self.suffix {
            attributes.push((TEXT, "suffix", value.clone()));
        }
        if let Some(value) = self.numbered_entries {
            attributes.push((TEXT, "numbered-entries", value.to_string()));
        }
        if let Some(value) = self.sort_by_position {
            attributes.push((TEXT, "sort-by-position", value.to_string()));
        }
        if let Some(value) = &self.language {
            attributes.push((FO, "language", value.clone()));
        }
        if let Some(value) = &self.country {
            attributes.push((FO, "country", value.clone()));
        }
        if let Some(value) = &self.script {
            attributes.push((FO, "script", value.clone()));
        }
        if let Some(value) = &self.rfc_language_tag {
            attributes.push((STYLE, "rfc-language-tag", value.clone()));
        }
        if let Some(value) = &self.sort_algorithm {
            attributes.push((TEXT, "sort-algorithm", value.clone()));
        }
        attributes.sort_by(|left, right| (left.0, left.1).cmp(&(right.0, right.1)));
        let mut output = String::from("<text:bibliography-configuration xmlns:fo=\"");
        escape_attribute(FO, &mut output);
        output.push_str("\" xmlns:style=\"");
        escape_attribute(STYLE, &mut output);
        output.push_str("\" xmlns:text=\"");
        escape_attribute(TEXT, &mut output);
        output.push('"');
        for (namespace, local, value) in attributes {
            output.push(' ');
            output.push_str(match namespace {
                TEXT => "text",
                FO => "fo",
                STYLE => "style",
                _ => unreachable!(),
            });
            output.push(':');
            output.push_str(local);
            output.push_str("=\"");
            escape_attribute(&value, &mut output);
            output.push('"');
        }
        if self.sort_keys.is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            for key in &self.sort_keys {
                output.push_str("<text:sort-key text:key=\"");
                output.push_str(key.field.as_str());
                output.push('"');
                if let Some(ascending) = key.ascending {
                    output.push_str(" text:sort-ascending=\"");
                    output.push_str(if ascending { "true" } else { "false" });
                    output.push('"');
                }
                output.push_str("/>");
            }
            output.push_str("</text:bibliography-configuration>");
        }
        if output.len() > MAX_AGGREGATE_BYTES {
            return invalid("serialized bibliography configuration exceeds 4 MiB");
        }
        Ok(output)
    }
}

fn escape_attribute(value: &str, output: &mut String) {
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

#[derive(Clone)]
struct Frame {
    namespace: Option<String>,
    local: String,
}

struct ActiveConfiguration {
    depth: usize,
    value: Configuration,
}

type Attributes = HashMap<(String, String), String>;

pub(crate) fn parse_bibliography_configuration(xml: &str) -> Result<Option<Configuration>> {
    parse_bibliography_configuration_parts(&[(xml, Part::Styles)])
}

pub(crate) fn parse_bibliography_configuration_parts(
    parts: &[(&str, Part)],
) -> Result<Option<Configuration>> {
    if !parts
        .iter()
        .any(|(xml, _)| xml.contains("bibliography-configuration"))
    {
        return Ok(None);
    }
    let total = parts.iter().try_fold(0usize, |total, (xml, _)| {
        total
            .checked_add(xml.len())
            .ok_or_else(|| make_error("bibliography configuration XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return invalid("bibliography configuration XML exceeds 64 MiB");
    }

    let mut result = None;
    let mut aggregate = 0usize;
    for (xml, part) in parts {
        parse_part(xml, *part, &mut result, &mut aggregate)?;
    }
    Ok(result)
}

fn parse_part(
    xml: &str,
    part: Part,
    result: &mut Option<Configuration>,
    aggregate: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut depth = 0usize;
    let mut active: Option<ActiveConfiguration> = None;
    let mut pending_sort_key: Option<usize> = None;

    loop {
        let (namespace, event) =
            reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|source| {
                    make_error(format!("invalid bibliography configuration XML: {source}"))
                })?;
        match event {
            Event::Start(ref element) => {
                if pending_sort_key.is_some() {
                    return invalid("text:sort-key cannot contain elements");
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if let Some(configuration) = active.as_mut() {
                    if namespace.as_deref() != Some(TEXT)
                        || local != "sort-key"
                        || depth != configuration.depth
                    {
                        return invalid(
                            "bibliography configuration may contain only text:sort-key elements",
                        );
                    }
                    add_sort_key(&reader, element, configuration, aggregate)?;
                    pending_sort_key = Some(depth + 1);
                } else if namespace.as_deref() == Some(TEXT)
                    && local == "bibliography-configuration"
                {
                    start_configuration(
                        &reader,
                        element,
                        part,
                        depth,
                        &stack,
                        result,
                        aggregate,
                        &mut active,
                    )?;
                }
                stack.push(Frame { namespace, local });
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| make_error("bibliography configuration depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid(format!(
                        "bibliography configuration exceeds {MAX_DEPTH} XML levels"
                    ));
                }
            },
            Event::Empty(ref element) => {
                if pending_sort_key.is_some() {
                    return invalid("text:sort-key cannot contain elements");
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if let Some(configuration) = active.as_mut() {
                    if namespace.as_deref() != Some(TEXT)
                        || local != "sort-key"
                        || depth != configuration.depth
                    {
                        return invalid(
                            "bibliography configuration may contain only text:sort-key elements",
                        );
                    }
                    add_sort_key(&reader, element, configuration, aggregate)?;
                } else if namespace.as_deref() == Some(TEXT)
                    && local == "bibliography-configuration"
                {
                    let mut temporary = None;
                    start_configuration(
                        &reader,
                        element,
                        part,
                        depth,
                        &stack,
                        result,
                        aggregate,
                        &mut temporary,
                    )?;
                    let configuration = temporary
                        .ok_or_else(|| make_error("missing bibliography configuration"))?
                        .value;
                    configuration.validate()?;
                    *result = Some(configuration);
                }
            },
            Event::End(_) => {
                if pending_sort_key.is_some_and(|pending_depth| pending_depth == depth) {
                    pending_sort_key = None;
                }
                if active
                    .as_ref()
                    .is_some_and(|configuration| configuration.depth == depth)
                {
                    let configuration = active
                        .take()
                        .ok_or_else(|| make_error("missing completed bibliography configuration"))?
                        .value;
                    configuration.validate()?;
                    *result = Some(configuration);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| make_error("bibliography configuration stack underflow"))?;
                stack
                    .pop()
                    .ok_or_else(|| make_error("bibliography frame stack underflow"))?;
            },
            Event::Text(ref text) => {
                let value = text.decode().map_err(|source| {
                    make_error(format!("invalid bibliography configuration text: {source}"))
                })?;
                if pending_sort_key.is_some() && !value.is_empty() {
                    return invalid("text:sort-key must be empty");
                }
                if active.is_some() && pending_sort_key.is_none() && !value.trim().is_empty() {
                    return invalid("bibliography configuration cannot contain text");
                }
            },
            Event::CData(ref value)
                if (active.is_some() || pending_sort_key.is_some()) && !value.is_empty() =>
            {
                return invalid("bibliography configuration cannot contain CDATA");
            },
            Event::GeneralRef(_) if active.is_some() || pending_sort_key.is_some() => {
                return invalid("bibliography configuration cannot contain entity references");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "DTDs and processing instructions are prohibited in bibliography XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || active.is_some() || pending_sort_key.is_some() {
        return invalid("incomplete bibliography configuration XML structure");
    }
    Ok(())
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
        .ok_or_else(|| make_error("invalid bibliography XML event boundary"))
}

fn locate_bibliography_configuration(xml: &str) -> Result<(Option<XmlSpan>, StylesSite)> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("bibliography configuration XML exceeds 64 MiB");
    }
    parse_bibliography_configuration(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut styles_site = None;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|source| {
                make_error(format!(
                    "invalid bibliography configuration XML while locating mutation site: {source}"
                ))
            })?;
        let namespace = namespace_uri(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| make_error("bibliography configuration depth overflow"))?;
                if namespace.as_deref() == Some(TEXT)
                    && local == "bibliography-configuration"
                    && matches!(
                        stack.last(),
                        Some(parent)
                            if parent.namespace.as_deref() == Some(OFFICE)
                                && parent.local == "styles"
                    )
                {
                    open_target = Some((depth, start));
                }
                stack.push(Frame { namespace, local });
            },
            Event::Empty(ref element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if namespace.as_deref() == Some(TEXT)
                    && local == "bibliography-configuration"
                    && matches!(
                        stack.last(),
                        Some(parent)
                            if parent.namespace.as_deref() == Some(OFFICE)
                                && parent.local == "styles"
                    )
                {
                    target = Some(XmlSpan { start, end });
                }
                if namespace.as_deref() == Some(OFFICE) && local == "styles" {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Empty(
                        XmlSpan { start, end },
                        decode(element.name().as_ref(), "office:styles QName")?,
                    ));
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let depth = stack.len();
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, target_start) = open_target
                        .take()
                        .ok_or_else(|| make_error("missing bibliography target start"))?;
                    target = Some(XmlSpan {
                        start: target_start,
                        end,
                    });
                }
                if matches!(
                    stack.last(),
                    Some(parent)
                        if parent.namespace.as_deref() == Some(OFFICE)
                            && parent.local == "styles"
                ) {
                    if styles_site.is_some() {
                        return invalid("multiple office:styles elements are not supported");
                    }
                    styles_site = Some(StylesSite::Content(start));
                }
                stack
                    .pop()
                    .ok_or_else(|| make_error("bibliography configuration stack underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || open_target.is_some() {
        return invalid("incomplete bibliography configuration XML structure");
    }
    Ok((
        target,
        styles_site.ok_or_else(|| make_error("document has no office:styles element"))?,
    ))
}

/// Insert or replace the document-wide bibliography configuration without
/// rewriting unrelated styles XML.
pub(crate) fn set_bibliography_configuration_xml(
    xml: &str,
    configuration: &Configuration,
) -> Result<String> {
    configuration.validate()?;
    let (target, site) = locate_bibliography_configuration(xml)?;
    let fragment = configuration.to_xml_fragment()?;
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
                .ok_or_else(|| make_error("invalid empty office:styles element"))?;
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

/// Remove the document-wide bibliography configuration without rewriting
/// unrelated styles XML.
pub(crate) fn remove_bibliography_configuration_xml(xml: &str) -> Result<String> {
    let (target, _) = locate_bibliography_configuration(xml)?;
    let Some(span) = target else {
        return Ok(xml.to_string());
    };
    Ok(format!("{}{}", &xml[..span.start], &xml[span.end..]))
}

#[allow(clippy::too_many_arguments)]
fn start_configuration(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: Part,
    depth: usize,
    stack: &[Frame],
    result: &Option<Configuration>,
    aggregate: &mut usize,
    active: &mut Option<ActiveConfiguration>,
) -> Result<()> {
    if result.is_some() {
        return invalid("duplicate document bibliography configuration");
    }
    if part == Part::Content {
        return invalid("bibliography configuration must reside in styles metadata");
    }
    let parent = stack
        .last()
        .ok_or_else(|| make_error("misplaced bibliography configuration"))?;
    if parent.namespace.as_deref() != Some(OFFICE) || parent.local != "styles" {
        return invalid("bibliography configuration must be a direct office:styles child");
    }
    let attributes = collect_attributes(reader, element, aggregate)?;
    reject_unexpected(
        &attributes,
        &[
            (TEXT, "prefix"),
            (TEXT, "suffix"),
            (TEXT, "numbered-entries"),
            (TEXT, "sort-by-position"),
            (TEXT, "sort-algorithm"),
            (FO, "language"),
            (FO, "country"),
            (FO, "script"),
            (STYLE, "rfc-language-tag"),
        ],
    )?;
    let configuration = Configuration {
        prefix: get_owned(&attributes, TEXT, "prefix"),
        suffix: get_owned(&attributes, TEXT, "suffix"),
        numbered_entries: get(&attributes, TEXT, "numbered-entries")
            .map(parse_bool)
            .transpose()?,
        sort_by_position: get(&attributes, TEXT, "sort-by-position")
            .map(parse_bool)
            .transpose()?,
        sort_algorithm: get_owned(&attributes, TEXT, "sort-algorithm"),
        language: get_owned(&attributes, FO, "language"),
        country: get_owned(&attributes, FO, "country"),
        script: get_owned(&attributes, FO, "script"),
        rfc_language_tag: get_owned(&attributes, STYLE, "rfc-language-tag"),
        sort_keys: Vec::new(),
    };
    *active = Some(ActiveConfiguration {
        depth: depth + 1,
        value: configuration,
    });
    Ok(())
}

fn add_sort_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    configuration: &mut ActiveConfiguration,
    aggregate: &mut usize,
) -> Result<()> {
    if configuration.value.sort_keys.len() >= MAX_SORT_KEYS {
        return invalid(format!(
            "bibliography configuration exceeds {MAX_SORT_KEYS} sort keys"
        ));
    }
    let attributes = collect_attributes(reader, element, aggregate)?;
    reject_unexpected(&attributes, &[(TEXT, "key"), (TEXT, "sort-ascending")])?;
    let field = Field::parse(
        get(&attributes, TEXT, "key")
            .ok_or_else(|| make_error("text:sort-key requires text:key"))?,
    )?;
    let ascending = get(&attributes, TEXT, "sort-ascending")
        .map(parse_bool)
        .transpose()?;
    configuration
        .value
        .sort_keys
        .push(SortKey { field, ascending });
    Ok(())
}

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|source| {
            make_error(format!(
                "invalid bibliography configuration attribute: {source}"
            ))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&namespace)?.unwrap_or_default();
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|source| {
                make_error(format!(
                    "invalid bibliography configuration attribute value: {source}"
                ))
            })?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("bibliography configuration attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("bibliography metadata size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("bibliography configuration metadata exceeds 4 MiB");
        }
        if attributes.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded bibliography configuration attribute");
        }
    }
    Ok(attributes)
}

fn reject_unexpected(attributes: &Attributes, allowed: &[(&str, &str)]) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) && matches!(namespace.as_str(), TEXT | FO | STYLE)
        {
            return invalid(format!(
                "unexpected bibliography configuration attribute {namespace}:{local}"
            ));
        }
    }
    Ok(())
}

fn get<'a>(attributes: &'a Attributes, namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn get_owned(attributes: &Attributes, namespace: &str, local: &str) -> Option<String> {
    get(attributes, namespace, local).map(str::to_string)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid bibliography boolean '{value}'")),
    }
}

fn reject_spoofed_name(namespace: Option<&str>, local: &str) -> Result<()> {
    if matches!(local, "bibliography-configuration" | "sort-key") && namespace != Some(TEXT) {
        return invalid("bibliography configuration vocabulary uses the wrong namespace");
    }
    Ok(())
}

fn namespace_uri(result: &ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        )),
    }
}

fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| make_error(format!("invalid UTF-8 {description}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
