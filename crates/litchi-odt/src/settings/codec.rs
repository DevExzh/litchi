//! Bounded, namespace-aware XML codec for ODF configuration settings.

use base64::Engine as _;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::HashSet;

use super::model::{
    ConfigItem, ConfigMap, ConfigMapEntry, ConfigNode, ConfigSet, ConfigValue, Settings,
};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CONFIG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_SCALAR_BYTES: usize = 4 * 1024 * 1024;
const MAX_NODES: usize = 65_536;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 32;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DocumentKind {
    Flat,
    Package,
}

enum Frame {
    Set {
        name: String,
        children: Vec<ConfigNode>,
    },
    Item {
        name: String,
        value_type: String,
        text: String,
    },
    IndexedMap {
        name: String,
        entries: Vec<ConfigMapEntry>,
    },
    NamedMap {
        name: String,
        entries: Vec<ConfigMapEntry>,
    },
    Entry {
        name: Option<String>,
        children: Vec<ConfigNode>,
    },
}

enum Finished {
    Node(ConfigNode),
    Entry(ConfigMapEntry),
}

impl Frame {
    fn local_name(&self) -> &'static [u8] {
        match self {
            Self::Set { .. } => b"config-item-set",
            Self::Item { .. } => b"config-item",
            Self::IndexedMap { .. } => b"config-item-map-indexed",
            Self::NamedMap { .. } => b"config-item-map-named",
            Self::Entry { .. } => b"config-item-map-entry",
        }
    }

    fn finish(self) -> Result<Finished> {
        Ok(match self {
            Self::Set { name, children } => {
                Finished::Node(ConfigNode::Set(ConfigSet { name, children }))
            },
            Self::Item {
                name,
                value_type,
                text,
            } => Finished::Node(ConfigNode::Item(ConfigItem {
                name,
                value: parse_value(&value_type, &text)?,
            })),
            Self::IndexedMap { name, entries } => {
                Finished::Node(ConfigNode::IndexedMap(ConfigMap { name, entries }))
            },
            Self::NamedMap { name, entries } => {
                Finished::Node(ConfigNode::NamedMap(ConfigMap { name, entries }))
            },
            Self::Entry { name, children } => Finished::Entry(ConfigMapEntry { name, children }),
        })
    }
}

pub(super) fn parse(xml: &str, kind: DocumentKind) -> Result<Settings> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("settings XML exceeds the configured size limit");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let mut buffer = Vec::new();
    let mut xml_depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut settings_seen = false;
    let mut in_settings = false;
    let mut settings_depth = 0usize;
    let mut node_count = 0usize;
    let mut stack = Vec::<Frame>::new();
    let mut namespace_scopes = Vec::<Vec<(Vec<u8>, Vec<u8>)>>::new();
    let mut result = Settings::default();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid settings XML: {error}")))?;
        match event {
            Event::Start(start) => {
                if root_closed {
                    return invalid("settings XML has more than one document element");
                }
                if xml_depth >= MAX_DEPTH {
                    return invalid("settings XML exceeds the configured depth limit");
                }
                namespace_scopes.push(namespace_declarations(&start, decoder)?);
                process_start(
                    decoder,
                    &namespace_scopes,
                    &namespace,
                    &start,
                    false,
                    xml_depth,
                    kind,
                    &mut root_seen,
                    &mut settings_seen,
                    &mut in_settings,
                    &mut settings_depth,
                    &mut node_count,
                    &mut stack,
                    &mut result,
                )?;
                xml_depth += 1;
            },
            Event::Empty(start) => {
                if root_closed {
                    return invalid("settings XML has more than one document element");
                }
                if xml_depth >= MAX_DEPTH {
                    return invalid("settings XML exceeds the configured depth limit");
                }
                namespace_scopes.push(namespace_declarations(&start, decoder)?);
                process_start(
                    decoder,
                    &namespace_scopes,
                    &namespace,
                    &start,
                    true,
                    xml_depth,
                    kind,
                    &mut root_seen,
                    &mut settings_seen,
                    &mut in_settings,
                    &mut settings_depth,
                    &mut node_count,
                    &mut stack,
                    &mut result,
                )?;
                namespace_scopes.pop();
                if xml_depth == 0 {
                    root_closed = true;
                }
            },
            Event::End(end) => {
                xml_depth = xml_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("unbalanced settings XML".to_string()))?;
                let local = end.local_name();
                if in_settings && xml_depth == settings_depth {
                    if !is_namespace(&namespace, OFFICE) || local.as_ref() != b"settings" {
                        return invalid("configuration element closed outside office:settings");
                    }
                    if !stack.is_empty() {
                        return invalid("unclosed configuration element");
                    }
                    in_settings = false;
                } else if in_settings {
                    if !is_namespace(&namespace, CONFIG) {
                        return invalid("configuration elements must use the config namespace");
                    }
                    let frame = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("orphan configuration end element".to_string())
                    })?;
                    if local.as_ref() != frame.local_name() {
                        return invalid("mismatched configuration end element");
                    }
                    attach(frame.finish()?, &mut stack, &mut result)?;
                }
                if xml_depth == 0 {
                    root_closed = true;
                }
                namespace_scopes.pop().ok_or_else(|| {
                    Error::InvalidFormat("unbalanced namespace scope".to_string())
                })?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid settings text encoding: {error}"))
                })?;
                let value = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::InvalidFormat(format!("invalid settings text escape: {error}"))
                })?;
                append_text(&value, root_seen, root_closed, in_settings, &mut stack)?;
            },
            Event::CData(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid settings CDATA encoding: {error}"))
                })?;
                append_text(&value, root_seen, root_closed, in_settings, &mut stack)?;
            },
            Event::GeneralRef(reference) => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_error| {
                    Error::InvalidFormat("invalid XML character reference".to_string())
                })?;
                let value = resolve_reference(name)?;
                append_text(&value, root_seen, root_closed, in_settings, &mut stack)?;
            },
            Event::Decl(decl) => {
                if root_seen {
                    return invalid("XML declaration must precede the document element");
                }
                if decl
                    .version()
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML declaration: {error}"))
                    })?
                    .as_ref()
                    != b"1.0"
                {
                    return invalid("only XML 1.0 settings documents are supported");
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid(
                    "DOCTYPE and processing instructions are prohibited in settings XML",
                );
            },
            Event::Eof => break,
            Event::Comment(_) => {},
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || xml_depth != 0 {
        return invalid("incomplete settings XML document");
    }
    if !namespace_scopes.is_empty() {
        return invalid("unclosed namespace scope in settings XML");
    }
    if kind == DocumentKind::Package && !settings_seen {
        return invalid("settings.xml has no office:settings element");
    }
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn process_start(
    decoder: Decoder,
    namespace_scopes: &[Vec<(Vec<u8>, Vec<u8>)>],
    namespace: &ResolveResult<'_>,
    start: &BytesStart<'_>,
    empty: bool,
    depth: usize,
    kind: DocumentKind,
    root_seen: &mut bool,
    settings_seen: &mut bool,
    in_settings: &mut bool,
    settings_depth: &mut usize,
    node_count: &mut usize,
    stack: &mut Vec<Frame>,
    result: &mut Settings,
) -> Result<()> {
    let local = start.local_name();
    if depth == 0 {
        if *root_seen || !is_namespace(namespace, OFFICE) {
            return invalid("invalid settings document element");
        }
        let expected = match kind {
            DocumentKind::Flat => b"document".as_slice(),
            DocumentKind::Package => b"document-settings".as_slice(),
        };
        if local.as_ref() != expected {
            return invalid("unexpected settings document element");
        }
        validate_no_semantic_attributes(decoder, start)?;
        *root_seen = true;
        return Ok(());
    }

    if is_namespace(namespace, OFFICE) && local.as_ref() == b"settings" {
        if depth != 1 || *settings_seen || *in_settings {
            return invalid(
                "office:settings must be a unique direct child of the document element",
            );
        }
        validate_no_semantic_attributes(decoder, start)?;
        *settings_seen = true;
        *settings_depth = depth;
        if !empty {
            *in_settings = true;
        }
        return Ok(());
    }

    if !*in_settings {
        return Ok(());
    }
    if !is_namespace(namespace, CONFIG) {
        return invalid("office:settings contains a non-config element");
    }
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("configuration node count overflow".to_string()))?;
    if *node_count > MAX_NODES {
        return invalid("settings exceed the configured node limit");
    }
    if stack.len() >= MAX_DEPTH {
        return invalid("settings exceed the configured semantic depth limit");
    }

    let attributes = semantic_attributes(decoder, namespace_scopes, start)?;
    let name = attribute(&attributes, b"name");
    let frame = match local.as_ref() {
        b"config-item-set" => Frame::Set {
            name: required(name, "config-item-set requires config:name")?,
            children: Vec::new(),
        },
        b"config-item" => Frame::Item {
            name: required(name, "config-item requires config:name")?,
            value_type: required(
                attribute(&attributes, b"type"),
                "config-item requires config:type",
            )?,
            text: String::new(),
        },
        b"config-item-map-indexed" => Frame::IndexedMap {
            name: required(name, "config-item-map-indexed requires config:name")?,
            entries: Vec::new(),
        },
        b"config-item-map-named" => Frame::NamedMap {
            name: required(name, "config-item-map-named requires config:name")?,
            entries: Vec::new(),
        },
        b"config-item-map-entry" => Frame::Entry {
            name,
            children: Vec::new(),
        },
        _ => return invalid("unknown configuration element"),
    };
    validate_placement(&frame, stack)?;
    if empty {
        attach(frame.finish()?, stack, result)?;
    } else {
        stack.push(frame);
    }
    Ok(())
}

fn validate_placement(frame: &Frame, stack: &[Frame]) -> Result<()> {
    match (stack.last(), frame) {
        (None, Frame::Set { .. }) => Ok(()),
        (None, _) => invalid("office:settings may contain only top-level config-item-set elements"),
        (Some(Frame::Set { .. } | Frame::Entry { .. }), Frame::Entry { .. }) => {
            invalid("config-item-map-entry must be a direct child of a config map")
        },
        (Some(Frame::Set { .. } | Frame::Entry { .. }), _) => Ok(()),
        (Some(Frame::IndexedMap { .. } | Frame::NamedMap { .. }), Frame::Entry { .. }) => Ok(()),
        (Some(Frame::IndexedMap { .. } | Frame::NamedMap { .. }), _) => {
            invalid("configuration maps may contain only map entries")
        },
        (Some(Frame::Item { .. }), _) => invalid("config-item elements cannot contain elements"),
    }
}

fn attach(finished: Finished, stack: &mut [Frame], result: &mut Settings) -> Result<()> {
    match (stack.last_mut(), finished) {
        (None, Finished::Node(ConfigNode::Set(set))) => {
            result.sets.push(set);
            Ok(())
        },
        (None, _) => invalid("invalid top-level configuration node"),
        (
            Some(Frame::Set { children, .. } | Frame::Entry { children, .. }),
            Finished::Node(node),
        ) => {
            children.push(node);
            Ok(())
        },
        (Some(Frame::IndexedMap { entries, .. }), Finished::Entry(entry)) => {
            if entry.name.is_some() {
                return invalid("indexed map entries must not have config:name");
            }
            entries.push(entry);
            Ok(())
        },
        (Some(Frame::NamedMap { entries, .. }), Finished::Entry(entry)) => {
            if entry.name.as_deref().is_none_or(str::is_empty) {
                return invalid("named map entries require a non-empty config:name");
            }
            entries.push(entry);
            Ok(())
        },
        _ => invalid("invalid configuration child placement"),
    }
}

fn append_text(
    value: &str,
    root_seen: bool,
    root_closed: bool,
    in_settings: bool,
    stack: &mut [Frame],
) -> Result<()> {
    if !in_settings {
        if root_seen && !root_closed {
            return Ok(());
        }
        if value.trim().is_empty() {
            return Ok(());
        }
        return invalid("unexpected text outside the settings document element");
    }
    if stack.is_empty() {
        if value.trim().is_empty() {
            return Ok(());
        }
        return invalid("unexpected text outside a configuration item");
    }
    match stack.last_mut() {
        Some(Frame::Item { text, .. }) => {
            if text.len().saturating_add(value.len()) > MAX_SCALAR_BYTES {
                return invalid("configuration scalar exceeds the configured size limit");
            }
            text.push_str(value);
            Ok(())
        },
        _ if value.trim().is_empty() => Ok(()),
        _ => invalid("unexpected text in a configuration container"),
    }
}

fn semantic_attributes(
    decoder: Decoder,
    namespace_scopes: &[Vec<(Vec<u8>, Vec<u8>)>],
    start: &BytesStart<'_>,
) -> Result<Vec<(Vec<u8>, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::<Vec<u8>>::new();
    let mut count = 0usize;
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid settings attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        count += 1;
        if count > MAX_ATTRIBUTES {
            return invalid("configuration element exceeds the attribute limit");
        }
        let (namespace, local) = resolve_attribute(raw, namespace_scopes)?;
        if namespace != CONFIG {
            return invalid("configuration attributes must use the config namespace");
        }
        if local != b"name" && local != b"type" {
            return invalid("unknown configuration attribute");
        }
        if !seen.insert(local.clone()) {
            return invalid("duplicate configuration attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid settings attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_SCALAR_BYTES {
            return invalid("configuration attribute exceeds the configured size limit");
        }
        values.push((local, value));
    }
    Ok(values)
}

fn validate_no_semantic_attributes(decoder: Decoder, start: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid settings attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        count += 1;
        if count > MAX_ATTRIBUTES {
            return invalid("settings element exceeds the attribute limit");
        }
        attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid settings attribute: {error}")))
            .map(|_| ())?;
    }
    Ok(())
}

fn attribute(attributes: &[(Vec<u8>, String)], name: &[u8]) -> Option<String> {
    attributes
        .iter()
        .find(|(local, _)| local.as_slice() == name)
        .map(|(_, value)| value.clone())
}

fn required(value: Option<String>, message: &'static str) -> Result<String> {
    match value {
        Some(value) if !value.is_empty() => Ok(value),
        _ => invalid(message),
    }
}

fn parse_value(value_type: &str, text: &str) -> Result<ConfigValue> {
    let trimmed = text.trim();
    Ok(match value_type {
        "boolean" => ConfigValue::Boolean(match trimmed {
            "true" | "1" => true,
            "false" | "0" => false,
            _ => return invalid("invalid boolean configuration value"),
        }),
        "short" => ConfigValue::Short(parse_integer(trimmed, "short")?),
        "int" => ConfigValue::Int(parse_integer(trimmed, "int")?),
        "long" => ConfigValue::Long(parse_integer(trimmed, "long")?),
        "double" => ConfigValue::Double(match trimmed {
            "INF" => f64::INFINITY,
            "-INF" => f64::NEG_INFINITY,
            "NaN" => f64::NAN,
            _ => {
                let value = trimmed.parse::<f64>().map_err(|_error| {
                    Error::InvalidFormat("invalid double configuration value".to_string())
                })?;
                if !value.is_finite() {
                    return invalid("invalid double configuration value");
                }
                value
            },
        }),
        "string" => ConfigValue::String(text.to_string()),
        "datetime" => {
            if !is_datetime(trimmed) {
                return invalid("invalid datetime configuration value");
            }
            ConfigValue::DateTime(trimmed.to_string())
        },
        "base64Binary" => ConfigValue::Base64Binary(decode_base64(trimmed)?),
        _ => return invalid("unknown configuration value type"),
    })
}

fn decode_base64(value: &str) -> Result<Vec<u8>> {
    let normalized = if value.bytes().any(|byte| byte.is_ascii_whitespace()) {
        Cow::Owned(
            value
                .bytes()
                .filter(|byte| !byte.is_ascii_whitespace())
                .collect::<Vec<_>>(),
        )
    } else {
        Cow::Borrowed(value.as_bytes())
    };
    base64::engine::general_purpose::STANDARD
        .decode(normalized.as_ref())
        .map_err(|_error| Error::InvalidFormat("invalid base64 configuration value".to_string()))
}

fn parse_integer<T>(value: &str, kind: &str) -> Result<T>
where
    T: std::str::FromStr,
{
    value
        .parse::<T>()
        .map_err(|_error| Error::InvalidFormat(format!("invalid {kind} configuration value")))
}

pub(super) fn is_datetime(value: &str) -> bool {
    let Some((date, time)) = value.split_once('T') else {
        return false;
    };
    let date = date.strip_prefix('-').unwrap_or(date);
    let mut date_parts = date.split('-');
    let (Some(year), Some(month), Some(day), None) = (
        date_parts.next(),
        date_parts.next(),
        date_parts.next(),
        date_parts.next(),
    ) else {
        return false;
    };
    if year.len() < 4
        || !year.bytes().all(|byte| byte.is_ascii_digit())
        || year.bytes().all(|b| b == b'0')
    {
        return false;
    }
    let Ok(month) = month.parse::<u8>() else {
        return false;
    };
    let Ok(day) = day.parse::<u8>() else {
        return false;
    };
    if !(1..=12).contains(&month) || day == 0 || day > days_in_month(year, month) {
        return false;
    }

    let (time, timezone) = if let Some(time) = time.strip_suffix('Z') {
        (time, Some("Z"))
    } else if let Some(index) = timezone_start(time) {
        (&time[..index], Some(&time[index..]))
    } else {
        (time, None)
    };
    if timezone.is_some_and(|timezone| !valid_timezone(timezone)) {
        return false;
    }
    let mut time_parts = time.split(':');
    let (Some(hour), Some(minute), Some(second), None) = (
        time_parts.next(),
        time_parts.next(),
        time_parts.next(),
        time_parts.next(),
    ) else {
        return false;
    };
    let Ok(hour) = hour.parse::<u8>() else {
        return false;
    };
    let Ok(minute) = minute.parse::<u8>() else {
        return false;
    };
    let (second, fraction) = second
        .split_once('.')
        .map_or((second, None), |(second, fraction)| {
            (second, Some(fraction))
        });
    if fraction.is_some_and(|fraction| {
        fraction.is_empty() || !fraction.bytes().all(|byte| byte.is_ascii_digit())
    }) {
        return false;
    }
    let Ok(second) = second.parse::<u8>() else {
        return false;
    };
    if minute > 59 || second > 59 || hour > 24 {
        return false;
    }
    hour < 24
        || (minute == 0
            && second == 0
            && fraction.is_none_or(|fraction| fraction.bytes().all(|byte| byte == b'0')))
}

fn timezone_start(value: &str) -> Option<usize> {
    value
        .char_indices()
        .skip_while(|(index, _)| *index < 8)
        .find_map(|(index, character)| matches!(character, '+' | '-').then_some(index))
}

fn valid_timezone(value: &str) -> bool {
    if value == "Z" {
        return true;
    }
    let Some(value) = value.strip_prefix('+').or_else(|| value.strip_prefix('-')) else {
        return false;
    };
    let Some((hour, minute)) = value.split_once(':') else {
        return false;
    };
    if hour.len() != 2 || minute.len() != 2 {
        return false;
    }
    let (Ok(hour), Ok(minute)) = (hour.parse::<u8>(), minute.parse::<u8>()) else {
        return false;
    };
    hour < 14 && minute <= 59 || hour == 14 && minute == 0
}

fn days_in_month(year: &str, month: u8) -> u8 {
    match month {
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 31,
    }
}

fn is_leap_year(year: &str) -> bool {
    let modulo = |divisor: u16| {
        year.bytes().fold(0u16, |remainder, byte| {
            (remainder * 10 + u16::from(byte - b'0')) % divisor
        }) == 0
    };
    modulo(400) || modulo(4) && !modulo(100)
}

fn resolve_reference(name: &str) -> Result<String> {
    if let Some(value) = quick_xml::escape::resolve_xml_entity(name) {
        return Ok(value.to_string());
    }
    let codepoint =
        if let Some(hex) = name.strip_prefix("#x").or_else(|| name.strip_prefix("#X")) {
            u32::from_str_radix(hex, 16)
        } else if let Some(decimal) = name.strip_prefix('#') {
            decimal.parse::<u32>()
        } else {
            return invalid("undeclared entity reference in settings XML");
        }
        .map_err(|_error| Error::InvalidFormat("invalid XML character reference".to_string()))?;
    let character = char::from_u32(codepoint)
        .filter(|character| is_xml_character(*character))
        .ok_or_else(|| Error::InvalidFormat("invalid XML character reference".to_string()))?;
    Ok(character.to_string())
}

fn is_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || matches!(character as u32, 0x20..=0xd7ff | 0xe000..=0xfffd | 0x10000..=0x10ffff)
}

fn is_namespace(result: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(result, ResolveResult::Bound(Namespace(namespace)) if *namespace == expected)
}

fn namespace_declarations(
    start: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut declarations = Vec::new();
    let mut seen = HashSet::<Vec<u8>>::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        let prefix = if raw == b"xmlns" {
            Vec::new()
        } else if let Some(prefix) = raw.strip_prefix(b"xmlns:") {
            prefix.to_vec()
        } else {
            continue;
        };
        if prefix == b"xml" || !seen.insert(prefix.clone()) {
            return invalid("invalid or duplicate namespace declaration");
        }
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid namespace URI: {error}")))?;
        declarations.push((prefix, uri.as_bytes().to_vec()));
    }
    Ok(declarations)
}

fn resolve_attribute<'a>(
    raw: &'a [u8],
    scopes: &'a [Vec<(Vec<u8>, Vec<u8>)>],
) -> Result<(&'a [u8], Vec<u8>)> {
    let Some(colon) = raw.iter().position(|byte| *byte == b':') else {
        return invalid("configuration attributes must be namespace-qualified");
    };
    if raw[colon + 1..].contains(&b':') {
        return invalid("invalid qualified configuration attribute name");
    }
    let prefix = &raw[..colon];
    let local = raw[colon + 1..].to_vec();
    let namespace = scopes
        .iter()
        .rev()
        .flat_map(|scope| scope.iter().rev())
        .find_map(|(candidate, uri)| (candidate.as_slice() == prefix).then_some(uri.as_slice()))
        .ok_or_else(|| {
            Error::InvalidFormat("unbound configuration attribute prefix".to_string())
        })?;
    Ok((namespace, local))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
