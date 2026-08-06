//! Bounded SpreadsheetML XML wire helpers for data validation.

use super::super::model::{Range, Source, Sqref};
use super::super::{
    CORE, EXTENSION_URI, MAX_ATTRIBUTE_BYTES, MAX_CAPTURED_COLLECTIONS, MAX_DEPTH, MAX_EVENTS,
    MAX_FRAGMENT_BYTES, MAX_NODES, MAX_REFERENCES, MAX_RETAINED_BYTES, MAX_XML_BYTES, STRICT, X14,
    XR,
};
use crate::error::{Error, Result};
use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::Writer;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::TryReserveError;
use std::fmt;

pub(super) struct Captured {
    pub(super) source: Source,
    pub(super) prefix: Vec<u8>,
    pub(super) bytes: Vec<u8>,
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

pub(crate) fn reserve_vec<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> Result<()> {
    values
        .try_reserve_exact(additional)
        .map_err(|source| allocation(resource, source))
}

pub(super) fn append_limited_text(
    value: &mut String,
    addition: &str,
    limit: usize,
    field: &str,
) -> Result<()> {
    let length = value
        .len()
        .checked_add(addition.len())
        .ok_or_else(|| invalid(format!("{field} length overflow")))?;
    if length > limit {
        return Err(invalid(format!("{field} is too large")));
    }
    value
        .try_reserve_exact(addition.len())
        .map_err(|source| allocation("data-validation text", source))?;
    value.push_str(addition);
    Ok(())
}

/// A fallible, bounded formatter used by the XML writer.
pub(crate) struct BoundedXml {
    value: String,
    allocation: Option<TryReserveError>,
    exceeded: bool,
}

impl BoundedXml {
    pub(crate) fn new() -> Self {
        Self {
            value: String::new(),
            allocation: None,
            exceeded: false,
        }
    }

    pub(crate) fn write_arguments(&mut self, arguments: fmt::Arguments<'_>) -> Result<()> {
        if fmt::write(self, arguments).is_ok() {
            return Ok(());
        }
        if let Some(source) = self.allocation.take() {
            return Err(allocation("data-validation XML output", source));
        }
        if self.exceeded {
            Err(invalid("data-validation XML output exceeds resource limit"))
        } else {
            Err(invalid("failed to format data-validation XML"))
        }
    }

    pub(crate) fn push_str(&mut self, value: &str) -> Result<()> {
        self.write_arguments(format_args!("{value}"))
    }

    pub(crate) fn finish(self) -> String {
        self.value
    }
}

impl fmt::Write for BoundedXml {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        let length = self
            .value
            .len()
            .checked_add(value.len())
            .ok_or(fmt::Error)?;
        if length > MAX_XML_BYTES {
            self.exceeded = true;
            return Err(fmt::Error);
        }
        if let Err(source) = self.value.try_reserve_exact(value.len()) {
            self.allocation = Some(source);
            return Err(fmt::Error);
        }
        self.value.push_str(value);
        Ok(())
    }
}

pub(crate) fn append_bounded_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    let length = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("data-validation XML output length overflow"))?;
    if length > MAX_XML_BYTES {
        return Err(invalid("data-validation XML output exceeds resource limit"));
    }
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation("data-validation XML output", source))?;
    output.extend_from_slice(bytes);
    Ok(())
}

type CaptureState = Option<(usize, Source, Vec<u8>, Writer<Vec<u8>>)>;

fn retain_capture(
    values: &mut Vec<Captured>,
    retained: &mut usize,
    captured: Captured,
) -> Result<()> {
    if values.len() >= MAX_CAPTURED_COLLECTIONS {
        return Err(invalid("too many data-validation collections"));
    }
    let size = captured
        .prefix
        .len()
        .checked_add(captured.bytes.len())
        .ok_or_else(|| invalid("data-validation retained-byte overflow"))?;
    *retained = retained
        .checked_add(size)
        .ok_or_else(|| invalid("data-validation retained-byte overflow"))?;
    if *retained > MAX_RETAINED_BYTES {
        return Err(invalid("data-validation content exceeds resource limit"));
    }
    reserve_vec(values, 1, "data-validation collections")?;
    values.push(captured);
    Ok(())
}

pub(super) fn capture_collections(xml: &[u8]) -> Result<Vec<Captured>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut values = Vec::new();
    let mut depth = 0usize;
    let mut extension_depth = None;
    let mut capture: CaptureState = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    let mut retained = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("data-validation XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("data-validation XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            if capture.is_some() || depth != 0 {
                return Err(invalid("unterminated data-validation worksheet XML"));
            }
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-validation XML node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-validation XML exceeds node limit"));
            }
        }
        if let Some((capture_depth, _, _, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                return Err(invalid("dataValidations fragment is too large"));
            }
            match event {
                Event::Start(_) => {
                    if *capture_depth >= MAX_DEPTH {
                        return Err(invalid("dataValidations nesting is too deep"));
                    }
                    *capture_depth = capture_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidations nesting overflow"))?;
                },
                Event::End(_) => {
                    *capture_depth = capture_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid dataValidations nesting"))?;
                },
                _ => {},
            }
            if *capture_depth == 0 {
                let Some((_, source, prefix, writer)) = capture.take() else {
                    return Err(invalid("dataValidations capture state disappeared"));
                };
                let bytes = writer.into_inner();
                retain_capture(
                    &mut values,
                    &mut retained,
                    Captured {
                        source,
                        prefix,
                        bytes,
                    },
                )?;
            }
            continue;
        }
        match event {
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth > 0
                    && depth == 1
                    && spreadsheet(&namespace) =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                capture = Some((1, Source::Core, prefix, writer));
            },
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth > 0
                    && exact(&namespace, X14)
                    && extension_depth.is_some() =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                capture = Some((1, Source::Office2010, prefix, writer));
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth == 1
                    && spreadsheet(&namespace) =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                retain_capture(
                    &mut values,
                    &mut retained,
                    Captured {
                        source: Source::Core,
                        prefix,
                        bytes: writer.into_inner(),
                    },
                )?;
            },
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if depth == 0 {
                    if root_seen
                        || !spreadsheet(&namespace)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("data-validation parser requires a worksheet root"));
                    }
                    root_seen = true;
                }
                if depth >= MAX_DEPTH {
                    return Err(invalid("worksheet nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet nesting overflow"))?;
                if spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"ext"
                    && optional_attr(&element, b"uri", decoder)?.as_deref() == Some(EXTENSION_URI)
                {
                    extension_depth = Some(depth);
                }
            },
            Event::Empty(element) if depth == 0 => {
                if root_seen
                    || !spreadsheet(&namespace)
                    || element.local_name().as_ref() != b"worksheet"
                {
                    return Err(invalid("data-validation parser requires a worksheet root"));
                }
                root_seen = true;
                root_closed = true;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid worksheet nesting"));
                }
                if depth == 1
                    && (!spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet")
                {
                    return Err(invalid("invalid worksheet closing element"));
                }
                if extension_depth == Some(depth) {
                    extension_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid worksheet nesting"))?;
                if depth == 0 {
                    // The namespace/local-name check is performed by the XML reader for
                    // qualified names; this branch only records the root boundary.
                    root_closed = true;
                }
            },
            Event::Text(value) => {
                if (!root_seen || root_closed)
                    && !value.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if depth == 1 && !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) if depth == 1 || !root_seen || root_closed => {
                return Err(invalid("worksheet XML contains unexpected CDATA"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::GeneralRef(reference) => {
                decode_xml_reference(&reference)?;
            },
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(invalid("incomplete worksheet data-validation XML"));
    }
    Ok(values)
}

pub(super) fn sqref_text(value: &Sqref) -> Result<String> {
    let mut text = String::new();
    for (index, range) in value.ranges.iter().enumerate() {
        if index != 0 {
            append_limited_text(&mut text, " ", MAX_FRAGMENT_BYTES, "data-validation sqref")?;
        }
        append_limited_text(
            &mut text,
            range.0.as_str(),
            MAX_FRAGMENT_BYTES,
            "data-validation sqref",
        )?;
    }
    Ok(text)
}

pub(crate) fn parse_sqref(
    value: &str,
    edited: bool,
    split: bool,
    adjusted: bool,
    adjust: bool,
) -> Result<Sqref> {
    if value.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("data-validation sqref is too large"));
    }
    if adjusted && !adjust {
        return Err(invalid("sqref adjusted requires adjust"));
    }
    let mut ranges = Vec::new();
    for item in value.split_whitespace() {
        if ranges.len() == MAX_REFERENCES {
            return Err(invalid("too many data-validation references"));
        }
        let mut parts = item.split(':');
        let Some(first) = parts.next() else {
            return Err(invalid("invalid empty data-validation range"));
        };
        let second = parts.next();
        if parts.next().is_some() || !valid_cell(first) || second.is_some_and(|v| !valid_cell(v)) {
            return Err(invalid(format!("invalid data-validation range '{item}'")));
        }
        reserve_vec(&mut ranges, 1, "data-validation references")?;
        ranges.push(Range(item.to_owned()));
    }
    if ranges.is_empty() {
        return Err(invalid("data-validation sqref is empty"));
    }
    Ok(Sqref {
        ranges,
        edited,
        split,
        adjusted,
        adjust,
    })
}

fn valid_cell(value: &str) -> bool {
    let raw = value.as_bytes();
    let mut i = 0;
    if i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let start = i;
    while i < raw.len() && raw[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return false;
    }
    let mut col = 0u32;
    for b in &raw[start..i] {
        col = col
            .saturating_mul(26)
            .saturating_add(u32::from(b.to_ascii_uppercase() - b'A' + 1));
    }
    if !(1..=16_384).contains(&col) {
        return false;
    }
    if i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let Ok(row_text) = std::str::from_utf8(&raw[i..]) else {
        return false;
    };
    let Ok(row) = row_text.parse::<u32>() else {
        return false;
    };
    (1..=1_048_576).contains(&row)
}

pub(super) fn sqref_flags(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<(bool, bool, bool, bool)> {
    Ok((
        optional_bool(element, b"edited", decoder)?.unwrap_or(false),
        optional_bool(element, b"split", decoder)?.unwrap_or(false),
        optional_bool(element, b"adjusted", decoder)?.unwrap_or(false),
        optional_bool(element, b"adjust", decoder)?.unwrap_or(false),
    ))
}
pub(super) fn encode_flags(v: (bool, bool, bool, bool)) -> String {
    format!("{}{}{}{}|", v.0 as u8, v.1 as u8, v.2 as u8, v.3 as u8)
}
pub(super) fn decode_flags(value: &str) -> Result<((bool, bool, bool, bool), String)> {
    let bytes = value.as_bytes();
    if bytes.len() < 5 || bytes[4] != b'|' {
        return Err(invalid("invalid sqref state"));
    }
    Ok((
        (
            bytes[0] == b'1',
            bytes[1] == b'1',
            bytes[2] == b'1',
            bytes[3] == b'1',
        ),
        value[5..].to_owned(),
    ))
}

pub(super) fn uid_attr(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        if attribute.key.local_name().as_ref() == b"uid"
            && (exact(&attribute_namespace, XR) || attribute.key.as_ref() == b"xr:uid")
        {
            if result.is_some() {
                return Err(invalid("duplicate data-validation uid"));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            if value.len() > MAX_ATTRIBUTE_BYTES {
                return Err(invalid("data-validation uid is too large"));
            }
            if !valid_guid(&value) {
                return Err(invalid("invalid data-validation uid"));
            }
            result = Some(value);
        }
    }
    Ok(result)
}
pub(crate) fn optional_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == name {
            if result.is_some() {
                return Err(invalid(format!(
                    "duplicate '{}' attribute",
                    String::from_utf8_lossy(name)
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            if value.len() > MAX_ATTRIBUTE_BYTES {
                return Err(invalid(format!(
                    "data-validation attribute '{}' is too large",
                    String::from_utf8_lossy(name)
                )));
            }
            result = Some(value);
        }
    }
    Ok(result)
}
pub(super) fn required_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<String> {
    optional_attr(element, name, decoder)?.ok_or_else(|| {
        invalid(format!(
            "missing '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}
pub(super) fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<u32>> {
    optional_attr(element, name, decoder)?
        .map(|v| {
            v.parse()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
pub(super) fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<bool>> {
    optional_attr(element, name, decoder)?
        .map(|v| match v.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
pub(super) fn bounded_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    max: usize,
) -> Result<Option<String>> {
    let value = optional_attr(element, name, decoder)?;
    if value.as_ref().is_some_and(|v| v.chars().count() > max) {
        return Err(invalid(format!(
            "{} exceeds {max} characters",
            String::from_utf8_lossy(name)
        )));
    }
    Ok(value)
}

pub(super) fn wrap(prefix: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("data-validation fragment is too large"));
    }
    let mut out = Vec::new();
    append_bounded_bytes(
        &mut out,
        br#"<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision""#,
    )?;
    if !prefix.is_empty() && !matches!(prefix, b"s" | b"x14") {
        append_bounded_bytes(&mut out, b" xmlns:")?;
        append_bounded_bytes(&mut out, prefix)?;
        append_bounded_bytes(
            &mut out,
            if prefix == b"x" {
                b"=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
            } else {
                b"=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\""
            },
        )?;
    }
    append_bounded_bytes(&mut out, b">")?;
    append_bounded_bytes(&mut out, fragment)?;
    append_bounded_bytes(&mut out, b"</root>")?;
    Ok(out)
}
fn prefix(name: &[u8]) -> Result<Vec<u8>> {
    let Some(index) = name.iter().position(|value| *value == b':') else {
        return Ok(Vec::new());
    };
    let mut value = Vec::new();
    reserve_vec(&mut value, index, "data-validation namespace prefix")?;
    value.extend_from_slice(&name[..index]);
    Ok(value)
}
pub(super) fn source_ns(source: Source, ns: &ResolveResult<'_>) -> bool {
    match source {
        Source::Core => spreadsheet(ns),
        Source::Office2010 => exact(ns, X14),
    }
}
pub(crate) fn spreadsheet(ns: &ResolveResult<'_>) -> bool {
    exact(ns, CORE) || exact(ns, STRICT)
}
pub(crate) fn exact(ns: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(ns,ResolveResult::Bound(value)if value.as_ref()==expected)
}
pub(crate) fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
