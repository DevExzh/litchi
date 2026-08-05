//! Shared XML boundary helpers for the data-pilot parser and writer.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesEnd, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::super::invalid_message;

pub(super) const TABLE_NAMESPACE: &[u8] = super::super::TABLE_NAMESPACE;
pub(super) const OFFICE_NAMESPACE: &[u8] = super::super::OFFICE_NAMESPACE;
pub(super) const TABLE_EXT_NAMESPACE: &[u8] = super::super::TABLE_EXT_NAMESPACE;
pub(super) const CALC_EXT_NAMESPACE: &[u8] = super::super::CALC_EXT_NAMESPACE;

pub(super) trait HasLocalName {
    fn local(&self) -> &[u8];
}

impl HasLocalName for BytesStart<'_> {
    fn local(&self) -> &[u8] {
        self.local_name().into_inner()
    }
}

impl HasLocalName for BytesEnd<'_> {
    fn local(&self) -> &[u8] {
        self.local_name().into_inner()
    }
}

pub(super) fn is_table(
    namespace: &ResolveResult<'_>,
    element: &impl HasLocalName,
    local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE)
        && element.local() == local
}

pub(super) fn is_office(
    namespace: &ResolveResult<'_>,
    element: &impl HasLocalName,
    local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE)
        && element.local() == local
}

pub(super) fn is_table_ext(
    namespace: &ResolveResult<'_>,
    element: &impl HasLocalName,
    local: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_EXT_NAMESPACE)
        && element.local() == local
}

pub(super) fn is_table_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE)
}

pub(super) fn consume_empty_extension(reader: &mut NsReader<&[u8]>, local: &[u8]) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::End(ref element) if is_table_ext(&namespace, element, local) => return Ok(()),
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated data-pilot extension element")),
            _ => {
                return Err(invalid_message(
                    "data-pilot grand-total extension must be empty",
                ));
            },
        }
        buffer.clear();
    }
}

pub(super) fn skip_foreign_element(
    reader: &mut NsReader<&[u8]>,
    _start: &BytesStart<'_>,
) -> Result<()> {
    let mut depth = 1usize;
    let mut buffer = Vec::new();
    while depth > 0 {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_message("data-pilot extension depth overflow"))?;
                if depth > 128 {
                    return Err(invalid_message(
                        "data-pilot extension nesting exceeds limit",
                    ));
                }
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) => {
                return Err(invalid_message(
                    "DOCTYPE is not allowed in data-pilot extensions",
                ));
            },
            Event::Eof => return Err(invalid_message("unterminated data-pilot extension")),
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

pub(super) fn optional_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_message(&format!("invalid XML attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TABLE_NAMESPACE)
            && name.as_ref() == local
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_message(&format!("invalid XML attribute value: {error}")))?
                .into_owned();
            if found.replace(value).is_some() {
                return Err(invalid_message("duplicate table attribute"));
            }
        }
    }
    Ok(found)
}

pub(super) fn optional_ns_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    wanted_namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_message(&format!("invalid XML attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == wanted_namespace)
            && name.as_ref() == local
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_message(&format!("invalid XML attribute value: {error}")))?
                .into_owned();
            if found.replace(value).is_some() {
                return Err(invalid_message("duplicate extension attribute"));
            }
        }
    }
    Ok(found)
}

pub(super) fn optional_ns_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<bool>> {
    optional_ns_attr(reader, element, namespace, local)?
        .map(|value| parse_bool(&value))
        .transpose()
}

pub(super) fn required_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    optional_attr(reader, element, local)?.ok_or_else(|| {
        invalid_message(&format!("missing table:{}", String::from_utf8_lossy(local)))
    })
}

pub(super) fn optional_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<bool>> {
    optional_attr(reader, element, local)?
        .map(|value| parse_bool(&value))
        .transpose()
}

pub(super) fn required_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<bool> {
    parse_bool(&required_attr(reader, element, local)?)
}

pub(super) fn optional_i64(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<i64>> {
    optional_attr(reader, element, local)?
        .map(|value| value.parse().map_err(|_| invalid("integer", &value)))
        .transpose()
}

pub(super) fn required_u64(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<u64> {
    let value = required_attr(reader, element, local)?;
    value
        .parse()
        .map_err(|_| invalid("non-negative integer", &value))
}

pub(super) fn required_f64(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<f64> {
    let value = required_attr(reader, element, local)?;
    value.parse().map_err(|_| invalid("number", &value))
}

pub(super) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("Boolean", value)),
    }
}

pub(super) fn text_is_whitespace(text: &quick_xml::events::BytesText<'_>) -> Result<bool> {
    Ok(text
        .xml_content(XmlVersion::Explicit1_0)
        .map_err(|error| invalid_message(&format!("invalid XML text: {error}")))?
        .trim()
        .is_empty())
}

pub(super) fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid_message(&format!("duplicate {name}")))
    } else {
        Ok(())
    }
}

fn invalid(kind: &str, value: &str) -> Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

fn xml_error(error: quick_xml::Error) -> Error {
    invalid_message(&format!("XML parsing error: {error}"))
}
