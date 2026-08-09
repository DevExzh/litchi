//! Bounded style-catalog projection.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::borrow::Cow;
use std::ops::Range;

use crate::style::{Definition, Origin};

const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 1_000_000;
const MAX_VALUE_BYTES: usize = 16 * 1024;

pub(crate) fn parse_catalog(content: &str, styles: Option<&str>) -> Result<Vec<Definition>> {
    let mut definitions = Vec::new();
    parse_part(content, Origin::Content, &mut definitions)?;
    if let Some(xml) = styles {
        parse_part(xml, Origin::Styles, &mut definitions)?;
    }
    Ok(definitions)
}

fn parse_part(xml: &str, origin: Origin, definitions: &mut Vec<Definition>) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("ODM style part exceeds the family limit"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut active = Vec::new();
    loop {
        let event_start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM style XML: {error}")))?;
        let style_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE);
        let event = event.into_owned();
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                if let Some(index) = observe_style(
                    &reader,
                    style_namespace,
                    &element,
                    origin,
                    definitions,
                    event_start..event_end,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM style event span is outside its XML part"))?,
                )? {
                    active.push((depth, index));
                }
            },
            Event::Empty(element) => {
                let _virtual_depth = checked_depth(depth)?;
                let _definition = observe_style(
                    &reader,
                    style_namespace,
                    &element,
                    origin,
                    definitions,
                    event_start..event_end,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM style event span is outside its XML part"))?,
                )?;
            },
            Event::End(_) => {
                if active
                    .last()
                    .is_some_and(|(style_depth, _)| *style_depth == depth)
                {
                    let (_style_depth, index) = active
                        .pop()
                        .ok_or_else(|| invalid("ODM style nesting underflow"))?;
                    definitions
                        .get_mut(index)
                        .ok_or_else(|| invalid("ODM style definition disappeared"))?
                        .source_span
                        .end = event_end;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODM style XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM style XML")),
            Event::GeneralRef(_) => {
                return Err(invalid(
                    "named XML entities are not allowed in ODM style XML",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if depth != 0 {
        return Err(invalid("ODM style XML nesting is incomplete"));
    }
    Ok(())
}

fn observe_style(
    reader: &NsReader<&[u8]>,
    style_namespace: bool,
    element: &BytesStart<'_>,
    origin: Origin,
    definitions: &mut Vec<Definition>,
    source_span: Range<usize>,
    tag: &[u8],
) -> Result<Option<usize>> {
    if !style_namespace || element.local_name().as_ref() != b"style" {
        return Ok(None);
    }
    let Some(name) = attribute(reader, element, b"name")? else {
        return Err(invalid("ODM style:style has no style:name"));
    };
    let family = attribute(reader, element, b"family")?;
    let parent = attribute(reader, element, b"parent-style-name")?;
    for (value, scope) in [
        (Some(name.as_str()), "ODM style name"),
        (family.as_deref(), "ODM style family"),
        (parent.as_deref(), "ODM parent style name"),
    ] {
        if value.is_some_and(|value| value.len() > MAX_VALUE_BYTES) {
            return Err(invalid(format!("{scope} exceeds the 16 KiB limit")));
        }
    }
    if definitions.len() >= MAX_STYLES {
        return Err(invalid("ODM style count exceeds the limit"));
    }
    definitions
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "ODM style catalog",
            source,
        })?;
    definitions.push(Definition {
        name,
        family,
        parent,
        origin,
        source_span,
        name_span: {
            let key = attribute_key(reader, element, b"name")?
                .ok_or_else(|| invalid("ODM style name source spelling disappeared"))?;
            let (start, end) = attribute_value_span(tag, &key)?;
            start..end
        },
    });
    let definition = definitions.len() - 1;
    let tag_start = definitions[definition].source_span.start;
    definitions[definition].name_span.start += tag_start;
    definitions[definition].name_span.end += tag_start;
    Ok(Some(definition))
}

fn attribute_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<Vec<u8>>> {
    let mut key = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODM style attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE)
            && name.as_ref() == local
            && key.replace(attribute.key.as_ref().to_vec()).is_some()
        {
            return Err(invalid(
                "duplicate namespace-equivalent ODM style attribute",
            ));
        }
    }
    Ok(key)
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("ODM style attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("ODM style attribute is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid("ODM style attribute is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    Err(invalid("ODM style attribute source span is missing"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_range_error| invalid("ODM style source position exceeds the platform range"))
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODM style attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE)
            && name.as_ref() == local
        {
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(Cow::into_owned)
                .map_err(|error| invalid(format!("invalid ODM style value: {error}")))?;
            if value.replace(decoded).is_some() {
                return Err(invalid(
                    "duplicate namespace-equivalent ODM style attribute",
                ));
            }
        }
    }
    Ok(value)
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODM style XML depth overflow"))?;
    if next > MAX_DEPTH {
        return Err(invalid("ODM style XML depth exceeds the limit"));
    }
    Ok(next)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
