//! Bounded style-catalog projection.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::borrow::Cow;

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
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM style XML: {error}")))?;
        let style_namespace =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE);
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe_style(&reader, style_namespace, &element, origin, definitions)?;
            },
            Event::Empty(element) => {
                let _virtual_depth = checked_depth(depth)?;
                observe_style(&reader, style_namespace, &element, origin, definitions)?;
            },
            Event::End(_) => {
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
) -> Result<()> {
    if !style_namespace || element.local_name().as_ref() != b"style" {
        return Ok(());
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
    });
    Ok(())
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
