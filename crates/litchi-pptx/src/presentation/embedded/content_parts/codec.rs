//! Bounded, namespace-aware codec for PresentationML content-part anchors.

use super::super::{
    MAX_ATTRIBUTE_BYTES, MAX_XML_ATTRIBUTES, MAX_XML_BYTES, MAX_XML_DEPTH, increment_nodes,
    invalid, is_presentationml_name, limit, relationship_value, validate_root,
};
use super::model::Anchor;
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2012/main";
const MAX_RELATIONSHIP_ID_BYTES: usize = 4 * 1024;

#[derive(Debug)]
struct Frame {
    start: usize,
    depth: usize,
    local: Vec<u8>,
}

#[derive(Debug)]
struct ContentFrame {
    start: usize,
    depth: usize,
    relationship_id: String,
}

/// Scan one slide owner and retain each active content-part anchor exactly.
pub(crate) fn scan_slide(xml: &[u8], maximum: usize) -> Result<Vec<Anchor>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("content-part slide XML bytes", MAX_XML_BYTES));
    }
    let processed = process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut frames = Vec::new();
    let mut content_frames = Vec::<ContentFrame>::new();
    let mut anchors = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let before = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                validate_attributes(&element)?;
                if depth == 0 {
                    if root_closed {
                        return Err(invalid("content-part slide has multiple roots"));
                    }
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("content-part XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("content-part XML depth", MAX_XML_DEPTH));
                }
                let local = element.name().local_name().as_ref().to_vec();
                frames.push(Frame {
                    start: before,
                    depth,
                    local,
                });
                if is_content_part(&namespace, element.name()) {
                    if anchors.len() >= maximum {
                        return Err(limit("content-part count", maximum));
                    }
                    let relationship_id = relationship_id(&element, reader.decoder(), &resolver)?;
                    content_frames.push(ContentFrame {
                        start: before,
                        depth,
                        relationship_id,
                    });
                }
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                validate_attributes(&element)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("content-part XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("content-part XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    if root_closed {
                        return Err(invalid("content-part slide has multiple roots"));
                    }
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                    root_closed = true;
                }
                if is_content_part(&namespace, element.name()) {
                    if anchors.len() >= maximum {
                        return Err(limit("content-part count", maximum));
                    }
                    let resolver = reader.resolver().clone();
                    let relationship_id = relationship_id(&element, reader.decoder(), &resolver)?;
                    anchors.push(Anchor {
                        relationship_id,
                        xml: processed[before..after].to_vec(),
                    });
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("content-part slide has an unmatched end element"));
                }
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("content-part XML frame stack underflow"))?;
                if frame.depth != depth {
                    return Err(invalid("content-part XML nesting depth is inconsistent"));
                }
                if frame.local != element.name().local_name().as_ref() {
                    return Err(invalid("content-part XML start/end names do not match"));
                }
                if content_frames
                    .last()
                    .is_some_and(|content| content.depth == depth)
                {
                    let content = content_frames
                        .pop()
                        .ok_or_else(|| invalid("content-part anchor stack underflow"))?;
                    anchors.push(Anchor {
                        relationship_id: content.relationship_id,
                        xml: processed[content.start..after].to_vec(),
                    });
                }
                depth -= 1;
                if depth == 0 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("content-part slide must close with p:sld"));
                    }
                    root_closed = true;
                }
                if frame.start >= after {
                    return Err(invalid("content-part XML frame has an invalid span"));
                }
            },
            Event::Text(value) => {
                if depth == 0
                    && value
                        .as_ref()
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(invalid("content-part slide has text outside its root"));
                }
            },
            Event::CData(value) => {
                if depth == 0
                    && value
                        .as_ref()
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(invalid("content-part slide has CDATA outside its root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "content-part slide rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 || !frames.is_empty() {
                    return Err(invalid("unterminated content-part slide"));
                }
                if !content_frames.is_empty() {
                    return Err(invalid("unterminated content-part anchor"));
                }
                return Ok(anchors);
            },
            _ => {},
        }
    }
}

fn relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<String> {
    let value = relationship_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| invalid("PresentationML contentPart is missing r:id"))?;
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("PresentationML contentPart has an invalid r:id"));
    }
    Ok(value)
}

fn validate_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| limit("content-part XML attributes", MAX_XML_ATTRIBUTES))?;
        if count > MAX_XML_ATTRIBUTES {
            return Err(limit("content-part XML attributes", MAX_XML_ATTRIBUTES));
        }
        if attribute.key.as_ref().len() > MAX_ATTRIBUTE_BYTES
            || attribute.value.len() > MAX_ATTRIBUTE_BYTES
        {
            return Err(limit(
                "content-part XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
            ));
        }
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("content-part XML offset does not fit usize"))
}

fn is_content_part(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    if name.local_name().as_ref() != b"contentPart" {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == super::super::PML
                || *value == super::super::STRICT_PML
                || *value == P14
                || *value == P15
        },
        ResolveResult::Unknown(prefix) => {
            prefix.as_slice() == b"p" || prefix.as_slice() == b"p14" || prefix.as_slice() == b"p15"
        },
        ResolveResult::Unbound => false,
    }
}
