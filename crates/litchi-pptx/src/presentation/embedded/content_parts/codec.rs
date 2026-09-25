//! Bounded, namespace-aware codec for `PresentationML` content-part anchors.

use super::super::{
    MAX_ATTRIBUTE_BYTES, MAX_XML_ATTRIBUTES, MAX_XML_BYTES, MAX_XML_DEPTH, increment_nodes,
    invalid, is_presentationml_name, limit, relationship_value, validate_root,
};
use super::model::{Anchor, BlackWhiteMode};
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;

const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_URI: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2012/main";
const MAX_RELATIONSHIP_ID_BYTES: usize = 4 * 1024;

/// One source span for a raw `p:contentPart` anchor.
#[derive(Debug, Clone)]
pub(crate) struct SourceAnchor {
    pub(crate) range: Range<usize>,
    pub(crate) relationship_id: String,
    pub(crate) relationship_span: Range<usize>,
}

#[derive(Debug)]
struct SourceFrame {
    start: usize,
    depth: usize,
    local: Vec<u8>,
    content: Option<(String, Range<usize>)>,
}

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
    bw_mode: Option<BlackWhiteMode>,
    bw_mode_key: Option<Box<[u8]>>,
}

/// Scan one slide owner and retain each active content-part anchor exactly.
pub(crate) fn scan_slide(xml: &[u8], maximum: usize) -> Result<Vec<Anchor>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("content-part slide XML bytes", MAX_XML_BYTES));
    }
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14_URI);
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    let origin = ReaderOrigin::of(processed.as_ref());
    let mut frames = Vec::new();
    let mut content_frames = Vec::<ContentFrame>::new();
    let mut anchors = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let before = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, resolved_event) = resolver.resolve_event(event);

        match resolved_event {
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
                    let (bw_mode, bw_mode_key) =
                        black_white_mode(&element, reader.decoder(), &resolver)?;
                    content_frames.push(ContentFrame {
                        start: before,
                        depth,
                        relationship_id,
                        bw_mode,
                        bw_mode_key,
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
                    let relationship_id = relationship_id(&element, reader.decoder(), &resolver)?;
                    let (bw_mode, bw_mode_key) =
                        black_white_mode(&element, reader.decoder(), &resolver)?;
                    anchors.push(Anchor {
                        relationship_id,
                        xml: processed[before..after].to_vec(),
                        bw_mode,
                        bw_mode_key,
                        bw_mode_namespace_added: false,
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
                        bw_mode: content.bw_mode,
                        bw_mode_key: content.bw_mode_key,
                        bw_mode_namespace_added: false,
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

/// Locate every raw content-part anchor, including equivalent MCE choice and
/// fallback branches. The transaction layer uses these spans to edit only the
/// selected relationship attribute while retaining branch-local opaque XML.
pub(crate) fn locate_content_parts(xml: &[u8]) -> Result<Vec<SourceAnchor>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("content-part slide XML bytes", MAX_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::new();
    let mut anchors = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let before = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = position(&reader, origin)?;
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
                let content = if is_content_part(&namespace, element.name()) {
                    let id = relationship_id(&element, decoder, &resolver)?;
                    let span = relationship_value_span(&element, before, decoder, &resolver)?;
                    Some((id, span))
                } else {
                    None
                };
                frames.push(SourceFrame {
                    start: before,
                    depth,
                    local: element.name().local_name().as_ref().to_vec(),
                    content,
                });
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
                    let relationship_id = relationship_id(&element, decoder, &resolver)?;
                    let relationship_span =
                        relationship_value_span(&element, before, decoder, &resolver)?;
                    anchors.push(SourceAnchor {
                        range: before..after,
                        relationship_id,
                        relationship_span,
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
                if frame.depth != depth
                    || frame.local != element.name().local_name().as_ref()
                    || frame.start >= after
                {
                    return Err(invalid("content-part XML nesting is inconsistent"));
                }
                if let Some((relationship_id, relationship_span)) = frame.content {
                    anchors.push(SourceAnchor {
                        range: frame.start..after,
                        relationship_id,
                        relationship_span,
                    });
                }
                depth -= 1;
                if depth == 0 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("content-part slide must close with p:sld"));
                    }
                    root_closed = true;
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
                return Ok(anchors);
            },
            _ => {},
        }
    }
}

/// Validate one detached anchor fragment and return its relationship ID.
pub(crate) fn validate_anchor_xml_with_key(xml: &[u8], key: Option<&[u8]>) -> Result<String> {
    if xml.is_empty() || xml.len() > MAX_XML_BYTES {
        return Err(invalid("content-part anchor XML is empty or too large"));
    }
    let prefix = detached_slide_prefix(key)?;
    let mut wrapped = Vec::with_capacity(prefix.len() + xml.len() + 8);
    wrapped.extend_from_slice(prefix.as_bytes());
    let offset = wrapped.len();
    wrapped.extend_from_slice(xml);
    wrapped.extend_from_slice(b"</p:sld>");
    let anchors = locate_content_parts(&wrapped)?;
    if anchors.len() != 1 {
        return Err(invalid(
            "content-part anchor must contain exactly one contentPart element",
        ));
    }
    let anchor = &anchors[0];
    if anchor.range.start != offset
        || anchor.range.end != offset.saturating_add(xml.len())
        || wrapped.get(anchor.range.clone()) != Some(xml)
    {
        return Err(invalid(
            "content-part anchor XML contains unsupported roots",
        ));
    }
    Ok(anchor.relationship_id.clone())
}

/// Read the typed `p14:bwMode` attribute from one detached anchor.
pub(crate) fn anchor_bw_mode(xml: &[u8]) -> Result<(Option<BlackWhiteMode>, Option<Box<[u8]>>)> {
    anchor_bw_mode_with_declaration(xml, None)
}

fn anchor_bw_mode_with_declaration(
    xml: &[u8],
    key: Option<&[u8]>,
) -> Result<(Option<BlackWhiteMode>, Option<Box<[u8]>>)> {
    if xml.is_empty() || xml.len() > MAX_XML_BYTES {
        return Err(invalid("content-part anchor XML is empty or too large"));
    }
    let prefix = detached_slide_prefix(key)?;
    let mut wrapped = Vec::with_capacity(prefix.len() + xml.len() + 8);
    wrapped.extend_from_slice(prefix.as_bytes());
    wrapped.extend_from_slice(xml);
    wrapped.extend_from_slice(b"</p:sld>");
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    let mut found = None;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_content_part(&namespace, element.name()) =>
            {
                if found.is_some() {
                    return Err(invalid(
                        "content-part anchor contains multiple contentPart elements",
                    ));
                }
                found = Some(black_white_mode(&element, reader.decoder(), &resolver)?);
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "content-part anchor rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    found.ok_or_else(|| invalid("content-part anchor has no contentPart element"))
}

fn detached_slide_prefix(key: Option<&[u8]>) -> Result<String> {
    let mut prefix = format!(
        "<p:sld xmlns:p=\"{}\" xmlns:r=\"{}\" xmlns:p14=\"{}\" xmlns:p15=\"{}\"",
        String::from_utf8_lossy(super::super::PML),
        String::from_utf8_lossy(super::super::REL),
        String::from_utf8_lossy(P14),
        String::from_utf8_lossy(P15),
    );
    if let Some(key) = key {
        let prefix_name = key
            .strip_suffix(b":bwMode")
            .filter(|prefix| !prefix.is_empty() && !prefix.contains(&b':'))
            .ok_or_else(|| invalid("content-part p14:bwMode attribute name is invalid"))?;
        let prefix_name = std::str::from_utf8(prefix_name)
            .map_err(|_error| invalid("content-part p14:bwMode prefix is not UTF-8"))?;
        if prefix_name != "p14" {
            prefix.push_str(" xmlns:");
            prefix.push_str(prefix_name);
            prefix.push_str("=\"");
            prefix.push_str(
                std::str::from_utf8(P14)
                    .map_err(|_error| invalid("content-part p14 namespace is not UTF-8"))?,
            );
            prefix.push('"');
        }
    }
    prefix.push('>');
    Ok(prefix)
}

pub(crate) fn anchor_bw_mode_with_key(
    xml: &[u8],
    key: Option<&[u8]>,
) -> Result<(Option<BlackWhiteMode>, Option<Box<[u8]>>)> {
    let parsed = anchor_bw_mode_with_declaration(xml, key)?;
    if parsed.0.is_some() || key.is_none() {
        return Ok(parsed);
    }
    let (start, end) = content_part_start_tag(xml)?;
    let tag = &xml[start..end];
    let Some(attribute) = raw_attribute_span(tag, key.unwrap_or_default()) else {
        return Ok((None, None));
    };
    let value = std::str::from_utf8(&tag[attribute.value])
        .map_err(|_err| invalid("content-part p14:bwMode value is not UTF-8"))?;
    let value =
        quick_xml::escape::unescape(value).map_err(|error| Error::Xml(error.to_string()))?;
    Ok((
        Some(BlackWhiteMode::parse(&value)?),
        Some(key.unwrap_or_default().into()),
    ))
}

/// Rewrite only the typed `p14:bwMode` attribute in one detached anchor.
pub(crate) fn rewrite_anchor_bw_mode(
    xml: &[u8],
    current_key: Option<&[u8]>,
    namespace_added: bool,
    value: Option<BlackWhiteMode>,
) -> Result<(Vec<u8>, Option<Box<[u8]>>, bool)> {
    let (parsed, parsed_key) = anchor_bw_mode_with_key(xml, current_key)?;
    let key = parsed_key.or_else(|| current_key.map(Into::into));
    if parsed.is_some() && key.is_none() {
        return Err(invalid("content-part p14:bwMode attribute name is missing"));
    }
    if value.is_none() && key.is_none() {
        return Ok((xml.to_vec(), None, false));
    }

    let (start, end) = content_part_start_tag(xml)?;
    let tag = &xml[start..end];
    let attribute = key.as_deref().and_then(|key| raw_attribute_span(tag, key));
    let mut output = xml.to_vec();
    let mut namespace_added = namespace_added;
    match (value, attribute) {
        (Some(value), Some(attribute)) => {
            let begin = start
                .checked_add(attribute.value.start)
                .ok_or_else(|| invalid("content-part p14:bwMode span overflow"))?;
            let finish = start
                .checked_add(attribute.value.end)
                .ok_or_else(|| invalid("content-part p14:bwMode span overflow"))?;
            output.splice(begin..finish, value.as_str().as_bytes().iter().copied());
        },
        (None, Some(attribute)) => {
            let begin = start
                .checked_add(attribute.full.start)
                .ok_or_else(|| invalid("content-part p14:bwMode span overflow"))?;
            let finish = start
                .checked_add(attribute.full.end)
                .ok_or_else(|| invalid("content-part p14:bwMode span overflow"))?;
            output.splice(begin..finish, std::iter::empty());
            if namespace_added {
                let (namespace_start, namespace_end) = content_part_start_tag(&output)?;
                let namespace_tag = &output[namespace_start..namespace_end];
                if let Some(namespace_attr) = raw_attribute_span(namespace_tag, b"xmlns:p14") {
                    let begin = namespace_start
                        .checked_add(namespace_attr.full.start)
                        .ok_or_else(|| invalid("content-part p14 namespace span overflow"))?;
                    let finish = namespace_start
                        .checked_add(namespace_attr.full.end)
                        .ok_or_else(|| invalid("content-part p14 namespace span overflow"))?;
                    output.splice(begin..finish, std::iter::empty());
                }
                namespace_added = false;
            }
        },
        (Some(value), None) => {
            let mut insertion = end
                .checked_sub(1)
                .ok_or_else(|| invalid("content-part start tag is truncated"))?;
            while insertion > start && output[insertion - 1].is_ascii_whitespace() {
                insertion -= 1;
            }
            if insertion > start && output[insertion - 1] == b'/' {
                insertion -= 1;
            }
            let mut addition = format!(" p14:bwMode=\"{}\"", value.as_str());
            if raw_attribute_span(tag, b"xmlns:p14").is_none() {
                addition.push_str(" xmlns:p14=\"");
                addition.push_str(
                    std::str::from_utf8(P14)
                        .map_err(|_err| invalid("content-part p14 namespace is not UTF-8"))?,
                );
                addition.push('"');
                namespace_added = true;
            }
            output.splice(insertion..insertion, addition.into_bytes());
        },
        (None, None) => {},
    }
    if output.len() > MAX_XML_BYTES {
        return Err(limit("content-part anchor XML bytes", MAX_XML_BYTES));
    }
    let (result, result_key) = anchor_bw_mode_with_key(&output, key.as_deref())?;
    if result != value {
        return Err(invalid(
            "content-part p14:bwMode rewrite did not round-trip",
        ));
    }
    Ok((output, result_key, namespace_added))
}

struct RawAttributeSpan {
    full: Range<usize>,
    value: Range<usize>,
}

fn content_part_start_tag(xml: &[u8]) -> Result<(usize, usize)> {
    let start = xml
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or_else(|| invalid("content-part anchor start tag is missing"))?;
    let mut quote = None;
    for (offset, byte) in xml.iter().enumerate().skip(start + 1) {
        match (quote, *byte) {
            (Some(value), byte) if byte == value => quote = None,
            (None, b'\'' | b'\"') => quote = Some(*byte),
            (None, b'>') => return Ok((start, offset + 1)),
            _ => {},
        }
    }
    Err(invalid("content-part anchor start tag is unterminated"))
}

fn raw_attribute_span(tag: &[u8], selected: &[u8]) -> Option<RawAttributeSpan> {
    let mut index = 0usize;
    while index < tag.len()
        && !tag[index].is_ascii_whitespace()
        && !matches!(tag[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < tag.len() {
        let whitespace_start = index;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= tag.len() || tag[index] == b'>' || tag[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < tag.len()
            && !tag[index].is_ascii_whitespace()
            && !matches!(tag[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &tag[name_start..index];
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if tag.get(index) != Some(&b'=') {
            return None;
        }
        index += 1;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *tag.get(index)?;
        if quote != b'\'' && quote != b'\"' {
            return None;
        }
        index += 1;
        let value_start = index;
        while index < tag.len() && tag[index] != quote {
            index += 1;
        }
        let value_end = index;
        if index >= tag.len() {
            return None;
        }
        index += 1;
        if name == selected {
            return Some(RawAttributeSpan {
                full: whitespace_start..index,
                value: value_start..value_end,
            });
        }
    }
    None
}

/// Rewrite only the relationship value in one detached anchor fragment.
pub(crate) fn rewrite_anchor_relationship_id(
    xml: &[u8],
    value: &str,
    bw_mode_key: Option<&[u8]>,
) -> Result<Vec<u8>> {
    let anchors = {
        let prefix = detached_slide_prefix(bw_mode_key)?;
        let mut wrapped = Vec::with_capacity(prefix.len() + xml.len() + 8);
        wrapped.extend_from_slice(prefix.as_bytes());
        let offset = wrapped.len();
        wrapped.extend_from_slice(xml);
        wrapped.extend_from_slice(b"</p:sld>");
        let anchors = locate_content_parts(&wrapped)?;
        if anchors.len() != 1 {
            return Err(invalid("content-part anchor has an invalid element count"));
        }
        let anchor = anchors
            .into_iter()
            .next()
            .ok_or_else(|| invalid("content-part anchor is missing"))?;
        if anchor.range.start != offset || anchor.range.end != offset + xml.len() {
            return Err(invalid("content-part anchor has an invalid source span"));
        }
        SourceAnchor {
            range: (anchor.range.start - offset)..(anchor.range.end - offset),
            relationship_id: anchor.relationship_id,
            relationship_span: (anchor.relationship_span.start - offset)
                ..(anchor.relationship_span.end - offset),
        }
    };
    let mut output = xml.to_vec();
    if anchors.relationship_span.end > output.len()
        || anchors.relationship_span.start > anchors.relationship_span.end
    {
        return Err(invalid("content-part relationship span escapes anchor XML"));
    }
    output.splice(
        anchors.relationship_span,
        escape_attribute(value).into_bytes(),
    );
    validate_anchor_xml_with_key(&output, bw_mode_key)?;
    Ok(output)
}

/// Find a safe insertion point immediately before the owning `p:spTree`
/// close tag. Content parts are valid group-shape children in this location.
pub(crate) fn shape_tree_insertion(source: &[u8]) -> Result<usize> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut tree_depth = None;
    let mut insertion = None;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let before = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                validate_attributes(&element)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("content-part XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("content-part XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree")
                    && tree_depth.replace(child_depth).is_some()
                {
                    return Err(invalid("content-part slide has multiple shape trees"));
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                validate_attributes(&element)?;
                if depth == 0 {
                    return Err(invalid("content-part slide has an empty root"));
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree") {
                    return Err(invalid("content-part slide has an empty shape tree"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("content-part slide has an unmatched end element"));
                }
                if tree_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"spTree")
                {
                    insertion = Some(before);
                    tree_depth = None;
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("content-part slide must close with p:sld"));
                    }
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "content-part slide rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(invalid("unterminated content-part slide"));
    }
    insertion.ok_or_else(|| invalid("content-part slide has no non-empty shape tree"))
}

/// Apply non-overlapping source replacements from right to left.
pub(crate) fn replace_spans(
    source: &[u8],
    mut replacements: Vec<(Range<usize>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    replacements.sort_by(|left, right| {
        right
            .0
            .start
            .cmp(&left.0.start)
            .then_with(|| right.0.end.cmp(&left.0.end))
    });
    let mut output = source.to_vec();
    let mut upper = source.len();
    for (range, value) in replacements {
        if range.start > range.end || range.end > source.len() || range.end > upper {
            return Err(invalid(
                "content-part source patch ranges overlap or escape XML",
            ));
        }
        output.splice(range.clone(), value);
        upper = range.start;
    }
    if output.len() > MAX_XML_BYTES {
        return Err(limit("content-part slide XML bytes", MAX_XML_BYTES));
    }
    Ok(output)
}

/// Return whether a raw slide still contains a relationship attribute for an
/// ID after a content-part anchor has been removed.
pub(crate) fn contains_relationship_reference(xml: &[u8], relationship_id: &str) -> bool {
    let double = format!(r#"r:id=\"{relationship_id}\""#);
    let single = format!("r:id='{relationship_id}'");
    xml.windows(double.len())
        .any(|window| window == double.as_bytes())
        || xml
            .windows(single.len())
            .any(|window| window == single.as_bytes())
}

fn relationship_value_span(
    element: &BytesStart<'_>,
    start: usize,
    _decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<Range<usize>> {
    let mut selected = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *super::super::REL || *value == *super::super::STRICT_REL)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r")
        {
            if selected.is_some() {
                return Err(invalid("content-part anchor has duplicate r:id attributes"));
            }
            selected = Some(attribute.key.as_ref().to_vec());
        }
    }
    let key = selected.ok_or_else(|| invalid("content-part anchor has no relationship ID span"))?;
    let local = attribute_span(element.as_ref(), &key)?;
    let base = start
        .checked_add(1)
        .ok_or_else(|| invalid("content-part relationship span overflow"))?;
    let begin = base
        .checked_add(local.start)
        .ok_or_else(|| invalid("content-part relationship span overflow"))?;
    let end = base
        .checked_add(local.end)
        .ok_or_else(|| invalid("content-part relationship span overflow"))?;
    Ok(begin..end)
}

struct LocalAttributeSpan {
    start: usize,
    end: usize,
}

fn attribute_span(raw: &[u8], selected: &[u8]) -> Result<LocalAttributeSpan> {
    let mut index = 0usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && !matches!(raw[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] == b'>' || raw[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[name_start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if raw.get(index) != Some(&b'=') {
            return Err(invalid("content-part anchor attribute has no value"));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| invalid("content-part anchor attribute value is missing"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("content-part anchor attribute value is not quoted"));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(invalid(
                "content-part anchor attribute value is unterminated",
            ));
        }
        if name == selected {
            return Ok(LocalAttributeSpan {
                start: value_start,
                end: index,
            });
        }
        index += 1;
    }
    Err(invalid(
        "content-part relationship attribute span is missing",
    ))
}

fn escape_attribute(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
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

fn black_white_mode(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<(Option<BlackWhiteMode>, Option<Box<[u8]>>)> {
    let mut value = None;
    let mut key = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"bwMode"
            || !matches!(
                namespace,
                ResolveResult::Bound(Namespace(value)) if *value == *P14
            )
        {
            continue;
        }
        if value.is_some() {
            return Err(invalid(
                "PresentationML contentPart has duplicate p14:bwMode attributes",
            ));
        }
        let text = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(BlackWhiteMode::parse(&text)?);
        key = Some(attribute.key.as_ref().into());
    }
    Ok((value, key))
}

fn validate_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
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

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("content-part XML offset does not fit usize"))
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
