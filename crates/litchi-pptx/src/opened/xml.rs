//! Namespace-aware, range-preserving XML edits for the transaction root.

use std::ops::Range;

use litchi_ooxml_common::xml::{DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Slide, invalid};
use crate::{Error, Result};

const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

#[derive(Debug)]
struct SlideIdElement {
    id: u32,
    relationship_id: String,
    span: Range<usize>,
}

#[derive(Debug)]
struct OpenElement {
    id: u32,
    relationship_id: String,
    start: usize,
    depth: usize,
}

#[derive(Debug)]
struct TextElement {
    span: Range<usize>,
    empty_name: Option<Vec<u8>>,
}

pub(crate) fn reorder_slides(xml: &[u8], current: &[Slide], ordered: &[u32]) -> Result<Vec<u8>> {
    if ordered.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide order is not a complete permutation",
        ));
    }
    let elements = slide_id_elements(xml)?;
    if elements.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    for (element, slide) in elements.iter().zip(current) {
        if element.id != slide.id || element.relationship_id != slide.relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
    }
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide permutation",
            source,
        })?;
    let mut seen = std::collections::HashSet::new();
    for id in ordered {
        if !seen.insert(*id) {
            return Err(invalid(
                "opened-presentation slide order repeats an identity",
            ));
        }
        let index = elements
            .iter()
            .position(|element| element.id == *id)
            .ok_or_else(|| invalid("opened-presentation slide order references an unknown ID"))?;
        selected.push(index);
    }
    if selected
        .iter()
        .enumerate()
        .all(|(left, right)| left == *right)
    {
        return Ok(xml.to_vec());
    }
    let first = elements
        .first()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let last = elements
        .last()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation reordered XML",
            source,
        })?;
    output.extend_from_slice(&xml[..first.span.start]);
    for (position, source_index) in selected.into_iter().enumerate() {
        let source = &elements[source_index];
        output.extend_from_slice(&xml[source.span.clone()]);
        if let Some(next) = elements.get(position + 1) {
            output.extend_from_slice(&xml[elements[position].span.end..next.span.start]);
        }
    }
    output.extend_from_slice(&xml[last.span.end..]);
    Ok(output)
}

pub(crate) fn rewrite_shape_text(
    xml: &[u8],
    shape: Range<usize>,
    text: &str,
    max_text_bytes: usize,
) -> Result<Vec<u8>> {
    if text.len() > max_text_bytes {
        return Err(Error::Limit {
            resource: "opened-presentation shape text bytes",
            limit: max_text_bytes,
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid(
            "opened-presentation shape text contains an invalid XML character",
        ));
    }
    if shape.start >= shape.end || shape.end > xml.len() {
        return Err(invalid("opened-presentation shape range is invalid"));
    }
    let spans = drawing_text_elements(xml, &shape)?;
    if spans.is_empty() {
        return Err(invalid(
            "opened-presentation selected shape has no DrawingML text run",
        ));
    }
    let escaped = quick_xml::escape::escape(text);
    let mut output = Vec::new();
    output
        .try_reserve(xml.len().saturating_add(escaped.len()))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text XML",
            source,
        })?;
    let mut cursor = 0usize;
    for (index, span) in spans.iter().enumerate() {
        output.extend_from_slice(&xml[cursor..span.span.start]);
        if index == 0 {
            if let Some(name) = &span.empty_name {
                let raw = &xml[span.span.clone()];
                let slash = raw
                    .iter()
                    .rposition(|byte| *byte == b'/')
                    .ok_or_else(|| invalid("opened-presentation empty text tag is malformed"))?;
                let mut open_end = slash;
                while open_end > 0 && raw[open_end - 1].is_ascii_whitespace() {
                    open_end -= 1;
                }
                output.extend_from_slice(&raw[..open_end]);
                output.push(b'>');
                output.extend_from_slice(escaped.as_bytes());
                output.extend_from_slice(b"</");
                output.extend_from_slice(name);
                output.push(b'>');
            } else {
                output.extend_from_slice(escaped.as_bytes());
            }
        } else if span.empty_name.is_some() {
            output.extend_from_slice(&xml[span.span.clone()]);
        }
        cursor = span.span.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

fn slide_id_elements(xml: &[u8]) -> Result<Vec<SlideIdElement>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut list_depth = None;
    let mut lists = 0usize;
    let mut open = None;
    let mut elements = Vec::new();
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml_namespace = is_presentation_namespace(&namespace);
        let event = event.into_owned();
        drop(namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "opened-presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                if pml_namespace && element.local_name().as_ref() == b"sldIdLst" {
                    lists = lists.saturating_add(1);
                    if lists != 1 {
                        return Err(invalid("opened-presentation has multiple slide-ID lists"));
                    }
                    list_depth = Some(depth + 1);
                } else if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                    if open.is_some() {
                        return Err(invalid("opened-presentation slide IDs overlap"));
                    }
                    let (id, relationship_id) = parse_slide_id(&element, &reader)?;
                    open = Some(OpenElement {
                        id,
                        relationship_id,
                        start,
                        depth: depth + 1,
                    });
                }
                depth += 1;
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                    let (id, relationship_id) = parse_slide_id(&element, &reader)?;
                    elements.push(SlideIdElement {
                        id,
                        relationship_id,
                        span: start..end,
                    });
                }
            },
            Event::End(element) => {
                if let Some(active) = &open
                    && active.depth == depth
                    && pml_namespace
                    && element.local_name().as_ref() == b"sldId"
                {
                    let active = open
                        .take()
                        .ok_or_else(|| invalid("opened-presentation slide ID disappeared"))?;
                    elements.push(SlideIdElement {
                        id: active.id,
                        relationship_id: active.relationship_id,
                        span: active.start..end,
                    });
                }
                if list_depth == Some(depth)
                    && pml_namespace
                    && element.local_name().as_ref() == b"sldIdLst"
                {
                    list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(value) if list_depth == Some(depth) => {
                if !value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("opened-presentation slide-ID list contains text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if list_depth == Some(depth) => {
                return Err(invalid(
                    "opened-presentation slide-ID list contains unsupported content",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || open.is_some() || list_depth.is_some() {
        return Err(invalid("opened-presentation XML is unterminated"));
    }
    if lists != 1 {
        return Err(invalid("opened-presentation has no slide-ID list"));
    }
    Ok(elements)
}

fn parse_slide_id(element: &BytesStart<'_>, reader: &NsReader<&[u8]>) -> Result<(u32, String)> {
    let value =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", reader.decoder())?
            .ok_or_else(|| invalid("opened-presentation slide ID lacks id"))?;
    let id = value
        .parse::<u32>()
        .map_err(|_err| invalid("opened-presentation slide ID is invalid"))?;
    let relationship_id = crate::parts::relationship_attribute(element, reader)?
        .ok_or_else(|| invalid("opened-presentation slide ID lacks r:id"))?;
    Ok((id, relationship_id))
}

fn drawing_text_elements(xml: &[u8], owner: &Range<usize>) -> Result<Vec<TextElement>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut active: Option<(usize, usize)> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let drawing_namespace = is_drawing(&namespace);
        let event = event.into_owned();
        drop(namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "opened-presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                depth += 1;
                if owner.contains(&start)
                    && drawing_namespace
                    && element.local_name().as_ref() == b"t"
                {
                    if active.replace((end, depth)).is_some() {
                        return Err(invalid(
                            "opened-presentation DrawingML text elements overlap",
                        ));
                    }
                } else if active.is_some() {
                    return Err(invalid(
                        "opened-presentation DrawingML text contains child markup",
                    ));
                }
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if owner.contains(&start)
                    && drawing_namespace
                    && element.local_name().as_ref() == b"t"
                {
                    spans.push(TextElement {
                        span: start..end,
                        empty_name: Some(element.name().as_ref().to_vec()),
                    });
                } else if active.is_some() {
                    return Err(invalid(
                        "opened-presentation DrawingML text contains child markup",
                    ));
                }
            },
            Event::End(element) => {
                if let Some((content_start, active_depth)) = active
                    && active_depth == depth
                    && drawing_namespace
                    && element.local_name().as_ref() == b"t"
                {
                    spans.push(TextElement {
                        span: content_start..start,
                        empty_name: None,
                    });
                    active = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Comment(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains a comment",
                ));
            },
            Event::Decl(_) | Event::PI(_) | Event::DocType(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains forbidden markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || active.is_some() {
        return Err(invalid("opened-presentation slide XML is unterminated"));
    }
    Ok(spans)
}

fn is_drawing(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

fn is_presentation_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
    )
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("opened-presentation XML position exceeds usize"))
}

fn bump(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("opened-presentation XML node count overflow"))?;
    if *nodes > MAX_XML_NODES {
        return Err(Error::Limit {
            resource: "opened-presentation XML nodes",
            limit: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x1_0000..=0x10_FFFF)
}
