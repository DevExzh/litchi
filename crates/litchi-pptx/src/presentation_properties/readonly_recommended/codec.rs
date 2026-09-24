//! Bounded namespace-aware codec and source splicer for `readonlyRecommended`.

use std::collections::HashSet;
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{
    Capabilities, Limits as MceLimits, NAMESPACE as MCE_NAMESPACE, OffsetLimits, active_offsets,
    process_markup_compatibility,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{XmlVersion, encoding::Decoder, reader::NsReader};

use super::{EXTENSION_URI, NAMESPACE};
use crate::presentation_properties::{P_NS, P_STRICT, Properties};
use crate::{Error, Result};

pub(crate) const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_OFFSETS: usize = 100_000;

#[derive(Clone, Debug)]
pub(crate) struct ElementSpan {
    pub(crate) range: Range<usize>,
    pub(crate) close_start: Option<usize>,
    pub(crate) qname: Vec<u8>,
    pub(crate) empty: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct Located {
    pub(crate) value: Option<bool>,
    pub(crate) root: ElementSpan,
    pub(crate) ext_list: Option<ElementSpan>,
    pub(crate) target_extension: Option<ElementSpan>,
    pub(crate) value_span: Option<Range<usize>>,
}

#[derive(Clone, Debug)]
struct Frame {
    span: ElementSpan,
    mce: bool,
    root: bool,
    ext_list: bool,
    target_extension: bool,
}

#[derive(Clone, Debug)]
struct Candidate {
    start: usize,
    ext_start: usize,
    value: Option<bool>,
    value_span: Option<Range<usize>>,
}

#[derive(Default)]
struct RawScan {
    root: Option<ElementSpan>,
    ext_lists: Vec<ElementSpan>,
    target_extensions: Vec<ElementSpan>,
    candidates: Vec<Candidate>,
    frames: Vec<Frame>,
    nodes: usize,
}

/// Parse one complete `presentationPr` part and locate the effective owner.
pub(crate) fn locate(source: &[u8]) -> Result<Located> {
    if source.len() > MAX_BYTES {
        return Err(limit(
            "presentation-properties readonlyRecommended bytes",
            MAX_BYTES,
        ));
    }
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(NAMESPACE);
    let mce_limits = MceLimits {
        max_input_bytes: MAX_BYTES,
        max_output_bytes: MAX_BYTES,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    let processed = process_markup_compatibility(source, &capabilities, &mce_limits)
        .map_err(Error::MarkupCompatibility)?
        .xml;
    let value = semantic_value(processed.as_ref())?;
    let mut raw = scan_raw(source)?;

    let mut offsets = Vec::with_capacity(raw.ext_lists.len().saturating_add(raw.candidates.len()));
    for element in &raw.ext_lists {
        offsets.push(u32::try_from(element.range.start).map_err(|_| {
            limit(
                "presentation-properties readonlyRecommended offsets",
                MAX_BYTES,
            )
        })?);
    }
    for candidate in &raw.candidates {
        offsets.push(u32::try_from(candidate.start).map_err(|_| {
            limit(
                "presentation-properties readonlyRecommended offsets",
                MAX_BYTES,
            )
        })?);
    }
    let active = if offsets.is_empty() {
        HashSet::new()
    } else {
        let limits = OffsetLimits {
            max_source_bytes: MAX_BYTES,
            max_offsets: MAX_OFFSETS,
            max_marked_bytes: MAX_BYTES,
            processing: mce_limits.clone(),
        };
        active_offsets(source, &offsets, &capabilities, &limits)
            .map_err(Error::MarkupCompatibility)?
            .into_iter()
            .collect::<HashSet<_>>()
    };

    let root = raw
        .root
        .take()
        .ok_or_else(|| invalid("presentation-properties root is missing"))?;
    let active_ext_lists: Vec<_> = raw
        .ext_lists
        .into_iter()
        .filter(|element| {
            u32::try_from(element.range.start)
                .ok()
                .is_some_and(|offset| active.contains(&offset))
        })
        .collect();
    if active_ext_lists.len() > 1 {
        return Err(invalid(
            "presentation properties contain multiple effective extLst elements",
        ));
    }
    let ext_list = active_ext_lists.into_iter().next();

    let active_candidates: Vec<_> = raw
        .candidates
        .drain(..)
        .filter(|candidate| {
            u32::try_from(candidate.start)
                .ok()
                .is_some_and(|offset| active.contains(&offset))
        })
        .collect();
    let target_extension = if value.is_some() {
        if active_candidates.len() != 1 {
            return Err(invalid(
                "effective readonlyRecommended payload has no unique source owner",
            ));
        }
        let candidate = &active_candidates[0];
        let extension = raw
            .target_extensions
            .iter()
            .find(|element| element.range.start == candidate.ext_start)
            .cloned()
            .ok_or_else(|| invalid("readonlyRecommended extension source is missing"))?;
        if candidate.value != value {
            return Err(invalid(
                "readonlyRecommended source value differs from semantic value",
            ));
        }
        Some(extension)
    } else {
        if !active_candidates.is_empty() {
            return Err(invalid(
                "inactive readonlyRecommended source was selected unexpectedly",
            ));
        }
        None
    };
    let value_span = if value.is_some() {
        active_candidates
            .first()
            .and_then(|candidate| candidate.value_span.clone())
    } else {
        None
    };

    Ok(Located {
        value,
        root,
        ext_list,
        target_extension,
        value_span,
    })
}

/// Rewrite only the recognized extension, preserving all other source bytes.
pub(crate) fn rewrite(source: &[u8], located: &Located, value: Option<bool>) -> Result<Vec<u8>> {
    if located.value == value {
        return Ok(source.to_vec());
    }
    let mut replacements = Vec::new();
    match (located.value, value) {
        (Some(_), Some(value)) => {
            let range = located
                .value_span
                .clone()
                .ok_or_else(|| invalid("readonlyRecommended value source span is missing"))?;
            replacements.push(Replacement {
                range,
                value: if value {
                    b"true".to_vec()
                } else {
                    b"false".to_vec()
                },
            });
        },
        (Some(_), None) => {
            let range = located
                .target_extension
                .as_ref()
                .ok_or_else(|| invalid("readonlyRecommended extension source span is missing"))?
                .range
                .clone();
            replacements.push(Replacement {
                range,
                value: Vec::new(),
            });
        },
        (None, Some(value)) => {
            let payload = format!(
                r#"<p1710:readonlyRecommended xmlns:p1710="{NAMESPACE}" val="{}"/>"#,
                if value { "true" } else { "false" }
            );
            let prefix = qname_prefix(
                located
                    .ext_list
                    .as_ref()
                    .map(|element| element.qname.as_slice())
                    .unwrap_or(&located.root.qname),
            );
            let ext_name = qualified(prefix, b"ext");
            let ext_list_name = located
                .ext_list
                .as_ref()
                .map(|element| element.qname.clone())
                .unwrap_or_else(|| qualified(prefix, b"extLst"));
            let extension = format!(
                r#"<{ext_name} uri="{EXTENSION_URI}">{payload}</{ext_name}>"#,
                ext_name = String::from_utf8(ext_name)
                    .map_err(|_| invalid("presentation extension QName is not UTF-8"))?,
            );
            if let Some(ext_list) = &located.ext_list {
                if ext_list.empty {
                    let opening = empty_opening(source, ext_list)?;
                    let replacement = format!(
                        "{opening}{extension}</{}>",
                        String::from_utf8(ext_list.qname.clone())
                            .map_err(|_| invalid("presentation extLst QName is not UTF-8"))?
                    );
                    replacements.push(Replacement {
                        range: ext_list.range.clone(),
                        value: replacement.into_bytes(),
                    });
                } else {
                    let at = ext_list
                        .close_start
                        .ok_or_else(|| invalid("presentation extLst closing span is missing"))?;
                    replacements.push(Replacement {
                        range: at..at,
                        value: extension.into_bytes(),
                    });
                }
            } else {
                let root = &located.root;
                let ext_list = format!(
                    "<{name}>{extension}</{name}>",
                    name = String::from_utf8(ext_list_name)
                        .map_err(|_| invalid("presentation extLst QName is not UTF-8"))?
                );
                if root.empty {
                    let opening = empty_opening(source, root)?;
                    let replacement = format!(
                        "{opening}{ext_list}</{}>",
                        String::from_utf8(root.qname.clone())
                            .map_err(|_| invalid("presentation root QName is not UTF-8"))?
                    );
                    replacements.push(Replacement {
                        range: root.range.clone(),
                        value: replacement.into_bytes(),
                    });
                } else {
                    let at = root
                        .close_start
                        .ok_or_else(|| invalid("presentation root closing span is missing"))?;
                    replacements.push(Replacement {
                        range: at..at,
                        value: ext_list.into_bytes(),
                    });
                }
            }
        },
        (None, None) => unreachable!("semantic no-op was handled above"),
    }
    let updated = apply_replacements(source, replacements)?;
    if updated.len() > MAX_BYTES {
        return Err(limit(
            "serialized presentation-properties readonlyRecommended bytes",
            MAX_BYTES,
        ));
    }
    Ok(updated)
}

fn semantic_value(xml: &[u8]) -> Result<Option<bool>> {
    let properties = Properties::parse(xml)?;
    let mut value = None;
    for extension in &properties.extensions {
        if let crate::presentation_properties::Extension::ReadonlyRecommended(candidate) = extension
        {
            if value.replace(*candidate).is_some() {
                return Err(invalid("duplicate readonlyRecommended extensions"));
            }
        }
    }
    Ok(value)
}

fn scan_raw(source: &[u8]) -> Result<RawScan> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    let mut raw = RawScan::default();
    let mut buffer = Vec::new();
    loop {
        let before = position(&reader, origin)?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let after = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                raw.nodes = raw.nodes.saturating_add(1);
                if raw.nodes > MAX_NODES {
                    return Err(limit(
                        "presentation-properties readonlyRecommended XML nodes",
                        MAX_NODES,
                    ));
                }
                let (frame, candidate) = start_frame(
                    &raw.frames,
                    &namespace,
                    &element,
                    reader.decoder(),
                    before,
                    after,
                )?;
                if let Some(candidate) = candidate {
                    raw.candidates.push(candidate);
                }
                if frame.root {
                    if raw.root.is_some() {
                        return Err(invalid("presentation-properties XML has multiple roots"));
                    }
                    raw.root = Some(frame.span.clone());
                }
                raw.frames.push(frame);
            },
            Event::Empty(element) => {
                raw.nodes = raw.nodes.saturating_add(1);
                if raw.nodes > MAX_NODES {
                    return Err(limit(
                        "presentation-properties readonlyRecommended XML nodes",
                        MAX_NODES,
                    ));
                }
                let (mut frame, candidate) = start_frame(
                    &raw.frames,
                    &namespace,
                    &element,
                    reader.decoder(),
                    before,
                    after,
                )?;
                if let Some(candidate) = candidate {
                    raw.candidates.push(candidate);
                }
                frame.span.empty = true;
                finish_frame(&mut raw, frame.clone())?;
            },
            Event::End(element) => {
                let mut frame = raw
                    .frames
                    .pop()
                    .ok_or_else(|| invalid("presentation-properties XML has an unmatched end"))?;
                if frame.span.qname.as_slice() != element.name().as_ref() {
                    return Err(invalid(
                        "presentation-properties XML start/end names differ",
                    ));
                }
                frame.span.close_start = Some(before);
                frame.span.range.end = after;
                finish_frame(&mut raw, frame)?;
            },
            Event::Decl(_) | Event::Comment(_) | Event::Text(_) | Event::CData(_) => {},
            Event::GeneralRef(_) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "presentation-properties XML cannot contain DTDs or processing instructions",
                ));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !raw.frames.is_empty() {
        return Err(invalid("unterminated presentation-properties XML"));
    }
    Ok(raw)
}

fn start_frame(
    stack: &[Frame],
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    before: usize,
    after: usize,
) -> Result<(Frame, Option<Candidate>)> {
    if stack.len() >= MAX_DEPTH {
        return Err(limit(
            "presentation-properties readonlyRecommended XML depth",
            MAX_DEPTH,
        ));
    }
    let (namespace, mce) = match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            (value.to_vec(), *value == MCE_NAMESPACE.as_bytes())
        },
        ResolveResult::Unknown(_) | ResolveResult::Unbound => (Vec::new(), false),
    };
    let local = element.local_name().as_ref().to_vec();
    let qname = element.name().as_ref().to_vec();
    let root = stack.is_empty()
        && namespace_matches(
            &namespace,
            element.name().local_name().as_ref(),
            b"presentationPr",
        );
    let non_mce_parent = stack.iter().rev().find(|frame| !frame.mce);
    let ext_list = !mce
        && namespace_matches(&namespace, &local, b"extLst")
        && non_mce_parent.is_some_and(|frame| frame.root);
    let target_extension = !mce
        && namespace_matches(&namespace, &local, b"ext")
        && non_mce_parent.is_some_and(|frame| frame.ext_list)
        && extension_uri(element, decoder)? == EXTENSION_URI;
    let payload = !mce
        && namespace == NAMESPACE.as_bytes()
        && local.as_slice() == b"readonlyRecommended"
        && non_mce_parent.is_some_and(|frame| frame.target_extension);
    let candidate = if payload {
        let (value, value_span) = parse_value_attr(element, decoder, before)?;
        Some(Candidate {
            start: before,
            ext_start: non_mce_parent
                .ok_or_else(|| invalid("readonlyRecommended extension parent is missing"))?
                .span
                .range
                .start,
            value,
            value_span,
        })
    } else {
        None
    };
    Ok((
        Frame {
            span: ElementSpan {
                range: before..after,
                close_start: None,
                qname,
                empty: false,
            },
            mce,
            root,
            ext_list,
            target_extension,
        },
        candidate,
    ))
}

fn finish_frame(raw: &mut RawScan, frame: Frame) -> Result<()> {
    if frame.root {
        match raw.root.as_ref() {
            Some(root) if root.range.start == frame.span.range.start => {
                raw.root = Some(frame.span.clone());
            },
            Some(_) => return Err(invalid("presentation-properties XML has multiple roots")),
            None => raw.root = Some(frame.span.clone()),
        }
    }
    if frame.ext_list {
        raw.ext_lists.push(frame.span.clone());
    }
    if frame.target_extension {
        raw.target_extensions.push(frame.span.clone());
    }
    Ok(())
}

fn parse_value_attr(
    element: &BytesStart<'_>,
    decoder: Decoder,
    event_start: usize,
) -> Result<(Option<bool>, Option<Range<usize>>)> {
    let raw = element.as_ref();
    let mut value = None;
    let mut span = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() != b"val" {
            continue;
        }
        if value.is_some() {
            return Err(invalid("readonlyRecommended has duplicate val attributes"));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        value = Some(parse_bool(&decoded)?);
        span = Some(
            attribute_value_span(raw, b"val")?
                .ok_or_else(|| invalid("readonlyRecommended val source span is missing"))?,
        );
    }
    let span = span.map(|range| (event_start + 1 + range.start)..(event_start + 1 + range.end));
    Ok((value, span))
}

fn namespace_matches(namespace: &[u8], local: &[u8], expected: &[u8]) -> bool {
    (namespace == P_NS.as_bytes() || namespace == P_STRICT.as_bytes()) && local == expected
}

fn extension_uri(element: &BytesStart<'_>, decoder: Decoder) -> Result<String> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"uri" {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| value.into_owned())
                .map_err(xml_error);
        }
    }
    Ok(String::new())
}

fn attribute_value_span(raw: &[u8], key: &[u8]) -> Result<Option<Range<usize>>> {
    let mut index = 0usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && raw[index] != b'>'
        && raw[index] != b'/'
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
        let start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] != b'=' {
            return Err(invalid("presentation-properties attribute has no value"));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| invalid("presentation-properties attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid(
                "presentation-properties attribute value is not quoted",
            ));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(invalid(
                "presentation-properties attribute value is unterminated",
            ));
        }
        if name == key {
            return Ok(Some(value_start..index));
        }
        index += 1;
    }
    Ok(None)
}

fn empty_opening(source: &[u8], span: &ElementSpan) -> Result<String> {
    let mut end = span.range.end;
    while end > span.range.start && source[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    if end < 2 || source[end - 1] != b'>' || source[end - 2] != b'/' {
        return Err(invalid(
            "self-closing presentation-properties tag is malformed",
        ));
    }
    String::from_utf8(source[span.range.start..end - 2].to_vec())
        .map(|mut opening| {
            opening.push('>');
            opening
        })
        .map_err(|_| invalid("presentation-properties source QName is not UTF-8"))
}

fn qname_prefix(qname: &[u8]) -> &[u8] {
    qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &qname[..index])
}

fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    if prefix.is_empty() {
        local.to_vec()
    } else {
        let mut value = Vec::with_capacity(prefix.len() + 1 + local.len());
        value.extend_from_slice(prefix);
        value.push(b':');
        value.extend_from_slice(local);
        value
    }
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

fn apply_replacements(source: &[u8], mut replacements: Vec<Replacement>) -> Result<Vec<u8>> {
    replacements.sort_by_key(|replacement| std::cmp::Reverse(replacement.range.start));
    let mut output = source.to_vec();
    let mut upper = source.len();
    for replacement in replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.end > upper
        {
            return Err(invalid(
                "presentation-properties patch ranges overlap or escape source",
            ));
        }
        output.splice(replacement.range.clone(), replacement.value);
        upper = replacement.range.start;
    }
    Ok(output)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("readonlyRecommended/@val is not an xsd:boolean")),
    }
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("presentation-properties XML offset does not fit usize"))
}

fn limit(resource: &'static str, maximum: usize) -> Error {
    Error::Limit {
        resource,
        limit: maximum,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
