#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the fixture helpers stay beside the differential tests they serve"
)]

//! Differential qualification for the private DOCX body-scan fusion.
//!
//! The range oracle below is a frozen copy of the pre-fusion namespace walk.
//! It intentionally does not call `crate::namespace::scan_word_element_ranges`:
//! a refactor that extracts a shared scanner state must still be checked against
//! the old event contract.  The candidate adapter is the only production symbol
//! that this file expects the parent module to wire.

use crate::alt::{Chunk, Rel, active};
use crate::error::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const MAX_CHUNKS: usize = 4096;
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 256;
const MAX_SCAN_DEPTH: usize = 128;
const MAX_SCAN_NODES: usize = 1_000_000;
const MAX_VISIBILITY_OFFSETS: usize = 1_000_000;

type BlockRange = (usize, u32, u32);

#[derive(Debug, PartialEq, Eq)]
struct ScanOutput {
    chunks: BTreeMap<u32, Chunk>,
    ranges: Vec<BlockRange>,
}

/// Adapter for the production helper added by the implementation lane.
///
/// The parent module wires this test file from `writer/doc/package.rs`.  The
/// implementation should expose the collected result as
/// `scan_alt_and_block_ranges(&[u8]) -> Result<(BTreeMap<u32, Chunk>,
/// Vec<(usize, u32, u32)>)>`.  Keeping this one-line adapter local means the
/// differential oracle never has to know how the writer stores its private
/// result while the expected output remains explicit here.
fn candidate_scan(xml: &[u8]) -> Result<ScanOutput> {
    let (chunks, ranges) = super::scan_alt_and_block_ranges(xml)?;
    Ok(ScanOutput { chunks, ranges })
}

fn frozen_scan(xml: &[u8]) -> Result<ScanOutput> {
    let chunks = frozen_alt_scan(xml)?;
    let ranges = frozen_active_block_ranges(xml)?;
    Ok(ScanOutput { chunks, ranges })
}

fn assert_parity(name: &str, xml: &[u8]) {
    let expected = frozen_scan(xml);
    let actual = candidate_scan(xml);
    match (expected, actual) {
        (Ok(expected), Ok(actual)) => assert_eq!(actual, expected, "{name}"),
        (Err(expected), Err(actual)) => {
            assert_eq!(
                error_fingerprint(&actual),
                error_fingerprint(&expected),
                "{name}"
            );
        },
        (expected, actual) => panic!(
            "{name}: result kind changed\nexpected={:?}\nactual={:?}",
            result_fingerprint(&expected),
            result_fingerprint(&actual)
        ),
    }
}

fn error_fingerprint(error: &Error) -> (String, String) {
    (error.to_string(), format!("{error:?}"))
}

fn result_fingerprint(result: &Result<ScanOutput>) -> String {
    match result {
        Ok(value) => format!("ok chunks={:?} ranges={:?}", value.chunks, value.ranges),
        Err(error) => format!("err={:?}", error_fingerprint(error)),
    }
}

fn span_after(xml: &[u8], literal: &str, offset: usize) -> (usize, u32) {
    let needle = literal.as_bytes();
    let start = xml
        .get(offset..)
        .unwrap_or_default()
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|relative| relative + offset)
        .unwrap_or_else(|| panic!("fixture does not contain {literal:?}"));
    (
        start,
        u32::try_from(needle.len()).expect("fixture span length fits u32"),
    )
}

fn expected_ranges(xml: &[u8], literals: &[(usize, &str)]) -> Vec<BlockRange> {
    let mut offset = 0;
    literals
        .iter()
        .map(|(target, literal)| {
            let (start, length) = span_after(xml, literal, offset);
            offset = start.saturating_add(usize::try_from(length).expect("fixture length fits"));
            (
                *target,
                u32::try_from(start).expect("fixture offset fits u32"),
                length,
            )
        })
        .collect()
}

fn expected_chunk(
    xml: &[u8],
    literal: &str,
    relationship: &str,
    match_source: Option<bool>,
) -> (u32, Chunk) {
    let start = xml
        .windows(literal.len())
        .position(|window| window == literal.as_bytes())
        .unwrap_or_else(|| panic!("fixture does not contain {literal:?}"));
    (
        u32::try_from(start).expect("fixture offset fits u32"),
        Chunk::new(
            Rel::new(relationship).expect("fixture relationship is valid"),
            match_source,
        ),
    )
}

fn expected_chunks(xml: &[u8], values: &[(&str, &str, Option<bool>)]) -> BTreeMap<u32, Chunk> {
    values
        .iter()
        .map(|(literal, relationship, match_source)| {
            expected_chunk(xml, literal, relationship, *match_source)
        })
        .collect()
}

fn assert_expected(
    name: &str,
    xml: &[u8],
    chunks: &[(&str, &str, Option<bool>)],
    ranges: &[(usize, &str)],
) {
    assert_parity(name, xml);
    let expected = frozen_scan(xml).expect("fixture is expected to succeed");
    assert_eq!(
        expected.chunks,
        expected_chunks(xml, chunks),
        "{name}: chunks"
    );
    assert_eq!(
        expected.ranges,
        expected_ranges(xml, ranges),
        "{name}: ranges"
    );
    let actual = candidate_scan(xml).expect("candidate is expected to succeed");
    assert_eq!(actual.chunks, expected.chunks, "{name}: candidate chunks");
    assert_eq!(actual.ranges, expected.ranges, "{name}: candidate ranges");
}

fn transitional(body: &str) -> String {
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}"><w:body>{body}</w:body></w:document>"#
    )
}

fn strict(body: &str) -> String {
    format!(
        r#"<s:document xmlns:s="{STRICT_WORD}" xmlns:r="{STRICT_RELATIONSHIPS}"><s:body>{body}</s:body></s:document>"#
    )
}

fn mce(body: &str) -> String {
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><w:body>{body}</w:body></w:document>"#
    )
}

fn strict_mce(body: &str) -> String {
    format!(
        r#"<s:document xmlns:s="{STRICT_WORD}" xmlns:r="{STRICT_RELATIONSHIPS}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><s:body>{body}</s:body></s:document>"#
    )
}

fn deep_transitional(wrapper_count: usize, tail: &str) -> String {
    let mut body = String::new();
    for _ in 0..wrapper_count {
        body.push_str("<u:wrap>");
    }
    body.push_str(tail);
    for _ in 0..wrapper_count {
        body.push_str("</u:wrap>");
    }
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}" xmlns:u="urn:wrapper"><w:body>{body}</w:body></w:document>"#
    )
}

fn unbound_fragment(body: &str) -> String {
    format!("<document><body>{body}</body></document>")
}

fn unknown_prefix_fragment(body: &str) -> String {
    format!("<q:document><q:body>{body}</q:body></q:document>")
}

fn malformed_tail() -> String {
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}"><w:body><w:p></w:body></w:document>"#
    )
}

fn alt_depth_probe(depth: usize) -> Vec<u8> {
    let mut xml = String::with_capacity(depth.saturating_mul(7));
    for _ in 0..depth {
        xml.push_str("<x:x>");
    }
    for _ in 0..depth {
        xml.push_str("</x:x>");
    }
    xml.into_bytes()
}

struct FrozenPendingChunk {
    root_depth: usize,
    start: u32,
    relationship: Rel,
    match_source: Option<bool>,
    saw_properties: bool,
    properties_depth: Option<usize>,
    opaque_depth: Option<usize>,
}

fn frozen_alt_scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    frozen_alt_validate_xml(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut pending: Option<FrozenPendingChunk> = None;
    let mut chunks = BTreeMap::new();

    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .and_then(|offset| u32::try_from(offset).ok())
            .ok_or_else(|| Error::Invalid("altChunk XML offset does not fit u32".into()))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                let event_depth = frozen_alt_next_depth(depth)?;
                if pending.is_none()
                    && frozen_alt_is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    pending = Some(FrozenPendingChunk {
                        root_depth: event_depth,
                        start: event_start,
                        relationship: frozen_alt_relationship(&element, decoder, &resolver)?,
                        match_source: None,
                        saw_properties: false,
                        properties_depth: None,
                        opaque_depth: None,
                    });
                } else if let Some(chunk) = pending.as_mut() {
                    frozen_alt_parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        false,
                    )?;
                }
                depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = frozen_alt_next_depth(depth)?;
                if pending.is_none()
                    && frozen_alt_is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    let chunk =
                        Chunk::new(frozen_alt_relationship(&element, decoder, &resolver)?, None);
                    frozen_alt_insert_chunk(&mut chunks, event_start, chunk)?;
                } else if let Some(chunk) = pending.as_mut() {
                    frozen_alt_parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        true,
                    )?;
                }
            },
            Event::End(_) => {
                if let Some(chunk) = pending.as_mut()
                    && chunk.opaque_depth == Some(depth)
                {
                    chunk.opaque_depth = None;
                }
                if let Some(chunk) = pending.as_mut()
                    && chunk.properties_depth == Some(depth)
                {
                    chunk.properties_depth = None;
                }
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.root_depth == depth)
                {
                    let chunk = pending
                        .take()
                        .ok_or_else(|| Error::Invalid("missing pending altChunk".into()))?;
                    frozen_alt_insert_chunk(
                        &mut chunks,
                        chunk.start,
                        Chunk::new(chunk.relationship, chunk.match_source),
                    )?;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("unexpected altChunk XML end element".into()))?;
            },
            Event::Text(text)
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none())
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::Invalid("altChunk contains unexpected text".into()));
            },
            Event::CData(_) | Event::GeneralRef(_)
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none()) =>
            {
                return Err(Error::Invalid(
                    "altChunk contains unexpected character data".into(),
                ));
            },
            Event::Eof => {
                if pending.is_some() {
                    return Err(Error::Invalid("unterminated altChunk".into()));
                }
                break;
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    let offsets = chunks.keys().copied().collect::<Vec<_>>();
    let active = active(xml, &offsets)?;
    let mut selected = active.into_iter();
    let mut next = selected.next();
    chunks.retain(|offset, _| {
        if next == Some(*offset) {
            next = selected.next();
            true
        } else {
            false
        }
    });
    Ok(chunks)
}

fn frozen_alt_validate_xml(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(frozen_alt_invalid(format!(
            "alternative-format scan input exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

fn frozen_alt_parse_child(
    chunk: &mut FrozenPendingChunk,
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    empty: bool,
) -> Result<()> {
    if chunk.opaque_depth.is_some() {
        return Ok(());
    }
    let is_word = frozen_alt_is_word_namespace(namespace);
    if !is_word {
        if !empty {
            chunk.opaque_depth = Some(depth);
        }
        return Ok(());
    }
    let properties_depth = chunk
        .root_depth
        .checked_add(1)
        .ok_or_else(|| frozen_alt_invalid("altChunk XML nesting is too deep"))?;
    let value_depth = properties_depth
        .checked_add(1)
        .ok_or_else(|| frozen_alt_invalid("altChunk XML nesting is too deep"))?;
    if depth == properties_depth
        && element.local_name().as_ref() == b"altChunkPr"
        && !chunk.saw_properties
    {
        chunk.saw_properties = true;
        if !empty {
            chunk.properties_depth = Some(depth);
        }
        return Ok(());
    }
    if depth == value_depth
        && chunk.properties_depth == Some(properties_depth)
        && element.local_name().as_ref() == b"matchSrc"
        && chunk.match_source.is_none()
    {
        chunk.match_source = Some(frozen_alt_parse_on_off(
            element,
            decoder,
            resolver,
            frozen_alt_is_transitional_word_namespace(namespace),
        )?);
        return Ok(());
    }
    Err(Error::Invalid("altChunk has invalid child content".into()))
}

fn frozen_alt_relationship(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Rel> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let valid_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if uri == RELATIONSHIPS.as_bytes() || uri == STRICT_RELATIONSHIPS.as_bytes()
        );
        if !valid_namespace {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid(
                "altChunk has duplicate relationship IDs".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    let value = value
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("altChunk lacks a relationship ID".into()))?;
    Rel::new(value)
}

fn frozen_alt_parse_on_off(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    allow_legacy_values: bool,
) -> Result<bool> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !frozen_alt_is_word_namespace(&namespace) && !matches!(namespace, ResolveResult::Unbound)
        {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid("matchSrc has duplicate values".into()));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    match value.as_deref() {
        None | Some("true" | "1") => Ok(true),
        Some("false" | "0") => Ok(false),
        Some("on") if allow_legacy_values => Ok(true),
        Some("off") if allow_legacy_values => Ok(false),
        Some(value) => Err(Error::Invalid(format!("invalid matchSrc value '{value}'"))),
    }
}

fn frozen_alt_insert_chunk(
    chunks: &mut BTreeMap<u32, Chunk>,
    start: u32,
    chunk: Chunk,
) -> Result<()> {
    if chunks.len() >= MAX_CHUNKS {
        return Err(frozen_alt_invalid(
            "alternative-format anchor limit exceeded",
        ));
    }
    if chunks.insert(start, chunk).is_some() {
        return Err(Error::Invalid("duplicate altChunk XML position".into()));
    }
    Ok(())
}

fn frozen_alt_is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri))
            if *uri == WORD.as_bytes() || *uri == STRICT_WORD.as_bytes()
    )
}

fn frozen_alt_is_transitional_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri)) if *uri == WORD.as_bytes()
    )
}

fn frozen_alt_next_depth(depth: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| frozen_alt_invalid("alternative-format XML nesting overflowed"))?;
    if next > MAX_XML_DEPTH {
        return Err(frozen_alt_invalid(format!(
            "alternative-format XML exceeds {MAX_XML_DEPTH} nesting levels"
        )));
    }
    Ok(next)
}

fn frozen_alt_invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn frozen_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == WORD.as_bytes() || *value == STRICT_WORD.as_bytes()
    )
}

fn frozen_fragment_word_namespace(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if frozen_word_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

fn frozen_scan_word_element_ranges(
    xml_bytes: &[u8],
    targets: &[&[u8]],
    mut emit: impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start(usize),
        NestedStart,
        Empty(usize),
        End,
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    let origin = ReaderOrigin::of(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut capture: Option<(usize, usize, usize)> = None;
    let mut nodes = 0usize;
    let mut total_depth = 0usize;

    loop {
        let event_start = origin.offset(reader.buffer_position()).ok_or_else(|| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_SCAN_NODES} elements"
                    )));
                }
            }
            // Total nesting is tracked separately from capture depth so
            // deeply nested non-target content is rejected before
            // quick-xml's own namespace resolver overflows (u16).
            if matches!(event, Event::Start(_)) {
                total_depth = total_depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if total_depth > MAX_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            }
            if matches!(event, Event::End(_)) {
                total_depth = total_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            }

            if fragment_prefix.is_none()
                && let Event::Start(element) = &event
                && !matches!(namespace, ResolveResult::Bound(_))
            {
                fragment_prefix = Some(
                    element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec()),
                );
            }

            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element)
                    if frozen_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                        .map_or(ScanEvent::Other, ScanEvent::Start)
                },
                Event::Empty(element)
                    if capture.is_none()
                        && frozen_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                        .map_or(ScanEvent::Other, ScanEvent::Empty)
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                Event::Start(_)
                | Event::End(_)
                | Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => ScanEvent::Other,
            }
        };
        let event_end = origin.offset(reader.buffer_position()).ok_or_else(|| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;

        match event {
            ScanEvent::Start(target) => capture = Some((target, event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word element nesting is too deep".to_string())
                })?;
                if *depth > MAX_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word element nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            },
            ScanEvent::Empty(target) => {
                frozen_emit_word_element_range(target, event_start, event_end, &mut emit)?;
            },
            ScanEvent::End => {
                let Some((_, _, depth)) = capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _)) = capture.take() else {
                        return Err(Error::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    frozen_emit_word_element_range(target, start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated Word element".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    Ok(())
}

fn frozen_emit_word_element_range(
    target: usize,
    start: usize,
    end: usize,
    emit: &mut impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidFormat("invalid Word element byte range".to_string()))?;
    let start = u32::try_from(start).map_err(|_source_error| {
        Error::InvalidFormat("Word element offset exceeds u32".to_string())
    })?;
    let length = u32::try_from(length).map_err(|_source_error| {
        Error::InvalidFormat("Word element length exceeds u32".to_string())
    })?;
    emit(target, start, length)
}

const MAX_DOCUMENT_SEMANTIC_VALUES: usize = 1_000_000;

fn frozen_reserve_document_value<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    if values.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn frozen_active_block_ranges(xml: &[u8]) -> Result<Vec<BlockRange>> {
    let mut ranges = Vec::new();
    frozen_scan_word_element_ranges(
        xml,
        &[b"p".as_slice(), b"tbl".as_slice(), b"altChunk".as_slice()],
        |target, start, length| {
            frozen_reserve_document_value(&mut ranges, "active document block ranges")?;
            ranges.push((target, start, length));
            Ok(())
        },
    )?;
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(ranges.len())
        .map_err(|source| Error::Allocation {
            resource: "active document block offsets",
            source,
        })?;
    for &(_, start, _) in &ranges {
        starts.push(start);
    }
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)
}

#[test]
fn ordinary_aliases_preserve_chunks_ranges_and_marker_like_text() {
    let xml = transitional(
        r#"<!--before--><w:p w:rsidR="A"><w:r><w:t>&lt;w:p&gt; marker-like text</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><x:foreign xmlns:x="urn:foreign"><x:p/><x:altChunk r:id="foreign"/></x:foreign><w:altChunk r:id="first"><w:altChunkPr><w:matchSrc w:val="0"/></w:altChunkPr></w:altChunk><w:sectPr/>"#,
    );
    assert_expected(
        "ordinary aliases, foreign markup, and marker-like text",
        xml.as_bytes(),
        &[(
            r#"<w:altChunk r:id="first"><w:altChunkPr><w:matchSrc w:val="0"/></w:altChunkPr></w:altChunk>"#,
            "first",
            Some(false),
        )],
        &[
            (
                0,
                r#"<w:p w:rsidR="A"><w:r><w:t>&lt;w:p&gt; marker-like text</w:t></w:r></w:p>"#,
            ),
            (
                1,
                r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
            ),
            (
                2,
                r#"<w:altChunk r:id="first"><w:altChunkPr><w:matchSrc w:val="0"/></w:altChunkPr></w:altChunk>"#,
            ),
        ],
    );
}

#[test]
fn strict_and_transitional_relationship_dialects_keep_typed_metadata() {
    let transitional_xml = transitional(
        r#"<w:altChunk r:id="legacy"><w:altChunkPr><w:matchSrc w:val="on"/></w:altChunkPr></w:altChunk>"#,
    );
    assert_expected(
        "transitional matchSrc legacy spelling",
        transitional_xml.as_bytes(),
        &[(
            r#"<w:altChunk r:id="legacy"><w:altChunkPr><w:matchSrc w:val="on"/></w:altChunkPr></w:altChunk>"#,
            "legacy",
            Some(true),
        )],
        &[(
            2,
            r#"<w:altChunk r:id="legacy"><w:altChunkPr><w:matchSrc w:val="on"/></w:altChunkPr></w:altChunk>"#,
        )],
    );

    let strict_xml = strict(
        r#"<s:altChunk r:id="strict"><s:altChunkPr><s:matchSrc s:val="false"/></s:altChunkPr></s:altChunk>"#,
    );
    assert_expected(
        "strict matchSrc spelling",
        strict_xml.as_bytes(),
        &[(
            r#"<s:altChunk r:id="strict"><s:altChunkPr><s:matchSrc s:val="false"/></s:altChunkPr></s:altChunk>"#,
            "strict",
            Some(false),
        )],
        &[(
            2,
            r#"<s:altChunk r:id="strict"><s:altChunkPr><s:matchSrc s:val="false"/></s:altChunkPr></s:altChunk>"#,
        )],
    );
}

#[test]
fn nested_anchor_metadata_is_retained_while_outer_ranges_suppress_nested_targets() {
    let xml = transitional(
        r#"<w:p><w:altChunk r:id="inside-paragraph"/></w:p><w:tbl><w:tr><w:tc><w:p/><w:altChunk r:id="inside-table"/></w:tc></w:tr></w:tbl><w:altChunk r:id="top-level"/>"#,
    );
    assert_expected(
        "nested paragraph and table anchors",
        xml.as_bytes(),
        &[
            (
                r#"<w:altChunk r:id="inside-paragraph"/>"#,
                "inside-paragraph",
                None,
            ),
            (r#"<w:altChunk r:id="inside-table"/>"#, "inside-table", None),
            (r#"<w:altChunk r:id="top-level"/>"#, "top-level", None),
        ],
        &[
            (0, r#"<w:p><w:altChunk r:id="inside-paragraph"/></w:p>"#),
            (
                1,
                r#"<w:tbl><w:tr><w:tc><w:p/><w:altChunk r:id="inside-table"/></w:tc></w:tr></w:tbl>"#,
            ),
            (2, r#"<w:altChunk r:id="top-level"/>"#),
        ],
    );
}

#[test]
fn mce_choices_keep_two_independent_visible_views() {
    let xml = mce(
        r#"<mc:AlternateContent><mc:Choice Requires="u"><w:p><w:altChunk r:id="inactive-choice"/></w:p></mc:Choice><mc:Fallback><w:p><w:altChunk r:id="active-fallback"/></w:p></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="w"><w:tbl><w:altChunk r:id="active-choice"/></w:tbl></mc:Choice><mc:Fallback><w:tbl><w:altChunk r:id="inactive-fallback"/></w:tbl></mc:Fallback></mc:AlternateContent><w:p/>"#,
    );
    assert_expected(
        "MCE fallback and supported choice",
        xml.as_bytes(),
        &[
            (
                r#"<w:altChunk r:id="active-fallback"/>"#,
                "active-fallback",
                None,
            ),
            (
                r#"<w:altChunk r:id="active-choice"/>"#,
                "active-choice",
                None,
            ),
        ],
        &[
            (0, r#"<w:p><w:altChunk r:id="active-fallback"/></w:p>"#),
            (1, r#"<w:tbl><w:altChunk r:id="active-choice"/></w:tbl>"#),
            (0, r#"<w:p/>"#),
        ],
    );

    let strict_xml = strict_mce(
        r#"<mc:AlternateContent><mc:Choice Requires="u"><s:altChunk r:id="strict-inactive"/></mc:Choice><mc:Fallback><s:altChunk r:id="strict-active"/></mc:Fallback></mc:AlternateContent>"#,
    );
    assert_expected(
        "strict MCE fallback",
        strict_xml.as_bytes(),
        &[(
            r#"<s:altChunk r:id="strict-active"/>"#,
            "strict-active",
            None,
        )],
        &[(2, r#"<s:altChunk r:id="strict-active"/>"#)],
    );
}

#[test]
fn malformed_first_mce_is_still_seen_by_the_second_active_selection() {
    let no_anchor = mce(r#"<mc:AlternateContent><w:p/></mc:AlternateContent><w:p/>"#);
    assert_parity(
        "malformed MCE with empty first active offsets",
        no_anchor.as_bytes(),
    );
    let error =
        frozen_scan(no_anchor.as_bytes()).expect_err("second active call must diagnose MCE");
    assert!(matches!(error, Error::Mce(_)), "{error:?}");

    let must_understand = mce(r#"<w:p mc:MustUnderstand="u"/><w:p/>"#);
    assert_parity(
        "unknown MustUnderstand with empty first active offsets",
        must_understand.as_bytes(),
    );
    let error = frozen_scan(must_understand.as_bytes())
        .expect_err("second active call must diagnose MustUnderstand");
    assert!(matches!(error, Error::Mce(_)), "{error:?}");
}

#[test]
fn first_active_error_wins_over_a_deferred_range_error() {
    let mut wrappers = String::new();
    for _ in 0..127 {
        wrappers.push_str("<u:wrap>");
    }
    wrappers.push_str(r#"<w:altChunk r:id="valid"/>"#);
    for _ in 0..127 {
        wrappers.push_str("</u:wrap>");
    }
    wrappers.push_str(r#"<mc:AlternateContent><w:p/></mc:AlternateContent>"#);
    let xml = mce(&wrappers);
    assert_parity(
        "first active MCE error beats deferred range depth",
        xml.as_bytes(),
    );
    let error = frozen_scan(xml.as_bytes()).expect_err("first active call must fail");
    assert!(matches!(error, Error::Mce(_)), "{error:?}");
}

#[test]
fn fragment_and_foreign_namespace_rules_remain_distinct_from_alt_namespace_rules() {
    let foreign = transitional(
        r#"<x:foreign xmlns:x="urn:foreign"><x:p/><x:tbl/><x:altChunk r:id="foreign"/></x:foreign><w:p/><w:altChunk r:id="bound"/>"#,
    );
    assert_expected(
        "foreign namespace is ignored",
        foreign.as_bytes(),
        &[(r#"<w:altChunk r:id="bound"/>"#, "bound", None)],
        &[(0, r#"<w:p/>"#), (2, r#"<w:altChunk r:id="bound"/>"#)],
    );

    let unbound = unbound_fragment(r#"<p/><altChunk/><foreign><p/><altChunk/></foreign>"#);
    assert_expected(
        "unbound fragment prefix heuristic",
        unbound.as_bytes(),
        &[],
        &[
            (0, r#"<p/>"#),
            (2, r#"<altChunk/>"#),
            (0, r#"<p/>"#),
            (2, r#"<altChunk/>"#),
        ],
    );

    let unknown = unknown_prefix_fragment(r#"<q:p/><q:altChunk/><q:foreign><q:p/></q:foreign>"#);
    assert_expected(
        "unknown first fragment prefix heuristic",
        unknown.as_bytes(),
        &[],
        &[(0, r#"<q:p/>"#), (2, r#"<q:altChunk/>"#), (0, r#"<q:p/>"#)],
    );
}

#[test]
fn typed_alt_child_grammar_and_relationship_errors_are_unchanged() {
    let cases = [
        (
            "missing relationship",
            transitional(r#"<w:altChunk/>"#),
            "lacks a relationship ID",
        ),
        (
            "unexpected text",
            transitional(r#"<w:altChunk r:id="bad">text</w:altChunk>"#),
            "unexpected text",
        ),
        (
            "duplicate relationship IDs",
            transitional(
                r#"<w:altChunk xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="one" q:id="two"/>"#,
            ),
            "duplicate relationship IDs",
        ),
        (
            "unsafe relationship ID",
            transitional(r#"<w:altChunk r:id="bad&amp;id"/>"#),
            "safe XML attribute value",
        ),
        (
            "duplicate properties",
            transitional(r#"<w:altChunk r:id="bad"><w:altChunkPr/><w:altChunkPr/></w:altChunk>"#),
            "invalid child content",
        ),
        (
            "invalid matchSrc",
            transitional(
                r#"<w:altChunk r:id="bad"><w:altChunkPr><w:matchSrc w:val="maybe"/></w:altChunkPr></w:altChunk>"#,
            ),
            "invalid matchSrc value",
        ),
    ];
    for (name, xml, detail) in cases {
        let expected = frozen_scan(xml.as_bytes()).expect_err(name);
        let actual = candidate_scan(xml.as_bytes()).expect_err(name);
        assert_eq!(
            error_fingerprint(&actual),
            error_fingerprint(&expected),
            "{name}"
        );
        assert!(
            expected.to_string().contains(detail),
            "{name}: expected detail {detail:?}, got {}",
            expected
        );
    }

    let opaque = transitional(
        r#"<w:altChunk r:id="opaque"><x:payload xmlns:x="urn:foreign"><w:altChunkPr/></x:payload></w:altChunk>"#,
    );
    assert_expected(
        "foreign alt child is opaque",
        opaque.as_bytes(),
        &[(
            r#"<w:altChunk r:id="opaque"><x:payload xmlns:x="urn:foreign"><w:altChunkPr/></x:payload></w:altChunk>"#,
            "opaque",
            None,
        )],
        &[(
            2,
            r#"<w:altChunk r:id="opaque"><x:payload xmlns:x="urn:foreign"><w:altChunkPr/></x:payload></w:altChunk>"#,
        )],
    );
}

#[test]
fn malformed_tail_uses_the_existing_reader_error() {
    let xml = malformed_tail();
    assert_parity("malformed tail", xml.as_bytes());
    let expected = frozen_scan(xml.as_bytes()).expect_err("malformed tail must fail");
    assert!(matches!(expected, Error::Xml(_)), "{expected:?}");
}

#[test]
fn later_alt_error_wins_over_an_earlier_deferred_range_error() {
    // The document/body plus 127 wrappers exceed the range scanner's depth
    // limit, while the alt scanner's 256-level limit still accepts the walk.
    // The malformed anchor comes after the range refusal point and therefore
    // proves that a fused reader cannot return the range error immediately.
    let xml = deep_transitional(127, r#"<w:altChunk/>"#);
    assert_parity(
        "later alt relationship error beats range depth",
        xml.as_bytes(),
    );
    let expected = frozen_scan(xml.as_bytes()).expect_err("malformed alt must fail");
    assert!(
        expected.to_string().contains("relationship ID"),
        "unexpected precedence oracle result: {expected:?}"
    );
}

#[test]
fn range_error_is_returned_after_alt_success() {
    let xml = deep_transitional(127, r#"<w:p/>"#);
    assert_parity(
        "range depth error after successful alt scan",
        xml.as_bytes(),
    );
    let expected = frozen_scan(xml.as_bytes()).expect_err("range depth must fail");
    assert!(
        expected
            .to_string()
            .contains("nesting exceeds the 128 depth limit"),
        "unexpected range error: {expected:?}"
    );
}

#[test]
fn exact_range_depth_boundary_remains_accepted() {
    let xml = deep_transitional(126, r#"<w:p/>"#);
    assert_expected(
        "range depth 128 boundary",
        xml.as_bytes(),
        &[],
        &[(0, r#"<w:p/>"#)],
    );
}

#[test]
fn alt_depth_boundary_remains_256_even_though_range_limit_is_lower() {
    let accepted = alt_depth_probe(MAX_XML_DEPTH);
    assert!(
        frozen_alt_scan(&accepted).is_ok(),
        "exact alt depth must remain accepted"
    );
    let rejected = alt_depth_probe(MAX_XML_DEPTH + 1);
    let error = frozen_alt_scan(&rejected).expect_err("one level beyond alt depth must fail");
    assert!(
        error.to_string().contains("256 nesting levels"),
        "unexpected alt depth error: {error:?}"
    );
}

#[test]
fn empty_event_depth_contract_does_not_share_the_range_counter() {
    for depth in [MAX_XML_DEPTH - 1, MAX_XML_DEPTH] {
        let mut xml = "<x:x>".repeat(depth);
        // Empty events check one extra level without retaining the increment.
        // Two siblings and a subsequent start must each use the same depth.
        xml.push_str("<x:empty/><x:empty/><x:probe></x:probe>");
        xml.push_str(&"</x:x>".repeat(depth));
        let result = frozen_alt_scan(xml.as_bytes());
        assert_eq!(result.is_ok(), depth < MAX_XML_DEPTH);
        assert_parity("empty event depth boundary", xml.as_bytes());
    }
}

#[test]
fn anchor_limit_error_precedes_any_range_result() {
    let mut body = String::new();
    for index in 0..=MAX_CHUNKS {
        body.push_str(&format!(r#"<w:altChunk r:id="rId{index}"/>"#));
    }
    let xml = transitional(&body);
    assert_parity("MAX_CHUNKS plus one", xml.as_bytes());
    let expected = frozen_scan(xml.as_bytes()).expect_err("anchor limit must fail");
    assert!(
        expected.to_string().contains("anchor limit"),
        "unexpected anchor-limit error: {expected:?}"
    );
}

#[test]
fn raw_xml_limit_is_checked_before_the_fused_walk() {
    let xml = vec![b' '; MAX_XML_BYTES + 1];
    assert_parity("source byte limit", &xml);
    let expected = frozen_scan(&xml).expect_err("oversized source must fail");
    assert!(
        expected
            .to_string()
            .contains("alternative-format scan input exceeds"),
        "unexpected source-limit error: {expected:?}"
    );
}

#[test]
fn one_million_structural_nodes_keep_the_range_limit_and_error_order() {
    // Two Start events belong to the document and body.  One fewer paragraph
    // would land exactly on MAX_SCAN_NODES; this case adds one and must fail at
    // the range scanner after the alt scanner has completed successfully.
    let paragraph_count = MAX_SCAN_NODES - 1;
    let mut body = String::with_capacity(paragraph_count.saturating_mul(6));
    for _ in 0..paragraph_count {
        body.push_str("<w:p/>");
    }
    let xml = transitional(&body);
    assert_parity("MAX_SCAN_NODES plus one", xml.as_bytes());
    let expected = frozen_scan(xml.as_bytes()).expect_err("node limit must fail");
    assert!(
        expected.to_string().contains("exceeds 1000000 elements"),
        "unexpected node-limit error: {expected:?}"
    );
}

#[test]
fn active_offset_limit_remains_typed_and_bounded() {
    // The structural scanner cannot naturally produce more than one million
    // offsets because it has the same one-million-node ceiling.  Exercise the
    // MCE boundary directly so a future fusion cannot silently loosen the
    // second active-offset call's limit.
    let xml = transitional(r#"<w:p/>"#);
    let offsets = vec![0_u32; MAX_VISIBILITY_OFFSETS + 1];
    let error = active(xml.as_bytes(), &offsets).expect_err("offset limit must fail");
    assert!(matches!(error, Error::Mce(_)), "{error:?}");
    assert!(error.to_string().contains("offset"), "{error:?}");
}
