use crate::{Error, Result};
use litchi_opc::constants::relationship_type;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;

use super::model::{
    Chunk, Conformance, MAX_CHUNKS, MAX_MARKED_XML_BYTES, MAX_VISIBILITY_OFFSETS, MAX_XML_BYTES,
    MAX_XML_DEPTH, Rel, STRICT_RELATIONSHIP, STRICT_RELATIONSHIP_NAMESPACE, STRICT_WORD_NAMESPACE,
    TRANSITIONAL_RELATIONSHIP_NAMESPACE, TRANSITIONAL_WORD_NAMESPACE,
};

impl Chunk {
    /// Serialize this anchor using an isolated, namespace-complete element.
    pub fn xml(&self, conformance: Conformance) -> String {
        let word_ns = conformance.word_namespace();
        let relationship_ns = conformance.relationship_namespace();
        let opening = format!(
            r#"<w:altChunk xmlns:w="{word_ns}" xmlns:r="{relationship_ns}" r:id="{}""#,
            self.relationship().as_str()
        );
        match self.match_source() {
            None => format!("{opening}/>"),
            Some(value) => format!(
                r#"{opening}><w:altChunkPr><w:matchSrc w:val="{}"/></w:altChunkPr></w:altChunk>"#,
                u8::from(value)
            ),
        }
    }
}

/// Whether `value` is a supported alternative-format relationship type.
pub fn is_relationship(value: &str) -> bool {
    matches!(
        value,
        relationship_type::ALTERNATIVE_FORMAT_IMPORT
            | relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
            | STRICT_RELATIONSHIP
    )
}

struct PendingChunk {
    root_depth: usize,
    start: u32,
    relationship: Rel,
    match_source: Option<bool>,
    saw_properties: bool,
    properties_depth: Option<usize>,
    opaque_depth: Option<usize>,
}

/// Retain offsets whose XML positions survive baseline markup-compatibility
/// processing.
///
/// The returned offsets always refer to `xml`, not to a rewritten MCE view.
/// Input order is preserved. This low-level helper lets package facades retain
/// exact source ranges while selecting only the active `mc:Choice` or
/// `mc:Fallback` branch.
pub fn active(xml: &[u8], offsets: &[u32]) -> Result<Vec<u32>> {
    validate_xml(xml)?;
    let limits = litchi_ooxml_common::mce::OffsetLimits {
        max_source_bytes: MAX_XML_BYTES,
        max_offsets: MAX_VISIBILITY_OFFSETS,
        max_marked_bytes: MAX_MARKED_XML_BYTES,
        processing: litchi_ooxml_common::mce::Limits {
            max_input_bytes: MAX_MARKED_XML_BYTES,
            max_output_bytes: MAX_MARKED_XML_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_namespace_bindings: 4096,
            max_directive_tokens: 4096,
            max_choices_per_alternate: 1024,
        },
    };
    litchi_ooxml_common::mce::active_offsets(
        xml,
        offsets,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )
    .map_err(Error::from)
}

/// Parse every altChunk anchor against the full namespace context.
pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut pending: Option<PendingChunk> = None;
    let mut chunks = BTreeMap::new();

    loop {
        let event_start = u32::try_from(reader.buffer_position())
            .map_err(|_| Error::Invalid("altChunk XML offset does not fit u32".into()))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                let event_depth = next_depth(depth)?;
                if pending.is_none()
                    && is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    pending = Some(PendingChunk {
                        root_depth: event_depth,
                        start: event_start,
                        relationship: relationship(&element, decoder, &resolver)?,
                        match_source: None,
                        saw_properties: false,
                        properties_depth: None,
                        opaque_depth: None,
                    });
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
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
                let event_depth = next_depth(depth)?;
                if pending.is_none()
                    && is_word_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    let chunk = Chunk::new(relationship(&element, decoder, &resolver)?, None);
                    insert_chunk(&mut chunks, event_start, chunk)?;
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
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
                    insert_chunk(
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
            _ => {},
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

fn validate_xml(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "alternative-format scan input exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    Ok(())
}

fn parse_child(
    chunk: &mut PendingChunk,
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
    let is_word = is_word_namespace(namespace);
    if !is_word {
        if !empty {
            chunk.opaque_depth = Some(depth);
        }
        return Ok(());
    }
    let properties_depth = chunk
        .root_depth
        .checked_add(1)
        .ok_or_else(|| invalid("altChunk XML nesting is too deep"))?;
    let value_depth = properties_depth
        .checked_add(1)
        .ok_or_else(|| invalid("altChunk XML nesting is too deep"))?;
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
        chunk.match_source = Some(parse_on_off(
            element,
            decoder,
            resolver,
            is_transitional_word_namespace(namespace),
        )?);
        return Ok(());
    }
    Err(Error::Invalid("altChunk has invalid child content".into()))
}

fn relationship(
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
                if uri == TRANSITIONAL_RELATIONSHIP_NAMESPACE
                    || uri == STRICT_RELATIONSHIP_NAMESPACE
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

fn parse_on_off(
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
        if !is_word_namespace(&namespace) && !matches!(namespace, ResolveResult::Unbound) {
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

fn insert_chunk(chunks: &mut BTreeMap<u32, Chunk>, start: u32, chunk: Chunk) -> Result<()> {
    if chunks.len() >= MAX_CHUNKS {
        return Err(invalid("alternative-format anchor limit exceeded"));
    }
    if chunks.insert(start, chunk).is_some() {
        return Err(Error::Invalid("duplicate altChunk XML position".into()));
    }
    Ok(())
}

fn is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri))
            if *uri == TRANSITIONAL_WORD_NAMESPACE || *uri == STRICT_WORD_NAMESPACE
    )
}

fn is_transitional_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(uri)) if *uri == TRANSITIONAL_WORD_NAMESPACE
    )
}

fn next_depth(depth: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("alternative-format XML nesting overflowed"))?;
    if next > MAX_XML_DEPTH {
        return Err(invalid(format!(
            "alternative-format XML exceeds {MAX_XML_DEPTH} nesting levels"
        )));
    }
    Ok(next)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
