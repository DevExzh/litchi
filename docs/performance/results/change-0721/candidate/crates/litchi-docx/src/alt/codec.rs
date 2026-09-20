#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::namespace::WordElementRangeScanner;
use crate::parts::document_part::reserve_document_value;
use crate::{Error, Result};
use litchi_opc::constants::relationship_type;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

use super::model::{
    Chunk, Conformance, MAX_CHUNKS, MAX_MARKED_XML_BYTES, MAX_VISIBILITY_OFFSETS, MAX_XML_BYTES,
    MAX_XML_DEPTH, Rel, STRICT_RELATIONSHIP, STRICT_RELATIONSHIP_NAMESPACE, STRICT_WORD_NAMESPACE,
    TRANSITIONAL_RELATIONSHIP_NAMESPACE, TRANSITIONAL_WORD_NAMESPACE,
};

impl Chunk {
    /// Serialize this anchor using an isolated, namespace-complete element.
    #[must_use]
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
#[must_use]
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

struct AltScanState {
    depth: usize,
    pending: Option<PendingChunk>,
    chunks: BTreeMap<u32, Chunk>,
}

impl AltScanState {
    fn new() -> Self {
        Self {
            depth: 0,
            pending: None,
            chunks: BTreeMap::new(),
        }
    }

    /// Consume one event using the established alternative-format grammar.
    /// Returns `true` after a complete EOF event.
    fn observe(
        &mut self,
        event_start: u32,
        decoder: quick_xml::encoding::Decoder,
        resolver: &NamespaceResolver,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
    ) -> Result<bool> {
        match event {
            Event::Start(element) => {
                let event_depth = next_depth(self.depth)?;
                if self.pending.is_none()
                    && is_word_namespace(namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    self.pending = Some(PendingChunk {
                        root_depth: event_depth,
                        start: event_start,
                        relationship: relationship(element, decoder, resolver)?,
                        match_source: None,
                        saw_properties: false,
                        properties_depth: None,
                        opaque_depth: None,
                    });
                } else if let Some(chunk) = self.pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        namespace,
                        element,
                        decoder,
                        resolver,
                        false,
                    )?;
                }
                self.depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = next_depth(self.depth)?;
                if self.pending.is_none()
                    && is_word_namespace(namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    let chunk = Chunk::new(relationship(element, decoder, resolver)?, None);
                    insert_chunk(&mut self.chunks, event_start, chunk)?;
                } else if let Some(chunk) = self.pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        namespace,
                        element,
                        decoder,
                        resolver,
                        true,
                    )?;
                }
            },
            Event::End(_) => {
                if let Some(chunk) = self.pending.as_mut()
                    && chunk.opaque_depth == Some(self.depth)
                {
                    chunk.opaque_depth = None;
                }
                if let Some(chunk) = self.pending.as_mut()
                    && chunk.properties_depth == Some(self.depth)
                {
                    chunk.properties_depth = None;
                }
                if self
                    .pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.root_depth == self.depth)
                {
                    let chunk = self
                        .pending
                        .take()
                        .ok_or_else(|| Error::Invalid("missing pending altChunk".into()))?;
                    insert_chunk(
                        &mut self.chunks,
                        chunk.start,
                        Chunk::new(chunk.relationship, chunk.match_source),
                    )?;
                }
                self.depth = self
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("unexpected altChunk XML end element".into()))?;
            },
            Event::Text(text)
                if self
                    .pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none())
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::Invalid("altChunk contains unexpected text".into()));
            },
            Event::CData(_) | Event::GeneralRef(_)
                if self
                    .pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.opaque_depth.is_none()) =>
            {
                return Err(Error::Invalid(
                    "altChunk contains unexpected character data".into(),
                ));
            },
            Event::Eof => {
                if self.pending.is_some() {
                    return Err(Error::Invalid("unterminated altChunk".into()));
                }
                return Ok(true);
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        Ok(false)
    }
}

/// Retain offsets whose XML positions survive baseline markup-compatibility
/// processing.
///
/// The returned offsets always refer to `xml`, not to a rewritten MCE view.
/// Input order is preserved. This low-level helper lets package facades retain
/// exact source ranges while selecting only the active `mc:Choice` or
/// `mc:Fallback` branch.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
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
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn scan(xml: &[u8]) -> Result<BTreeMap<u32, Chunk>> {
    validate_xml(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut scanner = AltScanState::new();

    loop {
        let event_start = u32::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::Invalid("altChunk XML offset does not fit u32".into())
        })?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if scanner.observe(event_start, decoder, &resolver, &namespace, &event)? {
            break;
        }
    }

    let offsets = scanner.chunks.keys().copied().collect::<Vec<_>>();
    let active_offsets = active(xml, &offsets)?;
    let mut selected = active_offsets.into_iter();
    let mut next = selected.next();
    scanner.chunks.retain(|offset, _| {
        if next == Some(*offset) {
            next = selected.next();
            true
        } else {
            false
        }
    });
    Ok(scanner.chunks)
}

/// Parse alternative-format anchors and source-preserving document block
/// ranges from one lexical event walk.
///
/// This is private to the mutable document writer. The two returned views
/// deliberately retain the independent offset vectors and the two ordered
/// markup-compatibility selections used by the legacy callers.
pub(crate) fn scan_with_block_ranges(
    xml: &[u8],
) -> Result<(BTreeMap<u32, Chunk>, Vec<(usize, u32, u32)>)> {
    validate_xml(xml)?;
    let targets = [b"p".as_slice(), b"tbl".as_slice(), b"altChunk".as_slice()];
    let mut range_scanner = WordElementRangeScanner::new(&targets);
    let mut ranges = Vec::new();
    let mut range_error = None;
    let mut reader = NsReader::from_reader(xml);
    let mut alt_scanner = AltScanState::new();

    loop {
        // The alternative-format parser's u32 conversion remains the first
        // refusal at this boundary, as it was in `scan`. The range scanner's
        // usize conversion is remembered and deferred when the platform can
        // represent fewer bytes than quick-xml's position counter.
        let raw_start = reader.buffer_position();
        let event_start = u32::try_from(raw_start).map_err(|_source_error| {
            Error::Invalid("altChunk XML offset does not fit u32".into())
        })?;
        let event_start_usize = if range_error.is_none() {
            match usize::try_from(raw_start) {
                Ok(value) => Some(value),
                Err(_source_error) => {
                    range_error = Some(Error::InvalidFormat(
                        "Word XML offset does not fit usize".to_string(),
                    ));
                    None
                },
            }
        } else {
            None
        };
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        // Alternative-format validation has the earlier observable stage.
        // Borrowing the event lets the range state consume the same resolved
        // event without changing either parser's owned-event behavior.
        let eof = alt_scanner.observe(event_start, decoder, &resolver, &namespace, &event)?;

        if range_error.is_none() {
            if let Some(event_start) = event_start_usize {
                let scan_event = match range_scanner.classify(&namespace, &event) {
                    Ok(scan_event) => Some(scan_event),
                    Err(error) => {
                        range_error = Some(error);
                        None
                    },
                };
                if let Some(scan_event) = scan_event {
                    let event_end = match usize::try_from(reader.buffer_position()) {
                        Ok(value) => Some(value),
                        Err(_source_error) => {
                            range_error = Some(Error::InvalidFormat(
                                "Word XML offset does not fit usize".to_string(),
                            ));
                            None
                        },
                    };
                    if let Some(event_end) = event_end {
                        let mut emit = |target: usize, start: u32, length: u32| {
                            reserve_document_value(&mut ranges, "active document block ranges")?;
                            ranges.push((target, start, length));
                            Ok(())
                        };
                        if let Err(error) =
                            range_scanner.finish(event_start, event_end, scan_event, &mut emit)
                        {
                            range_error = Some(error);
                        }
                    }
                }
            }
        }

        if eof {
            break;
        }
    }

    let offsets = alt_scanner.chunks.keys().copied().collect::<Vec<_>>();
    let active_offsets = active(xml, &offsets)?;
    let mut selected = active_offsets.into_iter();
    let mut next = selected.next();
    alt_scanner.chunks.retain(|offset, _| {
        if next == Some(*offset) {
            next = selected.next();
            true
        } else {
            false
        }
    });

    // The old range walk completed before its second MCE call. Preserve that
    // boundary, while allowing a later alt error or the first active call to
    // take precedence over a remembered range refusal.
    if let Some(error) = range_error {
        return Err(error);
    }
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
    Ok((alt_scanner.chunks, ranges))
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
