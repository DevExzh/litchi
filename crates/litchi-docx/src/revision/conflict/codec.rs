#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::similar_names,
    reason = "domain names mirror distinct OOXML roles"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
#![expect(
    clippy::too_many_arguments,
    reason = "the signature mirrors the corresponding OOXML record"
)]
//! Bounded two-pass, source-coordinate parser for Word 2010 conflict markup.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

use super::{AttributeSpan, Conflict, Id, Inventory, Kind, Limits, Metadata, Range, Scope, Span};

const W14: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const W: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
// `active_offsets` inserts `<!--` + its fixed 38-byte marker + decimal index
// + `-->` before every candidate.  Keep this coupled to common's published
// implementation bound and cap the marked input independently of raw XML.
const MCE_MARKER_FIXED_BYTES: usize = 45;
const MAX_MCE_MARKED_BYTES: usize = 256 * 1024 * 1024;
const MAX_MCE_PROCESSED_BYTES: usize = 512 * 1024 * 1024;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Ns {
    Word,
    Math,
    W14,
    Other,
}
#[derive(Clone)]
struct Frame {
    local: Vec<u8>,
    ns: Ns,
    scope: Option<Scope>,
    ignorable_w14: bool,
    parent_is_ppr: bool,
    paragraph_mark_rpr: bool,
    ruby_sdt: bool,
    process_directives: Vec<(Vec<u8>, Vec<u8>)>,
    transparent: bool,
    conflict: Option<usize>,
    active: bool,
    range_content_start: Option<usize>,
}

struct MceFrame {
    declarations: Vec<(Vec<u8>, Option<Vec<u8>>)>,
    ignorable_w14: bool,
}

struct Open {
    metadata: Metadata,
    start: Span,
    id: Option<AttributeSpan>,
    author: Option<AttributeSpan>,
    date: Option<AttributeSpan>,
}

pub(crate) fn parse(source: &[u8], limits: Limits) -> Result<Inventory> {
    if source.len() > limits.max_source_bytes {
        return Err(invalid("conflict source exceeds configured byte limit"));
    }
    let active = active_starts(source, limits)?;
    let mut reader = NsReader::from_reader(source);
    let mut frames = Vec::<Frame>::new();
    frames
        .try_reserve_exact(limits.max_depth.min(256))
        .map_err(alloc("conflict XML stack"))?;
    let mut inventory = Inventory::default();
    inventory
        .conflicts
        .try_reserve_exact(limits.max_conflicts.min(1024))
        .map_err(alloc("conflict inventory"))?;
    inventory
        .ranges
        .try_reserve_exact(limits.max_ranges.min(1024))
        .map_err(alloc("conflict ranges"))?;
    let mut open = HashMap::<(Kind, i32), Open>::new();
    let mut seen = HashSet::<(Kind, i32)>::new();
    let mut events = 0usize;
    let mut metadata_bytes = 0usize;
    let mut text_bytes = 0usize;
    let mut text_segment_count = 0usize;
    let mut text_segments = Vec::<Vec<Span>>::new();
    let max_process_targets = limits
        .max_attributes
        .checked_mul(limits.max_depth)
        .ok_or_else(|| invalid("conflict ProcessContent limit overflow"))?
        .min(4096);
    let mut active_process_content = HashMap::<(Vec<u8>, Vec<u8>), usize>::new();
    active_process_content
        .try_reserve(max_process_targets.min(256))
        .map_err(alloc("conflict MCE ProcessContent map"))?;

    loop {
        let begin = pos(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|e| Error::Xml(e.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = pos(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("conflict XML event counter overflow"))?;
        if events > limits.max_events {
            return Err(invalid("conflict XML exceeds configured event limit"));
        }
        match event {
            Event::Start(element) => {
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("conflict XML depth overflow"))?;
                if depth > limits.max_depth {
                    return Err(invalid("conflict XML exceeds configured nesting limit"));
                }
                let (kind, range) = classify(&namespace, &element);
                let selected = active.contains(&begin);
                let raw_parent = frames.last();
                let parent = effective_parent(&frames);
                let transparent = mce_transparent(
                    &namespace,
                    element.local_name().as_ref(),
                    &active_process_content,
                );
                let process_directives =
                    process_content_directives(&element, decoder, &resolver, limits, true)?;
                if let Some(kind) = kind.filter(|_| selected) {
                    require_ignorable(frames.last(), &element, decoder, &resolver)?;
                    let scope = scope(parent)?;
                    let parsed = metadata(
                        source,
                        begin,
                        end,
                        &element,
                        decoder,
                        &resolver,
                        limits,
                        &mut metadata_bytes,
                    )?;
                    if inventory.conflicts.len() >= limits.max_conflicts {
                        return Err(invalid("conflict count exceeds configured limit"));
                    }
                    inventory
                        .conflicts
                        .try_reserve(1)
                        .map_err(alloc("conflict inventory"))?;
                    inventory.conflicts.push(Conflict {
                        kind,
                        scope,
                        metadata: parsed.0,
                        span: Span::new(begin, begin)?,
                        start_tag: Span::new(begin, end)?,
                        id_span: parsed.1,
                        content: Span::new(end, end)?,
                        text: Arc::from([]),
                        author_span: parsed.2,
                        date_span: parsed.3,
                    });
                    text_segments
                        .try_reserve(1)
                        .map_err(alloc("conflict text segments"))?;
                    text_segments.push(Vec::new());
                    push_process_directives(
                        &mut active_process_content,
                        &process_directives,
                        max_process_targets,
                    )?;
                    frames.push(Frame {
                        local: element.local_name().as_ref().to_vec(),
                        ns: Ns::W14,
                        scope: Some(scope),
                        ignorable_w14: true,
                        parent_is_ppr: false,
                        paragraph_mark_rpr: false,
                        ruby_sdt: false,
                        process_directives,
                        transparent: false,
                        conflict: Some(inventory.conflicts.len() - 1),
                        active: true,
                        range_content_start: None,
                    });
                } else {
                    let range_content_start =
                        if let Some((kind, start)) = range.filter(|_| selected) {
                            require_ignorable(raw_parent, &element, decoder, &resolver)?;
                            range_parent(parent)?;
                            range_marker(
                                source,
                                begin,
                                end,
                                kind,
                                start,
                                &element,
                                decoder,
                                &resolver,
                                limits,
                                &mut metadata_bytes,
                                &mut open,
                                &mut seen,
                                &mut inventory,
                            )?;
                            Some(end)
                        } else {
                            None
                        };
                    let inherited = raw_parent.is_some_and(|frame| frame.ignorable_w14);
                    push_process_directives(
                        &mut active_process_content,
                        &process_directives,
                        max_process_targets,
                    )?;
                    frames.push(Frame {
                        local: element.local_name().as_ref().to_vec(),
                        ns: ns(&namespace),
                        scope: None,
                        ignorable_w14: inherited
                            || declares_w14_ignorable(&element, decoder, &resolver)?,
                        parent_is_ppr: parent.is_some_and(|parent| {
                            parent.ns == Ns::Word && parent.local.as_slice() == b"pPr"
                        }),
                        paragraph_mark_rpr: ns(&namespace) == Ns::Word
                            && match element.local_name().as_ref() {
                                b"rPr" => parent.is_some_and(|parent| {
                                    (parent.ns == Ns::Word && parent.local.as_slice() == b"pPr")
                                        || (parent.ns == Ns::Word
                                            && parent.local.as_slice() == b"rPrChange"
                                            && parent.paragraph_mark_rpr)
                                }),
                                b"rPrChange" => {
                                    parent.is_some_and(|parent| parent.paragraph_mark_rpr)
                                },
                                _ => false,
                            },
                        ruby_sdt: ns(&namespace) == Ns::Word
                            && element.local_name().as_ref() == b"sdt"
                            && parent.is_some_and(|parent| {
                                parent.ns == Ns::Word
                                    && matches!(
                                        parent.local.as_slice(),
                                        b"customXml"
                                            | b"fldSimple"
                                            | b"hyperlink"
                                            | b"rt"
                                            | b"rubyBase"
                                            | b"sdtContent"
                                    )
                            }),
                        transparent,
                        process_directives,
                        conflict: None,
                        active: selected,
                        range_content_start,
                    });
                }
            },
            Event::Empty(element) => {
                let (kind, range) = classify(&namespace, &element);
                process_content_directives(&element, decoder, &resolver, limits, false)?;
                if let Some(kind) = kind.filter(|_| active.contains(&begin)) {
                    require_ignorable(frames.last(), &element, decoder, &resolver)?;
                    let scope = scope(effective_parent(&frames))?;
                    let parsed = metadata(
                        source,
                        begin,
                        end,
                        &element,
                        decoder,
                        &resolver,
                        limits,
                        &mut metadata_bytes,
                    )?;
                    if inventory.conflicts.len() >= limits.max_conflicts {
                        return Err(invalid("conflict count exceeds configured limit"));
                    }
                    inventory
                        .conflicts
                        .try_reserve(1)
                        .map_err(alloc("conflict inventory"))?;
                    inventory.conflicts.push(Conflict {
                        kind,
                        scope,
                        metadata: parsed.0,
                        span: Span::new(begin, end)?,
                        start_tag: Span::new(begin, end)?,
                        id_span: parsed.1,
                        content: Span::new(end, end)?,
                        text: Arc::from([]),
                        author_span: parsed.2,
                        date_span: parsed.3,
                    });
                    text_segments
                        .try_reserve(1)
                        .map_err(alloc("conflict text segments"))?;
                    text_segments.push(Vec::new());
                } else if let Some((kind, start)) = range.filter(|_| active.contains(&begin)) {
                    require_ignorable(frames.last(), &element, decoder, &resolver)?;
                    range_parent(effective_parent(&frames))?;
                    range_marker(
                        source,
                        begin,
                        end,
                        kind,
                        start,
                        &element,
                        decoder,
                        &resolver,
                        limits,
                        &mut metadata_bytes,
                        &mut open,
                        &mut seen,
                        &mut inventory,
                    )?;
                }
            },
            Event::Text(text) => {
                if active.contains(&begin) {
                    retain_text_span(
                        &frames,
                        &mut text_segments,
                        begin,
                        end,
                        text.as_ref().len(),
                        limits,
                        &mut text_bytes,
                        &mut text_segment_count,
                    )?;
                }
            },
            Event::CData(text) => {
                if active.contains(&begin) {
                    retain_text_span(
                        &frames,
                        &mut text_segments,
                        begin,
                        end,
                        text.as_ref().len(),
                        limits,
                        &mut text_bytes,
                        &mut text_segment_count,
                    )?;
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
                pop_process_directives(&mut active_process_content, frame.process_directives)?;
                if frame
                    .range_content_start
                    .is_some_and(|start| start != begin)
                {
                    return Err(invalid("conflict range markers must be childless"));
                }
                if let Some(index) = frame.conflict {
                    let conflict = &mut inventory.conflicts[index];
                    conflict.content = Span::new(conflict.start_tag.end(), begin)?;
                    conflict.span = Span::new(conflict.span.start(), end)?;
                    if conflict.scope == Scope::Property && !conflict.content.is_empty() {
                        return Err(invalid("property conflict marker must be childless"));
                    }
                }
            },
            Event::Eof => {
                if !frames.is_empty() {
                    return Err(invalid("unterminated conflict XML"));
                }
                if !open.is_empty() {
                    return Err(invalid("orphaned conflict range start marker"));
                }
                break;
            },
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    for (conflict, spans) in inventory.conflicts.iter_mut().zip(text_segments) {
        conflict.set_text_spans(Arc::from(spans));
    }
    inventory
        .ranges
        .sort_by_key(|range| range.start_span.start());
    Ok(inventory)
}

fn effective_parent(frames: &[Frame]) -> Option<&Frame> {
    frames
        .iter()
        .rev()
        .find(|frame| frame.active && !frame.transparent)
}

fn namespace_bytes(namespace: &ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => (*value).to_vec(),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Vec::new(),
    }
}

fn process_content_directives(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    limits: Limits,
    retain: bool,
) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut names = Vec::new();
    let per_directive_limit = limits.max_attributes.min(4096);
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MC)
            || attribute.key.local_name().as_ref() != b"ProcessContent"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let mut tokens = 0usize;
        for token in value.split_ascii_whitespace() {
            tokens = tokens
                .checked_add(1)
                .ok_or_else(|| invalid("mc:ProcessContent token counter overflow"))?;
            if tokens > per_directive_limit {
                return Err(invalid("mc:ProcessContent exceeds configured token limit"));
            }
            let (namespace, name) = resolver.resolve_element(QName(token.as_bytes()));
            let ResolveResult::Bound(Namespace(namespace)) = namespace else {
                return Err(invalid("mc:ProcessContent target has no namespace binding"));
            };
            if retain {
                names
                    .try_reserve(1)
                    .map_err(alloc("conflict MCE ProcessContent directives"))?;
                names.push(((*namespace).to_vec(), name.as_ref().to_vec()));
            }
        }
    }
    Ok(names)
}

fn push_process_directives(
    active: &mut HashMap<(Vec<u8>, Vec<u8>), usize>,
    directives: &[(Vec<u8>, Vec<u8>)],
    max_targets: usize,
) -> Result<()> {
    for (namespace, local) in directives {
        let key = (namespace.clone(), local.clone());
        if let Some(count) = active.get_mut(&key) {
            *count = count
                .checked_add(1)
                .ok_or_else(|| invalid("mc:ProcessContent reference counter overflow"))?;
        } else {
            if active.len() >= max_targets {
                return Err(invalid("mc:ProcessContent exceeds active target limit"));
            }
            active
                .try_reserve(1)
                .map_err(alloc("conflict MCE ProcessContent map"))?;
            active.insert(key, 1);
        }
    }
    Ok(())
}

fn pop_process_directives(
    active: &mut HashMap<(Vec<u8>, Vec<u8>), usize>,
    directives: Vec<(Vec<u8>, Vec<u8>)>,
) -> Result<()> {
    for directive in directives.into_iter().rev() {
        let remove = {
            let count = active
                .get_mut(&directive)
                .ok_or_else(|| invalid("mc:ProcessContent scope underflow"))?;
            *count = count
                .checked_sub(1)
                .ok_or_else(|| invalid("mc:ProcessContent scope underflow"))?;
            *count == 0
        };
        if remove {
            active.remove(&directive);
        }
    }
    Ok(())
}

fn mce_transparent(
    namespace: &ResolveResult<'_>,
    local: &[u8],
    process_content: &HashMap<(Vec<u8>, Vec<u8>), usize>,
) -> bool {
    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MC)
        && matches!(local, b"AlternateContent" | b"Choice" | b"Fallback")
    {
        return true;
    }
    let namespace = namespace_bytes(namespace);
    process_content.contains_key(&(namespace, local.to_vec()))
}

/// A selected segment is retained once for every enclosing selected inline
/// conflict. Both byte and segment quotas charge that fan-out, so nesting
/// cannot bypass a budget or create unaccounted retained references.
fn retain_text_span(
    frames: &[Frame],
    segments: &mut [Vec<Span>],
    begin: usize,
    end: usize,
    bytes: usize,
    limits: Limits,
    text_bytes: &mut usize,
    text_segments: &mut usize,
) -> Result<()> {
    let references = frames
        .iter()
        .filter(|frame| {
            frame.active && frame.scope == Some(Scope::Inline) && frame.conflict.is_some()
        })
        .count();
    if references == 0 {
        return Ok(());
    }
    let retained_bytes = bytes
        .checked_mul(references)
        .ok_or_else(|| invalid("conflict text counter overflow"))?;
    let next_bytes = text_bytes
        .checked_add(retained_bytes)
        .ok_or_else(|| invalid("conflict text counter overflow"))?;
    if next_bytes > limits.max_text_bytes {
        return Err(invalid("conflict text exceeds configured limit"));
    }
    let next_segments = text_segments
        .checked_add(references)
        .ok_or_else(|| invalid("conflict text segment counter overflow"))?;
    if next_segments > limits.max_text_segments {
        return Err(invalid("conflict text segments exceed configured limit"));
    }
    let span = Span::new(begin, end)?;
    for index in frames
        .iter()
        .filter(|frame| frame.active && frame.scope == Some(Scope::Inline))
        .filter_map(|frame| frame.conflict)
    {
        let spans = segments
            .get_mut(index)
            .ok_or_else(|| invalid("missing conflict text spans"))?;
        spans.try_reserve(1).map_err(alloc("conflict text spans"))?;
        spans.push(span);
    }
    *text_bytes = next_bytes;
    *text_segments = next_segments;
    Ok(())
}

/// Select all structural and character-data event starts that survive MCE
/// branch selection while retaining coordinates in the immutable source.
fn active_starts(source: &[u8], limits: Limits) -> Result<HashSet<usize>> {
    let mut reader = NsReader::from_reader(source);
    let mut offsets = Vec::<u32>::new();
    // Every semantic event gets an offset. This preserves selected branch and
    // ProcessContent topology for the inventory pass; inactive content cannot
    // supply text, a parent, or a range boundary.
    let max_offsets = limits.max_events;
    offsets
        .try_reserve_exact(max_offsets.min(1024))
        .map_err(alloc("conflict MCE offsets"))?;
    let mut events = 0usize;
    let mut frames = Vec::<MceFrame>::new();
    frames
        .try_reserve_exact(limits.max_depth.min(256))
        .map_err(alloc("conflict MCE stack"))?;
    let max_namespace_bindings = limits
        .max_attributes
        .checked_mul(limits.max_depth)
        .ok_or_else(|| invalid("conflict MCE namespace limit overflow"))?
        .min(4096);
    let mut effective_namespaces = HashMap::<Vec<u8>, Vec<Option<Vec<u8>>>>::new();
    effective_namespaces
        .try_reserve(max_namespace_bindings.min(256))
        .map_err(alloc("conflict MCE namespace map"))?;
    let mut namespace_bindings = 0usize;
    loop {
        let start = pos(&reader)?;
        let event = reader
            .read_event()
            .map_err(|e| Error::Xml(e.to_string()))?
            .into_owned();
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("conflict MCE event counter overflow"))?;
        if events > limits.max_events {
            return Err(invalid(
                "conflict MCE prepass exceeds configured event limit",
            ));
        }
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("conflict MCE depth overflow"))?;
                if depth > limits.max_depth {
                    return Err(invalid(
                        "conflict MCE prepass exceeds configured nesting limit",
                    ));
                }
                validate_lexical_attributes(&element, limits)?;
                let declared = namespace_declarations(&element, max_namespace_bindings)?;
                namespace_bindings = push_namespace_declarations(
                    &mut effective_namespaces,
                    declared.as_slice(),
                    namespace_bindings,
                    max_namespace_bindings,
                )?;
                let (wrapper, range) = classify(&namespace, &element);
                let ignorable_w14 = frames.last().is_some_and(|frame| frame.ignorable_w14)
                    || declares_w14_ignorable(&element, reader.decoder(), &resolver)?;
                if (wrapper.is_some() || range.is_some()) && !ignorable_w14 {
                    return Err(invalid(
                        "W14 conflict markup requires an effective mc:Ignorable binding",
                    ));
                }
                push_offset(&mut offsets, start, max_offsets)?;
                frames.push(MceFrame {
                    declarations: declared,
                    ignorable_w14,
                });
            },
            Event::Empty(element) => {
                validate_lexical_attributes(&element, limits)?;
                let declared = namespace_declarations(&element, max_namespace_bindings)?;
                namespace_bindings = push_namespace_declarations(
                    &mut effective_namespaces,
                    declared.as_slice(),
                    namespace_bindings,
                    max_namespace_bindings,
                )?;
                let (wrapper, range) = classify(&namespace, &element);
                let ignorable_w14 = frames.last().is_some_and(|frame| frame.ignorable_w14)
                    || declares_w14_ignorable(&element, reader.decoder(), &resolver)?;
                if (wrapper.is_some() || range.is_some()) && !ignorable_w14 {
                    return Err(invalid(
                        "W14 conflict markup requires an effective mc:Ignorable binding",
                    ));
                }
                push_offset(&mut offsets, start, max_offsets)?;
                namespace_bindings = pop_namespace_declarations(
                    &mut effective_namespaces,
                    declared,
                    namespace_bindings,
                )?;
            },
            Event::Text(_) | Event::CData(_) => push_offset(&mut offsets, start, max_offsets)?,
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("conflict MCE prepass has unmatched end element"))?;
                namespace_bindings = pop_namespace_declarations(
                    &mut effective_namespaces,
                    frame.declarations,
                    namespace_bindings,
                )?;
            },
            Event::Eof => {
                if !frames.is_empty() {
                    return Err(invalid("unterminated conflict MCE XML"));
                }
                break;
            },
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::ooxml_baseline();
    capabilities.understand_namespace(String::from_utf8_lossy(W14).into_owned());
    // `active_offsets` reserves its source boundary for markers in older
    // common versions, whereas conflict limits are inclusive.  The validated
    // conflict hard cap is far below the common hard cap, so this checked
    // increment preserves the public inclusive contract.
    let mce_source_bytes = limits
        .max_source_bytes
        .checked_add(1)
        .ok_or_else(|| invalid("conflict MCE source limit overflow"))?;
    let digits = decimal_digits(offsets.len().saturating_sub(1));
    let marker_bytes = MCE_MARKER_FIXED_BYTES
        .checked_add(digits)
        .ok_or_else(|| invalid("conflict MCE marker budget overflow"))?;
    let marked_bytes = source
        .len()
        .checked_add(
            offsets
                .len()
                .checked_mul(marker_bytes)
                .ok_or_else(|| invalid("conflict MCE marker budget overflow"))?,
        )
        .ok_or_else(|| invalid("conflict MCE marker budget overflow"))?;
    if marked_bytes > MAX_MCE_MARKED_BYTES {
        return Err(invalid(
            "conflict MCE marked source exceeds hard byte limit",
        ));
    }
    // MCE may repeat effective namespace maps while serializing, but it must
    // not receive an event-times-source budget: that rejects shallow input
    // near its legal source limit. Keep a source/marker-derived request and
    // clamp it to common's independent finite hard ceiling.
    let processed_bytes = marked_bytes
        .checked_mul(32)
        .unwrap_or(MAX_MCE_PROCESSED_BYTES)
        .min(MAX_MCE_PROCESSED_BYTES);
    let processing = litchi_ooxml_common::mce::Limits {
        max_input_bytes: marked_bytes,
        max_output_bytes: processed_bytes,
        max_depth: limits.max_depth,
        max_namespace_bindings: limits
            .max_attributes
            .saturating_mul(limits.max_depth)
            .clamp(1, 4096),
        max_directive_tokens: limits.max_attributes.min(4096),
        max_choices_per_alternate: limits.max_attributes.min(1024),
    };
    let selected = litchi_ooxml_common::mce::active_offsets(
        source,
        &offsets,
        &capabilities,
        &litchi_ooxml_common::mce::OffsetLimits {
            max_source_bytes: mce_source_bytes,
            max_offsets,
            max_marked_bytes: marked_bytes,
            processing,
        },
    )
    .map_err(Error::from)?;
    let mut active = HashSet::new();
    active
        .try_reserve(selected.len())
        .map_err(alloc("active conflict marker offsets"))?;
    for offset in selected {
        active.insert(
            usize::try_from(offset)
                .map_err(|_source_error| invalid("active conflict offset does not fit usize"))?,
        );
    }
    Ok(active)
}

fn push_offset(offsets: &mut Vec<u32>, start: usize, max: usize) -> Result<()> {
    if offsets.len() >= max {
        return Err(invalid(
            "conflict MCE candidate count exceeds configured limit",
        ));
    }
    offsets
        .try_reserve(1)
        .map_err(alloc("conflict MCE offsets"))?;
    offsets.push(
        u32::try_from(start)
            .map_err(|_source_error| invalid("conflict XML offset does not fit u32"))?,
    );
    Ok(())
}

fn validate_lexical_attributes(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("attribute counter overflow"))?;
        if count > limits.max_attributes {
            return Err(invalid("XML element has too many attributes"));
        }
        if attribute.value.len() > limits.max_attribute_bytes {
            return Err(invalid("XML attribute exceeds configured byte limit"));
        }
    }
    Ok(())
}

fn decimal_digits(mut value: usize) -> usize {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

fn range_marker(
    source: &[u8],
    begin: usize,
    end: usize,
    kind: Kind,
    start: bool,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    limits: Limits,
    metadata_bytes: &mut usize,
    open: &mut HashMap<(Kind, i32), Open>,
    seen: &mut HashSet<(Kind, i32)>,
    inventory: &mut Inventory,
) -> Result<()> {
    if start {
        let parsed = metadata(
            source,
            begin,
            end,
            element,
            decoder,
            resolver,
            limits,
            metadata_bytes,
        )?;
        let key = (kind, parsed.0.id.get());
        if seen.contains(&key) || open.contains_key(&key) {
            return Err(invalid("duplicate conflict range identifier"));
        }
        if open.len() >= limits.max_open_ranges {
            return Err(invalid("open conflict ranges exceed configured limit"));
        }
        open.try_reserve(1).map_err(alloc("open conflict ranges"))?;
        seen.try_reserve(1)
            .map_err(alloc("conflict range identifiers"))?;
        seen.insert(key);
        open.insert(
            key,
            Open {
                metadata: parsed.0,
                start: Span::new(begin, end)?,
                id: parsed.1,
                author: parsed.2,
                date: parsed.3,
            },
        );
    } else {
        let (id, id_span) = range_end_id(source, begin, end, element, decoder, resolver, limits)?;
        let key = (kind, id.get());
        let Some(opened) = open.remove(&key) else {
            return Err(invalid(
                "orphaned, wrong-kind, or end-before-start conflict range marker",
            ));
        };
        if inventory.ranges.len() >= limits.max_ranges {
            return Err(invalid("conflict range count exceeds configured limit"));
        }
        inventory
            .ranges
            .try_reserve(1)
            .map_err(alloc("conflict ranges"))?;
        inventory.ranges.push(Range {
            kind,
            metadata: opened.metadata,
            start_span: opened.start,
            end_span: Span::new(begin, end)?,
            start_id_span: opened.id,
            end_id_span: id_span,
            author_span: opened.author,
            date_span: opened.date,
        });
    }
    Ok(())
}

/// Decode the `CT_Markup` end marker.  Unlike `CT_TrackChange` starts it carries
/// only `w:id`; author/date are neither required nor retained.
fn range_end_id(
    source: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    limits: Limits,
) -> Result<(Id, Option<AttributeSpan>)> {
    let mut id = None;
    let mut span = None;
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("attribute counter overflow"))?;
        if count > limits.max_attributes {
            return Err(invalid("conflict range end has too many attributes"));
        }
        if attribute.value.len() > limits.max_attribute_bytes {
            return Err(invalid(
                "conflict range-end attribute exceeds configured byte limit",
            ));
        }
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word(&namespace) {
            return Err(invalid(
                "CT_Markup conflict range end has an unsupported attribute",
            ));
        }
        match attribute.key.local_name().as_ref() {
            b"id" => {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                    .map_err(|e| Error::Xml(e.to_string()))?
                    .into_owned();
                if id.replace(value).is_some() {
                    return Err(invalid("duplicate conflict range-end id attribute"));
                }
                span = Some(find_attr(source, start, end, attribute.key.as_ref())?);
            },
            _ => {
                return Err(invalid(
                    "CT_Markup conflict range end has an unsupported attribute",
                ));
            },
        }
    }
    let value = id.ok_or_else(|| invalid("conflict range end lacks required w:id"))?;
    let id = value
        .parse::<i32>()
        .map_err(|_source_error| invalid("conflict range-end id is not a signed i32"))?;
    Ok((Id::new(id)?, span))
}
fn is_namespace_declaration(key: &[u8]) -> bool {
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

fn metadata(
    source: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    limits: Limits,
    total: &mut usize,
) -> Result<(
    Metadata,
    Option<AttributeSpan>,
    Option<AttributeSpan>,
    Option<AttributeSpan>,
)> {
    let mut id = None;
    let mut author = None;
    let mut date = None;
    let mut id_span = None;
    let mut author_span = None;
    let mut date_span = None;
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("attribute counter overflow"))?;
        if count > limits.max_attributes {
            return Err(invalid("conflict element has too many attributes"));
        }
        if attribute.value.len() > limits.max_attribute_bytes {
            return Err(invalid("conflict attribute exceeds configured byte limit"));
        }
        let (ns, _) = resolver.resolve_attribute(attribute.key);
        if !is_word(&ns) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|e| Error::Xml(e.to_string()))?
            .into_owned();
        let span = find_attr(source, start, end, attribute.key.as_ref())?;
        match attribute.key.local_name().as_ref() {
            b"id" => {
                if id.replace(value).is_some() {
                    return Err(invalid("duplicate conflict id attribute"));
                }
                id_span = Some(span);
            },
            b"author" => {
                if author.is_some() {
                    return Err(invalid("duplicate conflict author attribute"));
                }
                *total = total
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("metadata byte counter overflow"))?;
                if *total > limits.max_metadata_bytes {
                    return Err(invalid("conflict metadata exceeds configured limit"));
                }
                author = Some(value);
                author_span = Some(span);
            },
            b"date" => {
                if date.is_some() {
                    return Err(invalid("duplicate conflict date attribute"));
                }
                *total = total
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("metadata byte counter overflow"))?;
                if *total > limits.max_metadata_bytes {
                    return Err(invalid("conflict metadata exceeds configured limit"));
                }
                date = Some(value);
                date_span = Some(span);
            },
            _ => {},
        }
    }
    let raw = id.ok_or_else(|| invalid("conflict markup lacks required w:id"))?;
    let id = raw
        .parse::<i32>()
        .map_err(|_source_error| invalid("conflict markup id is not a signed i32"))?;
    let author = author.ok_or_else(|| invalid("conflict markup lacks required w:author"))?;
    Ok((
        Metadata::new(Id::new(id)?, author, date)?,
        id_span,
        author_span,
        date_span,
    ))
}

fn classify(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> (Option<Kind>, Option<(Kind, bool)>) {
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == W14) {
        return (None, None);
    }
    match element.local_name().as_ref() {
        b"conflictIns" => (Some(Kind::Insert), None),
        b"conflictDel" => (Some(Kind::Delete), None),
        b"customXmlConflictInsRangeStart" => (None, Some((Kind::Insert, true))),
        b"customXmlConflictInsRangeEnd" => (None, Some((Kind::Insert, false))),
        b"customXmlConflictDelRangeStart" => (None, Some((Kind::Delete, true))),
        b"customXmlConflictDelRangeEnd" => (None, Some((Kind::Delete, false))),
        _ => (None, None),
    }
}

fn scope(parent: Option<&Frame>) -> Result<Scope> {
    let parent = parent.ok_or_else(|| invalid("conflict markup has no WordprocessingML parent"))?;
    if let Some(scope) = parent.scope {
        return (scope == Scope::Inline)
            .then_some(scope)
            .ok_or_else(|| invalid("property conflict markers must be leaf-only"));
    }
    if parent.ns == Ns::Word
        && matches!(
            parent.local.as_slice(),
            b"bdo"
                | b"body"
                | b"customXml"
                | b"del"
                | b"dir"
                | b"docPartBody"
                | b"endnote"
                | b"fldSimple"
                | b"footnote"
                | b"ftr"
                | b"hdr"
                | b"hyperlink"
                | b"ins"
                | b"moveFrom"
                | b"moveTo"
                | b"p"
                | b"rt"
                | b"rubyBase"
                | b"sdtContent"
                | b"tbl"
                | b"tc"
                | b"tr"
                | b"txbxContent"
        )
    {
        return Ok(Scope::Inline);
    }
    if parent.ns == Ns::Word
        && (parent.local.as_slice() == b"trPr"
            || (parent.local.as_slice() == b"rPr"
                && (parent.parent_is_ppr || parent.paragraph_mark_rpr)))
    {
        return Ok(Scope::Property);
    }
    if parent.ns == Ns::Math
        && matches!(
            parent.local.as_slice(),
            b"deg"
                | b"den"
                | b"e"
                | b"fName"
                | b"lim"
                | b"num"
                | b"oMath"
                | b"oMathPara"
                | b"sub"
                | b"sup"
        )
    {
        return Ok(Scope::Inline);
    }
    if parent.ns == Ns::Math && parent.local.as_slice() == b"ctrlPr" {
        return Ok(Scope::Property);
    }
    if parent.ns == Ns::W14 && matches!(parent.local.as_slice(), b"conflictIns" | b"conflictDel") {
        return Ok(Scope::Inline);
    }
    Err(invalid("conflict markup has an unsupported parent"))
}
fn range_parent(parent: Option<&Frame>) -> Result<()> {
    let parent = parent.ok_or_else(|| invalid("conflict range has no legal parent"))?;
    if parent.scope == Some(Scope::Inline) {
        return Ok(());
    }
    if parent.ns == Ns::Word
        && matches!(
            parent.local.as_slice(),
            b"bdo"
                | b"body"
                | b"customXml"
                | b"del"
                | b"dir"
                | b"docPartBody"
                | b"endnote"
                | b"fldSimple"
                | b"footnote"
                | b"ftr"
                | b"hdr"
                | b"hyperlink"
                | b"ins"
                | b"moveFrom"
                | b"moveTo"
                | b"p"
                | b"rt"
                | b"rubyBase"
                | b"sdtContent"
                | b"tbl"
                | b"tc"
                | b"tr"
                | b"txbxContent"
        )
    {
        return Ok(());
    }
    if parent.ns == Ns::Word && parent.local.as_slice() == b"sdt" && parent.ruby_sdt {
        return Ok(());
    }
    if parent.ns == Ns::Math
        && matches!(
            parent.local.as_slice(),
            b"deg"
                | b"den"
                | b"e"
                | b"fName"
                | b"lim"
                | b"num"
                | b"oMath"
                | b"oMathPara"
                | b"sub"
                | b"sup"
        )
    {
        return Ok(());
    }
    if parent.ns == Ns::W14 && matches!(parent.local.as_slice(), b"conflictIns" | b"conflictDel") {
        return Ok(());
    }
    Err(invalid("conflict range has unsupported parent"))
}
const MC: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
fn require_ignorable(
    parent: Option<&Frame>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    if parent.is_some_and(|frame| frame.ignorable_w14)
        || declares_w14_ignorable(element, decoder, resolver)?
    {
        Ok(())
    } else {
        Err(invalid(
            "active W14 conflict markup requires an in-scope mc:Ignorable prefix bound to W14",
        ))
    }
}
fn declares_w14_ignorable(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MC)
            || attribute.key.local_name().as_ref() != b"Ignorable"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|e| Error::Xml(e.to_string()))?;
        for token in value.split_ascii_whitespace() {
            if token.is_empty() || token.contains(':') {
                return Err(invalid("mc:Ignorable contains an invalid prefix token"));
            }
            // An Ignorable token denotes a namespace *prefix*, not an
            // unqualified element name. Resolve a synthetic qualified name
            // so the resolver consults the current prefix binding rather
            // than the default namespace.
            let qualified = format!("{token}:_");
            let (namespace, _) = resolver.resolve_element(QName(qualified.as_bytes()));
            if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == W14) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}
fn namespace_declarations(
    element: &BytesStart<'_>,
    max: usize,
) -> Result<Vec<(Vec<u8>, Option<Vec<u8>>)>> {
    let mut declarations = Vec::new();
    declarations
        .try_reserve_exact(max.min(16))
        .map_err(alloc("conflict MCE namespace declarations"))?;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some(&[][..])
        } else {
            key.strip_prefix(b"xmlns:")
        };
        if let Some(prefix) = prefix {
            if declarations.len() >= max {
                return Err(invalid(
                    "conflict MCE prepass exceeds namespace declaration limit",
                ));
            }
            declarations
                .try_reserve(1)
                .map_err(alloc("conflict MCE namespace declarations"))?;
            declarations.push((
                prefix.to_vec(),
                (!attribute.value.is_empty()).then(|| attribute.value.as_ref().to_vec()),
            ));
        }
    }
    Ok(declarations)
}
fn push_namespace_declarations(
    bindings: &mut HashMap<Vec<u8>, Vec<Option<Vec<u8>>>>,
    declarations: &[(Vec<u8>, Option<Vec<u8>>)],
    mut count: usize,
    max: usize,
) -> Result<usize> {
    for (prefix, value) in declarations {
        let stack = bindings.entry(prefix.clone()).or_default();
        let before = stack.last().is_some_and(Option::is_some);
        stack.push(value.clone());
        let after = stack.last().is_some_and(Option::is_some);
        if !before && after {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("conflict MCE namespace counter overflow"))?;
            if count > max {
                return Err(invalid(
                    "conflict MCE prepass exceeds namespace declaration limit",
                ));
            }
        } else if before && !after {
            count = count
                .checked_sub(1)
                .ok_or_else(|| invalid("conflict MCE namespace counter underflow"))?;
        }
    }
    Ok(count)
}
fn pop_namespace_declarations(
    bindings: &mut HashMap<Vec<u8>, Vec<Option<Vec<u8>>>>,
    declarations: Vec<(Vec<u8>, Option<Vec<u8>>)>,
    mut count: usize,
) -> Result<usize> {
    for (prefix, _) in declarations.into_iter().rev() {
        let (before, after, empty) = {
            let stack = bindings
                .get_mut(&prefix)
                .ok_or_else(|| invalid("conflict MCE namespace stack underflow"))?;
            let before = stack.last().is_some_and(Option::is_some);
            stack
                .pop()
                .ok_or_else(|| invalid("conflict MCE namespace stack underflow"))?;
            (
                before,
                stack.last().is_some_and(Option::is_some),
                stack.is_empty(),
            )
        };
        if !before && after {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("conflict MCE namespace counter overflow"))?;
        } else if before && !after {
            count = count
                .checked_sub(1)
                .ok_or_else(|| invalid("conflict MCE namespace counter underflow"))?;
        }
        if empty {
            bindings.remove(&prefix);
        }
    }
    Ok(count)
}
fn is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == W || *value == W_STRICT)
}
fn ns(namespace: &ResolveResult<'_>) -> Ns {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == W || *value == W_STRICT => Ns::Word,
        ResolveResult::Bound(Namespace(value)) if *value == W14 => Ns::W14,
        ResolveResult::Bound(Namespace(value))
            if *value == b"http://schemas.openxmlformats.org/officeDocument/2006/math"
                || *value == b"http://purl.oclc.org/ooxml/officeDocument/math" =>
        {
            Ns::Math
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => Ns::Other,
    }
}
fn pos(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| invalid("XML offset does not fit usize"))
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn alloc(resource: &'static str) -> impl FnOnce(std::collections::TryReserveError) -> Error {
    move |source| Error::Allocation { resource, source }
}

fn find_attr(source: &[u8], start: usize, end: usize, wanted: &[u8]) -> Result<AttributeSpan> {
    let tag = source
        .get(start..end)
        .ok_or_else(|| invalid("invalid XML attribute span"))?;
    let mut i = 0;
    while i < tag.len() && tag[i] != b' ' && tag[i] != b'\t' && tag[i] != b'\n' && tag[i] != b'\r' {
        i += 1;
    }
    while i < tag.len() {
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i >= tag.len() || tag[i] == b'/' || tag[i] == b'>' {
            break;
        }
        let a = i;
        while i < tag.len() && !tag[i].is_ascii_whitespace() && tag[i] != b'=' {
            i += 1;
        }
        let name = &tag[a..i];
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if tag.get(i) != Some(&b'=') {
            return Err(invalid("malformed XML attribute"));
        }
        i += 1;
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        let quote = *tag
            .get(i)
            .ok_or_else(|| invalid("unterminated XML attribute"))?;
        if quote != b'\'' && quote != b'\"' {
            return Err(invalid("XML attribute is not quoted"));
        }
        i += 1;
        let value = i;
        while i < tag.len() && tag[i] != quote {
            i += 1;
        }
        if i == tag.len() {
            return Err(invalid("unterminated XML attribute value"));
        }
        let value_end = i;
        i += 1;
        if name == wanted {
            return Ok(AttributeSpan {
                attribute: Span::new(start + a, start + i)?,
                value: Span::new(start + value, start + value_end)?,
            });
        }
    }
    Err(invalid(
        "resolved conflict attribute has no lexical source span",
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ignores_inactive_mce_orphan_and_selects_w14_choice() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w14:conflictIns w:id="1" w:author="a">x</w14:conflictIns></w:p></mc:Choice><mc:Fallback><w14:customXmlConflictDelRangeEnd w:id="9"/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#;
        let parsed = parse(xml, Limits::default()).unwrap();
        assert_eq!(parsed.conflicts.len(), 1);
        assert_eq!(parsed.conflicts[0].metadata.id.get(), 1);
        assert!(parsed.ranges.is_empty());
    }

    #[test]
    fn accepts_direct_paragraph_inline_conflict() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w14:conflictDel w:id="3" w:author="a">x</w14:conflictDel></w:p></w:body></w:document>"#;
        assert_eq!(
            parse(xml, Limits::default()).unwrap().conflicts[0].scope,
            Scope::Inline
        );
    }

    #[test]
    fn accepts_exact_configured_source_limit() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w14:conflictIns w:id="7" w:author="a"/></w:p></w:body></w:document>"#;
        let mut limits = Limits::default();
        limits.max_source_bytes = xml.len();
        assert_eq!(parse(xml, limits).unwrap().conflicts.len(), 1);
    }

    #[test]
    fn mce_markers_fit_at_exact_source_limit_for_multiple_conflicts() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w14:conflictIns w:id="1" w:author="a"/><w14:conflictDel w:id="2" w:author="b"/></w:p></w:body></w:document>"#;
        let mut limits = Limits::default();
        limits.max_source_bytes = xml.len();
        assert_eq!(parse(xml, limits).unwrap().conflicts.len(), 2);
    }

    #[test]
    fn pairs_ct_track_change_start_with_ct_markup_end() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w14:customXmlConflictInsRangeStart w:id="8" w:author="a" w:date="2020-01-01T00:00:00Z"/>x<w14:customXmlConflictInsRangeEnd w:id="8"/></w:p></w:body></w:document>"#;
        let parsed = parse(xml, Limits::default()).unwrap();
        assert_eq!(parsed.ranges.len(), 1);
        assert_eq!(parsed.ranges[0].metadata.author, "a");
        assert!(parsed.ranges[0].end_id_span.is_some());
    }

    #[test]
    fn rejects_conflict_under_run() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:r><w14:conflictIns w:id="1" w:author="a"/></w:r></w:p></w:body></w:document>"#;
        assert!(parse(xml, Limits::default()).is_err());
    }

    fn document(content: &str) -> String {
        format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>{content}</w:body></w:document>"#,
        )
    }

    #[test]
    fn accepts_only_schema_run_and_leaf_conflict_parents() {
        let legal_inline = [
            "w:bdo",
            "w:body",
            "w:customXml",
            "w:del",
            "w:dir",
            "w:docPartBody",
            "w:endnote",
            "w:fldSimple",
            "w:footnote",
            "w:ftr",
            "w:hdr",
            "w:hyperlink",
            "w:ins",
            "w:moveFrom",
            "w:moveTo",
            "w:p",
            "w:rt",
            "w:rubyBase",
            "w:sdtContent",
            "w:tbl",
            "w:tc",
            "w:tr",
            "w:txbxContent",
            "m:deg",
            "m:den",
            "m:e",
            "m:fName",
            "m:lim",
            "m:num",
            "m:oMath",
            "m:oMathPara",
            "m:sub",
            "m:sup",
            "w14:conflictIns",
        ];
        for parent in legal_inline {
            let opening = if parent.starts_with("w14:") {
                format!("<{parent} w:id=\"99\" w:author=\"a\">")
            } else {
                format!("<{parent}>")
            };
            let xml = document(&format!(
                "{opening}<w14:conflictIns w:id=\"1\" w:author=\"a\"/></{parent}>"
            ));
            assert!(
                parse(xml.as_bytes(), Limits::default()).is_ok(),
                "legal parent {parent}"
            );
        }
        for parent in ["w:trPr", "m:ctrlPr"] {
            let xml = document(&format!(
                "<{parent}><w14:conflictDel w:id=\"2\" w:author=\"a\"/></{parent}>"
            ));
            assert!(
                parse(xml.as_bytes(), Limits::default()).is_ok(),
                "legal leaf parent {parent}"
            );
        }
        assert!(
            parse(
                document(
                    "<w:pPr><w:rPr><w14:conflictIns w:id=\"3\" w:author=\"a\"/></w:rPr></w:pPr>"
                )
                .as_bytes(),
                Limits::default()
            )
            .is_ok()
        );
    }

    #[test]
    fn rejects_non_schema_conflict_and_range_parents() {
        for parent in [
            "w:r",
            "w:proofErr",
            "w:bookmarkStart",
            "w:permStart",
            "w:tblPr",
            "w:tcPr",
            "w:numPr",
            "m:acc",
        ] {
            let xml = document(&format!(
                "<{parent}><w14:conflictIns w:id=\"1\" w:author=\"a\"/></{parent}>"
            ));
            assert!(
                parse(xml.as_bytes(), Limits::default()).is_err(),
                "illegal conflict parent {parent}"
            );
        }
        for parent in ["w:r", "w:tblPr", "w:tcPr", "m:ctrlPr", "m:acc"] {
            let xml = document(&format!(
                "<{parent}><w14:customXmlConflictInsRangeStart w:id=\"4\" w:author=\"a\"/><w14:customXmlConflictInsRangeEnd w:id=\"4\"/></{parent}>"
            ));
            assert!(
                parse(xml.as_bytes(), Limits::default()).is_err(),
                "illegal range parent {parent}"
            );
        }
        let xml = document(
            "<w:pPr><w:rPr><w:foo><w14:conflictIns w:id=\"5\" w:author=\"a\"/></w:foo></w:rPr></w:pPr>",
        );
        assert!(parse(xml.as_bytes(), Limits::default()).is_err());
    }

    #[test]
    fn accepts_normative_range_parents_and_refuses_renamed_markers() {
        for parent in ["w:sdt", "w:body", "m:oMath", "w14:conflictDel"] {
            let opening = if parent.starts_with("w14:") {
                format!("<{parent} w:id=\"99\" w:author=\"a\">")
            } else {
                format!("<{parent}>")
            };
            let xml = document(&format!(
                "{opening}<w14:customXmlConflictDelRangeStart w:id=\"8\" w:author=\"a\"/><w14:customXmlConflictDelRangeEnd w:id=\"8\"/></{parent}>"
            ));
            let expected = parent != "w:sdt";
            assert_eq!(
                parse(xml.as_bytes(), Limits::default()).is_ok(),
                expected,
                "range parent {parent}"
            );
        }
        let xml = document(
            "<w:p><w14:customXmlConflictInsertionRangeStart w:id=\"9\" w:author=\"a\"/></w:p>",
        );
        assert!(parse(xml.as_bytes(), Limits::default()).is_ok());
        assert!(
            parse(xml.as_bytes(), Limits::default())
                .unwrap()
                .ranges
                .is_empty()
        );
    }

    #[test]
    fn mce_requires_effective_w14_ignorable_namespace_and_honors_aliases() {
        let missing = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w14:conflictIns w:id="1" w:author="a"/></w:p></w:body></w:document>"#;
        assert!(parse(missing, Limits::default()).is_err());
        let alias = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"><w:body><w:p><x:conflictIns w:id="1" w:author="a"/></w:p></w:body></w:document>"#;
        assert!(parse(alias, Limits::default()).is_ok());
        let shadowed = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p xmlns:w14="urn:not-w14" xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml"><x:conflictIns w:id="1" w:author="a"/></w:p></w:body></w:document>"#;
        assert!(parse(shadowed, Limits::default()).is_ok());
    }

    #[test]
    fn mce_prepass_observes_event_limits() {
        let xml = document("<w:p><w14:conflictIns w:id=\"1\" w:author=\"a\"/></w:p>");
        let mut limits = Limits::default();
        limits.max_events = 3;
        assert!(parse(xml.as_bytes(), limits).is_err());
    }

    #[test]
    fn accepts_only_paragraph_mark_rpr_change_original_context() {
        let previous_paragraph_mark = document(
            "<w:pPr><w:rPr><w:rPrChange><w:rPr><w14:conflictIns w:id=\"1\" w:author=\"a\"/></w:rPr></w:rPrChange></w:rPr></w:pPr>",
        );
        assert!(parse(previous_paragraph_mark.as_bytes(), Limits::default()).is_ok());
        let ordinary_run_property = document(
            "<w:p><w:r><w:rPr><w:rPrChange><w:rPr><w14:conflictIns w:id=\"1\" w:author=\"a\"/></w:rPr></w:rPrChange></w:rPr></w:r></w:p>",
        );
        assert!(parse(ordinary_run_property.as_bytes(), Limits::default()).is_err());
        let ordinary_rpr_original = document(
            "<w:pPr><w:rPrChange><w:rPr><w14:conflictIns w:id=\"1\" w:author=\"a\"/></w:rPr></w:rPrChange></w:pPr>",
        );
        assert!(parse(ordinary_rpr_original.as_bytes(), Limits::default()).is_err());
    }

    #[test]
    fn text_segment_limit_is_exact() {
        let xml = document(
            "<w:p><w14:conflictIns w:id=\"1\" w:author=\"a\">a<![CDATA[b]]></w14:conflictIns></w:p>",
        );
        let mut exact = Limits::default();
        exact.max_text_segments = 2;
        assert!(parse(xml.as_bytes(), exact).is_ok());
        exact.max_text_segments = 1;
        assert!(parse(xml.as_bytes(), exact).is_err());
    }

    #[test]
    fn mce_namespace_limit_counts_effective_prefixes_not_shadow_occurrences() {
        let mut content = String::from("<w:p>");
        for _ in 0..8 {
            content.push_str("<w:customXml xmlns:x=\"urn:shadow\">");
        }
        content.push_str("<w14:conflictIns w:id=\"1\" w:author=\"a\"/>");
        for _ in 0..8 {
            content.push_str("</w:customXml>");
        }
        content.push_str("</w:p>");
        let mut limits = Limits::default();
        limits.max_depth = 16;
        limits.max_attributes = 5;
        assert!(parse(document(&content).as_bytes(), limits).is_ok());
    }

    #[test]
    fn mce_namespace_limit_rejects_many_distinct_effective_prefixes() {
        let mut xml = String::from(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14""#,
        );
        for index in 0..6 {
            xml.push_str(&format!(r#" xmlns:p{index}="urn:prefix:{index}""#));
        }
        xml.push_str(
            r#"><w:body><w:p><w14:conflictIns w:id="1" w:author="a"/></w:p></w:body></w:document>"#,
        );
        let mut limits = Limits::default();
        limits.max_depth = 4;
        limits.max_attributes = 2;
        assert!(parse(xml.as_bytes(), limits).is_err());
    }

    #[test]
    fn metadata_limit_counts_only_retained_author_and_date() {
        let xml = document(
            "<w:p><w14:conflictIns w:id=\"1\" w:author=\"a\" w:ignored=\"unretained\"/></w:p>",
        );
        let mut limits = Limits::default();
        limits.max_metadata_bytes = 1;
        assert!(parse(xml.as_bytes(), limits).is_ok());
        limits.max_metadata_bytes = 0;
        assert!(parse(xml.as_bytes(), limits).is_err());
    }

    #[test]
    fn lexical_attribute_limits_apply_to_non_conflict_ancestors() {
        let xml = document(
            "<w:p xmlns:x=\"urn:test\" x:large=\"this-ancestor-value-is-deliberately-too-large\"><w14:conflictIns w:id=\"1\" w:author=\"a\"/></w:p>",
        );
        let mut limits = Limits::default();
        limits.max_attribute_bytes = 32;
        assert!(parse(xml.as_bytes(), limits).is_err());
        limits.max_attribute_bytes = 96;
        limits.max_attributes = 5;
        assert!(parse(xml.as_bytes(), limits).is_ok());
    }

    #[test]
    fn mce_ignores_inactive_text_and_keeps_selected_text() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w14:conflictIns w:id="1" w:author="a">selected</w14:conflictIns></w:p></mc:Choice><mc:Fallback><w:p><w14:conflictIns w:id="2" w:author="a">inactive</w14:conflictIns></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#;
        let parsed = parse(xml, Limits::default()).unwrap();
        assert_eq!(parsed.conflicts.len(), 1);
        assert_eq!(parsed.conflicts[0].metadata.id.get(), 1);
        let span = parsed.conflicts[0].text[0];
        assert_eq!(&xml[span.start()..span.end()], b"selected");
    }

    #[test]
    fn mce_process_content_does_not_become_a_semantic_parent() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:process" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 x" mc:ProcessContent="x:wrapper"><w:body><x:wrapper><w:p><w14:conflictIns w:id="1" w:author="a">kept</w14:conflictIns></w:p></x:wrapper></w:body></w:document>"#;
        let parsed = parse(xml, Limits::default()).unwrap();
        assert_eq!(parsed.conflicts.len(), 1);
        assert_eq!(parsed.conflicts[0].scope, Scope::Inline);
        let span = parsed.conflicts[0].text[0];
        assert_eq!(&xml[span.start()..span.end()], b"kept");
    }

    #[test]
    fn mce_direct_process_content_wrapper_is_transparent_to_scope() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:process" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 x" mc:ProcessContent="x:wrapper"><w:body><w:p><x:wrapper><w14:conflictIns w:id="1" w:author="a"/></x:wrapper></w:p></w:body></w:document>"#;
        assert_eq!(
            parse(xml, Limits::default()).unwrap().conflicts[0].scope,
            Scope::Inline
        );
    }

    #[test]
    fn mce_direct_selected_choice_is_transparent_to_scope() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><mc:AlternateContent><mc:Choice Requires="w14"><w14:conflictIns w:id="1" w:author="a"/></mc:Choice><mc:Fallback><w14:conflictIns w:id="2" w:author="a"/></mc:Fallback></mc:AlternateContent></w:p></w:body></w:document>"#;
        let parsed = parse(xml, Limits::default()).unwrap();
        assert_eq!(parsed.conflicts.len(), 1);
        assert_eq!(parsed.conflicts[0].scope, Scope::Inline);
    }

    #[test]
    fn nested_repeated_process_content_is_scoped_without_growth_per_frame() {
        let mut xml = String::from(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:process" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 x" mc:ProcessContent="x:wrapper"><w:body><w:p>"#,
        );
        for _ in 0..32 {
            xml.push_str("<x:wrapper>");
        }
        xml.push_str("<w14:conflictIns w:id=\"1\" w:author=\"a\"/>");
        for _ in 0..32 {
            xml.push_str("</x:wrapper>");
        }
        xml.push_str("</w:p></w:body></w:document>");
        let mut limits = Limits::default();
        limits.max_depth = 40;
        assert_eq!(parse(xml.as_bytes(), limits).unwrap().conflicts.len(), 1);
    }

    #[test]
    fn many_distinct_active_process_content_targets_are_refused() {
        let names = (0..61).map(|index| format!("p{index}")).collect::<Vec<_>>();
        let ignorable = names.join(" ");
        let mut xml = String::from(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p>"#,
        );
        for level in 0..68 {
            xml.push_str("<w:customXml");
            for name in &names {
                xml.push_str(&format!(" xmlns:{name}=\"urn:target:{level}:{name}\""));
            }
            xml.push_str(&format!(
                " mc:Ignorable=\"{ignorable}\" mc:ProcessContent=\""
            ));
            for (index, name) in names.iter().enumerate() {
                if index != 0 {
                    xml.push(' ');
                }
                xml.push_str(name);
                xml.push_str(":wrapper");
            }
            xml.push_str("\">");
        }
        xml.push_str("<w14:conflictIns w:id=\"1\" w:author=\"a\"/>");
        for _ in 0..68 {
            xml.push_str("</w:customXml>");
        }
        xml.push_str("</w:p></w:body></w:document>");
        let mut limits = Limits::default();
        limits.max_attributes = 64;
        limits.max_depth = 80;
        assert!(parse(xml.as_bytes(), limits).is_err());
    }

    #[test]
    fn nested_inline_conflicts_charge_every_retained_text_reference() {
        let xml = document(
            "<w:p><w14:conflictIns w:id=\"1\" w:author=\"a\"><w14:conflictDel w:id=\"2\" w:author=\"a\">x</w14:conflictDel></w14:conflictIns></w:p>",
        );
        let mut limits = Limits::default();
        limits.max_text_segments = 2;
        limits.max_text_bytes = 2;
        let parsed = parse(xml.as_bytes(), limits).unwrap();
        assert_eq!(parsed.conflicts[0].text.len(), 1);
        assert_eq!(parsed.conflicts[1].text.len(), 1);
        limits.max_text_segments = 1;
        assert!(parse(xml.as_bytes(), limits).is_err());
    }

    #[test]
    fn expanded_range_markers_are_childless_and_ct_markup_end_is_exact() {
        let valid = document(
            "<w:p><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\"></w14:customXmlConflictInsRangeStart><w14:customXmlConflictInsRangeEnd w:id=\"1\"></w14:customXmlConflictInsRangeEnd></w:p>",
        );
        assert_eq!(
            parse(valid.as_bytes(), Limits::default())
                .unwrap()
                .ranges
                .len(),
            1
        );
        let content = document(
            "<w:p><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\">x</w14:customXmlConflictInsRangeStart></w:p>",
        );
        assert!(parse(content.as_bytes(), Limits::default()).is_err());
        let end_attribute = document(
            "<w:p><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\"/><w14:customXmlConflictInsRangeEnd w:id=\"1\" w:extra=\"x\"/></w:p>",
        );
        assert!(parse(end_attribute.as_bytes(), Limits::default()).is_err());
    }

    #[test]
    fn direct_sdt_range_parent_requires_ruby_sdt_ancestry() {
        for parent in ["w:rt", "w:rubyBase"] {
            let ruby = document(&format!(
                "<{parent}><w:sdt><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\"/><w14:customXmlConflictInsRangeEnd w:id=\"1\"/></w:sdt></{parent}>"
            ));
            assert_eq!(
                parse(ruby.as_bytes(), Limits::default())
                    .unwrap()
                    .ranges
                    .len(),
                1,
                "Ruby SDT parent {parent}"
            );
        }
        let direct_ruby = document(
            "<w:ruby><w:sdt><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\"/><w14:customXmlConflictInsRangeEnd w:id=\"1\"/></w:sdt></w:ruby>",
        );
        assert!(parse(direct_ruby.as_bytes(), Limits::default()).is_err());
        let ordinary = document(
            "<w:sdt><w14:customXmlConflictInsRangeStart w:id=\"1\" w:author=\"a\"/><w14:customXmlConflictInsRangeEnd w:id=\"1\"/></w:sdt>",
        );
        assert!(parse(ordinary.as_bytes(), Limits::default()).is_err());
    }
}
