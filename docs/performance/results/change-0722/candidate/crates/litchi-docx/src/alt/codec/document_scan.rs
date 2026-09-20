#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the writer scan keeps parser state beside its event loop"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]

use crate::namespace::is_wordprocessing_namespace;
use quick_xml::encoding::Decoder;
use std::collections::{BTreeMap, BTreeSet};

use super::{
    Chunk, Error, Event, NamespaceResolver, PendingChunk, ResolveResult, Result, active,
    insert_chunk, is_word_namespace, next_depth, parse_child, relationship, validate_xml,
};

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
        decoder: Decoder,
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

/// Parse the two writer-owned structural views from one lexical event walk.
pub(crate) fn scan_with_block_ranges(
    xml: &[u8],
) -> Result<(BTreeMap<u32, Chunk>, Vec<(usize, u32, u32)>)> {
    validate_xml(xml)?;
    let targets = [b"p".as_slice(), b"tbl".as_slice(), b"altChunk".as_slice()];
    let mut range_scanner = WordElementRangeScanner::new(&targets);
    let mut ranges = Vec::new();
    let mut range_error = None;
    let mut reader = quick_xml::reader::NsReader::from_reader(xml);
    let mut alt_scanner = AltScanState::new();

    loop {
        // Keep the established alternative-format offset refusal ahead of the
        // range walk. A range-side usize refusal is remembered so a later
        // alternative-format error retains its existing precedence.
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
        let eof = alt_scanner.observe(event_start, decoder, &resolver, &namespace, &event)?;

        if range_error.is_none()
            && let Some(event_start) = event_start_usize
        {
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

        if eof {
            break;
        }
    }

    // Release parser scratch before either MCE call retains selected offsets.
    drop(reader);
    drop(range_scanner);

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

const MAX_DOCUMENT_SEMANTIC_VALUES: usize = 1_000_000;

fn reserve_document_value<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    if values.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

const MAX_SCAN_DEPTH: usize = 128;
const MAX_SCAN_NODES: usize = 1_000_000;

struct WordElementRangeScanner<'a, 'b> {
    targets: &'a [&'b [u8]],
    fragment_prefix: Option<Option<Vec<u8>>>,
    capture: Option<(usize, usize, usize)>,
    nodes: usize,
    total_depth: usize,
}

#[derive(Clone, Copy)]
enum WordElementScanEvent {
    Start(usize),
    NestedStart,
    Empty(usize),
    End,
    Eof,
    Other,
}

impl<'a, 'b> WordElementRangeScanner<'a, 'b> {
    fn new(targets: &'a [&'b [u8]]) -> Self {
        Self {
            targets,
            fragment_prefix: None,
            capture: None,
            nodes: 0,
            total_depth: 0,
        }
    }

    fn classify(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
    ) -> Result<WordElementScanEvent> {
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Word XML element counter overflow".to_string())
            })?;
            if self.nodes > MAX_SCAN_NODES {
                return Err(Error::InvalidFormat(format!(
                    "Word XML exceeds {MAX_SCAN_NODES} elements"
                )));
            }
        }
        if matches!(event, Event::Start(_)) {
            self.total_depth = self
                .total_depth
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("Word XML nesting is too deep".to_string()))?;
            if self.total_depth > MAX_SCAN_DEPTH {
                return Err(Error::InvalidFormat(format!(
                    "Word XML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                )));
            }
        }
        if matches!(event, Event::End(_)) {
            self.total_depth = self
                .total_depth
                .checked_sub(1)
                .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
        }

        if self.fragment_prefix.is_none()
            && let Event::Start(element) = event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            self.fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        Ok(match event {
            Event::Start(_) if self.capture.is_some() => WordElementScanEvent::NestedStart,
            Event::Start(element)
                if is_fragment_word_namespace(namespace, &self.fragment_prefix) =>
            {
                self.targets
                    .iter()
                    .position(|target| element.local_name().as_ref() == *target)
                    .map_or(WordElementScanEvent::Other, WordElementScanEvent::Start)
            },
            Event::Empty(element)
                if self.capture.is_none()
                    && is_fragment_word_namespace(namespace, &self.fragment_prefix) =>
            {
                self.targets
                    .iter()
                    .position(|target| element.local_name().as_ref() == *target)
                    .map_or(WordElementScanEvent::Other, WordElementScanEvent::Empty)
            },
            Event::End(_) if self.capture.is_some() => WordElementScanEvent::End,
            Event::Eof => WordElementScanEvent::Eof,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => WordElementScanEvent::Other,
        })
    }

    fn finish(
        &mut self,
        event_start: usize,
        event_end: usize,
        scan_event: WordElementScanEvent,
        emit: &mut impl FnMut(usize, u32, u32) -> Result<()>,
    ) -> Result<()> {
        match scan_event {
            WordElementScanEvent::Start(target) => {
                self.capture = Some((target, event_start, 1));
            },
            WordElementScanEvent::NestedStart => {
                let Some((_, _, depth)) = self.capture.as_mut() else {
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
            WordElementScanEvent::Empty(target) => {
                emit_word_element_range(target, event_start, event_end, emit)?;
            },
            WordElementScanEvent::End => {
                let Some((_, _, depth)) = self.capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _)) = self.capture.take() else {
                        return Err(Error::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    emit_word_element_range(target, start, event_end, emit)?;
                }
            },
            WordElementScanEvent::Eof if self.capture.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated Word element".to_string(),
                ));
            },
            WordElementScanEvent::Eof | WordElementScanEvent::Other => {},
        }
        Ok(())
    }
}

fn is_fragment_word_namespace(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
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

fn emit_word_element_range(
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
