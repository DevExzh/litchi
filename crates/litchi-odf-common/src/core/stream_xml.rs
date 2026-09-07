//! Bounded, borrowing XML token streaming for source-preserving ODF paths.
//!
//! The ordinary `quick_xml::Reader` API is deliberately small and fast, but a
//! caller that supplies an untrusted `BufRead` cannot use its event buffer as a
//! resource boundary: the reader appends every byte of one token before it
//! returns the event.  This module puts a finite window in front of that
//! reader.  The window is reset for every event, so a token that crosses any
//! number of source chunks remains streamable while a token beyond the limit
//! receives at most one bounded lookahead byte and then fails before further
//! parser-buffer growth.
//!
//! The module is kept private by `core`; the `core::private` re-export is the
//! intentionally unstable seam used by ODF family crates.  Events borrow one
//! reusable buffer and must not escape the visitor callback.

use super::binding_tracker::{BindingTracker, BindingTrackerError};
use crate::validation::valid_xml_reference;
use litchi_core::{Error, Resource, ResourceLimit, Result};
use quick_xml::{
    Decoder,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, PrefixDeclaration, ResolveResult},
    reader::Reader,
};
use std::{
    collections::{HashMap, HashSet},
    fmt,
    io::{self, BufRead, Read},
    mem::size_of,
    ops::Range,
    str,
    sync::Arc,
};

const UTF8_BOM: &[u8] = b"\xef\xbb\xbf";
const UTF8_BOM_LEN: usize = 3;

// These are deliberately the same order of magnitude as the finite XML
// ceilings already used by the common generated-XML and validation paths.
// Callers may narrow them with `XmlStreamLimits::new`, but an execution
// context must never be asked to admit an unbounded parser plan.
const MAX_BYTES_CEILING: u64 = 256 * 1024 * 1024;
const MAX_DEPTH_CEILING: usize = 4_096;
const MAX_EVENTS_CEILING: usize = 4_000_000;
const MAX_ATTRIBUTES_CEILING: usize = 1_000_000;
const MAX_TOKEN_BYTES_CEILING: usize = 64 * 1024 * 1024;

const DEFAULT_MAX_BYTES: u64 = 256 * 1024 * 1024;
const DEFAULT_MAX_DEPTH: usize = 256;
const DEFAULT_MAX_EVENTS: usize = 4_000_000;
const DEFAULT_MAX_ATTRIBUTES: usize = 1_000_000;
const DEFAULT_MAX_TOKEN_BYTES: usize = 64 * 1024;

const MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT: u64 = 256;
const XMLNS_NAMESPACE_URI: &[u8] = b"http://www.w3.org/2000/xmlns/";
// BindingTracker's private binding record currently consists of three machine
// words and one u32 level (32 bytes on the supported 64-bit targets). Keep the
// envelope independent of that private layout by charging a conservative
// 32-byte record and a geometric Vec-capacity factor below.
const BINDING_RECORD_BYTES: u64 = 32;
const MACHINE_WORD_BYTES: u64 = size_of::<usize>() as u64;
const GEOMETRIC_VEC_FACTOR: u64 = 2;
// One expanded-name duplicate check charges its borrowed key, the projection
// record, the distinct raw-namespace cache, quick-xml's raw-name ranges/hash
// prefilter, and hash-table control bytes. The estimate is intentionally
// rounded up before applying the Vec and hash-table geometric factor below.
const EXPANDED_ATTRIBUTE_SCRATCH_BYTES: u64 = 256;
const FIXED_SCRATCH_BYTES: u64 = 64 * 1024;
const MIN_ATTRIBUTE_SYNTAX_BYTES: u64 = 5;

const TOKEN_RESOURCE: &str = "bounded XML token buffer";
const INPUT_RESOURCE: &str = "bounded XML input";
const EVENT_RESOURCE: &str = "bounded XML events";
const ATTRIBUTE_RESOURCE: &str = "bounded XML attributes";
const ATTRIBUTE_SET_RESOURCE: &str = "bounded XML expanded attribute set";
const NAMESPACE_RESOURCE: &str = "bounded XML normalized namespace buffer";
const DEPTH_RESOURCE: &str = "bounded XML depth";

/// Explicit finite limits for one streaming XML document.
///
/// Local ceiling failures identify `max_bytes` as `Resource::InputBytes`,
/// `max_events` as `Resource::Work`, `max_attributes` as `Resource::Objects`,
/// and `max_depth`/`max_token_bytes` as `Resource::Depth`/`Resource::Memory`.
/// These local checks do not debit hierarchical execution budgets.
/// The helper itself does not own a [`litchi_core::ExecutionContext`]; callers can use
/// [`Self::memory_upper_bound`] to reserve the parser envelope in their
/// context before starting a scan.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct XmlStreamLimits {
    max_bytes: u64,
    max_depth: usize,
    max_events: usize,
    max_attributes: usize,
    max_token_bytes: usize,
}

impl XmlStreamLimits {
    /// Construct a checked finite XML stream profile.
    ///
    /// The immutable ceilings are common safety boundaries.  A family may
    /// select any smaller values for a particular source or publication path.
    pub fn new(
        max_bytes: u64,
        max_depth: usize,
        max_events: usize,
        max_attributes: usize,
        max_token_bytes: usize,
    ) -> Result<Self> {
        if max_bytes == 0 || max_bytes > MAX_BYTES_CEILING {
            return Err(invalid_limits("max_bytes", max_bytes, MAX_BYTES_CEILING));
        }
        if max_depth == 0 || max_depth > MAX_DEPTH_CEILING {
            return Err(invalid_limits(
                "max_depth",
                max_depth as u64,
                MAX_DEPTH_CEILING as u64,
            ));
        }
        if max_events == 0 || max_events > MAX_EVENTS_CEILING {
            return Err(invalid_limits(
                "max_events",
                max_events as u64,
                MAX_EVENTS_CEILING as u64,
            ));
        }
        if max_attributes == 0 || max_attributes > MAX_ATTRIBUTES_CEILING {
            return Err(invalid_limits(
                "max_attributes",
                max_attributes as u64,
                MAX_ATTRIBUTES_CEILING as u64,
            ));
        }
        if max_token_bytes == 0 || max_token_bytes > MAX_TOKEN_BYTES_CEILING {
            return Err(invalid_limits(
                "max_token_bytes",
                max_token_bytes as u64,
                MAX_TOKEN_BYTES_CEILING as u64,
            ));
        }
        Ok(Self {
            max_bytes,
            max_depth,
            max_events,
            max_attributes,
            max_token_bytes,
        })
    }

    /// Return the total input-byte ceiling.
    #[must_use]
    pub const fn max_bytes(self) -> u64 {
        self.max_bytes
    }

    /// Return the element-depth ceiling.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Return the event-count ceiling, including the final EOF event.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Return the aggregate attribute-count ceiling.
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }

    /// Return the per-event token-byte ceiling.
    #[must_use]
    pub const fn max_token_bytes(self) -> usize {
        self.max_token_bytes
    }

    /// Return a conservative parser-memory envelope for execution admission.
    ///
    /// The estimate includes the exact event-buffer reservation (the public
    /// token limit plus one lookahead byte), geometric capacity for
    /// quick-xml's open-name stack, geometric capacity for the namespace
    /// binding bytes and records, open-depth indexes, the per-token expanded
    /// attribute-name projections and duplicate set, a bounded retained-binding
    /// arena for normalized namespace URI references, and fixed parser scratch. The
    /// attribute state is capped by both `max_attributes` and the number of
    /// minimum-sized attributes that fit in one token. It intentionally does
    /// not include memory allocated by a visitor. Every arithmetic step is
    /// checked and an overflow is reported as a typed invalid-limit error.
    pub fn memory_upper_bound(self) -> Result<u64> {
        let depth = self.max_depth as u64;
        // Keep one byte of lookahead available so quick-xml can distinguish an
        // exact-limit text token from a token that continues at the next byte.
        // The scanner checks the returned span against the public limit before
        // invoking the visitor.
        let token = (self.max_token_bytes as u64)
            .checked_add(1)
            .ok_or_else(invalid_memory_bound)?;
        let attributes = self.max_attributes as u64;
        let max_attributes_per_token = attributes.min(token / MIN_ATTRIBUTE_SYNTAX_BYTES);

        let open_name_bytes = checked_mul(
            checked_mul(
                depth.checked_add(1).ok_or_else(invalid_memory_bound)?,
                token,
                "open-name bytes",
            )?,
            GEOMETRIC_VEC_FACTOR,
            "open-name Vec capacity",
        )?;
        let namespace_bytes = checked_mul(
            checked_mul(depth, token, "namespace bytes")?,
            GEOMETRIC_VEC_FACTOR,
            "namespace Vec capacity",
        )?;
        // A failed or unknown binding carries one borrowed-name copy while it
        // is converted into the common error; charge that transient alongside
        // the retained namespace stack.
        let transient_namespace_bytes = token;
        // Namespace URI references are normalized only for expanded-name
        // duplicate comparison. Distinct inherited bindings can all be used
        // by one event, so charge the retained binding-byte bound plus the
        // current scope, with geometric Vec capacity.
        let normalized_namespace_bytes = checked_mul(
            checked_mul(
                depth.checked_add(1).ok_or_else(invalid_memory_bound)?,
                token,
                "normalized namespace bytes",
            )?,
            GEOMETRIC_VEC_FACTOR,
            "normalized namespace Vec capacity",
        )?;
        let namespace_records = checked_mul(
            checked_mul(
                depth
                    .checked_mul(MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT)
                    .and_then(|count| count.checked_add(2))
                    .ok_or_else(invalid_memory_bound)?,
                BINDING_RECORD_BYTES,
                "namespace binding record bytes",
            )?,
            GEOMETRIC_VEC_FACTOR,
            "namespace binding Vec capacity",
        )?;
        let depth_indexes = checked_mul(
            checked_mul(
                depth.checked_add(1).ok_or_else(invalid_memory_bound)?,
                MACHINE_WORD_BYTES,
                "open-depth indexes",
            )?,
            GEOMETRIC_VEC_FACTOR,
            "open-depth Vec capacity",
        )?;
        let attribute_state = checked_mul(
            checked_mul(
                max_attributes_per_token,
                EXPANDED_ATTRIBUTE_SCRATCH_BYTES,
                "expanded attribute state",
            )?,
            GEOMETRIC_VEC_FACTOR,
            "expanded attribute state capacity",
        )?;

        let mut total = FIXED_SCRATCH_BYTES;
        for amount in [
            token,
            open_name_bytes,
            namespace_bytes,
            transient_namespace_bytes,
            namespace_records,
            depth_indexes,
            normalized_namespace_bytes,
            attribute_state,
        ] {
            total = total.checked_add(amount).ok_or_else(invalid_memory_bound)?;
        }
        Ok(total)
    }
}

impl Default for XmlStreamLimits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_BYTES,
            max_depth: DEFAULT_MAX_DEPTH,
            max_events: DEFAULT_MAX_EVENTS,
            max_attributes: DEFAULT_MAX_ATTRIBUTES,
            max_token_bytes: DEFAULT_MAX_TOKEN_BYTES,
        }
    }
}

/// One borrowed parser event and its exact source span.
///
/// The event borrows the scanner's reusable token buffer.  It is valid only
/// for the duration of the `scan_xml` visitor callback.
#[derive(Debug)]
pub struct XmlStreamEvent<'a> {
    event: Event<'a>,
    start: u64,
    end: u64,
    decoder: Decoder,
}

impl<'a> XmlStreamEvent<'a> {
    /// Borrow the quick-xml event.
    #[must_use]
    pub fn event(&self) -> &Event<'a> {
        &self.event
    }

    /// Return the absolute byte offset at which this event begins.
    #[must_use]
    pub const fn start(&self) -> u64 {
        self.start
    }

    /// Return the absolute byte offset immediately after this event.
    #[must_use]
    pub const fn end(&self) -> u64 {
        self.end
    }

    /// Return the event's half-open absolute source span.
    #[must_use]
    pub fn span(&self) -> Range<u64> {
        self.start..self.end
    }

    /// Return quick-xml's active decoder for attribute value projection.
    #[must_use]
    pub const fn decoder(&self) -> Decoder {
        self.decoder
    }
}

/// Accounting returned after a bounded XML scan.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct XmlStreamReport {
    bytes: u64,
    events: usize,
    attributes: usize,
    max_depth: usize,
}

impl XmlStreamReport {
    /// Total logical input bytes consumed, including a UTF-8 BOM if present.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.bytes
    }

    /// Number of parser events delivered, including EOF.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Total attributes parsed across all start and empty elements and the XML
    /// declaration.
    #[must_use]
    pub const fn attributes(self) -> usize {
        self.attributes
    }

    /// Greatest element nesting depth observed.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
}

/// Scan one complete XML document under explicit finite limits.
///
/// The callback receives a borrowed event and the namespace tracker while the
/// event's lexical scope is active.  A `Start` scope remains active for its
/// children.  An `Empty` or `End` scope is popped immediately after its
/// callback returns.  The callback must inspect or copy any event data before
/// returning; the next event reuses the same token buffer.
pub fn scan_xml<R, F>(reader: R, limits: XmlStreamLimits, mut visit: F) -> Result<XmlStreamReport>
where
    R: BufRead,
    F: for<'event> FnMut(&'event XmlStreamEvent<'event>, &BindingTracker) -> Result<()>,
{
    // Reserve the public maximum plus one lookahead byte once. The guarded
    // BufRead never exposes more than this window to quick-xml, so the parser
    // cannot grow this token buffer past the admitted memory envelope. The
    // extra byte lets a text token end exactly at the public ceiling when the
    // next source byte is `<` or EOF; the returned span is checked below.
    let token_window = limits
        .max_token_bytes
        .checked_add(1)
        .ok_or_else(invalid_memory_bound)?;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(token_window)
        .map_err(|source| Error::Allocation {
            resource: TOKEN_RESOURCE,
            source,
        })?;

    let guarded = GuardedBufRead::new(reader, limits.max_bytes, limits.max_token_bytes)
        .map_err(|error| map_window_io(error, 0))?;
    let mut parser = Reader::from_reader(guarded);
    parser.config_mut().trim_text(false);
    parser.config_mut().check_end_names = true;
    parser.config_mut().check_comments = true;

    let mut tracker = BindingTracker::new().map_err(|error| tracker_error(error, 0))?;
    let mut depth = 0usize;
    let mut max_depth = 0usize;
    let mut events = 0usize;
    let mut attributes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut saw_any_event = false;
    let mut pending_pop = false;

    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }

        if events >= limits.max_events {
            return Err(limit_error(
                Resource::Work,
                events.saturating_add(1),
                limits.max_events,
                EVENT_RESOURCE,
            ));
        }

        buffer.clear();
        parser.get_mut().begin_token();
        let start = parser.get_ref().position();
        let event = parser
            .read_event_into(&mut buffer)
            .map_err(|error| map_reader_error(error, start))?;
        let end = parser.get_ref().position();
        events = events.checked_add(1).ok_or_else(|| {
            limit_error_u64(
                Resource::Work,
                u64::MAX,
                limits.max_events as u64,
                EVENT_RESOURCE,
            )
        })?;

        let token_bytes = end
            .checked_sub(start)
            .ok_or_else(|| invalid_xml(start, "XML source position moved backwards"))?;
        if token_bytes > limits.max_token_bytes as u64 {
            return Err(limit_error_u64(
                Resource::Memory,
                token_bytes,
                limits.max_token_bytes as u64,
                TOKEN_RESOURCE,
            ));
        }

        validate_event_bytes(&event, start)?;
        let stream_event = XmlStreamEvent {
            decoder: parser.decoder(),
            event,
            start,
            end,
        };

        match stream_event.event() {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid_xml(start, "content appears after the XML root"));
                }
                if depth >= limits.max_depth {
                    return Err(limit_error(
                        Resource::Depth,
                        depth.saturating_add(1),
                        limits.max_depth,
                        DEPTH_RESOURCE,
                    ));
                }
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_xml(start, "XML document has more than one root"));
                    }
                    root_seen = true;
                }
                tracker
                    .push(element)
                    .map_err(|error| tracker_error(error, start))?;
                let namespace_arena_limit = normalized_namespace_limit(depth, token_window)?;
                validate_start_element(
                    &tracker,
                    element,
                    &mut attributes,
                    limits.max_attributes,
                    namespace_arena_limit,
                    start,
                )?;
                depth += 1;
                max_depth = max_depth.max(depth);
                visit(&stream_event, &tracker)?;
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid_xml(start, "content appears after the XML root"));
                }
                if depth >= limits.max_depth {
                    return Err(limit_error(
                        Resource::Depth,
                        depth.saturating_add(1),
                        limits.max_depth,
                        DEPTH_RESOURCE,
                    ));
                }
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_xml(start, "XML document has more than one root"));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                tracker
                    .push(element)
                    .map_err(|error| tracker_error(error, start))?;
                let namespace_arena_limit = normalized_namespace_limit(depth, token_window)?;
                validate_start_element(
                    &tracker,
                    element,
                    &mut attributes,
                    limits.max_attributes,
                    namespace_arena_limit,
                    start,
                )?;
                max_depth = max_depth.max(depth + 1);
                visit(&stream_event, &tracker)?;
                tracker.pop();
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid_xml(start, "XML document has an unmatched end tag"));
                }
                validate_end_element(&tracker, element, start)?;
                visit(&stream_event, &tracker)?;
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
                pending_pop = true;
            },
            Event::Decl(declaration) => {
                if saw_any_event || declaration_seen || root_seen || depth != 0 {
                    return Err(invalid_xml(
                        start,
                        "XML declaration is not the first document event",
                    ));
                }
                declaration_seen = true;
                declaration.xml_version().map_err(|error| {
                    invalid_xml(start, format!("invalid XML declaration: {error}"))
                })?;
                if let Some(encoding) = declaration.encoding() {
                    let encoding = encoding.map_err(|error| {
                        invalid_xml(start, format!("invalid XML declaration encoding: {error}"))
                    })?;
                    if !encoding.eq_ignore_ascii_case(b"utf-8") {
                        return Err(invalid_xml(
                            start,
                            "XML declaration names an encoding other than UTF-8",
                        ));
                    }
                }
                if let Some(standalone) = declaration.standalone() {
                    let standalone = standalone.map_err(|error| {
                        invalid_xml(
                            start,
                            format!("invalid XML declaration standalone: {error}"),
                        )
                    })?;
                    if !matches!(standalone.as_ref(), b"yes" | b"no") {
                        return Err(invalid_xml(
                            start,
                            "XML declaration standalone must be 'yes' or 'no'",
                        ));
                    }
                }
                validate_decl_attributes(
                    &tracker,
                    declaration,
                    &mut attributes,
                    limits.max_attributes,
                    start,
                )?;
                visit(&stream_event, &tracker)?;
            },
            Event::DocType(_) => {
                return Err(invalid_xml(
                    start,
                    "DTD and DOCTYPE declarations are prohibited",
                ));
            },
            Event::GeneralRef(reference) => {
                if depth == 0 || root_closed {
                    return Err(invalid_xml(
                        start,
                        "entity reference appears outside the XML root",
                    ));
                }
                if !valid_xml_reference(reference) {
                    return Err(invalid_xml(
                        start,
                        "undeclared or invalid XML entity reference",
                    ));
                }
                visit(&stream_event, &tracker)?;
            },
            Event::CData(_) => {
                if depth == 0 || root_closed {
                    return Err(invalid_xml(start, "CDATA appears outside the XML root"));
                }
                visit(&stream_event, &tracker)?;
            },
            Event::Text(text) => {
                let bytes: &[u8] = text;
                if depth == 0
                    && !bytes
                        .iter()
                        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return Err(invalid_xml(
                        start,
                        "non-whitespace text appears outside the XML root",
                    ));
                }
                visit(&stream_event, &tracker)?;
            },
            Event::Comment(_) | Event::PI(_) => {
                if let Event::PI(instruction) = stream_event.event() {
                    let target = instruction.target();
                    validate_name(target, start, "processing-instruction target")?;
                    if target.eq_ignore_ascii_case(b"xml") {
                        return Err(invalid_xml(
                            start,
                            "processing-instruction target 'xml' is reserved",
                        ));
                    }
                }
                visit(&stream_event, &tracker)?;
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 {
                    return Err(invalid_xml(
                        start,
                        "XML document has no single complete root element",
                    ));
                }
                visit(&stream_event, &tracker)?;
                break;
            },
        }
        saw_any_event = true;
    }

    Ok(XmlStreamReport {
        bytes: parser.get_ref().position(),
        events,
        attributes,
        max_depth,
    })
}

/// A `BufRead` facade that exposes at most the remaining total and per-token
/// windows.  It deliberately does not retain a second source buffer: source
/// taps and spliced readers remain the owner of their own byte flow.
struct GuardedBufRead<R> {
    inner: R,
    // A three-byte fixed lookahead makes BOM detection independent of the
    // source's BufRead chunk size. Bytes held here have already been removed
    // from `inner` but are still logically unconsumed by the parser.
    prefix: [u8; UTF8_BOM_LEN],
    prefix_len: usize,
    prefix_pos: usize,
    total: u64,
    token: usize,
    max_total: u64,
    max_token: usize,
    max_token_window: usize,
}

impl<R: BufRead> GuardedBufRead<R> {
    fn new(inner: R, max_total: u64, max_token: usize) -> io::Result<Self> {
        let mut guarded = Self {
            inner,
            prefix: [0; UTF8_BOM_LEN],
            prefix_len: 0,
            prefix_pos: 0,
            total: 0,
            token: 0,
            max_total,
            max_token,
            max_token_window: max_token.checked_add(1).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "XML token window overflows usize",
                )
            })?,
        };
        // Never pull fixed-prefix bytes past the total input ceiling. A
        // partial BOM is left as ordinary pending source and will hit the
        // total window when the parser asks for the next byte.
        let prefix_limit = max_total.min(UTF8_BOM.len() as u64) as usize;
        while guarded.prefix_len < prefix_limit {
            let available = guarded.inner.fill_buf()?;
            if available.is_empty() {
                break;
            }
            let count = available
                .len()
                .min(prefix_limit.saturating_sub(guarded.prefix_len));
            guarded.prefix[guarded.prefix_len..guarded.prefix_len + count]
                .copy_from_slice(&available[..count]);
            guarded.inner.consume(count);
            guarded.prefix_len += count;
        }
        if guarded.prefix_len == UTF8_BOM.len() && guarded.prefix == UTF8_BOM {
            if max_total < UTF8_BOM.len() as u64 {
                return Err(window_total_error(UTF8_BOM.len() as u64, max_total));
            }
            guarded.prefix_pos = guarded.prefix_len;
            guarded.total = UTF8_BOM.len() as u64;
        }
        Ok(guarded)
    }

    fn begin_token(&mut self) {
        self.token = 0;
    }

    const fn position(&self) -> u64 {
        self.total
    }
}

impl<R: BufRead> Read for GuardedBufRead<R> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: BufRead> BufRead for GuardedBufRead<R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        let max_token = self.max_token;
        let max_total = self.max_total;
        if self.token > max_token {
            return Err(window_token_error(self.token.saturating_add(1), max_token));
        }
        if self.total >= max_total {
            if !self.pending().is_empty() {
                return Err(window_total_error(self.total.saturating_add(1), max_total));
            }
            let available = self.inner.fill_buf()?;
            if available.is_empty() {
                return Ok(available);
            }
            return Err(window_total_error(self.total.saturating_add(1), max_total));
        }
        // Compute all accounting scalars before borrowing the source slice;
        // the returned `BufRead` view may borrow either `prefix` or `inner`.
        let total_remaining = usize::try_from(max_total - self.total).unwrap_or(usize::MAX);
        let token_remaining = self.max_token_window.saturating_sub(self.token);
        let token_exhausted = self.token > self.max_token;
        let token_observed = self.token.saturating_add(1);
        let total_observed = self.total.saturating_add(1);
        let available = self.available_source()?;
        if available.is_empty() {
            if token_exhausted {
                return Err(window_token_error(token_observed, max_token));
            }
            return Ok(available);
        }
        if token_remaining == 0 {
            // The one-byte window has already been consumed into the parser
            // buffer. Do not expose a second lookahead byte: `BufRead::consume`
            // cannot account for a byte beyond this window, and replaying it
            // would let quick-xml append it repeatedly on malformed markup.
            return Err(window_token_error(token_observed, max_token));
        }
        let visible = available.len().min(total_remaining).min(token_remaining);
        if visible == 0 {
            if total_remaining == 0 {
                return Err(window_total_error(total_observed, max_total));
            }
            return Err(window_token_error(token_observed, max_token));
        }
        Ok(&available[..visible])
    }

    fn consume(&mut self, amount: usize) {
        // `Reader` consumes only bytes returned by `fill_buf`.  If another
        // caller violates BufRead's contract, saturating accounting avoids a
        // panic and the next fill reports the finite window failure.
        let amount = amount.min(self.max_token_window.saturating_sub(self.token));
        let amount_u64 = amount as u64;
        let pending = self.pending().len();
        let from_prefix = amount.min(pending);
        self.prefix_pos = self.prefix_pos.saturating_add(from_prefix);
        let from_inner = amount.saturating_sub(from_prefix);
        if from_inner != 0 {
            self.inner.consume(from_inner);
        }
        self.token = self.token.saturating_add(amount);
        self.total = self.total.saturating_add(amount_u64);
    }
}

impl<R> GuardedBufRead<R> {
    fn pending(&self) -> &[u8] {
        &self.prefix[self.prefix_pos..self.prefix_len]
    }
}

impl<R: BufRead> GuardedBufRead<R> {
    fn available_source(&mut self) -> io::Result<&[u8]> {
        if !self.pending().is_empty() {
            Ok(self.pending())
        } else {
            self.inner.fill_buf()
        }
    }
}

/// Namespace bytes used while checking one start tag's expanded attribute
/// names. Ordinary bindings borrow the tracker's stable namespace buffer;
/// bindings containing an entity reference use a bounded per-event arena.
#[derive(Clone, Copy)]
enum NamespaceProjection<'a> {
    Borrowed(&'a [u8]),
    Normalized { start: usize, end: usize },
}

#[derive(Clone, Copy)]
struct NamespaceRange {
    start: usize,
    end: usize,
}

#[derive(Clone, Copy, Eq, Hash, PartialEq)]
struct ExpandedAttributeName<'a> {
    namespace: &'a [u8],
    local: &'a [u8],
}

fn validate_event_bytes(event: &Event<'_>, offset: u64) -> Result<()> {
    let bytes: &[u8] = event;
    let text = str::from_utf8(bytes)
        .map_err(|error| invalid_xml(offset, format!("XML is not UTF-8: {error}")))?;
    if text.chars().any(|character| !is_xml10_character(character)) {
        return Err(invalid_xml(
            offset,
            "XML contains a character forbidden by XML 1.0",
        ));
    }
    Ok(())
}

fn validate_start_element(
    tracker: &BindingTracker,
    element: &BytesStart<'_>,
    attributes: &mut usize,
    max_attributes: usize,
    namespace_arena_limit: usize,
    offset: u64,
) -> Result<()> {
    validate_qname(element.name().as_ref(), offset, "element")?;
    if element
        .name()
        .prefix()
        .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
    {
        return Err(invalid_xml(
            offset,
            "the xmlns prefix is reserved for namespace declarations",
        ));
    }
    let resolved_element = tracker
        .resolve_element(element.name())
        .map_err(|error| tracker_error(error, offset))?;
    reject_unknown_namespace(&resolved_element.0, offset)?;

    let attribute_capacity = max_attributes.min(
        (element.attributes_raw().len() as u64 / MIN_ATTRIBUTE_SYNTAX_BYTES).min(usize::MAX as u64)
            as usize,
    );
    let mut namespace_projections = Vec::new();
    namespace_projections
        .try_reserve(attribute_capacity)
        .map_err(|source| Error::Allocation {
            resource: ATTRIBUTE_SET_RESOURCE,
            source,
        })?;
    let mut normalized_namespaces = Vec::new();
    if element.attributes_raw().contains(&b'&') {
        normalized_namespaces
            .try_reserve(element.attributes_raw().len())
            .map_err(|source| Error::Allocation {
                resource: NAMESPACE_RESOURCE,
                source,
            })?;
    }
    let mut normalized_namespace_cache = HashMap::new();

    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_xml(offset, format!("malformed XML attribute: {error}")))?;
        *attributes = checked_count(*attributes, max_attributes, ATTRIBUTE_RESOURCE)?;
        validate_qname(attribute.key.as_ref(), offset, "attribute")?;
        if let Some(PrefixDeclaration::Named(prefix)) = attribute.key.as_namespace_binding()
            && prefix.is_empty()
        {
            return Err(invalid_xml(
                offset,
                "namespace declaration has an empty prefix",
            ));
        }
        let value = attribute.value.as_ref();
        str::from_utf8(value)
            .map_err(|error| invalid_xml(offset, format!("XML attribute is not UTF-8: {error}")))?;
        if value.contains(&b'<') {
            return Err(invalid_xml(
                offset,
                "XML attribute values cannot contain an unescaped '<'",
            ));
        }
        if value
            .iter()
            .any(|byte| !matches!(*byte, 0x09 | 0x0A | 0x0D | 0x20..=0x7F) && *byte < 0x80)
        {
            return Err(invalid_xml(
                offset,
                "XML attribute contains a forbidden control",
            ));
        }
        validate_references(value, offset)?;

        let resolved = tracker
            .resolve_attribute(attribute.key)
            .map_err(|error| tracker_error(error, offset))?;
        reject_unknown_namespace(&resolved.0, offset)?;
        if namespace_projections.len() == namespace_projections.capacity() {
            namespace_projections
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: ATTRIBUTE_SET_RESOURCE,
                    source,
                })?;
        }
        namespace_projections.push(namespace_projection(
            attribute.key.as_namespace_binding(),
            resolved.0,
            &mut normalized_namespaces,
            &mut normalized_namespace_cache,
            namespace_arena_limit,
            offset,
        )?);
    }

    let mut expanded_names = HashSet::new();
    expanded_names
        .try_reserve(namespace_projections.len())
        .map_err(|source| Error::Allocation {
            resource: ATTRIBUTE_SET_RESOURCE,
            source,
        })?;
    for (index, attribute) in element.attributes().with_checks(false).enumerate() {
        let attribute = attribute
            .map_err(|error| invalid_xml(offset, format!("malformed XML attribute: {error}")))?;
        let projection = namespace_projections.get(index).ok_or_else(|| {
            invalid_xml(
                offset,
                "XML attribute projection count changed during validation",
            )
        })?;
        let local = tracker
            .resolve_attribute(attribute.key)
            .map_err(|error| tracker_error(error, offset))?
            .1
            .into_inner();
        let namespace = match projection {
            NamespaceProjection::Borrowed(namespace) => *namespace,
            NamespaceProjection::Normalized { start, end } => {
                normalized_namespaces.get(*start..*end).ok_or_else(|| {
                    invalid_xml(offset, "normalized namespace projection is out of bounds")
                })?
            },
        };
        let name = ExpandedAttributeName { namespace, local };
        if expanded_names.len() == expanded_names.capacity() {
            expanded_names
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: ATTRIBUTE_SET_RESOURCE,
                    source,
                })?;
        }
        if !expanded_names.insert(name) {
            return Err(invalid_xml(offset, "duplicate XML attribute expanded name"));
        }
    }
    Ok(())
}

fn namespace_projection<'a>(
    declaration: Option<PrefixDeclaration<'a>>,
    resolved: ResolveResult<'a>,
    normalized_namespaces: &mut Vec<u8>,
    normalized_namespace_cache: &mut HashMap<&'a [u8], NamespaceRange>,
    namespace_arena_limit: usize,
    offset: u64,
) -> Result<NamespaceProjection<'a>> {
    if declaration.is_some() {
        // `BindingTracker::resolve_attribute` intentionally leaves the
        // unprefixed `xmlns` declaration unbound because the default
        // namespace never applies to attributes. Namespace declaration
        // attributes nevertheless have the XMLNS expanded namespace.
        return Ok(NamespaceProjection::Borrowed(XMLNS_NAMESPACE_URI));
    }
    let namespace = match resolved {
        ResolveResult::Unbound => &[][..],
        ResolveResult::Bound(Namespace(namespace)) => namespace,
        ResolveResult::Unknown(_) => {
            return Err(invalid_xml(offset, "unknown namespace prefix"));
        },
    };
    if memchr::memchr(b'&', namespace).is_none() {
        return Ok(NamespaceProjection::Borrowed(namespace));
    }
    if let Some(&range) = normalized_namespace_cache.get(namespace) {
        return Ok(NamespaceProjection::Normalized {
            start: range.start,
            end: range.end,
        });
    }
    let projection = normalize_namespace_uri(
        namespace,
        normalized_namespaces,
        namespace_arena_limit,
        offset,
    )?;
    if normalized_namespace_cache.len() == normalized_namespace_cache.capacity() {
        normalized_namespace_cache
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: ATTRIBUTE_SET_RESOURCE,
                source,
            })?;
    }
    if let NamespaceProjection::Normalized { start, end } = projection {
        normalized_namespace_cache.insert(namespace, NamespaceRange { start, end });
    }
    Ok(projection)
}

fn normalize_namespace_uri<'a>(
    raw: &'a [u8],
    normalized_namespaces: &mut Vec<u8>,
    namespace_arena_limit: usize,
    offset: u64,
) -> Result<NamespaceProjection<'a>> {
    let start = normalized_namespaces.len();
    let mut remaining = raw;
    while let Some(index) = memchr::memchr(b'&', remaining) {
        append_namespace_bytes(
            normalized_namespaces,
            &remaining[..index],
            namespace_arena_limit,
        )?;
        let after_ampersand = &remaining[index + 1..];
        let end = memchr::memchr(b';', after_ampersand)
            .ok_or_else(|| invalid_xml(offset, "unterminated namespace URI entity reference"))?;
        let reference_text = str::from_utf8(&after_ampersand[..end]).map_err(|error| {
            invalid_xml(
                offset,
                format!("namespace URI entity is not UTF-8: {error}"),
            )
        })?;
        let reference = BytesRef::new(reference_text);
        if reference.is_char_ref() {
            let character = reference
                .resolve_char_ref()
                .map_err(|error| {
                    invalid_xml(offset, format!("invalid namespace URI reference: {error}"))
                })?
                .ok_or_else(|| invalid_xml(offset, "invalid namespace URI character reference"))?;
            let mut encoded = [0_u8; 4];
            let encoded = character.encode_utf8(&mut encoded);
            append_namespace_bytes(
                normalized_namespaces,
                encoded.as_bytes(),
                namespace_arena_limit,
            )?;
        } else {
            let replacement = quick_xml::escape::resolve_predefined_entity(reference_text)
                .ok_or_else(|| invalid_xml(offset, "undeclared namespace URI entity reference"))?;
            append_namespace_bytes(
                normalized_namespaces,
                replacement.as_bytes(),
                namespace_arena_limit,
            )?;
        }
        remaining = &after_ampersand[end + 1..];
    }
    append_namespace_bytes(normalized_namespaces, remaining, namespace_arena_limit)?;
    Ok(NamespaceProjection::Normalized {
        start,
        end: normalized_namespaces.len(),
    })
}

fn append_namespace_bytes(
    normalized_namespaces: &mut Vec<u8>,
    bytes: &[u8],
    namespace_arena_limit: usize,
) -> Result<()> {
    let observed = normalized_namespaces
        .len()
        .checked_add(bytes.len())
        .ok_or_else(invalid_memory_bound)?;
    if observed > namespace_arena_limit {
        return Err(limit_error(
            Resource::Memory,
            observed,
            namespace_arena_limit,
            NAMESPACE_RESOURCE,
        ));
    }
    normalized_namespaces
        .try_reserve(bytes.len())
        .map_err(|source| Error::Allocation {
            resource: NAMESPACE_RESOURCE,
            source,
        })?;
    normalized_namespaces.extend_from_slice(bytes);
    Ok(())
}

fn validate_decl_attributes(
    tracker: &BindingTracker,
    declaration: &quick_xml::events::BytesDecl<'_>,
    attributes: &mut usize,
    max_attributes: usize,
    offset: u64,
) -> Result<()> {
    let mut state = 0u8;
    let content = str::from_utf8(declaration.as_ref())
        .map_err(|error| invalid_xml(offset, format!("XML declaration is not UTF-8: {error}")))?;
    let declaration_start = BytesStart::from_content(content, 3);
    for attribute in declaration_start.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_xml(offset, format!("malformed XML declaration: {error}")))?;
        *attributes = checked_count(*attributes, max_attributes, ATTRIBUTE_RESOURCE)?;
        validate_qname(attribute.key.as_ref(), offset, "declaration attribute")?;
        if attribute.key.prefix().is_some() {
            return Err(invalid_xml(
                offset,
                "XML declaration attributes must be unprefixed",
            ));
        }
        state = match (state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return Err(invalid_xml(
                    offset,
                    "XML declaration attributes are missing, duplicated, or out of order",
                ));
            },
        };
        let resolved = tracker
            .resolve_attribute(attribute.key)
            .map_err(|error| tracker_error(error, offset))?;
        reject_unknown_namespace(&resolved.0, offset)?;
        validate_references(attribute.value.as_ref(), offset)?;
    }
    if state == 0 {
        return Err(invalid_xml(
            offset,
            "XML declaration is missing its version",
        ));
    }
    Ok(())
}

fn validate_end_element(
    tracker: &BindingTracker,
    element: &quick_xml::events::BytesEnd<'_>,
    offset: u64,
) -> Result<()> {
    validate_qname(element.name().as_ref(), offset, "end element")?;
    let resolved_element = tracker
        .resolve_element(element.name())
        .map_err(|error| tracker_error(error, offset))?;
    reject_unknown_namespace(&resolved_element.0, offset)
}

fn validate_qname(bytes: &[u8], offset: u64, kind: &str) -> Result<()> {
    let Some(colon) = bytes.iter().position(|byte| *byte == b':') else {
        return validate_name(bytes, offset, kind);
    };
    if colon == 0 || colon + 1 == bytes.len() || bytes[colon + 1..].contains(&b':') {
        return Err(invalid_xml(
            offset,
            format!("malformed {kind} namespace name"),
        ));
    }
    validate_name(&bytes[..colon], offset, kind)?;
    validate_name(&bytes[colon + 1..], offset, kind)
}

fn validate_name(bytes: &[u8], offset: u64, kind: &str) -> Result<()> {
    let text = str::from_utf8(bytes)
        .map_err(|error| invalid_xml(offset, format!("{kind} name is not UTF-8: {error}")))?;
    let mut chars = text.chars();
    let Some(first) = chars.next() else {
        return Err(invalid_xml(offset, format!("empty {kind} name")));
    };
    if !is_xml_name_start(first) || chars.any(|character| !is_xml_name_char(character)) {
        return Err(invalid_xml(offset, format!("malformed {kind} name")));
    }
    Ok(())
}

fn validate_references(bytes: &[u8], offset: u64) -> Result<()> {
    let mut remaining = bytes;
    while let Some(index) = memchr::memchr(b'&', remaining) {
        let after_ampersand = &remaining[index + 1..];
        let Some(end) = memchr::memchr(b';', after_ampersand) else {
            return Err(invalid_xml(offset, "unterminated XML entity reference"));
        };
        let reference_text = str::from_utf8(&after_ampersand[..end])
            .map_err(|error| invalid_xml(offset, format!("XML entity is not UTF-8: {error}")))?;
        let reference = BytesRef::new(reference_text);
        if !valid_xml_reference(&reference) {
            return Err(invalid_xml(
                offset,
                "undeclared or invalid XML entity reference",
            ));
        }
        remaining = &after_ampersand[end + 1..];
    }
    Ok(())
}

fn reject_unknown_namespace(result: &ResolveResult<'_>, offset: u64) -> Result<()> {
    if matches!(result, ResolveResult::Unknown(_)) {
        return Err(invalid_xml(offset, "unknown namespace prefix"));
    }
    Ok(())
}

fn is_xml10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(
            character as u32,
            0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
}

fn is_xml_name_start(character: char) -> bool {
    matches!(
        character as u32,
        0x3A | 0x5F | 0x41..=0x5A | 0x61..=0x7A | 0xC0..=0xD6 | 0xD8..=0xF6
            | 0xF8..=0x2FF
            | 0x370..=0x37D
            | 0x37F..=0x1FFF
            | 0x200C..=0x200D
            | 0x2070..=0x218F
            | 0x2C00..=0x2FEF
            | 0x3001..=0xD7FF
            | 0xF900..=0xFDCF
            | 0xFDF0..=0xFFFD
            | 0x10000..=0xEFFFF
    )
}

fn is_xml_name_char(character: char) -> bool {
    is_xml_name_start(character)
        || matches!(
            character as u32,
            0x2D | 0x2E | 0x30..=0x39 | 0xB7 | 0x300..=0x36F | 0x203F..=0x2040
        )
}

fn checked_count(current: usize, limit: usize, scope: &'static str) -> Result<usize> {
    let observed = current.saturating_add(1);
    if observed > limit {
        return Err(limit_error(Resource::Objects, observed, limit, scope));
    }
    Ok(observed)
}

fn checked_mul(left: u64, right: u64, label: &'static str) -> Result<u64> {
    left.checked_mul(right)
        .ok_or_else(|| invalid_memory_bound_with_label(label))
}

fn normalized_namespace_limit(depth: usize, token_window: usize) -> Result<usize> {
    depth
        .checked_add(1)
        .and_then(|levels| levels.checked_mul(token_window))
        .ok_or_else(invalid_memory_bound)
}

fn invalid_limits(name: &str, observed: u64, maximum: u64) -> Error {
    Error::InvalidFormat(format!(
        "XML stream limit {name}={observed} is outside the finite range 1..={maximum}"
    ))
}

fn invalid_memory_bound() -> Error {
    Error::InvalidFormat("XML stream memory envelope overflows u64".to_string())
}

fn invalid_memory_bound_with_label(label: &str) -> Error {
    Error::InvalidFormat(format!(
        "XML stream memory envelope overflows u64 at {label}"
    ))
}

fn invalid_xml(offset: u64, message: impl Into<String>) -> Error {
    Error::InvalidFormat(format!(
        "malformed XML at byte {offset}: {}",
        message.into()
    ))
}

fn limit_error(resource: Resource, observed: usize, limit: usize, scope: &'static str) -> Error {
    limit_error_u64(resource, observed as u64, limit as u64, scope)
}

fn limit_error_u64(resource: Resource, observed: u64, limit: u64, scope: &'static str) -> Error {
    Error::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from(scope),
    })
}

fn tracker_error(error: BindingTrackerError, offset: u64) -> Error {
    error.into_litchi_error_with_context(|| format!("malformed XML namespace at byte {offset}"))
}

fn map_reader_error(error: quick_xml::Error, offset: u64) -> Error {
    match error {
        quick_xml::Error::Io(source) => map_window_io_arc(source, offset),
        quick_xml::Error::Encoding(_) => invalid_xml(offset, "XML is not UTF-8"),
        other => invalid_xml(offset, format!("{other}")),
    }
}

fn map_window_io_arc(source: Arc<io::Error>, offset: u64) -> Error {
    if let Some(window) = source
        .get_ref()
        .and_then(|source| source.downcast_ref::<WindowError>())
    {
        return window.to_error();
    }
    Error::Io(io::Error::new(
        source.kind(),
        format!("XML input at byte {offset}: {source}"),
    ))
}

fn map_window_io(source: io::Error, offset: u64) -> Error {
    if let Some(window) = source
        .get_ref()
        .and_then(|source| source.downcast_ref::<WindowError>())
    {
        return window.to_error();
    }
    Error::Io(io::Error::new(
        source.kind(),
        format!("XML input at byte {offset}: {source}"),
    ))
}

#[derive(Debug)]
enum WindowError {
    Total { observed: u64, limit: u64 },
    Token { observed: usize, limit: usize },
}

impl fmt::Display for WindowError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Total { observed, limit } => {
                write!(
                    formatter,
                    "XML input byte window exceeded: observed {observed}, limit {limit}"
                )
            },
            Self::Token { observed, limit } => {
                write!(
                    formatter,
                    "XML token window exceeded: observed {observed}, limit {limit}"
                )
            },
        }
    }
}

impl std::error::Error for WindowError {}

impl WindowError {
    fn to_error(&self) -> Error {
        match self {
            Self::Total { observed, limit } => Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed: *observed,
                limit: *limit,
                scope: Arc::from(INPUT_RESOURCE),
            }),
            Self::Token { observed, limit } => Error::ResourceLimit(ResourceLimit {
                resource: Resource::Memory,
                observed: *observed as u64,
                limit: *limit as u64,
                scope: Arc::from(TOKEN_RESOURCE),
            }),
        }
    }
}

fn window_total_error(observed: u64, limit: u64) -> io::Error {
    io::Error::other(WindowError::Total { observed, limit })
}

fn window_token_error(observed: usize, limit: usize) -> io::Error {
    io::Error::other(WindowError::Token { observed, limit })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn limits() -> XmlStreamLimits {
        XmlStreamLimits::new(1 << 20, 32, 256, 256, 256).expect("test limits are finite")
    }

    fn scan(xml: &[u8], limits: XmlStreamLimits) -> Result<XmlStreamReport> {
        scan_xml(Cursor::new(xml), limits, |_event, _tracker| Ok(()))
    }

    #[test]
    fn token_window_rejects_before_parser_buffer_growth() {
        let limits = XmlStreamLimits::new(1024, 8, 64, 64, 16).expect("test limits are finite");
        let error = scan(br#"<root attr="01234567890123456789"/>"#, limits)
            .expect_err("long start tag must exceed the token window");
        assert!(matches!(
            error,
            Error::ResourceLimit(ResourceLimit {
                resource: Resource::Memory,
                ..
            })
        ));
    }

    #[test]
    fn depth_window_rejects_deep_input() {
        let mut xml = String::new();
        for _ in 0..8 {
            xml.push_str("<a>");
        }
        for _ in 0..8 {
            xml.push_str("</a>");
        }
        let limits = XmlStreamLimits::new(1024, 4, 128, 128, 32).expect("test limits are finite");
        let error =
            scan(xml.as_bytes(), limits).expect_err("deep XML must exceed the depth window");
        assert!(matches!(
            error,
            Error::ResourceLimit(ResourceLimit {
                resource: Resource::Depth,
                ..
            })
        ));
    }

    #[test]
    fn unknown_namespace_prefix_is_rejected() {
        let error = scan(br#"<p:root/>"#, limits()).expect_err("unknown prefix must be rejected");
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("unknown namespace prefix"))
        );
    }

    #[test]
    fn reserved_namespace_binding_is_rejected() {
        let error = scan(br#"<root xmlns:xmlns="urn:bad"/>"#, limits())
            .expect_err("the xmlns prefix cannot be rebound");
        assert!(matches!(error, Error::InvalidFormat(message) if message.contains("namespace")));
    }

    #[test]
    fn expanded_attribute_names_detect_namespace_aliases_and_allow_shadowing() {
        let alias = scan(
            br#"<root xmlns:a="urn:x" xmlns:b="urn:x" a:z="1" b:z="2"/>"#,
            limits(),
        )
        .expect_err("attributes with one expanded name must be rejected");
        assert!(
            matches!(alias, Error::InvalidFormat(message) if message.contains("expanded name"))
        );

        let entity_alias = scan(
            br#"<root xmlns:a="urn:&#120;" xmlns:b="urn:x" a:z="1" b:z="2"/>"#,
            limits(),
        )
        .expect_err("normalized namespace URI aliases must be rejected");
        assert!(
            matches!(entity_alias, Error::InvalidFormat(message) if message.contains("expanded name"))
        );

        scan(
            br#"<root xmlns:a="urn:x"><child xmlns:a="urn:y" a:z="1"/></root>"#,
            limits(),
        )
        .expect("a shadowed prefix with a different URI has a distinct expanded name");
    }

    #[test]
    fn inherited_entity_namespace_is_normalized_once_per_event() {
        let limits = XmlStreamLimits::new(1024, 8, 64, 64, 64).expect("test limits are finite");
        scan(
            br#"<root xmlns:a="urn:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx&#120;"><child a:a="1" a:b="2" a:c="3" a:d="4"/></root>"#,
            limits,
        )
        .expect("reusing one inherited URI must fit the bounded normalization arena");
    }

    #[test]
    fn chunk_boundaries_preserve_utf8_and_absolute_spans() {
        let xml = b"<root>  e\xC3\xA9 \xF0\x9F\x99\x82 </root>";
        let mut spans = Vec::new();
        let report = scan_xml(Chunked::new(xml, 1), limits(), |event, _tracker| {
            spans.push((
                event.start(),
                event.end(),
                matches!(event.event(), Event::Text(_)),
            ));
            Ok(())
        })
        .expect("one-byte chunks must parse");
        assert_eq!(report.bytes(), xml.len() as u64);
        assert!(spans.iter().any(|(_, _, text)| *text));
        assert_eq!(spans.last().map(|(_, end, _)| *end), Some(xml.len() as u64));
    }

    #[test]
    fn exact_token_limit_allows_text_before_markup_and_rejects_one_more_byte() {
        let exact = b"<root>1234567890123456</root>";
        let limits = XmlStreamLimits::new(1024, 8, 64, 64, 16).expect("test limits are finite");
        scan_xml(Chunked::new(exact, 1), limits, |_event, _tracker| Ok(()))
            .expect("a text token at the exact ceiling must parse");

        let over = b"<root>12345678901234567</root>";
        let error = scan_xml(Chunked::new(over, 1), limits, |_event, _tracker| Ok(()))
            .expect_err("one byte beyond the token ceiling must fail");
        assert!(matches!(
            error,
            Error::ResourceLimit(ResourceLimit {
                resource: Resource::Memory,
                ..
            })
        ));
    }

    #[test]
    fn guarded_token_boundary_does_not_replay_an_unconsumed_marker() {
        let mut guarded = GuardedBufRead::new(Cursor::new(b"abcde<<"), 1024, 4)
            .expect("the fixed token window must be representable");
        guarded.begin_token();
        let mut observed = Vec::new();
        let error = loop {
            let chunk_len = match guarded.fill_buf() {
                Ok(bytes) => {
                    assert!(!bytes.is_empty(), "the source still contains markers");
                    assert!(
                        !bytes.contains(&b'<'),
                        "the marker beyond the admitted token window must stay hidden"
                    );
                    observed.extend_from_slice(bytes);
                    bytes.len()
                },
                Err(error) => break error,
            };
            guarded.consume(chunk_len);
        };
        assert_eq!(observed, b"abcde");
        assert!(matches!(
            error
                .get_ref()
                .and_then(|source| source.downcast_ref::<WindowError>()),
            Some(WindowError::Token {
                observed: 6,
                limit: 4
            })
        ));
    }

    #[test]
    fn unterminated_token_at_marker_boundary_returns_without_replaying_input() {
        let limits = XmlStreamLimits::new(1024, 8, 64, 64, 12).expect("test limits are finite");
        // The first '<' after the quoted attribute is the single lookahead
        // byte. The second '<' is available source at the exhausted boundary;
        // the guard must fail immediately instead of returning that same byte
        // with a zero-byte consume, which would make quick-xml loop forever.
        let error = scan(br#"<root a="1234<<"#, limits)
            .expect_err("an unterminated token must hit the finite token window");
        assert!(matches!(
            error,
            Error::ResourceLimit(ResourceLimit {
                resource: Resource::Memory,
                ..
            })
        ));
    }

    #[test]
    fn split_bom_is_counted_and_split_utf8_is_validated() {
        let xml = b"\xEF\xBB\xBF<root>e\xC3\xA9\xF0\x9F\x99\x82</root>";
        let mut root_span = None;
        let report = scan_xml(Chunked::new(xml, 1), limits(), |event, _tracker| {
            if matches!(event.event(), Event::Start(_)) {
                root_span = Some(event.span());
            }
            Ok(())
        })
        .expect("a BOM and split UTF-8 code points must parse");
        assert_eq!(report.bytes(), xml.len() as u64);
        assert_eq!(root_span, Some(3..9));

        let malformed = b"<root>\xF0\x9F\x99</root>";
        let error = scan_xml(Chunked::new(malformed, 1), limits(), |_event, _tracker| {
            Ok(())
        })
        .expect_err("a truncated UTF-8 code point must fail");
        assert!(matches!(error, Error::InvalidFormat(message) if message.contains("UTF-8")));
    }

    #[test]
    fn invalid_entity_trailing_root_and_unterminated_root_are_rejected() {
        let entity = scan(br#"<root>&unknown;</root>"#, limits())
            .expect_err("undeclared entity must be rejected");
        assert!(matches!(entity, Error::InvalidFormat(message) if message.contains("entity")));

        let trailing =
            scan(br#"<root/><second/>"#, limits()).expect_err("trailing root must be rejected");
        assert!(
            matches!(trailing, Error::InvalidFormat(message) if message.contains("more than one root") || message.contains("after the XML root"))
        );

        let eof = scan(br#"<root>"#, limits()).expect_err("unterminated root must be rejected");
        assert!(
            matches!(eof, Error::InvalidFormat(message) if message.contains("XML") || message.contains("malformed"))
        );
    }

    #[test]
    fn aggregate_attribute_window_is_enforced() {
        let limits = XmlStreamLimits::new(1024, 8, 64, 1, 128).expect("test limits are finite");
        let error = scan(br#"<root a="1" b="2"/>"#, limits)
            .expect_err("attribute aggregate must be bounded");
        assert!(matches!(
            error,
            Error::ResourceLimit(ResourceLimit {
                resource: Resource::Objects,
                ..
            })
        ));
    }

    struct Chunked<'a> {
        source: &'a [u8],
        position: usize,
        chunk: usize,
    }

    impl<'a> Chunked<'a> {
        fn new(source: &'a [u8], chunk: usize) -> Self {
            Self {
                source,
                position: 0,
                chunk: chunk.max(1),
            }
        }
    }

    impl Read for Chunked<'_> {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let count = available.len().min(output.len());
            output[..count].copy_from_slice(&available[..count]);
            self.consume(count);
            Ok(count)
        }
    }

    impl BufRead for Chunked<'_> {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            let end = self
                .position
                .saturating_add(self.chunk)
                .min(self.source.len());
            Ok(&self.source[self.position..end])
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.source.len());
        }
    }
}
