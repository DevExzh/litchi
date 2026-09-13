//! Layered, namespace-aware parser for sparse `SpreadsheetML` worksheet data.
//!
//! The facade keeps worksheet consumers independent from the streaming codec;
//! raw records, semantic materialization, and validation rules live in focused
//! sibling modules so the hot path remains allocation-conscious.

mod codec;
pub(crate) mod edit;
mod model;
pub(crate) mod selected;
mod semantic;
mod validation;
mod x14ac;

#[cfg(test)]
mod tests;

pub(crate) use model::merge_successor;
pub(crate) use validation::{
    optional_bool, optional_u32, parse_a1, parse_one_based_row, required_u32,
};

use crate::cell::{Store, Text};
use crate::error::{Result, invalid};
use crate::layout::Defaults;
use litchi_ooxml_common::mce::{self, process_ooxml};
use quick_xml::encoding::Decoder;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;

/// Keep speculative source parsing below the aggregate multi-sheet limit. A
/// larger worksheet takes the established validation-then-parse path.
pub(crate) const MAX_SHARED_SOURCE_BYTES: usize = 8 * 1024 * 1024;

/// Bound the parser state retained while validation is still provisional.
/// This is deliberately below the ordinary parser event limit: after this
/// count the driver returns a provisional failure, drops the parser, and the
/// caller repeats the authoritative passes.
pub(crate) const MAX_SHARED_PROVISIONAL_EVENTS: usize = 131_072;

// Probe a short sparse prefix with the two-byte search before switching to
// the pinned memchr single-byte iterator's bulk count path.  These are
// private implementation controls; the event cap above remains the only
// resource-policy value exposed to the surrounding parser.
const SPARSE_MARKER_HIT_LIMIT: usize = 16;
const MARKER_COUNT_CHUNK_BYTES: usize = 64 * 1024;

#[expect(
    clippy::large_enum_variant,
    reason = "short-lived owned parser result avoids an extra heap allocation"
)]
pub(crate) enum SourceParseAttempt {
    Complete(Result<Store>),
    ProvisionalFailed,
    ReaderFailed,
}

/// Byte positions for one event in the original, byte-identical worksheet.
/// The optional proof observer receives these positions before the ordinary
/// parser advances to its next event.
#[derive(Debug, Clone, Copy)]
pub(crate) struct SourceEventSpan {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) decoder: Decoder,
}

/// A resolved address handoff emitted after the ordinary parser accepts an
/// event. Inferred addresses consequently come from the parser's checked
/// state rather than from a second address decoder.
#[derive(Debug, Clone, Copy)]
pub(crate) enum SourceHandoff {
    CellStart {
        address: litchi_sheet::Cell,
        span: SourceEventSpan,
    },
    CellEmpty {
        address: litchi_sheet::Cell,
        span: SourceEventSpan,
    },
    CellEnd {
        address: litchi_sheet::Cell,
        span: SourceEventSpan,
    },
    RowStart {
        number: u32,
        span: SourceEventSpan,
    },
    RowEmpty {
        number: u32,
        span: SourceEventSpan,
    },
    RowEnd {
        number: u32,
        span: SourceEventSpan,
    },
}

/// Admit only byte-identical, UTF-8, no-MCE/no-x14ac worksheets to the shared
/// reader. The original source bytes remain owned by `SourcePayload`; this
/// predicate only decides whether a temporary borrowed traversal is safe.
pub(crate) fn source_stream_eligible(content: &[u8]) -> bool {
    content.len() <= MAX_SHARED_SOURCE_BYTES
        && within_mce_limits(content)
        && std::str::from_utf8(content).is_ok()
        && !contains(content, mce::NAMESPACE.as_bytes())
        && !contains(content, b"AlternateContent")
        && !contains(content, x14ac::NAMESPACE)
        && !contains(content, b"dyDescent")
}

fn within_mce_limits(content: &[u8]) -> bool {
    let limits = mce::Limits::default();
    content.len() <= limits.max_input_bytes && content.len() <= limits.max_output_bytes
}

fn contains(content: &[u8], marker: &[u8]) -> bool {
    memchr::memmem::find(content, marker).is_some()
}

/// Return whether a conservative lexical upper bound fits the provisional
/// shared-reader event cap. This runs before constructing `NsReader` or the
/// raw parser. False positives fall back to the authoritative two-pass path;
/// this predicate can never admit a source whose reader event count exceeds
/// the cap.
pub(crate) fn shared_event_bound_within_cap(content: &[u8]) -> bool {
    let mut bound = 1usize; // The reader emits one terminal `Event::Eof`.

    if let Some(&first) = content.first()
        && !matches!(first, b'<' | b'&')
        && !add_shared_event_bound(&mut bound, 1)
    {
        return false;
    }

    // Every markup or general reference event begins at one of these bytes.
    // Keep the two-byte search for the sparse prefix; after the fixed number
    // of hits, count the disjoint marker classes in bounded chunks.
    let mut remaining = content;
    for _ in 0..SPARSE_MARKER_HIT_LIMIT {
        let Some(index) = memchr::memchr2(b'<', b'&', remaining) else {
            remaining = &[];
            break;
        };
        if !add_shared_event_bound(&mut bound, 1) {
            return false;
        }
        remaining = &remaining[index + 1..];
    }
    for chunk in remaining.chunks(MARKER_COUNT_CHUNK_BYTES) {
        // `<` and `&` are disjoint, so the subtotal cannot overflow usize.
        let subtotal =
            memchr::memchr_iter(b'<', chunk).count() + memchr::memchr_iter(b'&', chunk).count();
        if !add_shared_event_bound(&mut bound, subtotal) {
            return false;
        }
    }

    // A text event can begin after markup (`>`) or a completed reference
    // (`;`) when the next source byte is neither another delimiter nor EOF.
    // Delimiters inside attributes, comments, CDATA, and ordinary text make
    // this deliberately over-count, which only causes a safe fallback.
    for index in memchr::memchr2_iter(b'>', b';', content) {
        if content
            .get(index + 1)
            .is_some_and(|&next| !matches!(next, b'<' | b'&'))
            && !add_shared_event_bound(&mut bound, 1)
        {
            return false;
        }
    }

    true
}

#[inline]
fn add_shared_event_bound(bound: &mut usize, amount: usize) -> bool {
    let Some(next) = bound.checked_add(amount) else {
        return false;
    };
    if next > MAX_SHARED_PROVISIONAL_EVENTS {
        return false;
    }
    *bound = next;
    true
}

pub(crate) fn parse_source_with_observer<'a, F, O>(
    content: &[u8],
    strings: F,
    observer: O,
) -> SourceParseAttempt
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
    O: for<'event> FnMut(&ResolveResult<'event>, &Event<'event>) -> bool,
{
    codec::parse_source_with_observer(content, strings, observer)
}

/// Parse an eligible source with the ordinary parser and an optional
/// position-aware proof observer.
///
/// The event observer remains at the historical pre-transition point and can
/// request the established provisional fallback. The handoff runs only after
/// a transition succeeds and is deliberately infallible: a proof
/// implementation may disable itself, but it cannot alter parser errors or
/// force a replay.
pub(crate) fn parse_source_with_observer_and_handoff<'a, F, C, O, H>(
    content: &[u8],
    strings: F,
    context: &mut C,
    observer: O,
    handoff: H,
) -> SourceParseAttempt
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
    C: ?Sized,
    O: for<'event> FnMut(&mut C, &ResolveResult<'event>, &Event<'event>, SourceEventSpan) -> bool,
    H: FnMut(&mut C, SourceHandoff),
{
    if !shared_event_bound_within_cap(content) {
        return SourceParseAttempt::ProvisionalFailed;
    }
    codec::parse_source_with_observer_and_handoff(content, strings, context, observer, handoff)
}

/// Finish a source-backed raw parse while retaining the historical x14ac
/// retry that follows a plain worksheet parser failure.
pub(crate) fn complete_source_parse(content: &[u8], parsed: Result<Store>) -> Result<Store> {
    if parsed.is_ok() {
        // Eligibility already proved that the shared source is marker-free;
        // a successful completed parse needs no redundant x14ac scan.
        return parsed;
    }
    let needs_extension_capture = x14ac::may_contain_descent(content);
    complete_source_parse_with_extension_state(content, needs_extension_capture, parsed)
}

fn complete_source_parse_with_extension_state(
    content: &[u8],
    needs_extension_capture: bool,
    parsed: Result<Store>,
) -> Result<Store> {
    if parsed.is_err() && !needs_extension_capture {
        // The extension scan historically ran first. Repeat it only on a
        // rejected plain worksheet so its typed error and error precedence
        // remain unchanged without charging successful no-extension reads.
        x14ac::capture(content)?;
    }
    parsed
}

pub(crate) fn parse<'a, F>(content: &[u8], strings: F) -> Result<Store>
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
{
    let needs_extension_capture = x14ac::may_contain_descent(content);
    let extensions = if needs_extension_capture {
        x14ac::capture(content)?
    } else {
        x14ac::Values::default()
    };
    let parsed = (|| {
        let processed = process_ooxml(content)?;
        let content = std::str::from_utf8(processed.as_ref())
            .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
        model::Parser::parse(content, strings, extensions)
    })();

    complete_source_parse_with_extension_state(content, needs_extension_capture, parsed)
}

pub(crate) fn parse_defaults(content: &[u8]) -> Result<Option<Defaults>> {
    let mut descent = x14ac::capture_defaults(content)?;
    let processed = process_ooxml(content)?;
    let content = std::str::from_utf8(processed.as_ref())
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    codec::parse_processed_defaults(content, descent.take())
}
