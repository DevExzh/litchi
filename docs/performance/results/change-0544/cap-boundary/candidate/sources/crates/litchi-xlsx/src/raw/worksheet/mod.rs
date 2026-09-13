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

#[expect(
    clippy::large_enum_variant,
    reason = "short-lived owned parser result avoids an extra heap allocation"
)]
pub(crate) enum SourceParseAttempt {
    Complete(Result<Store>),
    ProvisionalFailed,
    ReaderFailed,
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
        && !add_shared_event_bound(&mut bound)
    {
        return false;
    }

    // Every markup or general reference event begins at one of these bytes.
    for _ in memchr::memchr2_iter(b'<', b'&', content) {
        if !add_shared_event_bound(&mut bound) {
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
            && !add_shared_event_bound(&mut bound)
        {
            return false;
        }
    }

    true
}

#[inline]
fn add_shared_event_bound(bound: &mut usize) -> bool {
    let Some(next) = bound.checked_add(1) else {
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
    if !shared_event_bound_within_cap(content) {
        return SourceParseAttempt::ProvisionalFailed;
    }
    codec::parse_source_with_observer(content, strings, observer)
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
