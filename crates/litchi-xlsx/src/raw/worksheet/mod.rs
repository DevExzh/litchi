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

pub(crate) use edit::{EventSpan, FactsBuilder, SourceFacts};
pub(crate) use model::merge_successor;
pub(crate) use validation::{
    optional_bool, optional_u32, parse_a1, parse_one_based_row, required_u32,
};

use crate::cell::{Store, Text};
use crate::error::{Result, invalid};
use crate::layout::Defaults;
use litchi_ooxml_common::mce::{self, process_ooxml};
use quick_xml::events::{BytesStart, Event};
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

/// How the markup-compatibility preprocessor would treat an admitted source.
///
/// The authoritative fallback preprocesses the source with [`process_ooxml`]
/// and parses the *processed* bytes. The shared traversal parses the *source*
/// bytes, so it is admissible only where the two event streams agree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SourceAdmission {
    /// The MCE namespace does not occur, so the preprocessor borrows its input
    /// and the parser sees exactly these bytes on both paths.
    Borrowed,
    /// The MCE namespace occurs, so the preprocessor rewrites. Equivalence is
    /// not lexical, and [`MceRewriteEquivalence`] must prove it event by event.
    Rewritten,
}

/// Admit only UTF-8 worksheets without extension values to the shared reader.
///
/// The original source bytes remain owned by `SourcePayload`; this predicate
/// only decides whether a temporary borrowed traversal is safe, and which
/// evidence the traversal still has to collect.
///
/// `dyDescent` stays refused because the fallback captures those values in a
/// separate x14ac pass and hands them to the parser, which the shared reader
/// does not do. `AlternateContent` stays refused because the preprocessor
/// selects one branch and drops the rest.
pub(crate) fn source_stream_admission(content: &[u8]) -> Option<SourceAdmission> {
    if content.len() > MAX_SHARED_SOURCE_BYTES
        || !within_mce_limits(content)
        || std::str::from_utf8(content).is_err()
        || contains(content, b"AlternateContent")
        || contains(content, b"dyDescent")
    {
        return None;
    }
    Some(if contains(content, mce::NAMESPACE.as_bytes()) {
        SourceAdmission::Rewritten
    } else {
        SourceAdmission::Borrowed
    })
}

/// Return whether the shared traversal may be attempted at all.
#[cfg(test)]
pub(crate) fn source_stream_eligible(content: &[u8]) -> bool {
    source_stream_admission(content).is_some()
}

fn within_mce_limits(content: &[u8]) -> bool {
    let limits = mce::Limits::default();
    content.len() <= limits.max_input_bytes && content.len() <= limits.max_output_bytes
}

fn contains(content: &[u8], marker: &[u8]) -> bool {
    memchr::memmem::find(content, marker).is_some()
}

/// Cap the namespace declarations a rewritten start tag may carry.
///
/// The preprocessor re-declares every in-scope binding on every start tag it
/// writes and admits up to 4,096 of them, but `quick_xml` refuses more than 256
/// declarations on one element. A source spread thinly enough over its
/// ancestors parses, while its rewrite would not, so the shared traversal must
/// not admit a source past this bound.
pub(crate) const MAX_REWRITTEN_DECLARATIONS: usize = 256;

/// Grow one source byte to the longest escape the preprocessor can emit.
const MAX_ESCAPE_GROWTH: usize = 6;

/// Prove, event by event, that preprocessing this source would give the
/// worksheet parser the same events the source itself gives it.
///
/// [`process_ooxml`] re-tokenizes and re-emits any part that mentions the MCE
/// namespace, so admitting such a part to the shared traversal replaces the
/// processed event stream with the source event stream. The rewrite differs
/// from its input in four ways, none of which the worksheet parser can observe:
///
/// * it re-declares every in-scope namespace on every emitted start tag. The
///   parser resolves names through the reader and never reads an `xmlns`
///   attribute, and re-declaring a binding already in scope resolves alike.
/// * it expands `<a/>` into `<a></a>`. [`Parser::transition`] answers `Empty`
///   with the same `start` and `finish` pair that `Start` and `End` run, and
///   `finish(Context::Worksheet)` is `Ok(())`, so an empty root agrees too.
/// * it drops character data, CDATA, comments and references outside the root.
///   The parser has no text target and no leaf context there and ignores all
///   four.
/// * it copies text, CDATA, comments and references inside the root verbatim,
///   and normalizes and re-escapes attribute values with exactly the
///   normalization `unqualified_attribute_value` applies when the parser reads
///   one.
///
/// What remains are the refusals the rewrite adds. Each is checked below, and a
/// failed check leaves the traversal through the provisional failure that
/// repeats the authoritative passes, so a rejected proof costs one fallback and
/// can never change a result.
struct MceRewriteEquivalence {
    max_output_bytes: usize,
    root_started: bool,
    declarations: usize,
    declaration_bytes: usize,
    emitted_bytes: usize,
}

impl MceRewriteEquivalence {
    fn new(content: &[u8]) -> Self {
        Self {
            max_output_bytes: mce::Limits::default().max_output_bytes,
            root_started: false,
            declarations: 0,
            declaration_bytes: 0,
            // Every source byte is copied into the rewrite, escaped.
            emitted_bytes: content.len().saturating_mul(MAX_ESCAPE_GROWTH),
        }
    }

    fn observe(&mut self, event: &Event<'_>) -> bool {
        match event {
            // A declaration after the root has opened is refused as late.
            Event::Decl(_) => !self.root_started,
            // Both are refused outright, while the parser ignores them.
            Event::PI(_) | Event::DocType(_) => false,
            // Only the predefined names and character references survive; any
            // other reference is refused as a custom entity.
            Event::GeneralRef(value) => match value.resolve_char_ref() {
                Ok(Some(_)) => true,
                Ok(None) => value
                    .decode()
                    .is_ok_and(|name| matches!(&*name, "amp" | "lt" | "gt" | "apos" | "quot")),
                Err(_) => false,
            },
            Event::Start(element) | Event::Empty(element) => self.observe_start(element),
            Event::End(_) | Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::Eof => {
                true
            },
        }
    }

    fn observe_start(&mut self, element: &BytesStart<'_>) -> bool {
        self.root_started = true;
        // A prefixed element name can name the MCE vocabulary itself, name a
        // namespace the rewrite drops or unwraps, or have no binding at all,
        // which the preprocessor refuses. No worksheet element is prefixed.
        // The preprocessor also refuses a name the reader accepts but that is
        // not a qualified name, which for a colon-free name is an NCName.
        if !is_unprefixed_ncname(element.name().as_ref()) {
            return false;
        }
        for attribute in element.attributes().with_checks(true) {
            let Ok(attribute) = attribute else {
                // Duplicate or malformed attributes are refused while reading.
                return false;
            };
            // The preprocessor decodes every attribute value where the parser
            // decodes only the ones it reads, so an undecodable value is a
            // refusal the source alone would not produce.
            if attribute.value.contains(&b'&') {
                return false;
            }
            let key = attribute.key.as_ref();
            if key == b"xmlns" {
                if !self.declare(0, attribute.value.len()) {
                    return false;
                }
            } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                // An empty value or a prefix that is not an NCName is refused
                // as an invalid namespace, while the reader undeclares or
                // resolves it.
                if attribute.value.is_empty()
                    || !std::str::from_utf8(prefix)
                        .is_ok_and(litchi_ooxml_common::xml_name::is_ncname)
                    || !self.declare(prefix.len(), attribute.value.len())
                {
                    return false;
                }
            } else if key.starts_with(b"xml:") {
                // The `xml` prefix is bound by definition.
            } else if !is_unprefixed_ncname(key) {
                // Any other prefixed attribute may be an MCE directive, may
                // belong to a namespace the rewrite drops, or may be unbound;
                // an unprefixed one still has to be a qualified name.
                return false;
            }
        }
        // Every in-scope declaration is re-emitted on this tag.
        self.emitted_bytes = self.emitted_bytes.saturating_add(self.declaration_bytes);
        self.emitted_bytes <= self.max_output_bytes
    }

    /// Account for one namespace declaration, conservatively treating every
    /// declaration seen so far as still in scope.
    fn declare(&mut self, prefix_len: usize, value_len: usize) -> bool {
        self.declarations = self.declarations.saturating_add(1);
        if self.declarations > MAX_REWRITTEN_DECLARATIONS {
            return false;
        }
        // ` xmlns:<prefix>="<value>"`, with the value escaped.
        self.declaration_bytes = self
            .declaration_bytes
            .saturating_add(10)
            .saturating_add(prefix_len)
            .saturating_add(value_len.saturating_mul(MAX_ESCAPE_GROWTH));
        true
    }
}

/// Return whether this name is a prefix-free qualified name.
fn is_unprefixed_ncname(name: &[u8]) -> bool {
    std::str::from_utf8(name).is_ok_and(litchi_ooxml_common::xml_name::is_ncname)
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
    admission: SourceAdmission,
    strings: F,
    observer: O,
) -> SourceParseAttempt
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
    O: for<'event> FnMut(&ResolveResult<'event>, &Event<'event>, EventSpan) -> bool,
{
    if !shared_event_bound_within_cap(content) {
        return SourceParseAttempt::ProvisionalFailed;
    }
    codec::parse_source_with_observer(content, admission, strings, observer)
}

/// Finish a source-backed raw parse while retaining the historical x14ac
/// retry that follows a plain worksheet parser failure.
pub(crate) fn complete_source_parse(content: &[u8], parsed: Result<Store>) -> Result<Store> {
    if parsed.is_ok() {
        // Admission already proved that the shared source carries no extension
        // value; a successful completed parse needs no redundant x14ac scan.
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
