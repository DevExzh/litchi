# 0544 one-scan preflight review

This document reviews an unmeasured follow-up to the rejected 0544 candidate.
It proposes replacing the two delimiter scans in
`shared_event_bound_within_cap` with one scan that charges two events for every
`<` or `&`:

```text
B(content) = 1                         // terminal Event::Eof
           + initial_nonmarker_text    // zero or one Event::Text
           + 2 * count('<' or '&')
```

The proposal is design-only. It has no patch, build, test, allocation, or
runtime evidence and must be evaluated in a fresh campaign. The 0544 source
and measurements remain unchanged and are not evidence for this follow-up.

## Soundness under the pinned reader

For `quick-xml 0.41.0` with the current `NsReader::from_reader(&[u8])` call,
`read_event()`, and `resolver().resolve_event(event)`, the intended proof is
that every emitted non-EOF event is charged by the formula:

* `Start`, `End`, `Empty`, `Comment`, `CData`, `DocType`, `Decl`, and `PI`
  events originate at a `<` byte;
* `GeneralRef` originates at an `&` byte;
* after any markup or reference event, the reader can emit at most one text
  event before it reaches the next `<` or `&`;
* a text prefix before the first delimiter contributes at most one initial
  text event; and
* EOF contributes one event.

The two-event charge per delimiter covers its markup/reference event and the
possible following text event. Delimiters inside an attribute, comment, CDATA,
doctype, processing instruction, or ordinary text only increase the charge or
produce no additional reader event. Consecutive delimiters do not create an
intervening text event. Namespace resolution returns a view of the existing
event and does not create another event. Therefore the admission implication
to prove is:

```text
actual emitted events > MAX_SHARED_PROVISIONAL_EVENTS
    implies B(content) > MAX_SHARED_PROVISIONAL_EVENTS
```

The proof must remain tied to these reader defaults: `allow_dangling_amp =
false`, `allow_unmatched_ends = false`, `check_comments = false`,
`check_end_names = true`, `expand_empty_elements = false`,
`trim_markup_names_in_closing_tags = true`, and both text-trimming options
false. In particular, `expand_empty_elements = true` could produce a
synthetic end event without another `<`, so it would invalidate this simple
charge. A dependency, reader call, or configuration change requires a new
proof review.

Malformed markup, malformed or dangling references, namespace failures, and
other reader errors do not weaken the safety argument: the shared attempt
returns a provisional failure and the authoritative validator/parser owns the
final result. The checked counter must also return false on arithmetic
overflow, even when the helper is called independently of the 8 MiB source
eligibility fence.

## Required edge oracle

The follow-up needs an independent direct-reader oracle, rather than a test
that reuses the production counter. The oracle should use the same reader
construction and configuration, resolve each event, count EOF, and retain the
emitted prefix when `read_event` returns an error. It should compare the
one-scan predicate with an independently calculated `B` and assert that every
observed prefix is at most `B`.

The cases should include empty and whitespace-only input, a UTF-8 BOM, XML
declarations and processing instructions, all markup event kinds, quoted
attribute delimiters, comments/CDATA/doctypes with embedded delimiters,
ordinary text, adjacent text/reference segments, valid named/decimal/hex
references, dangling and unterminated references, consecutive delimiters,
nested and self-closing markup, and malformed tails. The exact-cap fixture
must have a direct stream of exactly 131,072 events while the one-scan bound
still fits; the cap-plus-one fixture must be declined. Repeated adjacent
comments alone are insufficient for the exact-admission case because the new
bound charges two events per comment while the reader emits only one. The
fixture should interleave a legal text event after markup, such as whitespace
between comments inside a worksheet, so the two-event charge is exercised.

A false-positive case is required. A source such as delimiter-heavy comments
can keep the direct stream below the cap while the conservative bound exceeds
it. The source-backed operation must then use the existing authoritative
validation-then-parse path, preserve the original source bytes, and retain its
empty no-op patch. This demonstrates that over-counting is a performance cost,
not a semantic refusal.

## Fallback and resource fences

The one-scan helper must run before construction of the provisional reader,
validator, or raw parser. A false result returns the existing
`SourceParseAttempt::ProvisionalFailed`; the caller drops provisional state and
runs full validation followed by the authoritative raw parse on the original
bytes. The runtime event cap remains as a second defense if a source passes
the lexical preflight. The change must preserve the validation-first error
order, post-EOF materialization boundary, successful x14ac scan avoidance, and
failed-parse x14ac retry.

The source owner, aggregate worksheet budget, 8 MiB source ceiling, UTF-8 and
MCE/x14ac eligibility markers, parser depth/record/scalar/formula limits,
higher-ranked observer lifetime, and publication/version/cancellation fences
must remain unchanged. The preflight should allocate no source copy, event
queue, or diagnostic list. These are bounded-input/state properties and do
not constitute a process-wide memory budget or OOM guarantee.

## Performance hypothesis and coverage cost

One lexical pass can reduce the scanner work before the shared traversal and
can decline an over-cap source before it constructs and discards partial
validator/parser state. That may remove part of the cap-boundary cliff. The
tradeoff is a looser bound for ordinary inputs: admission requires roughly
`count('<' or '&') <= 65,535` when there is no initial text, compared with the
actual event cap of 131,072. Valid sources with many markup tokens can be sent
to the authoritative path even when their reader stream is below the cap.
The cost must be measured separately for the primary shapes and the
cap-boundary sizes; no timing conclusion follows from this design.

For the current private `worksheet_numeric_grid` generator, each cell emits
four `<` bytes (`<c>`, `<v>`, `</v>`, `</c>`), each row contributes two, and
the wrapper contributes four. Thus the static marker count is
`4 * side^2 + 2 * side + 4`, and the proposed bound is:

| Side | Marker count | Proposed bound including EOF |
| ---: | ---: | ---: |
| 96 | 37,060 | 74,121 |
| 128 | 65,796 | 131,593 |
| 160 | 102,724 | 205,449 |

These are generator-derived counts, not runtime measurements. They show that
the stated formula admits the 96 grid but declines the 128 and 160 grids at a
131,072 cap. Any alternate estimates such as roughly 55k/98k/154k must first
identify a different marker count or formula; they do not follow from
`1 + initial + 2 * count('<' or '&')` for this generator. The intended fixture
set and coverage requirement should be reconciled before a patch is frozen.

The smaller bound may still be appropriate if preserving 128-sized admission
is not required, but that is a product/performance decision for the fresh
campaign. A tighter parser-aware bound or a different cap would require its
own proof and measurement.

## Review gate

This follow-up is suitable for isolated implementation review only after the
formula, reader-default assumptions, exact-cap fixture, false-positive
fallback, source/resource fences, and intended coverage threshold are settled.
No code should be applied to the measured 0544 candidate, and ODF work remains
deferred while OLE2/OOXML optimization is active.
