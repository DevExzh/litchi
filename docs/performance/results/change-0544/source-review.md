# 0544 candidate source review

This is a read-only review of the frozen 0544 candidate. The reviewed patch is
`candidate.patch` with SHA-256
`36bb9c06c9a25136dd940550cb8cef8fc7c8aa33a8deb55a80fb6d70d3328ef8`.
The reviewed design is `design.md` with SHA-256
`18e69fd72a60f8d5f7d055864f0329fb463ba9fd990cafc56faa568bff2c0068`.
The five candidate source hashes are bound separately by
`candidate-source-hashes.json`. `git apply --check` and the prepared-source
comparison were read-only checks; this review ran no Rust build, test, or
runtime measurement.

## Static disposition

No new source-level freeze blocker was found. The candidate is conditionally
approved for the isolated baseline/candidate campaign. Runtime retention still
depends on the direct oracle, inherited differential/error-order tests, fresh
performance gates, and final quality checks. The 0543 measurements cannot be
reused for this revision.

## Lexical admission proof

`raw::worksheet::shared_event_bound_within_cap` runs before either the
provisional `NsReader` or raw parser is constructed. It starts with one event
for EOF, adds one for a possible initial text event, one for every `<` or `&`,
and one for every `>` or `;` followed by a byte that is neither `<`, `&`, nor
EOF. Each increment is checked against
`MAX_SHARED_PROVISIONAL_EVENTS`.

Under the pinned `quick-xml 0.41.0` `NsReader::from_reader(&[u8])` call and its
current defaults (`allow_dangling_amp = false`, `expand_empty_elements =
false`, `trim_text_start = false`, `trim_text_end = false`, and
`check_end_names = true`), every emitted non-EOF event is charged by this
bound:

* a markup event begins at `<`;
* a general-reference event begins at `&`;
* a text event is either the initial text prefix or begins after markup (`>`)
  or a completed reference (`;`), and is charged only when a following byte
  exists outside the two event-start delimiters;
* EOF is charged once.

The bound intentionally scans lexical bytes without interpreting XML context.
Delimiters in attributes, comments, CDATA, doctypes, processing instructions,
or ordinary text can only add charges. Malformed references and markup return a
reader error before producing more events, so they cannot make the actual
prefix exceed the bound. A checked increment failure also declines admission.
The proof is one-way: an emitted stream over the cap cannot be admitted. It
does not claim that every admitted stream is semantically valid.

The proof is tied to the reader configuration and dependency version above. A
change to that call, any of those defaults, or the quick-xml version requires
another proof review.

## Direct-reader oracle and boundary coverage

The candidate-only test constructs the same borrowed `NsReader`, sets
`check_end_names = true`, resolves each event through the namespace resolver,
counts EOF, and records the number of events before a reader error. Its
independent lexical implementation is compared with the production predicate,
and every observed prefix is asserted to be no larger than the independent
bound.

The edge matrix covers empty and whitespace input, BOM, declarations and
processing instructions, comments/CDATA/doctypes containing delimiters,
quoted attribute delimiters, ordinary text, named/decimal/hex/chained
references, malformed and dangling references, nested/self-closing markup,
and malformed tails. It constructs valid worksheet streams whose direct event
count is exactly 131,072 and 131,073, asserting admission at the exact limit
and refusal immediately above it. The deliberate delimiter-heavy comment case
shows a lexical false positive: its direct stream is below the cap while its
bound declines it, and the source-backed no-op still succeeds with unchanged
source bytes and an empty patch. The 96x96 and 128x128 numeric-grid controls
also remain admitted.

These tests are a meaningful oracle for the bound, while the malformed cases
only establish the reader-error prefix and do not treat reader acceptance as
worksheet validity. The ordinary source-backed tests remain responsible for
the full SpreadsheetML grammar and publication behavior.

## Validation, fallback, and error order

The source-backed loader chooses this path only after source ownership and the
existing aggregate worksheet-byte check. On eligible bytes,
`parse_source_with_observer` performs the lexical preflight, then gives each
borrowed event to `Validator::observe` before the shared ordinary parser
transition. Validator rejection, reader failure, parser-transition failure,
or the runtime event cap returns a provisional failure immediately. The
caller drops the provisional validator and runs the established full
`worksheet_xml` validation followed by authoritative `raw::worksheet::parse`.
Thus a speculative parser error cannot outrank a later validation error, and a
false-positive preflight has the same safe authoritative fallback.

Only an observer-accepted EOF can produce `SourceParseAttempt::Complete`.
EOF observation requires a root and zero validator depth; parser transition
also requires the historical closed worksheet root. The parser then consumes
its owned state in `finish_parse`, including formula resolution, string lookup,
materialization, and store construction. `validator.finish()` is run before
the result is forwarded, and currently repeats only the already-enforced root
and depth invariant. If a future `finish` check becomes substantive, it must
move before parser materialization to preserve this boundary.

For `Complete(Err(_))`, `complete_source_parse` retains the historical x14ac
retry: a failed marker-free parse still performs the extension capture needed
to preserve its diagnostic precedence. A capture error wins as before; a
successful capture leaves the original raw error. Successful eligible parses
skip the redundant scan only because eligibility excludes the x14ac markers.
The ordinary `parse` facade still captures extension state before parsing and
uses the stateful completion helper, so its existing x14ac order is unchanged.

## Ownership, resource, and lint fences

Eligibility retains the original source bytes and requires valid UTF-8, at
most 8 MiB, the default MCE input/output limits, and absence of the MCE,
`AlternateContent`, x14ac, and `dyDescent` markers. Ineligible sources keep
the established MCE preprocessing, validator, x14ac capture, and raw-parser
sequence. No reader event or borrowed payload escapes the callback. The
higher-ranked callback lifetime and the owned `Complete(Result<Store>)`
boundary prevent reader-backed data from entering the snapshot.

The event cap, source limit, existing parser depth/record/scalar/formula
limits, and checked collection reservations bound the speculative work. The
validator owns grammar-local element names and can remain live while owned
materialization runs. These are finite input/state fences, not a process-wide
memory budget or an OOM guarantee; namespace/attribute decoding and existing
post-validation collections retain their ordinary allocation behavior.

The only lint-specific source addition is a scoped
`#[expect(clippy::large_enum_variant, reason = "short-lived owned parser result avoids an extra heap allocation")]`
on the private `SourceParseAttempt`. It documents the deliberate choice to
carry the owned post-EOF result without an extra hot-path allocation. The
expectation is local to that private enum and does not alter the public API,
dependencies, unsafe surface, or ODF paths.

The candidate remains an unmeasured OOXML/XLSX follow-up. Fresh campaign
evidence must cover both the below-cap shared path and the 160/164/256
cap-boundary fixtures before any retention decision; ODF work remains deferred
under the stated OLE2/OOXML priority.
