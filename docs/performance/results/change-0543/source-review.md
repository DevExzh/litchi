# 0543 candidate source review

This is a read-only review of the frozen 0543 production candidate identified
by the SHA-256 of `candidate.patch`:

`bd3252181951573b738dfb01a86e64387b9be89932c5520097670e04c606c496`

The candidate has not been applied to the live checkout and has no runtime,
allocation, or build evidence. The corrected public post-EOF raw-error test is
bound by [`differential-tests-final.patch`](differential-tests-final.patch);
the original `differential-tests.patch` is retained as a failed attempt whose
unaffected-sheet assertion used an unsupported API. The private cap tests are
included in `candidate.patch`. This review covers the five production/private-
test files in the candidate and the inherited 0542 next-candidate
requirements.

## Static disposition

No source-level freeze blocker was found. The candidate is conditionally
approved for a fresh isolated campaign only. Runtime retention remains blocked
until the paired public differential test, fresh native and allocation guards,
the retained error-order matrix, and final quality checks pass. The 0542
measurements cannot be reused as evidence for this revision.

## Error order and materialization boundary

`Snapshot::from_source_selected` selects the shared path only after source
ownership and the aggregate worksheet-byte check. For an eligible worksheet,
`parse_source_with_observer` reads the original bytes through one borrowed
`NsReader`. Each event is given to `validation::Validator::observe` before the
ordinary `Parser::transition` function. A validator rejection, parser
transition error, reader error, or provisional event-cap refusal immediately
returns a provisional failure. Both failure arms drop the provisional
validator before running the established full validator and authoritative raw
parser. Therefore an early raw/parser failure cannot outrank a later
value-only validation error, and no provisional diagnostic is exposed.

The shared loop reaches `SourceParseAttempt::Complete` only after the observer
has accepted the EOF event. EOF observation checks that a root was seen and the
validator depth is zero. The parser then consumes its owned state in
`finish_parse`, including shared-formula resolution, string lookup, scalar
materialization, and `Store` construction. This is after validation has
accepted the complete event stream, so a parser/materialization failure is a
post-EOF raw result rather than an unvalidated early result. The caller then
consumes `validator.finish()` before forwarding the owned result.

There is a boundary invariant to preserve: `finish_parse` currently runs
inside `parse_source_with_observer` before the caller invokes
`validator.finish()`. That ordering is safe because the observer already
performed every substantive current validation and rejected incomplete
root/depth state at EOF; `finish()` is presently a redundant state check. A
future substantive validation added to `finish()` would violate the
validation-before-materialization boundary and must instead run before
`finish_parse`.

For `Complete(Err(error))`, `complete_source_parse` preserves the historical
x14ac retry: it checks the descent marker and runs `x14ac::capture` when the
parser failed without a pre-captured extension state. A capture error retains
the historical precedence; a successful capture leaves the original raw error
in place. The shared successful path skips this scan only because its
eligibility proof already excludes the x14ac markers. The ordinary eager
`parse` facade still captures extension state before parsing and uses the
stateful completion helper with the same old retry condition, so its error
order is unchanged.

## Bounds, ownership, and lifetimes

The source-backed loader retains the original `SourcePayload`; the shared
reader borrows its bytes and stores no reader-backed event. The existing
aggregate 64 MiB worksheet budget is checked before traversal, and the shared
path adds an 8 MiB source ceiling, the default MCE input/output limits, and a
131,072-event provisional ceiling. The raw parser keeps its existing XML depth,
scalar, formula, record, and checked allocation limits. The candidate adds
checked reservations for the parser element stack and row-index set.

The validator owns only grammar-local element names, one dialect name, root and
depth state, and one first error. Its allowed worksheet grammar bounds nesting;
events, namespace views, decoded event payloads, and the reader never escape
the callback. The callback uses a higher-ranked event lifetime
(`for<'event> FnMut(...)`), and `Complete` carries only an owned
`Result<Store>`. The provisional validator is explicitly dropped before every
authoritative fallback.

These are finite input and state bounds, not a process memory budget or an OOM
guarantee. In the valid complete path, validator state remains live while
`finish_parse` builds the owned store. `resolve_shared_formulas`, namespace and
attribute decoding, and existing parser collections can also allocate through
their established paths. Fresh allocation captures must include this overlap
and late materialization errors; no source-level claim should describe the
candidate as fixed-memory.

## Plain-source eligibility and surrounding fences

The shared reader is admitted only when all of these byte-level conditions
hold:

- the original source is at most 8 MiB and within the default MCE input/output
  limits;
- the bytes are valid UTF-8;
- the MCE namespace URI, `AlternateContent`, x14ac namespace URI, and
  `dyDescent` markers are absent.

This is conservative: a marker in otherwise harmless text can force the
authoritative path, but no concrete false-positive admission was found. The
predicate matches the no-op conditions required by `process_ooxml`, and it
prevents direct parsing with extension values that could be needed by x14ac or
MCE processing. Sources outside the predicate retain validation, MCE
preprocessing, x14ac capture, and raw parsing in their established order.

`complete_source_parse` is correct only for the marker-free, eligible call
path. Its current call graph has the shared eligible validator as its sole
caller; ordinary raw parsing uses
`complete_source_parse_with_extension_state`. Keep this call-site invariant
explicit when reviewing later changes, or carry an eligibility proof into the
helper rather than using it as a general raw facade.

Snapshot selection, worksheet-relationship refusal, execution checks, source
identity/version fences, style and scalar validation, source retention, and
publication remain outside the shared parser at their existing boundaries.

## Differential coverage and admission blockers

The new public post-EOF test uses medium and dense-sparse plain worksheets
with a long valid prefix and a final `t="b"` value of `maybe`. The validator
accepts the complete XML through EOF, while raw materialization returns the
typed `invalid worksheet boolean 'maybe'` error. The test repeats the refused
edit, checks source bytes and provenance, reads the unaffected sheet, and
checks byte-exact no-op publication. This covers the `Complete(Err)` branch
that the inherited duplicate-`<v>` private fixture does not exercise.

Before runtime retention, the coordinator must:

1. Apply the frozen production and public-test inputs only from their bound
   manifests, then run the public post-EOF differential test against both the
   restored baseline and candidate.
2. Run fresh native and operation-scoped allocation captures for every required
   shape/repeat. Preserve the unchanged valid workflow/planning gates, require
   each candidate invalid late-raw p50 to remain at most 2x the same-shape
   baseline valid p50, require the invalid peak to remain at most 1.10x the
   same-shape baseline valid peak, and report same-invalid deltas separately
   with individual review of every adverse record above five percent.
3. Run the full 0541 early-parser/late-validator/error-precedence matrix
   alongside the existing source-backed owner integration tests and the
   candidate cap tests. The 0541 matrix covers its defined first-error cases;
   the existing owner tests retain invalid-UTF-8, reader, MCE/x14ac, and retry
   coverage, while the candidate tests cover source and provisional-event cap
   fallback. Keep all of those test inputs in their respective source
   manifests.
4. If the pilot passes, run the required instruction profiles and ordinary
   eager-read controls, then final quality checks and read-only evidence
   replay. Retain the candidate only after all gates pass; otherwise restore
   the source revision while preserving useful test/evidence artifacts.

The 0543 campaign remains part of the OLE2/OOXML priority; ODF optimization is
deferred until that goal is complete.
