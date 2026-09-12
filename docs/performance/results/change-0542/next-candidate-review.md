# 0542 next-candidate source review

This is a read-only review of the proposed follow-up in
[`next-candidate.patch`](next-candidate.patch). The patch is relative to the
rejected, fully applied 0542 candidate, not to the restored production source.
Its SHA-256 is `83e325982227cbc2122f2bee8ec7b3aacda7cae2a18006f10c1cf12a1c708da8`.
It has no build, benchmark, or allocation evidence in this review.

The follow-up is suitable for an isolated next-campaign trial. It is not
admitted for runtime retention. Admission still requires a fresh matched
source campaign, the unchanged quality checks, and a differential test for
the post-EOF raw error that this patch changes.

## What the follow-up changes

The old shared driver converted every `Parser::finish_parse` error into
`ProvisionalFailed`, which caused the caller to run full validation and the
ordinary raw parse again. The follow-up changes the result enum to carry
`Complete(Result<Store>)`. The caller first runs `validator.finish()` and then
passes that result through `complete_source_parse`.

The extracted completion helper retains the old `raw::worksheet::parse`
behavior: when a plain worksheet parser result is an error, it runs the
historical x14ac capture retry; a failed capture replaces the parser error and
a successful capture leaves the parser error in place. The helper is also
used by the ordinary raw facade, so that refactoring is textually equivalent
to the previous retry block.

## Error order and ownership

The event loop still gives the validator the borrowed event first. A validator
error, parser transition error, event-cap refusal, or reader error returns a
provisional failure immediately. The caller explicitly drops the validator
before running `worksheet_xml(content)` and then `raw::worksheet::parse`, so
the authoritative fallback does not overlap the provisional diagnostic or its
element stack.

The `Complete` arm is reached only after the observer has accepted every
event, including EOF. `Validator::observe` returns `false` for its first
validation error and for an incomplete root/depth state at EOF. Therefore a
`finish_parse` error cannot outrank a validation error: a validation failure
takes the provisional fallback, while a `Complete(Err(_))` result means the
observer has already accepted the complete validation stream. The final
`validator.finish()` is still the right ownership check before forwarding the
raw result.

There is a resource-order detail that must stay explicit. `finish_parse` is
executed inside `parse_source_with_observer` before the caller consumes the
validator with `validator.finish()`. This is safe for the current state
machine because the observer has already accepted EOF, but the method does
resolve shared formulas, invoke the strings closure, materialize cells, and
possibly construct a `Store`; it is more than an error-only finalization. The
implementation and design text should describe the boundary as “after the
observer accepts the complete source, before the result is forwarded after
`validator.finish()`.” It must not imply that materialization runs only after
the final validator method if that method is later changed to do substantive
validation.

For an eligible source with a late raw materialization error, the parser is
consumed and its temporary output is dropped, `validator.finish()` succeeds,
and `complete_source_parse` performs the x14ac retry without repeating the
full validator/raw fallback. For a validation error, `finish_parse` is never
reached because the observer stops first. This is the intended reduction in
late-raw work and does not change the first-error owner.

## Eligibility and surrounding fences

The follow-up does not widen the shared-reader eligibility predicate. The
inherited candidate still requires valid UTF-8, the 8 MiB source cap, the MCE
input/output limits, and absence of the MCE namespace, `AlternateContent`,
x14ac namespace, and `dyDescent` markers. Sources outside that proof continue
through the established validation, x14ac preprocessing, MCE processing, and
raw parser path.

`complete_source_parse` is correct only as the completion step for the plain,
marker-free source path. It must not become a general replacement for
`raw::worksheet::parse` on MCE or x14ac-bearing input. The existing 0541 MCE,
DTD/PI, namespace, and extension cases should remain in the retained matrix;
their preprocessing and typed-error ownership are unchanged by this patch.

The follow-up changes no snapshot selection, `SourcePayload` ownership,
source identity, cancellation checks, source-version fences, style/scalar
closure, or publication point. Those inherited boundaries remain required:
the source-backed loader must still check execution around its existing work,
construct no snapshot from a provisional store, and preserve the source bytes
after both accepted and refused transactions.

## Lifetime and trait boundary

The callback remains higher-ranked over each event lifetime:

```text
for<'event> FnMut(&ResolveResult<'event>, &Event<'event>) -> bool
```

No event, namespace view, decoded text, or reader reference is stored in the
returned value. `Complete(Result<Store>)` contains only owned data and is
crate-private plumbing. The `FnOnce` strings callback is consumed only after
the observer has accepted EOF, and no borrowed parser state crosses the
`codec`/worksheet facade boundary. No public trait, unsafe block, dependency,
or source-payload lifetime contract is introduced.

The lifetime conclusion depends on keeping `Parser::finish_parse` private and
returning an owned `Result<Store>`. A future version must not expose a parser
or an event-bearing continuation merely to defer materialization.

## Resource boundary

The follow-up adds no event log, normalized-token buffer, second XML source,
or new unbounded collection. The inherited provisional limits remain 8 MiB of
source, 131,072 reader events, the model event/depth checks, bounded scalar
and formula text, checked parser row/stack reservations, and the validator's
owned grammar stack. These are finite input/state bounds, not a process-wide
memory budget or an OOM guarantee. Namespace resolution, attribute decoding,
shared-formula resolution, and existing ordinary collections still use their
established allocation behavior.

The explicit `drop(validator)` fixes the previously identified overlap on
provisional fallback. In the complete valid path, the validator has accepted
EOF and its empty/capped stack remains alive while `finish_parse` returns its
owned result; this finite valid-path overlap must be included in the fresh
allocation and invalid-late measurements. A `Complete(Err(_))` path no longer
falls back and therefore avoids the old duplicate full traversal, but it still
may allocate the preceding materialized cell vector before the late error.
The 1.10x invalid peak envelope is useful corpus evidence and does not replace
the source-level bounds or the existing no-OOM qualification.

## Missing differential coverage

The retained `shared_traversal_tests::LATE_RAW_TAIL` fixture uses duplicate
`<v>` elements. That error is found during `Parser::transition`; it does not
exercise the new `Complete(Err(_))` branch. The 0541 public matrix has a typed
boolean case, but its small case is an expected-error assertion rather than a
paired baseline/follow-up check of this post-EOF path. The standalone guard's
late-raw row checks the expected message and a repeated candidate attempt, but
does not by itself compare the restored two-pass implementation with this
follow-up.

Before retention, add or run a paired public source-editor differential with
a substantial valid prefix and a final typed materialization fault such as
`t="b"` with `<v>maybe</v>` (the guard fixture is suitable). Assert, for the
restored baseline and the follow-up, the same typed error and message, source
byte/provenance identity, no published snapshot, and the same result on a
retry and on a later unaffected-sheet transaction. Run the case for both
medium and dense-sparse shapes. Keep the existing early-parser/late-validator
cross-product because it exercises the separate immediate-fallback branch.

The retained matrix must also continue to prove that MCE/x14ac-bearing,
invalid-UTF-8, reader-error, and cap-refused inputs use the authoritative path.
For the plain eligible path, include a parser/materialization error where the
x14ac retry succeeds and leaves the original error intact. Any test of a
failed x14ac preflight belongs to the ineligible path and must assert that the
preprocessing error still wins before raw parsing.

## Measurement and disposition

`next-candidate.patch` is unmeasured relative to the full applied candidate.
The rejected 0542 late-raw rows, including the three ratios above the frozen
2.0x valid-planning p50 envelope, cannot be reused as evidence for this
follow-up. Rebuild a fresh baseline and follow-up from source-bound manifests
and rerun normal and allocator guard captures with matched fixtures and
repeats.

Keep all existing gates: ordinary native workflow/planning limits, ordinary
allocation peak/bytes limits, invalid late-error p50 no more than 2x its
same-shape valid planning baseline, incremental invalid peak no more than
1.10x, and individual review of every greater-than-5% adverse record. The
late-raw rows are the blocking measurement for this proposal; valid controls,
late-validator rows, cap fallbacks, and medium/dense-sparse repeats must be
reported alongside them. Allocation evidence from a binary without counters
must remain unavailable rather than being treated as zero.

Static disposition: the follow-up preserves the intended validation-first
error owner, private callback lifetime, eligibility boundary, and x14ac retry
when the observer invariant is maintained. It is approved only as a bounded
next-campaign proposal. Do not retain it in the restored runtime until the
post-EOF differential test and fresh native/allocation guard evidence pass.
