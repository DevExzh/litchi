# 0722 writer-local structural fusion design review

This is a bounded, read-only design review for the next DOCX structural-scan
pilot. It is based on the restored baseline at `a1b259201d57b5a98c6ac053ac5fe65c45a3fa7e`,
whose five relevant production-file hashes are the same as the `source-final`
manifest in change 0721. No build, benchmark, or source edit was performed for
this review. The unrelated `docs/UNIFIED_OPS_API_DESIGN.md` worktree change is
outside this review.

## Finding

The next candidate should fuse only the two writer-owned structural walks and
should keep the public alternative-format scan and the shared read-only
namespace scanner byte-for-byte unchanged. The smallest safe boundary is a
private child module declared by `alt/codec.rs`, for example
`crates/litchi-docx/src/alt/codec/document_scan.rs`. It can own a writer-only
dual-state event loop and expose one `pub(crate)` adapter through `alt/mod.rs`.
`writer/doc/package.rs` would replace its adjacent calls to `scan` and
`active_block_ranges` with that adapter. `namespace.rs` and
the shared read-only scanner body in `parts/document_part.rs` should not change
in this pilot; the now-dead writer-only adapter and its unused imports may be
removed.

This isolates the new code from the regression signal in 0721. That candidate
changed the shared `WordElementRangeScanner` path used by paragraph and table
read controls. Its structural trace did halve the two writer-side reader counts
(2,820 to 1,410 on generated-medium and 356 to 178 on NumberedList), but the
read controls regressed in every pair: generated paragraph listing by
6.72–8.60% at p50 and 7.14–8.33% at the mean, and the pinned-media control by
5.59–6.05% at p50 and 5.69–5.93% at the mean. The candidate was rejected.
Those figures are prior-pilot evidence, not a speed claim for 0722.

The proposed file split is intentionally duplicative. Extracting a generic
scanner from `namespace.rs` would alter the optimized body of a read consumer
again. A private writer copy gives the pilot a measurable one-reader mechanism
while keeping the established read path and its compiler surface intact. The
copy must be guarded by differential tests and should not become a new public
XML API.

## Current writer path and exact fusion boundary

`DocumentBody::from_xml` strips the UTF-8 BOM, then currently does the following
on the BOM-adjusted bytes:

1. `crate::alt::scan(bytes)` parses all `w:altChunk` anchors and performs its
   first active MCE selection.
2. `active_block_ranges(bytes)` runs the shared `scan_word_element_ranges`
   scanner for `p`, `tbl`, and `altChunk`, then performs its second active MCE
   selection.
3. The existing body reader walks the bytes again to preserve the body prefix,
   suffix, unknown content, and typed body elements.

Only the first two walks are in scope. The final body reader must remain
unchanged: it has different body-depth and preservation responsibilities and is
not equivalent to either structural scanner. The helper should return the same
logical pair as the two existing calls: the sorted `BTreeMap<u32, Chunk>` from
the alt scan and the source-ordered `(target, start, length)` block ranges after
MCE selection.

The production delta should be limited to:

* one `mod document_scan;` declaration in `alt/codec.rs`, while leaving the
  existing public `Chunk` implementation, `active` body, `scan` body, and
  their private grammar functions in that file unchanged;
* one `pub(crate)` re-export in `alt/mod.rs`;
* the new private `alt/codec/document_scan.rs`, containing the writer-local
  copied dual-state grammar, bounded range admission, and one-reader loop;
* the import and two-call replacement in `writer/doc/package.rs`, plus removal
  of the now-unused writer-only `active_block_ranges` adapter and its imports
  from `parts/document_part.rs`;
* unit or differential tests owned by the new child module (or a writer test
  module), with no production changes to the `scan_word_element_ranges` body
  or any read-only consumer.

The child may call the existing private alt model types and the stable public
`active` function through its parent module, but it should not make
`reserve_document_value` visible across `parts` and `alt`. A local copy of that
small bounded admission helper, retaining its exact count and allocation
labels, avoids a new module cycle and keeps this delta narrow.

The staged 0722 snapshot should be checked against this boundary before it is
treated as the candidate. Moving the public `active` and `scan` implementations
out of `alt/codec.rs` into the new child may preserve their text, but it changes
their source ownership and makes the public/read path part of the candidate
diff. Keep those bodies in `codec.rs`; the child should contain only the
writer-local adapter and its copied dual-state machinery. Removing the
writer-only `active_block_ranges` function from `document_part.rs` is a separate
cleanup and does not alter the shared read-only scanner body.

## Event-loop contract

The writer adapter should use one `NsReader` over the original byte slice. On
each iteration it must preserve the current alt reader sequence:

```text
buffer_position -> u32 conversion -> decoder -> read_event -> into_owned
-> resolver clone -> resolve_event
```

The resolved event is presented to the alt state first. Only if that state
accepts the event does the writer-local range state observe it. This preserves
the established alt-first error precedence and avoids handing a rejecting alt
event to the range side. The reader must be scoped inside the walk and dropped
before either MCE call; this is discussed under resource lifetime below.

The alt state is a mechanical private copy of the current grammar. It keeps
its own depth, pending anchor, opaque-subtree depth, properties depth, chunk
map, and `MAX_CHUNKS` admission. It must retain the current relationship
attribute namespace checks, duplicate and empty ID errors, strict versus
transitional `matchSrc` values, character-data refusals, and the current
`next_depth` behavior. In particular, an `Empty` event at current alt depth
256 still refuses because the established alt grammar checks `depth + 1` even
though an empty element does not persistently increase depth. A later sibling
empty element at depth 255 must remain valid.

The range state is a mechanical private copy of the current
`scan_word_element_ranges` behavior, with the same three target indexes:

```text
0 = p, 1 = tbl, 2 = altChunk
```

It must retain the independent 128-level total/capture depth limits, the
one-million element-node limit, checked `usize`/`u32` offset and length
conversions, capture suppression for nested targets, EOF and malformed nesting
errors, and the existing callback order. The writer copy must retain the
fragment namespace rule exactly:

* a bound transitional or strict Word namespace is accepted;
* if the first unbound root has an unknown prefix, only that same unknown
  prefix is treated as the fragment Word namespace;
* if the first root is unbound, only unbound elements are accepted;
* a bound foreign namespace is never accepted merely because its local name is
  `p`, `tbl`, or `altChunk`.

The first root-prefix observation is made only for a `Start` event, as in the
current scanner. It must not be broadened to `Empty`, and it must not be
replaced with a local-name-only test. The copied range state may consume the
already resolved event from the shared reader, but it must not change the
read-only `read_resolved_event` implementation or its caller.

The range side has a separate state and separate failure slot. Once its first
failure occurs, it stops accumulating ranges and continues allowing the alt
state to consume later events. At EOF the adapter must apply the following
ordering:

1. Run the first `active(bytes, alt_offsets)` call using the alt map's sorted
   keys, even when the vector is empty. Filter the map in the established
   sorted order.
2. If that MCE call fails, return its `Error::from` result before any deferred
   range failure.
3. If a range failure was remembered, return that original range error now.
4. Reserve the separate block-offset vector with the existing
   `active document block offsets` resource label, call `active(bytes, starts)`,
   convert the result to a set, and retain ranges in their original order.

This ordering reproduces the old call sequence: the alt scan and its MCE
decision precede the block-range walk and its MCE decision. It also preserves
the useful behavior that a later malformed alt event wins over an earlier
range-side refusal, while a successful alt MCE wins over a deferred range
failure. Do not return a range error immediately from the event loop, and do
not invoke the second MCE call after a range-side failure.

The helper must retain the current error classes and text families. Alt XML
reader failures remain `Error::Xml`; alt grammar and offset failures remain
`Error::Invalid`; range grammar, offset, depth, and node failures remain
`Error::InvalidFormat`; fallible range growth remains `Error::Allocation` with
`active document block ranges`; and MCE failures continue through the existing
`active` conversion. No new wrapper error should hide the original category or
message.

The BOM boundary is part of this contract. `DocumentBody::from_xml` removes
the BOM before invoking the helper, so every returned offset is relative to the
same BOM-adjusted slice used by the existing calls. The final body reader adds
the BOM back to the preserved prefix. The helper must not strip or re-add it.

## Resource lifetime and the +2,480-byte observation

The 0721 allocator diagnostic showed the important cost of the fused lifetime
on NumberedList edit: peak live bytes above the region start changed from
21,310 to 23,790 bytes, an increase of 2,480 bytes (11.64%). Net live bytes,
allocation requests, and requested bytes improved or stayed aligned. This was
not an RSS measurement. Source inspection gives a lifetime hypothesis for the
difference: the fused implementation kept the range-side reader/state and the
growing range vector alive while the alt-side first MCE operation ran. In the
sequential baseline, those range-walk locals did not coexist with the alt MCE
boundary. The receipt did not isolate individual suballocations, so it does not
prove how much of the 2,480 bytes came from retained range records and how much
came from reader or resolver scratch.

The next implementation should reduce that lifetime without weakening any
semantic or resource cap:

* Put the complete one-reader event loop in an inner scope and move out only
  the alt map, the retained raw range records, and the optional deferred range
  error. This drops the `NsReader`, resolver buffers, and range state before
  the first MCE call.
* Keep the range records bounded and source ordered. A private compact record
  containing a small target discriminant plus two `u32` values is compatible
  with the writer adapter and can reduce capacity pressure; it must preserve
  the same count reservation and output triples. This is an optional measured
  refinement after the scoped lifetime is in place, not a reason to change
  offset types or remove checks.
* Do not use `shrink_to_fit` as a correctness mechanism. It changes allocator
  timing and may copy the records; any such experiment needs its own
  allocator evidence.
* Do not stream ranges to an ambient file, grow an unbounded side buffer, or
  call MCE incrementally. Those choices would change the explicit-source and
  bounded-resource contract or the ordered MCE semantics.

The fused pass necessarily retains enough range metadata to return it after
the first MCE call; that lifetime cannot be eliminated while also retaining a
single lexical walk and the existing `active` implementation. The compatible
goal is to drop the reader and parser scratch at the MCE boundary and make the
remaining record storage compact. The allocator lane must report peak above
start separately for generated-medium and NumberedList. The frozen hypothesis
gate is peak-above-start no more than 3% above baseline for each corpus and
phase, together with no increase in net live bytes. A result outside either
bound requires rejection or revision after the scoped implementation; this does
not impose a zero-peak-change policy.

## Frozen differential coverage

The new child module needs an independent, reproducible differential matrix;
passing only the ordinary writer tests would be insufficient. Each fixture
should run the existing baseline pair (`scan` plus `active_block_ranges`) and
the writer helper on the same bytes. Successful results must compare the exact
chunk map, target order, offsets, lengths, and MCE-selected ranges. Error
results must compare a stable category plus display and debug fingerprints.
The expected values for boundary cases must be frozen independently rather than
generated from the candidate.

The matrix should retain the 0721 controls, including:

* transitional and strict namespaces, plain blocks, valid `Choice`/`Fallback`,
  inactive branches, nested anchors, and foreign lookalikes;
* unbound fragments and unknown-prefix fragments, including an empty root;
* malformed alternate content, malformed XML tails, unknown
  `MustUnderstand`, duplicate/missing relationship IDs, invalid strict and
  transitional `matchSrc` values, and unexpected character data;
* BOM-adjusted input and exact source offsets;
* alt depth 255/256/257, the depth-256 empty-event refusal, two empty siblings
  at depth 255, range depth 127/128/129, and nested anchor capture;
* the 32 MiB raw XML cap, 4,096/4,097 alt anchors, one-million range nodes,
  one-million-plus-one range values, active-offset count boundaries, checked
  offset/length conversion, and both allocation resource labels;
* range-before-bad-alt and alt-before-range-error precedence, including an
  error on the final EOF event.

The 19 synthetic public-oracle cases and both complete package corpora from
change 0721 should remain the external oracle: generated-medium's 21,517-byte
main XML and the real NumberedList document's 4,563-byte main XML. The oracle
must compare public scan/active results and the published writer bytes exactly.
The final body walk must remain in the writer result so the differential test
also catches ordering, unknown-block retention, and section-prefix/suffix
changes.

The trace lane should separately prove that successful representatives use one
reader for the two fused structural owners and that the observer count matches
the baseline event count. It must not treat reader-event count as a timing or
physical-I/O claim. On refusing inputs, the admission point and the fact that
the alt side saw the rejecting event first should remain explicit.

## Pilot gates and disposition

Before native capture, run the DOCX format, all-target tests, warning-denied
Clippy, doctests, and rustdoc on the candidate, then run the public oracle and
the frozen differential matrix. The read-only paragraph and eager-count
controls must be captured even though their source bodies are unchanged; this
guards against an unintended release-code or module-layout effect. The 0721
ABBA native gates remain the appropriate primary comparison: each edit p50 and
mean must improve by at least 3%, while lifecycle p50 and mean may regress by at
most 3%. Allocator request counts and requested bytes retain their 3% gate.

For this candidate, peak-above-start is a first-class review gate because the
prior implementation exposed a measured lifetime regression. The frozen
hypothesis permits at most a 3% peak-above-start regression for each corpus and
phase, and requires net-live bytes to show no increase. Both values must be
visible per corpus and phase; a result outside either gate requires rejection or
revision. This is a bounded regression gate and does not require zero peak
change. Tail and repeat-drift flags remain visible and never justify dropping
samples.

If any correctness, parity, read-control, lifecycle, allocator, or peak gate
fails, restore the baseline production files and retain the candidate, tests,
measurements, and failure analysis in the packet. If all gates pass, run the
full final-source evidence checks before retaining the five-path production
delta (four existing files plus the new child module). No source-level speed
claim should be made before those measurements;
the design only establishes a narrower hypothesis than 0721.

## Independent trace review

The terminal trace packet `trace-analysis.json` reports `status: pass` for all
21 trace documents: the 19 synthetic cases plus generated-medium and
NumberedList. The public oracle reports are byte-identical at 32,260 bytes
(`4540ce8b6e236dae30f92f8d26b3a0068c0e866dd6b5695adde455badacb9721`). An
independent comparison of the raw trace records found exact equality for the
active-call metadata and for every raw and MCE-selected chunk and range list.
The three early-failure cases where the baseline has no range scope and the
candidate has an empty incomplete scope are scope-lifetime differences; they
contain no semantic metadata.

The successful-read counts show the intended fusion separately from the range
observer count:

| corpus | baseline alt + range reads | candidate alt + range reads |
| --- | ---: | ---: |
| generated-medium | 1,410 + 1,410 | 1,410 + 0 |
| NumberedList | 178 + 178 | 178 + 0 |
| all 21 trace documents | 2,525 + 1,877 | 2,525 + 0 |

Every candidate range-reader count is zero, and its alt reader start/end spans
match the baseline alt spans for all 21 documents. The baseline trace records
`end: null` for range-reader calls because that instrumentation boundary has no
end position; those unknown ends are not inferred or compared. Candidate end
positions are available for its fused reader calls. The all-document totals are
4,402 baseline reader calls and 2,525 candidate reader calls; the split is kept
visible so it is not mistaken for an observer-event count. The corresponding
observer-event totals are 1,876 baseline and 2,138 candidate.

The observer callback has a different refusal boundary from the successful
reader count. In the baseline, the range observer runs after the event-end
offset conversion and structural admission. In the candidate, it runs
immediately before `classify`, after the fused loop has admitted the event to
the range phase but before `classify` applies its node and depth checks. Thus
`range-depth-only` records 129 baseline range reads but 128 baseline observer
events: the last read is refused before the baseline observer. The candidate
records 129 observer events, including that event, while still making only its
260 alt reader calls. The candidate observer count therefore must not be used
as a successful-read count.

The same distinction explains the early refusals. Baseline range scopes are
absent for `empty-at-depth-256`, `range-depth-before-bad-alt`, and
`malformed-tail` because the sequential alt walk refuses before the old range
walk starts. The fused candidate has already created its range state, so it
records 129, 129, and 3 observer events respectively while making no range
reader calls. These are admission-boundary observations; they do not indicate
additional source reads or a change to error precedence. The trace preserves
the ordered MCE calls and raw/selected metadata for every case where those
stages are reached.

## Read-control tail retention review

I independently read all 16 frozen read-control reports and recomputed each
`result.elapsed_ns.samples` vector. Every vector contains 200 retained samples,
and the recomputed p50, p95, p99, maximum, and mean match the recorded
statistics. The paired raw-vector deltas are:

| pair | control | p50 delta | mean delta | p99 delta | max delta |
| --- | --- | ---: | ---: | ---: | ---: |
| pair-1 | generated-medium-list-paragraphs | -0.061% | -0.198% | +0.602% | -12.607% |
| pair-1 | pinned-media-eager-paragraph-count | -0.828% | -0.448% | -0.255% | -4.465% |
| pair-2 | generated-medium-list-paragraphs | -0.672% | -0.508% | +0.544% | +1.459% |
| pair-2 | pinned-media-eager-paragraph-count | -0.694% | -0.229% | +1.457% | -10.243% |
| pair-3 | generated-medium-list-paragraphs | -0.351% | -0.182% | -0.014% | -1.486% |
| pair-3 | pinned-media-eager-paragraph-count | -1.719% | -0.576% | +2.595% | +36.537% |
| pair-4 | generated-medium-list-paragraphs | -0.486% | -0.337% | -3.538% | +1.307% |
| pair-4 | pinned-media-eager-paragraph-count | -0.741% | +0.156% | +18.583% | +21.727% |

The two pinned-media tail flags are present in the raw sorted vectors. In
pair-3, the baseline p99 and maximum are 118,291 and 124,530 ns; the candidate
values are 121,361 and 170,030 ns. Five candidate samples exceed the baseline
p99 and two exceed the baseline maximum. In pair-4, the corresponding baseline
values are 115,751 and 117,731 ns and the candidate values are 137,261 and
143,310 ns; six candidate samples exceed both baseline thresholds. No p95
comparison exceeds the 5% review threshold.

The frozen within-variant stage groups retain ten repeat flags, all at p99 or
maximum:

| stages | flagged control and metric | spread |
| --- | --- | ---: |
| baseline-A1 → baseline-A2 | generated maximum | +33.152% |
| baseline-A1 → baseline-A2 | pinned-media maximum | +46.055% |
| candidate-B1 → candidate-B2 | generated maximum | +14.693% |
| candidate-B1 → candidate-B2 | pinned-media maximum | +37.222% |
| baseline-A3 → baseline-A4 | generated p99 / maximum | +6.453% / +27.929% |
| baseline-A3 → baseline-A4 | pinned-media maximum | +5.775% |
| candidate-B3 → candidate-B4 | generated maximum | +31.556% |
| candidate-B3 → candidate-B4 | pinned-media p99 / maximum | +13.101% / +18.645% |

The p50 and mean repeat spreads remain below 5% for every read-control group;
the largest observed spreads are 1.12% for p50 and 0.84% for mean. The primary
native lane has no native tail or repeat flags over 5%. The read-control
analysis reports all p50/mean hard gates passing, exact normalized output
parity, and raw statistics recomputed from all retained samples.

The unchanged-read-path custody checks also pass: `source-guard.json` reports
the whole namespace file and guarded public/read-consumer regions unchanged,
and every read child records the candidate source unchanged before and after
capture with no changed paths. These checks establish custody only; they do
not assign a cause to any tail value.

Under the frozen rule, p50 and mean are the hard gates, while every tail and
repeat result above 5% remains visible and cannot justify dropping samples,
stages, or pairs. The bounded recommendation is to retain the candidate and
retain all of these read-tail flags in the packet. No causal explanation or
exclusion is made here; the final source disposition remains with root.
