# Change 0700 source review

Review status: source pass; final measurement review recommends rejecting the
direct `memmem` candidate for retention in the shared path. Reviewed candidate
codec SHA `d030f0671164fe861a6ae74f02842317d8822ef3429dd533066404bc8cbb4453`
and the exact 4,883-byte source diff recorded by the packet.

## Scope and proposed edit

Change 0699 identified the marker-presence precheck in
`crates/litchi-ooxml-common/src/mce/codec.rs::process_markup_compatibility` as
the measured fixed-window hot loop. The safe candidate is exactly:

```rust
if memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_none() {
```

replacing the existing `xml.windows(...).any(...)` predicate. The crate already
has a direct locked `memchr` dependency. No import, Cargo, feature, API, or
resource-policy change is needed.

This is a byte-for-byte lexical dispatch test. `NAMESPACE` is non-empty, so
`memchr::memmem::find` and the current fixed-size window comparison have the
same boolean result for an empty or short input, a match at every valid byte
boundary, repeated occurrences, near matches, and arbitrary bytes. The safe
library search has no access to XML structure and therefore preserves the
historical behavior that a URI in text, comments, or malformed bytes still
selects the reader path.

## Required source invariants

The candidate is acceptable only if the exact diff is limited to that
predicate (or an equivalent existing-module import plus the predicate), with
no changes to:

- the input-limit check or its precedence;
- the borrowed no-marker return, output-limit check, or default `Report`;
- `Reader`, `BoundedOutput`, `Frame`, `Ctx`, namespace layers, or MCE branch
  selection;
- error construction/order, output limits, allocation charging, or `Cow`
  ownership;
- the separate streaming MCE implementation and active-offset marker helper;
- public APIs, dependency versions, package ownership, or unsafe policy.

The source diff shows that marker-present inputs still enter precisely the old
reader path. The codec change is only the three-line predicate replacement;
the reader, frame, namespace, output, error, and report body is unchanged.
The five added focused tests cover short and near matches with arbitrary
bytes, input-before-output limits, marker bytes in comments and text, a final
position with truncated XML, and malformed UTF-8 in an attribute. The baseline
and candidate focused MCE runs each pass 101 tests with warning denial. These
tests do not replace the required end-to-end controls below.

The bounded disassembly confirms the intended loop removal and exposes a real
cost that the timing review must carry. The
`process_markup_compatibility` symbol is 6,305 bytes in the baseline and
7,565 bytes in the candidate (+1,260 bytes). The baseline entry reserves
`0x248` bytes after its prologue; the candidate reserves `0x320`, an increase
of 216 bytes in this function's stack frame. The old repeated `bcmp`/one-byte
advance loop is absent from the candidate processor body; the candidate calls
the existing `memchr::memmem::FinderBuilder::build_forward_with_ranker` and
uses its searcher result. This is a material code and stack tradeoff, not a
whole-program stack bound or a per-XML-depth bound.

## Required evidence

Before retention, the packet should bind both source states and run the focused
test matrix against each. It must include empty/short inputs, first and final
match positions, repeated prefixes and occurrences, first/middle/final-byte
near matches, arbitrary non-UTF-8 bytes, URI text/comment/malformed inputs,
and a late URI after a long prefix. It must also check input and output limit
precedence, exact output/error identity, complete `Report`, and `Cow` ownership.

The end-to-end controls must include marker-free refusal, real edit, real
no-op, and MCE-positive workflows. The positive workflow is necessary to
measure any setup cost that occurs before the existing reader and frame path.
The allocation/output/oracle checks must pass on the same frozen binaries;
timing claims require balanced repeated process legs with warm-up and retained
distributions. A microbenchmark or one favorable leg is insufficient.

## Frozen coverage audit

The frozen packet covers the main dispatch and preservation boundary. Both
focused receipts pass 101 tests. The added tests exercise empty and shorter
inputs, an arbitrary non-UTF-8 marker-free byte slice, final-byte near misses,
the input-before-output limit order on the borrowed path, exact URI bytes in
text and comments, a URI ending at the final input byte, and a malformed UTF-8
attribute that contains the URI. Existing MCE tests add transformed output
limits, parser refusals, report checks, and the unchanged streaming and active
offset paths.

The differential oracle covers 192 deterministic cases across five capability
and limit profiles (1,920 invocations, zero mismatches). Its corpus contains
36 identity, 36 directive, 36 duplicate, 36 QName, 36 unbound, and 12
synthetic cases. It exercises repeated exact occurrences (31 cases have two
and two cases have three), output/report/borrowed identity, and the existing
reader/refusal behavior. The separate marker controls cover a 245,766-byte
adversarial buffer with 4,096 repeated 58-byte near-prefixes, a 131,145-byte
late comment hit, a short marker-free input, and short/root comment hits.

The following are coverage gaps rather than observed semantic blockers. There
is no dedicated exact URI at byte offset zero, no valid XML case with the URI
at the last valid window, and no first- or middle-byte near miss; the focused
near misses differ at the final byte or end before the needle, while the final
window case is intentionally truncated. All 1,920 oracle subprocesses
complete, with 1,046 exact `OK` records and 874 exact typed `ERR` records
across the two sides (523 `OK` and 437 `ERR` per side). Those errors include
input-byte, depth, directive, and choice
limits plus malformed duplicate-attribute, unbound-prefix, and invalid-QName
refusals, so the differential receipt does exercise marker-present limit and
parser-error paths. The oracle corpus contains no non-UTF-8 bytes; focused
tests cover arbitrary bytes and malformed UTF-8. The marker controls are
lexical mechanism controls, not substitutes for semantic Office documents;
the real/declaration controls and the oracle cover that separate boundary.

## Static source decision

The one-line replacement is semantically compatible and follows the measured
unnecessary-work hypothesis. I found no source blocker: limit ordering,
borrowed return, report construction, parser entry, frame state, and the
separate streaming and active-offset paths are preserved. The conditional
retention criterion is not met: the terminal controls show recurring real-work
regressions alongside the +1,260-byte processor symbol and +216-byte local
stack reservation. This static source pass makes no retained production
performance claim; the final follow-up section records the rejection.

## Initial measurement assessment

The first frozen controls confirm both sides of the tradeoff. The long
repeated-prefix marker-free case improves by about 99.0% and the late comment
hit by about 96.4–96.8% in the two candidate pairs. Tiny marker-present cases
slow by about 43–53% (roughly 210–250 ns), while the tiny marker-free case is
too small for a useful median gain. The full native matrix improves the
marker-stripped, generated, and notes controls by roughly 3.8–17.7%, but the
three real workflows regress by about 2.0–3.7% at the median. The shared real
and declaration controls likewise show mostly positive candidate deltas, and
the existing MCE-positive refusal case has a p99 trigger above 40% even though
its median is close. Allocation counts and bytes are unchanged, so they do
not offset the measured latency or code-size cost.

These results make the direct `memmem` replacement a useful measured
hypothesis for long marker-free payloads, but they do not yet justify retaining
it as a general shared MCE optimization. The follow-up matrix and final gates
must decide whether the real-work regression is acceptable for the goal's
common Office CRUD scope. If it is rejected, a future experiment may measure a
lower-setup first-byte search with an exact `starts_with` check; that is only a
hypothesis and is not part of this review or the current source.

## Final follow-up decision

The terminal follow-up audit passes 24 native rows, 40 refusal rows, and 138
recomputed comparisons. It uses four balanced process legs and 300 samples
with ten warmups. For the primary `b0/a0` and `b1/a3` pairs, total median
latency changes are:

- `one-real`: +2.225% and +1.466%;
- `noop-real`: +2.094% and +3.442%;
- `two-real`: +2.418% and +1.709%;
- `one-control`: −13.064% and −13.619%;
- `one-generated`: −10.497% and −10.970%;
- `one-notes-poi`: −13.525% and −13.182%.

The real edit and no-op regressions remain across both balanced pairings,
while the large mechanism controls improve. Refusal totals improve in every
ordinary case; the MCE-positive `late-missing-relationship` case is +0.578%
and +0.069% at the median, and the MCE-positive late-root error is −1.728%
and −2.917%. The packet retains all 37 raw follow-up triggers, including
phase and maximum-value triggers, rather than treating them as semantic
failures. The final integration and evidence chains also pass, and the
follow-up audit confirms semantic identities, outputs, and source/binary
hashes.

The direct candidate should therefore be rejected for the shared production
path. It delivers the intended long marker-free scan improvement, but the
common real workflows regress and the implementation adds 1,260 bytes to the
processor symbol, 216 bytes to that function's stack reservation, and 7,064
bytes of native text. These costs are not justified by the measured gains on
marker-free and refusal paths.
The exact one-line candidate remains useful as a rejected experiment. A future
lower-setup first-byte search followed by an exact `starts_with` check may be
measured under the same controls; it has no result or retention claim here.
