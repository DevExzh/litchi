# 0966 DOC SPRM census review

Status: pre-seal review.  This note records the acceptance conditions for the
bounded diagnostic described by the coordinator; it records no build,
workload, timing, or production result.  The reviewed checkout is
`a499d34f5119a034bc92abde1e1dfcf529159b8d`.  The uninstrumented export was
recorded before applying the external diagnostic patch (`crates/litchi-doc/src/sprm.rs`
SHA-256 `f7a5499cbfb7b97a493c551f4fc33c1518c64beaaf47010f07e3644b4b1baaa0`;
`tools/perf-cfb-profile/src/lib.rs` SHA-256
`b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc`).  The
patched export identities are recorded separately in
`source-provenance-v1.json` and must be rebound if the design is resealed.

The hypothesis is narrow: `parse_sprms_two_byte` always requests at least
eight `Vec<Sprm>` elements, including for an empty `grpprl`; a bounded census
can establish how often that shape occurs in the already measured DOC edit
lifecycle.  The census may justify a later source-level candidate, but it
cannot establish that the capacity is removable, that it dominates native
time, or that a candidate would improve end-to-end latency or memory.

## Review disposition

The batch is admissible as evidence-only diagnostic work if the sealed design
and final packet satisfy the requirements below.  It must remain external
source instrumentation.  No production parser, public API, capacity policy,
or retention policy should be changed from this census.

The scope is the existing `tools/perf-cfb-profile` measured lifecycle around
`Snapshot::open_bounded`, one `Edit`, all staged paragraph-zero replacements,
retention release behavior, `commit`, output copying, and local destruction.
The two fixtures are `NoHeadFoot.doc` and `FloatingPictures.doc`, with 1, 2,
8, and 32 replacements and `zero`, `default_8mib`, and `release_8mib`
retention.  Two fresh child processes per cell give 48 cells.  The intended
`warmup=0, samples=1` setting is suitable for a deterministic call census;
affinity, cache-state, host-noise, and timing qualification controls are not
needed and should not be presented as such.

## Parser instrumentation requirements

Instrumentation must be outside the repository and must preserve the original
parser result exactly.  The least risky shape is a wrapper around the existing
parser body, with a fixed diagnostic observer; it must not replace the parser,
change its error ordering, alter `SmallVec` operands, or bypass validation.
If observing a valid prefix on an error requires moving the body into a local
closure or helper, the uninstrumented and instrumented results must be
compared independently, including malformed/error controls.  A source diff
and a successful build are not enough to prove error identity.

Count only one logical call at the `parse_sprms_two_byte` owner.  Do not count
both `parse_sprms` and its delegate.  The fixed counter record should include
the minimum fields needed by this question:

* total calls, successful calls, error calls, empty-input calls, and nonempty
  calls;
* total input bytes for all calls and total successful output SPRM records;
* successful returned `Vec<Sprm>` capacity in elements, when the unchanged
  result wrapper exposes it; and
* a fixed error-kind histogram keyed by the existing `Error` variants.

The error count must include an error after a valid prefix.  Reporting
`Result::is_err()` is sufficient for this census provided the parser body is
left unchanged; do not add intrusive prefix bookkeeping merely to recover a
discarded vector.  Output counts for an error must not be presented as
returned output.  Successful empty parses must remain distinguishable from
malformed nonempty calls.  The unchanged parser's initial request formula,
`max(grpprl.len() / 4, 8)`, may be reported as a derived interpretation of
input length; it is not a measured allocation counter and does not apply to a
discarded error result.  Do not add growth or prefix-length instrumentation
merely to recover that formula.

The capacity fields are logical returned-vector element counts only.  A report
may include the explicitly modeled value
`successful_output_capacity_total * size_of::<Sprm>()`; label it as logical
modeled bytes, never allocator traffic, live peak, or a lower bound on all
requests (errors have no returned vector).  Do not mix this census with the
0730 counting allocator or native timing lane.  If a histogram is used, define
all bins and include an explicit overflow bin.  Histogram conservation must
hold exactly:

```
calls = successful_calls + error_calls
calls = empty_calls + nonempty_calls
histogram_count_sum = the corresponding call count
error_kind_histogram_sum = error_calls
```

Use checked or saturating diagnostic arithmetic with an explicit overflow
failure/flag; never allow a counter wrap to make the packet appear valid.
Static atomics or another fixed, allocation-free observer are appropriate.
The observer must not allocate, retain document bytes, or expose paths,
credentials, or payload content.

## Measurement-region controls

The observer must be disabled and reset before `prepare` starts.  Preparation,
fixture hashing, all route/reference controls, public-reader reopen checks,
the 0730 oracle, output comparison, JSON serialization, and any post-sample
validation must contribute zero to the census.  Enable it immediately around
the `profiled_measured_case` call, take the counter snapshot, and disable it
on both success and error paths with an RAII guard.  A panic or early `Result`
return must not leave the process globally enabled for later work.

The measured wrapper may return the counters alongside the unchanged
`CaseResult`, but the counters must be captured before the result is consumed
by reporting or correctness code.  The selected sample's output and all
existing correctness gates remain mandatory: cross-route byte equality,
semantic readback, public reopen, exact no-op/source sharing, patch forward and
inverse/refusal checks, atomic refusal behavior, retention ceilings/release,
untouched streams and directory metadata, semantic oracle success, and exact
edit-count publication.

The packet should use a new diagnostic schema and state `timing_claim: false`
explicitly.  Any inherited `elapsed_ns` field must be ignored by reducers and
must not be used to compare retention routes.  A parser counter result is
valid only when all correctness gates for that child pass.

## Fresh-child protocol and controls

Each child must run exactly one fixture/edit/retention cell with the same
source export, toolchain, build flags, harness revision, fixture byte count,
fixture SHA-256, and diagnostic source hash.  The launcher should reject a
missing, malformed, duplicate, or extra child result.  It should retain the
child identity and command, but it need not pin CPUs or collect physical I/O
or host-noise evidence for deterministic counts.

The two repetitions of every cell must match exactly for every census counter,
histogram, selected output identity, and correctness gate.  Per-child route
and stage observations remain subject to the existing gates; include them in
the repeat identity only if the sealed schema declares them part of this
census.  Do not retry a mismatch or average it away.  A mismatch invalidates
that cell and is a diagnostic finding about the observer or the lifecycle,
not a performance result.  A complete packet needs all 48 valid cells; no
replacement process may be substituted after launch.

At minimum, the reducer should check these shape relations per child and per
cell:

* parse-call and histogram conservation;
* input-byte and successful-output counters are nonnegative and fit their
  declared integer widths;
* any capacity histogram reconciles with its successful-output capacity total,
  including overflow bins; and
* error-kind counts sum to the error count, including errors after a valid
  prefix; and
* the two fresh repetitions are byte-for-byte equal after removing only
  process identity and explicitly non-semantic metadata.

The edit-count sweep is useful because an intercept/slope pattern can reveal
whether calls repeat per staged replacement.  Report that as a descriptive
count relationship only.  Retention routes must remain separate: a different
SPRMs count between zero/default/release is a route observation, not evidence
that one route is faster or that work can be removed.

## Interpretation boundary

The strongest permitted conclusion is a bounded, current-source statement
such as: “In this named lifecycle and fixture matrix, N successful and M
error `parse_sprms_two_byte` calls occurred; K had empty input, and the
logical initial-capacity census recorded X element slots.”  It may also report
an exact repeated-count relation across 1/2/8/32 edits when the child identity
and conservation checks pass.

The packet must not claim a percentage of native latency, allocation bytes,
RSS, instructions, or removable work. It must not infer that empty calls are
allocations saved, because the logical capacity counter does not observe the
allocator and the parser's ownership/error contract is not changed. A later
candidate would require a separate source-preserving implementation, focused
correctness/adversarial tests, and fresh native plus allocator qualification;
this census alone authorizes none of those actions.

The evidence is compatible with ADR 0001's ownership boundary because the
instrumentation stays in the DOC owner/export, with ADR 0003 because source,
commit, patch, no-op, and atomic refusal behavior remain exercised, with ADR
0005 because the observer is bounded and evidence-only, with ADR 0006 because
validation/error and preservation behavior are retained, and with ADR 0008
because the diagnostic makes no unsupported performance claim.  The result
does not amend an accepted ADR or establish a production resource policy.

## Acceptance checklist for the sealed packet

Before accepting the final packet, verify:

1. A sealed design lists the exact source export and intentional source diff;
   after sealing, the author does not alter protocol, bins, thresholds, or
   interpretation without a new version and reason.
2. The external source export and original parser/harness identities are
   recorded, and the instrumented build is reproducible with no repository
   binaries or large evidence files.
3. The counter observer is fixed-size and allocation-free, counts every
   parser error including errors after a valid prefix, and reports logical
   capacity only.
4. The active guard excludes all preparation/oracle/post-sample work and is
   disabled on errors and unwinding.
5. All 48 child cells pass existing output/correctness gates and both
   repetitions match exactly; the reducer proves counter conservation.
6. The report says `timing_claim: false`, labels the result diagnostic, and
   makes no CPU, allocator, RSS, or causal removable-work claim.
