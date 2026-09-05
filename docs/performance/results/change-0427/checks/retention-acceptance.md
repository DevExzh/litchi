# 0427 retention acceptance gate

This note separates the bounded 0427 allocator checkpoint from the fuller
source-backed retention lifecycle still required afterward. It is an
acceptance checklist, not a result or a performance claim. The overall
non-iWork goal remains open.

The 0427 checkpoint may report absolute allocator callback live-byte values at
explicit post-drop points. Those values are valid numeric observations when
their callback-order and scope are recorded. They do not, by themselves,
identify object ownership, prove cache eviction, prove managed-budget release,
or measure RSS release. The near-limit cache/budget and phase-aware RSS matrix
below remains a separate follow-up gate. Passing the checkpoint authorizes no
causal optimization or retention claim.

## What the retained evidence does not show

The 0424 media lifecycle proves a reduction in operation allocator request
volume and in the absolute V3 operation-region peak. Its `live_after` values,
whole-process RSS, and Heaptrack traces do not identify which caller-owned
objects or source-cache entries remain alive. The 0424 record therefore
correctly makes no managed-budget, cache-retention, or post-drop claim. Its
remaining-work record explicitly calls for plan, publication, and caller-drop
snapshots. See [the resource decision](../../change-0424/resource-review.md)
and [the remaining-work record](../../change-0424/next-work.md).

The 0086/0088 OPC cache work supplies useful correctness and structural
invariants for reservation, pinning, eviction, bypass, single-flight, and
package release. It is a separate synthetic cache matrix, however; it does
not bind those counters to the 0424 source-backed PPTX plan/publication/drop
lifecycle. The 0405 lock observer is likewise a scoped diagnostic and does
not turn lock counters into a retention or latency measurement.

### Cache-diagnostic availability

The public PPTX source and editor methods at
`crates/litchi-pptx/src/presentation/source.rs:1108` and `:1396` currently
forward the infallible compatibility `SourceBackedPackage::cache_diagnostics`
snapshot. The OPC owner also exposes
`SourceBackedPackage::try_cache_diagnostics` at
`crates/litchi-opc/src/source_backed.rs:4250`, which fails closed on a poisoned
cache-state mutex or checked counter overflow. The PPTX seam does not forward
that fallible path, so a lifecycle report obtained through the public format
owner could mistake a recovered poison or invalid counter state for valid cache
evidence.

Accordingly, the 0427 protocol keeps `production_changes=false` and may retain
the scoped allocator checkpoint without cache-phase observations. Cache and
managed-budget phases remain deferred until the format owner forwards a
fail-closed diagnostic result, or another reviewed route proves the same
failure behavior and source ownership. This is an availability and evidence
boundary, not an assertion that the OPC owner has no counters.

The missing evidence for a retention/cache conclusion is therefore a
phase-aware ownership record that keeps these scopes separate:

| Missing fact | Required artifact |
| --- | --- |
| Which source, plan, publication, commit, and caller handles are still held at each boundary | A per-sample lifecycle journal with named ownership sets and explicit nested-scope drop points |
| What allocator and process state remains after each drop | Phase snapshots for live bytes, `region_peak_live_bytes`, allocation requests, RSS, and any independent Heaptrack evidence, with timer/setup scope recorded |
| What the source cache and managed budget retain after publication and drops | Per-phase cache hits, cold loads, bypasses, evictions, retained entries/bytes, in-flight loads, and budget reservations/usage |
| Whether near-limit behavior is bounded and fail-closed | Exact-limit, one-byte-under, pinned-handle, unpinned-eviction, and oversized-bypass rows with source-I/O and output counters |
| Whether a control/candidate comparison is comparable | Frozen protocol, source/binary/corpus/output identities, cache and budget configuration, observer revision, raw journals, and a standalone verifier/summary |

## Full retention lifecycle capture

For the full retention gate, use the existing 0424 plain and media corpora and
the same source-backed publication boundary when comparing roles. A sample
must record, in order, at least these explicit points:

1. baseline after setup and before the measured lifecycle;
2. after source-backed open and selected-part preparation;
3. after plan staging while the source and plan are both held;
4. after publication and candidate reopen while the returned result is held;
5. after dropping publication temporaries and the plan, while each declared
   caller-visible source/result handle remains held; and
6. after dropping those caller-visible handles and the package-owned cache
   owner, before process teardown.

The harness must create these boundaries with explicit scopes and capture the
snapshot before the next scope starts. A process-exit RSS or allocator value
is not a caller-drop snapshot. The journal must state which handles are
intentionally retained at every point, so a nonzero value is distinguishable
from an accidental lifetime extension.

The immediate 0427 allocator checkpoint may retain only the callback live-byte
points required by its frozen protocol. It must label those points as
absolute, scope-limited observations and leave the ownership, cache, managed
budget, and RSS portions of this section open.

For every checkpoint point, retain the operation allocator vectors and phase
deltas. The 0427 numeric checkpoint may stop there and report absolute
callback live bytes with their exact scope. A full retention conclusion also
requires whole-process RSS with its setup/teardown scope and the source/cache
diagnostics available for the selected owner. `region_peak_live_bytes` is an
operation maximum and must not be presented as retained endpoint memory. RSS
includes work outside the operation unless the protocol proves otherwise. A
cache counter is not a byte-copy or physical-I/O counter.

If the lifecycle uses managed OPC state, the journal must also retain the
caller-visible budget diagnostics. After dropping returned handles, any
remaining reservation must equal the declared clean cache retention. After
dropping the package/cache owner, package-owned budget use must be zero. If a
row intentionally keeps a snapshot or payload pinned, its retained entry and
reservation must be declared rather than treated as a leak. Unmanaged rows
must report cache occupancy separately and must not be compared as if they had
managed budget accounting.

## Near-limit rows

The next capture should include bounded rows that exercise the same lifecycle
or its directly exercised cache owner:

- an exact managed budget that admits the required payload and records the
  expected reservation through the cache after the returned handle is
  dropped;
- a one-byte-under budget that refuses before payload I/O or output, reports
  the expected reservation failures, and retains no reservation;
- a pinned-handle row where eviction cannot detach the payload, followed by an
  explicit handle drop and a row that proves the expected clean-entry
  eviction/release;
- an oversized-payload row that takes the documented cache-bypass path and
  reports zero false cache retention; and
- one repeated publication row at the selected limit, proving that retained
  entries, bytes, reservations, and in-flight work do not grow beyond the
  configured ceilings.

Each row needs exact cache-limit and budget configuration, expected counter
values, payload-I/O/read-call deltas, output bytes, typed error or success,
and zero-output/source-identity checks on refusal. The existing 0088
one-byte-under, pinning, bypass, and package-release invariants are the
appropriate model; a successful structural smoke alone is not a retention
measurement.

## Acceptance conditions

Retain the raw per-sample journal and a machine-readable summary with hashes
for the protocol, source revision, binaries, corpus, output sink, cache/budget
configuration, and verifier. The verifier must reject missing or reordered
phase snapshots, a mismatched ownership set, inconsistent cache/budget
counters, impossible live-byte relations, source/output identity changes, and
refusal rows that performed payload I/O or wrote output.

For an A/B retention comparison, use matched fresh processes and the frozen
0424 lane/sample contract unless a new protocol explains the change. Keep
normal elapsed timing separate from allocator and retention vectors; do not
authorize a latency claim from the drop snapshots. Any resource conclusion
must preserve the full vectors, repeat drift, RSS scope, and observer
overhead, and must be withheld when phase boundaries or ownership differ.

The lifecycle must also pass the semantic output, preservation, cancellation,
source-version, and bounded-sink checks already required by the 0423/0424
source-backed path. These checks establish that the measured object lifetime
belongs to the intended operation and that a cache or budget experiment did
not change publication behavior.

These requirements follow the accepted contracts in [ADR 0003](../../../../adr/0003-snapshots-edits-and-patches.md),
[ADR 0005](../../../../adr/0005-io-memory-and-performance.md), and [ADR 0006](../../../../adr/0006-validation-security-and-compatibility.md):
snapshots remain immutable, active handles pin clean cached values, budgets
remain finite and observable, validation stays separate from mutation, and
failure remains typed and atomic. Until the phase, near-limit, and provenance
artifacts pass these conditions, the 0424 allocation/region result remains
descriptive and no post-drop or cache-retention claim is authorized.
