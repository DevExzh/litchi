# 0427: explicit PPTX allocator drop checkpoints

The established PPTX lifecycle allocator brackets end while caller handles are
still alive. A separate allocator-only `retention` subcommand now records
absolute callback counters at input preparation, opening, planning,
publication, and ordered drops. Owned snapshots and packages are released
before the sink. Source-backed publication consumes the editor inside the
call; the remaining view and plan are dropped before the caller source Arcs,
then the sink. The source Arc strong counts are observed while handles remain
available. No zero count is invented after their release.

The new whole-probe region includes input clones and the matched bounded sink
reservation. It has a distinct peak field and cannot replace the earlier
operation-region metric. Fixed scalar row storage is allocated before the
loop. All existing corpus correctness/refusal gates run before observations;
each iteration compares its output with independently validated exact bytes
before any drop. Reports reject unavailable, overflowed, unbalanced or
incomplete counters. Existing selectors, allocator callbacks and production
code remain unchanged.

The [frozen protocol](../results/change-0427/protocol.json) uses eight fresh
release processes on CPU 2 with one worker, 30 retained samples and three
warmups per process, in owned/source-backed/source-backed/owned order for each
plain/media-rich corpus. Source revision, source manifests, binary, protocol,
corpus and repeated output identities are bound in the
[bundle](../results/change-0427/README.md). This is a descriptive observation
of current APIs, with no before/after optimization comparison.

All eight reports validate, retaining 240 samples and 24 checked warmups. Every
final sink-drop point equals its own entry callback live-byte value. Phase
changes match within processes and across repeats, with no 5% repeat-drift
trigger. Media-rich publication deltas are 160.719714 MiB for owned and
118.309136 MiB for source-backed, with different held owner sets. The
[resource review](../results/change-0427/resource-review.md) explains the
source-caller and sink releases and retains exact bytes. These values are
current-API descriptions, not a matched memory-reduction claim. Portable replay
rejects 160 report mutations plus altered validator and repeated-output
identity probes.

Validation retains the full harness command with 287 passes, one failure and
one ignored test. The stale selector-count assertion reproduced on the exact
prior library source and passes after correction to 429. Composite coverage is
288 passes and one ignored, without counting the focused rerun twice or
claiming a new full-suite pass. Four debug CLI cases and their refusal probes
pass. Warning-denied documentation, final formatting, boundaries, CRUD index
and nine registered claim replays pass. Strict harness Clippy retains its
same 29 prior findings; no suppression is added. The failed format attempt
and selector-count runs remain available in the bundle.

ADR 0003 ownership/publication behavior is unchanged; the journal follows the
existing public owner lifetimes. ADR 0005 scope is explicit: callback live
bytes and peak counts are separate from RSS, object-owned bytes, cache gauges
and managed budgets. ADR 0006 semantic, preservation and refusal gates stay
outside observation intervals, with exact output comparisons inside them.
There is no new architectural optimization or unsafe code.

This batch makes no latency, throughput, physical-copy, cache-eviction,
managed-budget, RSS-release or leak claim. The
[next-work record](../results/change-0427/next-work.md) identifies additive
fallible PPTX cache-diagnostic forwarding and separate near-limit captures.
Native/cold/range sources, broader CRUD, bounded semantic streaming/append,
scaling and global strict-gate debt remain open. The full non-iWork performance
goal is incomplete.
