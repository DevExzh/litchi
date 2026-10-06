# 0936 — borrowed DOCX footnote scanner

The candidate is rejected under both prospective improvement gates. Borrowing
transient XML events and namespace state reduces construction-requested bytes
by 5.48–10.96%, but both small cells miss the required 10% minimum. The primary
latency ratio is 1.002845, below 1.05. The production parser is restored; five
regression tests and the measurement evidence remain. The comparison starts
from the adopted 0933 implementation and excludes the rejected 0935 change.

## Hypothesis and constraints

[0934](0934-docx-post-scanner-allocation-profile.md) attributes the largest
remaining nearest-physical-Litchi allocation bucket to footnote validation:
1,091,634,180 requested bytes with shared sources and 1,554,744,898 with
independent sources. Those are whole-process profile buckets, not removable
allocation estimates or construction-phase measurements.

The parser reads immutable byte slices, so its event and namespace resolver
can be borrowed for each match without copying them. The candidate removes
the per-event `into_owned()` and resolver clone. It retains both owned resolver
clones stored in returned root metadata, along with owned names and opening
markup. That state must survive the reader and source-buffer lifetimes.

The change stays inside one existing parser function. Namespace resolution,
qualified end-name checks, extension parsing, error order, source offsets,
node/depth limits and authored XML remain unchanged. There is no new public
API, cache, dependency, unsafe code or concurrency. Section and footnote
validation remain separate grammar passes.

Accepted ADRs 0001/0004 constrain public behavior and typed errors;
0002/0010/0011/0024 constrain ownership and dependencies; 0003/0006 require
preservation and atomic publication; 0005 requires reproducible measurements
and bounded resources; 0008 requires verification; 0030/0032 constrain lazy
state and memoization; 0031 constrains execution budgets. iWork is outside
this batch.

## Prospective comparison

The [plan](results/change-0936/comparison-plan-v1.json) fixes 48 normal-release children: six interleaved baseline/candidate
pairs for each of generic ReadAt, owned bytes, owned reader and short-range
ReadAt, with five warmups and 50 samples per child. A primary latency claim
requires a median paired baseline/candidate p50 ratio of at least 1.05 and
bootstrap 95% lower bound above 1.0 (10,000 draws, seed 936). Route p50
regressions must remain within 5%; tails, spread and RSS have separate review
flags.

The allocation lane fixes 24 children: 16 or 8,192 sections, shared or
independent sources, three repetitions per version, eight concurrent commits
and exact retention. Its separate adoption gate requires at least 10% fewer
construction-requested bytes and strictly fewer allocation calls in every
cell, passing correctness and resolved latency/memory guardrails. Construction
deltas end at `after_snapshot_drop`. Absolute same-phase peaks and process RSS
are evaluated separately; increases above 5% require review.

Both lanes use CPU 12, immutable source/build bindings and terminal-first
records. Formal children cannot be selectively retried or replaced.

The synthetic fixtures exercise this parser while validating ordinary section
fragments; they contain no `w12:footnoteColumns` extension. The normal corpus
retains three main-story sections, including unknown `x:` children. Allocation
fixtures contain `w:pgSz` and `w:cols` in every section. They support a claim
about transient validation work on these fragments, not extension-specific or
general real-world DOCX performance. Normal timing covers the complete
open/query/edit/commit/publication scenario, and construction gauges include
opening, inventory and commits; neither isolates this parser's cost.

## Source and regression coverage

Matched external source trees contain identical test changes and differ only
in `footnote_columns/codec.rs`. Retained 0933 candidate binaries provide the
baseline after runtime-equivalence verification against current HEAD
`0eae43dd007e63df5b7f406a46bed356a080d1be`. Their historical build identity is
retained; equivalence does not imply a new build at that commit. Candidate
normal and allocator binaries are built from the frozen candidate tree.

Five new tests cover inherited namespace context with nested prefix rebinding
and restoration, qualified mismatched end names, element-count overflow,
parsed values after source-buffer drop, and package snapshot/context use after
the source part is dropped. Existing depth, preservation, publication and
inverse tests remain required. The new tests compile against both versions.

## Correctness validation

The [independent quality receipt](results/change-0936/independent-quality-audit-v1.json)
records all checks. Baseline focused testing passes all 24 footnote tests. Candidate all-targets
testing passes 1,936 tests across 73 suites, with one existing test ignored.
Doctests pass 78 cases with 31 ignored. Warning-denied Clippy and rustdoc,
package-scoped formatting, and the full boundary scan pass. The boundary scan
covers 65 packages, 244 internal dependency declarations and 11 explicit debt
items.

A redundant standalone rustfmt invocation omitted the Rust 2024 edition and
failed to parse existing let-chain syntax; its terminal is retained as a
non-verdict. The package-aware formatting check passes. An early checkout
`cargo check --locked` stopped before compilation because the checkout's
ignored root lock differs from the canonical isolated lock; isolated quality
uses the bound canonical lock instead.

## Build custody

Fresh normal and allocator release builds both complete successfully with
Rust 1.95.0, offline locked dependencies and two Cargo jobs per build. The
normal build unsets both Rust flag variables; the allocator build sets
`RUSTFLAGS=-C force-frame-pointers=yes` and unsets encoded flags. Build wrappers
inherit the process environment with recorded overrides; the formal children
use a fully specified clean environment.

The initial normal-build wrapper fails with a Python `NameError` before Cargo
starts. Its failure record remains. Successful v2 build records are retained
unchanged, and separate adapters bind their source, locks and environment to
the comparison runner's schema. The first adapter attempt omits a required
exit-code field; it is preserved as superseded derived evidence. These are
pre-capture preparation corrections, not benchmark retries.

Embedded compiler metadata is compared across all four binaries. Its current
collection revision is not the historical source revision of the retained
baseline executables; their original build receipts and separate runtime
equivalence remain authoritative.

The [final tooling receipt](results/change-0936/comparison-tooling-validation-v3.json)
binds the prospective runner, plan, manifests, source listings, builds and
compiler metadata. Pre-capture corrections remove stale seed/threshold
references and normalize receipt fields. A missing scratch directory causes
an inert control attempt to fail; the failure remains, and corrected controls
pass before authorization. No document workload runs during these controls.
The independent allocator checker also passes a historical-fixture positive
check and rejects eight corruptions before capture. That check uses retained
0935 evidence with a disclosed in-memory terminal-schema adaptation only;
it is not validation of a new 0936 packet.

## Results and decision

All 72 planned children pass capture validation, with no retry or replacement.
Each allocator cell has identical construction counters in all three
repetitions:

| Sections | Ownership | Baseline requested bytes | Candidate requested bytes | Reduction | Baseline calls | Candidate calls |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 16 | Shared | 2,108,662 | 1,934,718 | 8.2490% | 28,398 | 26,342 |
| 16 | Independent | 5,607,808 | 5,300,584 | 5.4785% | 47,928 | 44,304 |
| 8,192 | Shared | 800,608,673 | 712,869,769 | 10.9590% | 11,875,584 | 10,843,352 |
| 8,192 | Independent | 1,425,467,600 | 1,269,489,336 | 10.9423% | 21,109,592 | 19,274,544 |

Allocation calls decrease in all cells, but the gate requires at least 10%
fewer requested bytes in every cell. The two small cells fail. Named absolute
and region peaks and allocator-process RSS all remain within the 5% review
limit; the largest paired RSS ratio is 1.0282. This does not establish a live
memory saving or satisfy the failed allocation improvement gate.

| Normal ingress | Baseline pooled p50, ms | Candidate pooled p50, ms | Median paired baseline/candidate p50 |
| --- | ---: | ---: | ---: |
| Generic ReadAt | 4.046042 | 4.033346 | 1.002845 |
| Owned bytes | 3.267344 | 3.258368 | 1.003093 |
| Owned reader | 3.404659 | 3.396619 | 1.002512 |
| Short-range ReadAt | 4.083572 | 4.074862 | 1.002531 |

The primary paired ratio has bootstrap 95% interval [1.001446, 1.005200],
but its median misses 1.05. Pooled percentiles and paired statistics remain
distinct. Owned-reader block 1 also regresses 6.0264% at p50 and 8.1502% in
process RSS. Thirteen pairs trigger spread review flags; no individual p95
or p99 pair exceeds the 5% regression threshold. The reader observation is
retained, and the candidate fails the individual p50 safety check.

Neither improvement gate passes, so the one production-file change is not
adopted. The frozen external candidate sources and binaries remain available.
The checkout retains the five behavior regressions against the restored
baseline. No threshold is relaxed and no failed or noisy pair is replaced.

Both 0935 and 0936 reduce distinct allocation work but fail their individual
all-cell gates. A combined candidate may be evaluated in a separately frozen
comparison; the two measured reductions cannot simply be added because
their interaction with parsing and temporary ownership requires a fresh
end-to-end measurement.

## Independent replay and retained state

The [normal audit](results/change-0936/independent-normal-audit-v2.json)
recomputes all 48 normal reports and 24 pairs and verifies all 72 terminal
identities. Its initial version incorrectly requires zero observation polls
for allocator children too. Those longer children legitimately need repeated
observation while the same process continues. The failed audit is preserved;
v2 accepts and records those polls while retaining terminal/deadline checks.
This is an offline checker correction, not a workload retry. All normal
children have zero observation polls.

The [allocator audit](results/change-0936/independent-allocator-audit-v1.json)
validates all 24 allocator reports, same-phase peaks,
fixture identities and conservation checks, and rejects all eight negative
controls on both packet copies. Its failed allocation-gate status is preserved
separately from successful evidence validation. Canonical and moved-packet
runner results agree after normalizing only the packet path; independent
normal and allocator results also agree across copies. The
[postcapture receipt](results/change-0936/comparison-postcapture-validation-v1.json)
binds both manifests, the original outer terminal, reduction and replay
commands and their results.

The [rejection decision](results/change-0936/rejection-decision-v1.json)
binds the restored production file, retained tests and external candidate
patch. The production file matches both HEAD and the frozen baseline source.
The [final retained-state check](results/change-0936/final-retained-baseline-quality-audit-v1.json)
passes 1,936 tests with one ignored. All 450 DOCX files match the source tree
used by that run, including the five new tests.

## Evidence retention and cleanup

The [archive receipt](results/change-0936/canonical-evidence-archive-v1.json)
binds 689 explicitly listed files: capture data, original terminals, replay and
audit outputs, preparation failures, source listings, scripts and a snapshot
of compact repository evidence. The external archive is
`/home/zhuhe/litchi-performance-0936/archives/canonical-evidence-v1.tar.zst`,
803,923 bytes, SHA-256
`8e3a0da6a29784cc71b303dcc37693a3d1af160be2c09de549aa8a47c9975131`.
The receipt contains exact regeneration arguments, member/input manifests and
archive command terminals. Source trees and measured executables remain
separately retained; no executable is committed.

[Cleanup](results/change-0936/cleanup-v1.json) removes 16,706,369,219 logical
bytes of owned quality and release intermediates and rehashes both retained
candidate binaries. This is a file-length sum, not measured physical disk
recovery. The shared checkout build cache and unrelated work remain untouched.
