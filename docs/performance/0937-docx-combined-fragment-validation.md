# 0937 — combined DOCX fragment validation work reduction

The combined candidate is rejected under the frozen normal p50 safety rule.
Requested allocation bytes fall 15.16–30.45%, with fewer calls in every cell,
but one reader pair regresses 5.38% at p50. The primary latency ratio is also
below 1.05. All four production changes are restored; the evidence is retained.

This comparison evaluates private inventory parse reuse together with
borrowed transient footnote-parser state. The separate candidates in
[0935](0935-docx-inventory-fragment-parse-reuse.md) and
[0936](0936-docx-borrowed-footnote-scanner.md) failed their prospective
improvement gates and were not adopted. This is a new frozen comparison
against the retained 0933 implementation, with unchanged improvement
thresholds. No combined performance result is inferred by adding the earlier
reductions.

## Mechanism and semantic boundary

The private inventory codec entry point parses raw section topology once,
validates footnote extensions, then decodes semantic state and tolerance data.
Inventory consumes those values without constructing and discarding a public
`Section` snapshot. Within footnote validation, events and namespace resolver
state are borrowed for each iteration. Both owned resolver clones stored in
returned root metadata remain required and retained.

The combined diff changes four private implementation files. Public `Section`
construction and lazy semantic errors remain unchanged. Raw validation still
precedes footnote validation, which precedes semantic decoding. Inventory
markup-compatibility processing remains upstream of all three stages.
Namespace materialization, source offsets, reference checks, node/depth/byte
limits, exact authored bytes, publication validation and inverses remain
required. No new API, cache, dependency, unsafe code or concurrency is added.

Accepted ADRs 0001/0004 constrain API and error behavior;
0002/0010/0011/0024 constrain ownership and dependencies; 0003/0006 require
preservation and atomic publication; 0005 requires bounded resources and
measured evidence; 0008 governs verification; 0030/0032 constrain lazy state
and memoization; 0031 constrains execution budgets. These ADRs are unchanged
from the preceding reviewed batches. iWork remains outside this work.

## Prospective measurement

The normal lane fixes 48 children: six interleaved baseline/candidate pairs
for generic ReadAt, owned bytes, owned reader and short-range ReadAt, with
five warmups and 50 samples per child. A primary latency claim requires a
median paired baseline/candidate p50 ratio of at least 1.05 and a bootstrap
95% lower bound above 1.0, using 10,000 draws and seed 937.
The p50 safety check applies to every paired block on every route:
candidate/baseline p50 must not exceed 1.05. Tail, spread and RSS increases
above 5% are retained as explicit review flags.

The allocation lane fixes 24 children: 16 or 8,192 sections, shared or
independent sources, three repetitions per version, eight concurrent commits
and exact retention. Its separate adoption gate requires at least 10% fewer
construction-requested bytes and strictly fewer allocation calls in every
cell, passing correctness and resolved normal-route and memory guardrails.
The construction interval ends at `after_snapshot_drop`; same-phase absolute
and region peaks, live/retained bytes and process RSS are separate metrics.
Increases above 5% require explicit review.

Both lanes use CPU 12, terminal-first records and fixed schedules without
selective retries or replacement. The synthetic fixtures validate ordinary
section fragments and contain no `w12:footnoteColumns` extension. Their
results cannot establish extension-heavy or general real-world DOCX
performance. Normal timings cover the complete named CRUD scenario, and
allocation construction gauges include opening, inventory and commits.

## Source and correctness coverage

The baseline source is copied from the retained 0936 baseline tree. Candidate
sources differ only in the four reviewed production files. Retained 0933
candidate binaries provide the baseline after runtime-equivalence validation
against review HEAD `e3b5acdd91280b9926f94a61e5d736ad8063422e`; their historical
build provenance remains explicit. New normal and allocator builds use the
frozen combined candidate and canonical locks.

The existing eleven regressions from 0935/0936 cover the combined interaction:
raw/footnote/semantic error order, MCE preservation and stripping, inherited
namespace context and prefix restoration, limits, qualified end names,
source-buffer lifetime and authored bytes. Existing section and source-backed
tests cover publication, inverse restoration and source conflicts. No new
test is added solely to mirror the combined implementation. In particular,
an unprocessed detached `Section` is not used as an error oracle for inventory
input that has undergone MCE processing.

The [independent quality receipt](results/change-0937/quality-independent-audit-v1.json)
binds the frozen sources and validation. Candidate all-targets testing passes 1,936 tests with one ignored. Doctests
pass 78 cases with 31 ignored; warning-denied Clippy and rustdoc and
package-scoped formatting also pass. Baseline footnote testing passes all
24 cases. One complete boundary scan passes too. Sources remain unchanged
through validation. These correctness checks do not establish a performance
benefit.

## Preparation and custody

Both fresh candidate release builds complete successfully with Rust 1.95.0,
offline locked dependencies and two Cargo jobs per build. Normal compilation
unsets Rust flags; allocator compilation sets frame pointers and unsets encoded
flags. The build wrappers inherit the environment with recorded overrides;
formal measurement children use a fully specified clean environment.

The [tooling receipt](results/change-0937/comparison-tooling-validation-v1.json)
binds source listings, the exact combined patch, historical baseline
equivalence, build terminals, binary hashes, compiler metadata, plan and
pre-capture controls. The independent allocator checker validates a retained
0936 fixture and rejects all eight corruptions before new capture; only an
in-memory terminal-schema label is adapted, with original evidence unchanged.

Two preparation corrections remain explicit. The normal audit checker's
input schemas are updated consistently to 0937 before capture; its earlier
adaptation is retained externally. During source-listing script adaptation,
an old 0936 differential file is briefly regenerated at its historical path;
its archived canonical bytes are restored and the incident is recorded in
the tooling receipt. This does not change either measured source tree or
authorize modifying historical evidence.

## Results and decision

All 72 formal children pass capture validation with no retry or replacement.
Construction counters are identical across all three repetitions in each
allocator cell:

| Sections | Ownership | Baseline requested bytes | Candidate requested bytes | Reduction | Baseline calls | Candidate calls |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 16 | Shared | 2,108,662 | 1,628,830 | 22.7553% | 28,398 | 21,302 |
| 16 | Independent | 5,607,808 | 4,757,592 | 15.1613% | 47,928 | 35,344 |
| 8,192 | Shared | 800,608,673 | 556,786,553 | 30.4546% | 11,875,584 | 8,262,872 |
| 8,192 | Independent | 1,425,467,600 | 992,008,872 | 30.4082% | 21,109,592 | 14,687,024 |

The allocation improvement gate passes in every cell. Absolute same-phase
peaks, region peaks and allocator-process RSS have no increases above the
5% review threshold. Requested-byte reductions do not establish equivalent
live-memory or RSS savings.

| Normal ingress | Baseline pooled p50, ms | Candidate pooled p50, ms | Median paired baseline/candidate p50 |
| --- | ---: | ---: | ---: |
| Generic ReadAt | 4.042082 | 4.013326 | 1.006838 |
| Owned bytes | 3.274979 | 3.245153 | 1.009801 |
| Owned reader | 3.403969 | 3.384348 | 1.007003 |
| Short-range ReadAt | 4.078152 | 4.053777 | 1.006383 |

The primary ratio's bootstrap 95% interval is [1.004961, 1.009071], but its
median misses 1.05. Pooled statistics do not replace the frozen paired safety
rule. Owned-reader block 3, position 3 has candidate/baseline p50 ratio
1.0538246832 and RSS ratio 1.0571879937. This p50 safety failure prevents
allocation-based adoption even though the allocation improvement gate passes.

Two individual p99 pairs also trigger review: short-range block 1, position 4
at 1.0526754641 and owned-reader block 6, position 2 at 1.0611420004. Eleven
pairs trigger spread flags; no individual p95 pair exceeds 5%. Every flag
remains in the [independent normal audit](results/change-0937/independent-normal-audit-v1.json).
No mean or pooled-route result is used to hide a failed individual check.

The decision is to reject this cohort and restore all four production files.
No new tests were needed or retained in this batch; the existing eleven
regressions remain. The external candidate sources and binaries stay available
for reproduction. Neither the primary latency threshold nor the per-pair
safety threshold is waived.

## Replay and retained-state verification

The [postcapture receipt](results/change-0937/comparison-postcapture-validation-v1.json)
binds the original capture terminal, authorized and packet manifests, exact
replay commands and results. Original and moved packets contain the same
521 files, totaling 4,380,460 bytes. Runner reductions agree after normalizing
only the packet path; independent normal and allocator audits also agree.

The [allocator audit](results/change-0937/independent-allocator-audit-v1.json)
validates all 72 terminals and 24 allocator reports, including fixture/cap
identity, conservation checks and eight negative controls. Its initial audit
outputs came from direct tool completions without separate command-terminal
files. An earlier compound command was rejected before checker launch and
created no output. These facts are disclosed; an additional immutable offline
revalidation writes proper terminal records and reproduces both output hashes.
No benchmark workload is rerun.

The [rejection decision](results/change-0937/rejection-decision-v1.json)
binds all four restored files and the measured candidate patch. The
[retained-source binding](results/change-0937/final-retained-source-binding-v1.json)
verifies that all 450 DOCX files equal the baseline used by the final 0936
test run: 1,936 passed, zero failed, one ignored. That earlier test result is
explicitly reused rather than represented as a new full-suite run; no code
or test changes remain in this batch. Current baseline focused tests also pass.

## Follow-up evidence

Reader p50 outliers occurred at different blocks in 0935, 0936 and 0937.
This repeated observation warrants an independently planned identical-binary
control to characterize process and route variability before evaluating more
candidates with this protocol. It does not prove that the candidate caused
no regression, and future calibration cannot retroactively qualify 0937.

The broader program remains open. Latest non-iWork reports also motivate
operation-local CPU/I/O attribution of the restored CFB repeated-edit path
after the rejected 0920–0922 candidates, and current XLSX full-cell fallback
attribution after the 0891 improvement and 0894/0896/0898 rejections. Existing
OLE2 range-source controls already exist; new work must build on that coverage
instead of recreating it from historical gap descriptions.

## Evidence retention and cleanup

The [canonical archive receipt](results/change-0937/canonical-evidence-archive-v1.json)
binds an explicit 657-file list with 10,094,416 logical input bytes. The archive
is `/home/zhuhe/litchi-performance-0937/archives/canonical-evidence-v1.tar.zst`,
748,268 bytes, SHA-256
`11ee88a534cbf202d4cc11a098cfc45b038eb0b0886fbf35726ea51fee86a82f`.
Its receipt supplies exact regeneration arguments and member/input manifests.
The [integrity check](results/change-0937/archive-integrity-verify-v1.json)
rehashes all inputs, verifies exact regular-file membership and passes zstd
validation. Sources and measured binaries remain separately retained.

[Packaging history](results/change-0937/archive-preparation-history-v1.json)
records a rejected preflight, an interrupted overbroad archive attempt, and a
verifier correction for the intentionally included compact-evidence snapshot.
The corrected archive excludes boundary sources, targets and binary
directories. Interrupted partial-output identities are retained; redundant
partial archives are removed. Later verification and cleanup records remain
external with committed descriptors, rather than being silently inserted into
the frozen archive.

[Cleanup](results/change-0937/cleanup-v1.json) removes 16,707,487,030 logical
bytes of owned build intermediates and rehashes the retained binaries. This
is a file-length total, not measured physical disk recovery. The shared
checkout build cache, measured sources, raw evidence and unrelated work remain
untouched.
