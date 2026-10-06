# 0935 — DOCX inventory fragment parse reuse

The candidate is rejected under the prospective gates. The small independent
ownership cell reduces requested bytes by 9.6828%, below the required 10%; the
primary latency ratio is also below the required 1.05. The regression tests and
evidence are retained, while the three production changes are not adopted.

## Hypothesis and scope

[0934](0934-docx-post-scanner-allocation-profile.md) identifies fragment codec
and footnote validation as the largest remaining allocation buckets after
borrowed scanner adoption. Source inspection shows that inventory constructs a
public `Section`, validates a raw codec parse, validates footnote extensions,
and then parses the same fragment again to obtain semantic state. Temporary
snapshot ownership and cloned state are discarded immediately afterward.

The candidate uses a private inventory codec entry point to retain one raw
parse across borrowed footnote validation, then decode semantic state and
return it with its tolerance report. The public `Section` constructor and its
lazy semantic errors remain unchanged. Footnote validation is retained; it is
an independent grammar pass, not redundant section codec work.

Namespace materialization, source spans, reference preflight, section and byte
limits, tolerance offset mapping, fresh source checks, publication validation
and inverse checks remain required. Holding raw state across footnote parsing
may change temporary peak memory; allocation traffic alone cannot justify
adoption without reviewing that effect.

## ADR constraints

| Constraint | Candidate boundary |
| --- | --- |
| 0001, 0004: typed APIs and errors | Private helper only; public construction and getters retain error timing. |
| 0002, 0010, 0011, 0024: ownership | Section and footnote grammars stay in DOCX; no new dependency. |
| 0003, 0006: preservation and atomic publication | Original source spans and unknown XML remain authoritative; existing publication and inverse validation remain. |
| 0005: measured performance and bounds | Paired normal and allocator measurements, explicit resource limits, peak-memory review. |
| 0008: verification | Differential regressions and existing DOCX, documentation, lint and boundary gates. |
| 0030, 0032: lazy state and memo policy | No new cache or memo; public lazy semantic construction is unchanged. |
| 0031: execution budgets | No new concurrency, executor, or ambient provider. |

## Prospective comparison

The [plan](results/change-0935/comparison-plan-v1.json) fixes 48 normal-release
children: six interleaved pairs for each of generic ReadAt, owned bytes, owned
reader and short-range input, with five warmups and 50 samples per child.
Generic ReadAt is primary. A latency improvement claim requires a median paired
baseline/candidate p50 ratio of at least 1.05 and a bootstrap 95% lower bound
above 1.0. Route p50 regressions must stay within 5%; tail and RSS increases
above 5% require explicit review.

A separate allocator lane fixes 24 children: three repetitions for each version
in four cells (16 or 8,192 sections, shared snapshots or independent sources,
eight live commits, exact retention). Its prospective allocation adoption gate
requires at least 10% fewer construction-requested bytes and strictly fewer
allocation calls in every cell, passing correctness and resolved normal-route
and memory guardrails. Construction deltas end at `after_snapshot_drop`.
Absolute same-phase observer peaks, live/retained bytes and process RSS are
reported separately; peak-live or RSS increases above 5% require review.

Both lanes use CPU 12 and terminal-first evidence. Failed formal children are
preserved without selective replacement. Normal-release timing is not inferred
from allocator instrumentation, and allocation traffic is not live memory.

## Source and test binding

The [source binding](results/change-0935/source-binding-v3.json) records complete
external manifests for matched baseline/candidate trees. Only three production
files differ between them. Both contain the same six new regression tests,
registered only for tests. The [test design](results/change-0935/test-design-receipt-v3.json)
covers public semantic laziness, inventory raw/footnote/semantic error order,
reference preflight, namespace-injected tolerance offsets and footnote handling
after markup-compatibility processing. Existing inventory limits, publication and inverse tests remain gates.

Retained 0933 candidate binaries provide the baseline after source-equivalence
verification. Their original build identity is retained; the current review
HEAD is not represented as a historical rebuild. The candidate gets fresh
normal and allocator builds from the bound source tree with the same toolchain,
profiles and canonical standalone lock. The checkout's different ignored root
lock is excluded from this measurement identity.

## Baseline-oracle correction

The first baseline test run passed four cases and rejected two test assumptions.
Inventory performs markup-compatibility processing before fragment parsing and
understands the Word 2010 extension namespace, not the Word 2012 footnote
namespace. Ignorable footnote markup can therefore be removed before the
fragment reaches footnote validation. Comparing that inventory error directly
to a public `Section` built from unprocessed XML was not a valid oracle.
Likewise, inherited ignorable markup did not imply the inventory refusal
suggested by the 0934 design review. The failed run is retained, and focused
tests must distinguish processed inventory input from raw detached fragments.
This observation changes the test assumptions, not production behavior.

## Validation and build history

All six corrected baseline regressions pass. Candidate all-targets testing
passes 1,931 tests with one existing corpus test ignored. Warning-denied Clippy,
78 doctests (31 ignored) and warning-denied rustdoc also pass. The [independent quality receipt](results/change-0935/independent-quality-audit-v1.json)
binds those results and the full boundary pass for 65 packages and 244 internal
dependency declarations. Full-workspace formatting still reports unrelated
existing drift; the DOCX package passes formatting. A standalone rustfmt command
used the wrong edition for recursively parsed files and is superseded by the
package check. The initial missing boundary-tool attempt and cancellation of an
optional baseline boundary scan remain explicit non-verdict records.

Build preparation failures remain explicit. The first build wrapper could not
find Cargo after restricting PATH and launched no compiler child. Corrected v2
builds completed, but their original full-source binding was invalidated when
a corrected test-only file was copied into both source trees during compilation.
No production file changed. Those original records are retained as superseded;
a new full-source freeze and terminal-backed v3 Cargo builds/checks pass. Both
resulting binary hashes equal the v2 outputs, consistent with the changed file
being excluded from release compilation by `cfg(test)`. Formal measurement
uses the final v3 bindings.

The build receipt's `env_policy` wording is broader than the wrapper behavior:
build wrappers inherit the process environment and explicitly set/unset the
listed values. Formal measurement children use the fully specified clean
environment. [Embedded compiler metadata](results/change-0935/compiler-metadata-v1.json)
matches across all four measured binaries; it does not reconstruct unrecorded
historical build environment values.

## Results and decision

All 72 planned children pass their capture validation with no retry or
replacement. Each allocator cell has identical construction counters in its
three repetitions:

| Sections | Ownership | Baseline requested bytes | Candidate requested bytes | Reduction | Baseline calls | Candidate calls |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 16 | Shared | 2,108,662 | 1,802,774 | 14.5063% | 28,398 | 23,358 |
| 16 | Independent | 5,607,808 | 5,064,816 | 9.6828% | 47,928 | 38,968 |
| 8,192 | Shared | 800,608,673 | 644,525,457 | 19.4956% | 11,875,584 | 9,295,104 |
| 8,192 | Independent | 1,425,467,600 | 1,147,987,136 | 19.4659% | 21,109,592 | 16,522,072 |

Calls decrease in every cell, but the allocation adoption gate requires at least
10% fewer requested bytes in **every** cell. The 16-section independent case
fails that criterion. Its proximity to the threshold does not change the gate.

| Normal ingress | Baseline pooled p50, ms | Candidate pooled p50, ms | Median paired baseline/candidate p50 |
| --- | ---: | ---: | ---: |
| Generic ReadAt | 4.044641 | 4.010137 | 1.008732 |
| Owned bytes | 3.271040 | 3.235044 | 1.011575 |
| Owned reader | 3.418940 | 3.367065 | 1.040024 |
| Short-range ReadAt | 4.080293 | 4.054147 | 1.007100 |

The primary ratio has bootstrap 95% interval [1.007296, 1.009646], but its
median is below 1.05, so the separate latency gate fails too. Pooled and paired
statistics are different; the larger reader median ratio does not establish a
qualified improvement on the primary route.

An individual owned-reader pair (block 2) regresses 5.1278% at p50, 5.4210% at
p95 and 5.0818% at p99. Nine pairs trigger the spread review flag. These
observations remain in the record rather than being replaced. No individual
normal RSS pair exceeds the 5% review threshold. All allocator named absolute
and region-peak comparisons and process RSS remain within the 5% guardrail.
Memory safety alone does not satisfy either improvement gate.

The final decision is to retain the behavior regressions and reject the private
parse-reuse implementation. The external candidate source and binaries remain
available for reproducing the result. Future work may investigate the still
measured footnote scanner copies, but must establish its own prospective
comparison; this failed cohort is not a basis for relaxing the threshold.

The [rejection record](results/change-0935/rejection-decision-v1.json) binds the
three restored production files and confirms that all 450 DOCX files match the
source tree used for the final retained-state test run. That
[final validation](results/change-0935/final-retained-baseline-quality-audit-v1.json)
passes 1,931 tests with one ignored, including the six new behavior cases.

## Replay and analysis corrections

The [postcapture record](results/change-0935/comparison-postcapture-validation-v1.json)
binds the original capture terminal, frozen runner replay and moved-packet
replay. Their numerical metrics agree after normalizing only the packet path.
The [independent normal audit](results/change-0935/independent-normal-audit-v1.json)
recomputes paired ratios, uncertainty and individual regression flags.

The frozen reducer retained a stale `ge_20_percent` field from 0933 even though
the prospective 0935 plan and authorized manifest require 10%. A separate
offline reduction applies the correct frozen 10% criterion; original files are
unchanged, and the small independent cell still fails. The frozen independent
allocator checker also exposed two defects during replay: an incorrect
negative-control source-revision lookup and missing expected fixture comparison.
Failed checker versions and their terminals remain preserved. Corrections are
postcapture audit work, not workload retries or altered adoption criteria.

The [corrected v3 allocator audit](results/change-0935/independent-allocator-audit-v3.json)
restores the expected source/output/cap fixture checks from the retained 0933
contract and fixes the lookup. It validates all 72 terminals and 24 allocator
reports and rejects all eight negative controls on both packet copies. Results
match after removing only the packet path; manifest and checker hashes are
identical. Validation succeeds while the allocation improvement gate remains
failed. This is disclosed postcapture verification, not a retroactive claim
that the frozen v1 checker passed.

## Evidence retention and cleanup

The [final archive receipt](results/change-0935/canonical-evidence-archive-v2.json)
binds an explicit immutable list of 703 files, including raw capture reports,
original terminals, failed and corrected offline checks, quality/build history,
source manifests and compact evidence. The archive is
`/home/zhuhe/litchi-performance-0935/archives/canonical-evidence-v2.tar.zst`,
849,980 bytes, SHA-256
`f48b5b61977f59e6af47273375c227aa7457f7586f98d0487e16f9f3508c4832`.
Its receipt provides exact regeneration arguments, file-list/input-manifest
identities and command terminals. The earlier archive remains unchanged.
Source trees and both candidate executables are retained separately at their
recorded external paths; no executable is committed.

[Cleanup](results/change-0935/cleanup-v1.json) removes 17,275,940,844 logical
bytes of owned quality and release intermediates, then rehashes the retained
candidate binaries. This is a file-length sum, not physical disk recovery.
The shared checkout build cache, external evidence and unrelated work remain
untouched.
