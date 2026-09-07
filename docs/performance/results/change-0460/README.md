# ODP shared staging traversal

The fused staging candidate is **accepted and retained**. Sharing XML tokenization
and namespace maintenance between metadata and source-fragment scans reduces
normal medium/large ordinary append p50 latency by 5.068–6.085% on this host.
All four targets pass the predeclared 3% improvement gate with bootstrap upper
bounds below zero. All individual results remain in [summary.json](summary.json).

## Change and correctness

Transaction setup now supplies one reader/event stream to the existing settings,
declaration, page-metadata and source-fragment state machines. It removes one
complete XML tokenization/namespace traversal; it does not remove validation or
source retention. Source structural errors are deferred behind metadata errors,
then applied after MIME/styles setup. The independent old scanners remain as
test references. Direct source parsing retains its borrowed reader path.

The candidate passes 371 ODP tests, 387 harness tests (one ignored), strict
all-target Clippy, warning-denied rustdoc and scoped formatting. Differential
checks cover namespace rebinding, malformed input, settings/declaration/page
error precedence, later XML errors, limits, BOM prologue/page bytes, and five
native fixture source projections. Final candidate readback, compact/raw XML,
media, no-op, reversible patch and source mismatch checks remain intact. See
[source review](source-review.md) and [integration notes](integration-notes.md).

## Primary measurement

The unchanged harness measures owned ordinary `Snapshot` → transaction → add →
commit → sequential sink publication. Each variant has twelve reports: normal
and allocator binaries, 64/4,096/8,192 source slides, two repeats, three warmups
and thirty retained samples. All 24 reports / 720 operations pass the independent
oracle. A1/B1/B2/A2 ordering, one workload at a time, CPU 2, Rust 1.98.1 and equal
release flags are frozen in [protocol.json](protocol.json). Baseline revision is
`c16fd8cb5`; both compiled source manifests and exact executable hashes are bound.

| Repeat | Shape | Baseline p50 ms | Candidate p50 ms | Delta | 95% median-delta interval |
|---|---|---:|---:|---:|---:|
| R1 | tiny | 1.903 | 1.835 | -3.543% | [-3.848%, -3.320%] |
| R1 | medium | 75.720 | 71.113 | -6.085% | [-6.408%, -5.885%] |
| R1 | large | 151.613 | 143.315 | -5.473% | [-5.660%, -5.290%] |
| R2 | large | 152.225 | 143.355 | -5.827% | [-5.951%, -5.551%] |
| R2 | medium | 75.248 | 71.435 | -5.068% | [-5.281%, -4.859%] |
| R2 | tiny | 1.905 | 1.832 | -3.826% | [-4.112%, -3.695%] |

All normal p95/p99 comparisons improve. Allocator p50 deltas range from -7.320%
to -4.131%. There are no >5% adverse elapsed/RSS flags and no allocation increase
flags. Every allocator lane saves a fixed 4,642 requested bytes, 16 allocation
calls, 12 reallocation calls and 4 deallocation calls. Regional peak above entry
and retained live bytes are exactly unchanged. Process-lifetime maximum RSS
varies from -1.707% to +3.636%; no RSS reduction is claimed.

Quantiles use midpoint median and nearest-rank p95/p99. Confidence intervals use
10,000 independent median-ratio bootstrap resamples with recorded seeds. They
do not eliminate run-order effects or correct for multiple comparisons. Tiny
results are reported but are not gate targets. The decision reviews individual
lanes without a geometric mean. See [decision.json](decision.json).

## Mechanism diagnostics

Four separate normal-large phase reports retain 120 operations. Transaction
setup p50 falls 24.268% / 24.070% in R1/R2, consistent with removing a repeated
scan. Commit p50 varies +1.498% / +1.328%, snapshot opening +0.341% / +0.012%,
and publication -0.114% / -0.310%. These clocks are supplementary and do not
replace the unsegmented primary matrix or establish causality for small shifts.
See [phase-summary.json](phase-summary.json).

Two separate 100-sample whole-process counter runs include fixture setup,
warmups, correctness checks and reporting. Instructions fall 6.196%, cycles
4.910%, branches 5.731% and cache misses 6.031%; branch misses rise 0.550%.
Raw event runtime and perf scaling percentages are retained. These are not
operation-only counters. No new sampled build is needed: the [0459 profile](../change-0459/diagnostic-summary.json)
identified repeated transaction scans; candidate slide readback remains a larger
independent target. This change introduces no workers, locks or concurrency claim.

## Evidence and replay

Build/check logs, source manifests, full before/after source text artifacts,
all raw reports and receipts, frozen capture drivers, derived summaries and
acceptance are retained. [source-delta.json](source-delta.json) authenticates
exactly the five changed Rust files against both compiled epochs; `.txt` copies
do not add Rust build inputs. The first compile failure and incorrect test-fixture
attempt are retained separately from the passing final-source gates.

Run `python3 -B verify.py --portable` in a complete copied bundle to authenticate
its seal, all build/source/binary/argv/row identities and recompute both summaries.
Before cleanup, `--precleanup` additionally checks retained executable files and
the accepted current source epoch. `negative.py` verifies that changing a summary
and resealing still fails recomputation. `finalize.py` records precleanup and a
fresh-copy replay, then inventories and removes only `/tmp/litchi-goal-0460`.
`seal.py` refreshes the seal between receipt-producing steps. Final proof and
cleanup outcomes are recorded in their receipts: precleanup, fresh-copy replay
and tamper rejection pass; cleanup removes four temporary executables totaling
233,040,632 bytes.

All applicable final gates pass: owner/harness release tests, warning-denied
owner Clippy/rustdoc, scoped formatting and crate boundaries. No Office GUI roundtrip,
cold-cache, range-source or multiworker scaling result is claimed. Registry
coverage remains 439 selectors / 36 defaults, and the full non-iWork goal remains
open. Candidate readback and broader selective CRUD remain follow-up work.
