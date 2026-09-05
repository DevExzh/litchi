# Change 0412: isolate XLS read observer overhead

Date: 2026-09-05

Status: measurement enabler and plain-source diagnostic baseline; `performance_claim:
none`.

The XLS diagnostic adapter previously scanned one classification range per CFB sector on
every read and maintained a repeated-read union that XLS reports never consumed. This
change sorts each of the five classification catalogs and coalesces only valid, exactly
adjacent spans during source setup. Overlapping, duplicate, empty and reversed spans
retain their accounting semantics. It disables only the XLS repeated-read union
observer; other formats retain their existing instrumentation. A versioned
source-counter scope prevents the general performance comparator from silently treating
the two observers as identical.

Three opt-in selectors add matched plain `OwnedSource` open, open/list and open/one-cell
observations on the same deterministic comments/opaque-heavy fixture. They report
semantic digests and operation allocation observations, with source counters explicitly
not applicable. Input cloning and source setup precede both timing and allocation
regions. Workbook open and the named query are measured; validation, reporting and owner
drop follow the regions. Full instrumented locality observations remain separate.

This changes the benchmark harness only. The default 36-case/198-row contract is
unchanged; the registry grows from 422 to 425 selectors. Independent opt-in evidence
does not promote the default-bound CRUD coverage index. The broad non-iWork goal remains
open.

Control/candidate captures use the declared ABBA order, four fresh normal children
retaining 500 samples/case after 20 warmups, and four separate allocator children
retaining 30 samples after three warmups. Plain-source observations use four normal and
two allocator children at the same sample sizes. Samples within each child share its
heap. All captures run sequentially on CPU 2 with no competing build/test workload.

Operation allocation calls include successful reallocations. Peak/live counters are
process snapshots; RSS, PMU and CPU profiles include input clones, setup, warmups,
oracle checks and reporting. Native L2 request/hit events use the previously validated
host encodings; exact L1/LLC, physical I/O, cold-cache, remote input, lock wait and
scaling remain unmeasured here.

Normal p50 ranges below are the minimum–maximum across fresh child reports, in
milliseconds. Raw reports retain p95/p99, all samples and sample order;
`uncertainty.json` adds exact IID median intervals and eager review triggers. These
intervals describe within-child sampling only, excluding host drift and shared
heap/cache dependence.

| Lifecycle | Control observer p50 ms | Candidate observer p50 ms | Plain OwnedSource p50 ms | Plain allocator calls | Plain allocated bytes |
|---|---:|---:|---:|---:|---:|
| Open | 5.571505–5.609687 | 0.174466–0.174931 | 0.163961–0.165120 | 124 | 223,742 |
| Open + list | 5.574780–5.604942 | 0.174741–0.174951 | 0.164370–0.165856 | 127 | 223,807 |
| Open + one cell | 6.032486–6.072834 | 0.177676–0.179200 | 0.165870–0.168560 | 126 | 223,774 |

Both candidate instrumented allocation runs and both plain-source runs have the same
per-operation allocation-call and allocated-byte values shown above. The old observer
records 458 / 461 / 488 calls and 274,206 / 274,271 / 280,542 bytes for open / list /
one-cell. Removing its unused union eliminates one allocation per read: 334 for
open/list and 362 for one-cell. These are observer allocations; the production library
allocation path did not change.

Logical reads remain 334 calls / 138,459 bytes for open/list and 362 calls / 138,593
bytes for one-cell. Every retained instrumented observation reads zero opaque and
unselected worksheet payload bytes. The one-cell operation reads 134 selected worksheet
bytes. Corpus/output hashes, version observations and all remaining source evidence
match across the observer versions. The three eager controls remain below the
predeclared 5% p50/p95/p99/mean review trigger in both ABBA directions; this is a scoped
review result, not a production equivalence proof.

The fixed fixture remains 16,995,840 archive bytes with an 80,946-byte Workbook stream,
two sheets, eight 2 MiB opaque streams and `Untouched!E21 = 42`. Archive SHA-256 is
`6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53`; Workbook SHA-256 is
`c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041`.

All 12 focused harness tests pass after the unused-import correction; these cover range
multiplicity, locality, dedicated routing, plain-source metric absence, allocator
isolation and the 425-selector registry. All 91 comparator tests pass with Python
warnings treated as errors. Formatting, crate boundaries, the representative coverage
index, all seven existing strict claims and report classification pass. Repository
corpus-binding and operation-metric validators accept all 18 reports (37,900
observations including the separate profile/PMU runs).

The implementation is `b8f61970d`, followed by the unused-import cleanup in captured
candidate `63c95bc22d5883c8ecab0872030757e5584254f7`. Control is
`70756ae67e6763428759e8f446718ce68a528976`. Comparator scope enforcement is committed
separately as `8a49faa62`; it does not change either measured binary. The initial
candidate build warning and initial toolchain-selection failure are retained with the
final passing records.

Whole-process RSS is 144,636–145,404 KiB for the four original-family normal children
and 144,804–145,148 KiB for the four plain-source normal children. The allocator
children retain separate RSS records. Neither normal ABBA RSS direction crosses the 5%
review trigger. Three plain-source PMU repeats yield whole-process scaled IPC
1.888–1.930; hardware events are multiplexed at 83% scheduling, with software events at
100%. All raw event counts and scheduling data are retained. These measurements include
the untimed input clones and do not represent operation-only RSS, IPC or cache miss
rates.

The plain-source CPU export has 5,793 stack blocks and 24,451,673,686 weighted
`cycles:u` periods, with zero lost, malformed or empty-stack samples. Unknown frames
occur in 1.225% of weighted stack periods. Whole-process leaf weight remains 63.91%
`memmove`, including surrounding input-copy work. The observed source-open ancestor has
1,718 blocks / 30.131% of whole-process period; `SectorChainScratch::collect_exact` is
24.969% of its leaf weight and `try_push<u32>` is 24.870%. The selected-cell ancestor
has only 21 blocks / 0.369% of whole-process period, too little evidence for a
fine-grained optimization ranking. Ancestor subsets are sampled CPU attribution, not
wall-clock phases.

The next production experiment is narrow: `SectorChainScratch::collect_exact` already
reserves `expected_count` entries fallibly before walking exactly that many sectors, yet
each scratch push calls the generic fallible reserve helper again. Test removing only
that redundant scratch reservation while preserving cycle detection, marker validation,
cleanup, allocation refusal and source checks. Other generic `try_push` callers still
grow dynamically and need their existing behavior. This batch nominates the experiment
without implementing or claiming its benefit.

All source/resource observations, build identities, command journals, verifiers, CPU
exports and flame graph are retained in the [0412 evidence
bundle](../results/change-0412/). The host remains AMD EPYC 9R45 / Linux 7.0.0-1011-aws,
Rust/Cargo 1.98.1, CPU 2, standalone-workspace release builds with debug level 1, frame
pointers and unwind tables. The full non-iWork performance program remains open.

| Constraint | Application in this change | Evidence |
|---|---|---|
| ADR 0002 / 0024 ownership | All runtime changes remain in the standalone benchmark harness; production crate dependencies are unchanged. | Passing crate-boundary gate. |
| ADR 0005 I/O and measurement | Real logical reads remain separate from plain-source timing; observer versions and timer/allocation boundaries are explicit. | Matched logical vectors, independent source-free reports, comparator mismatch tests. |
| ADR 0006 preservation and bounded resources | No library validation is removed; catalog overlap multiplicity is preserved and only an unused harness union is disabled. | Differential interval tests, source locality/oracle checks. |
| ADR 0008 verification state | Descriptive opt-in evidence does not promote representative default coverage or close the program. | Passing coverage-index checker; open goal audit. |

Published replay passes after lossless compression. Six report/resource corruptions and
seven attribution corruptions are rejected with mutation-specific diagnostics. The final
attributor requires exact capture identity, positive `cycles:u` periods, complete parsed
stacks and the captured OwnedSource route; initial unverified attribution output remains
a labelled diagnostic.
