# Change 0413: reuse the exact CFB chain reservation

Date: 2026-09-05

`performance_claim: scoped`

`claim_authorized: true`

The retained change reduces plain-source XLS open/list/one-cell p50 by
1.65–3.18% across the two matched pairs; eager workflows improve 6.95–8.54%.
All nine XLS selectors pass the predeclared direction/drift policy for p50,
mean, p95 and p99. The strict registry entry is
`claim-0413-cfb-chain-scratch`, limited to this corpus, machine, builds and
warm lifecycle scope. The large-stream CFB guard is about 3% slower; its
original threshold crossing and follow-up are reviewed below.

## Mechanism and safety

The 0412 plain XLS source profile identified `SectorChainScratch::collect_exact`
and its `try_push<u32>` helper as 24.97% and 24.87% of period-weighted leaf
samples under observed source-open ancestors. The collector already reserves
the exact expected chain length fallibly before walking. This change replaces
only its per-sector `try_push` call with `Vec::push` into that reservation.
The generic dynamically growing helper remains unchanged.

`reset()` makes the vector length zero. The existing fallible reservation
establishes capacity at least `expected_count`; exactly that many iterations
each add one sector. Thus the replacement cannot allocate. The visited-map
reservation remains fallible and in the same order. Invalid start/count,
sector bounds, cycle, early/late markers, allocation ownership, physical
layout, source identity and error cleanup retain their existing checks.
No public API, dependency, unsafe code, parallelism or cache policy changes.

ADRs 0005/0006 retain bounded allocation, typed refusals and full validation.
The private CFB owner stays within ADRs 0001/0002/0010/0011/0024; immutable
source/edit/publication contracts under ADR 0003 are unaffected. Existing
scratch tests cover capacity growth before walking, successful/empty reuse,
cycle cleanup and reuse after error, and early/late termination errors.

## Measurement protocol

Control: `ceba0345220c1ca6a7f61f3fac86145b5afc55ca`.
Candidate: `bf5b7f50f` (full revision and executable hashes in build sidecars).
Both builds use Rust 1.98.1 release, debug level 1, forced frame pointers and
unwind tables, no incremental compilation, and four build jobs. Measurements
use CPU 2 on AMD EPYC 9R45, Linux 7.0.0-1011-aws, KVM, roughly 128 GiB RAM.
The guest is shared; task builds, tests and measurements are serialized.

Normal ABBA runs retain 1,000 samples after 50 warmups in each fresh process.
Separate allocator ABBA runs use 30 samples and 3 warmups; their timings are
excluded from latency claims. Nine XLS selectors share the deterministic
16,995,840-byte comments/opaque-heavy fixture, two sheets and the target
Untouched!E21 = 42.0. Plain OwnedSource and instrumented source-backed runners
have matching timer and oracle boundaries; eager workflows remain guards.
CFB guards separately open tiny MiniFAT and few-large regular FAT fixtures.

Input cloning/source setup, fixture generation and semantic oracle checks are
outside the operation timer. Process RSS and hardware counters include those
costs. The observer version is v2 in both roles. Source metrics are absent
for plain OwnedSource; separate instrumented runs establish logical locality.
No cold-cache, filesystem, remote-source, concurrency or broad-format claim
follows from these warm generated in-memory fixtures.

The predeclared policy requires both AB pairs to improve with same-role drift
at most 5% for p50/mean, 10% for p95, and 15% for p99. Any paired latency or
process RSS regression above 5% triggers individual review. Uncertainty from
shared-host drift remains distinct from within-process sampling uncertainty.

## Validation

CFB/XLS/DOC/PPT all-feature/all-target tests: 4,004 passed, 5 ignored.
CFB no-default-feature tests and doctests: 305 passed, 1 ignored. Formatting
and warning-denied CFB rustdoc pass. Rust 1.98 Clippy exposes six pre-existing
`chunks_exact_to_as_chunks` findings; a command-scoped exemption for that
single lint permits the remaining warning-denied all-feature/all-target check.
The production diff stays confined to the proven scratch reservation.

## Results and adverse-case review

All individual results remain in the ABBA packages. The following p50 values
are microseconds; percentages compare A1/B1 and A2/B2 separately. A positive
reduction is faster. Instrumented source cases retain the same v2 observer in
both roles; plain cases omit it.

| XLS selector | A1 | B1 | B2 | A2 | Reduction, pair 1 / pair 2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `xls_semantic_open` | 570.292 | 524.827 | 524.522 | 565.822 | 7.97% / 7.30% |
| `xls_source_backed_open` | 180.091 | 173.896 | 176.491 | 178.031 | 3.44% / 0.87% |
| `xls_eager_open_list_worksheets` | 567.837 | 528.362 | 526.097 | 569.807 | 6.95% / 7.67% |
| `xls_source_backed_open_list_worksheets` | 180.231 | 174.861 | 176.885 | 178.711 | 2.98% / 1.02% |
| `xls_eager_open_one_cell` | 567.092 | 518.882 | 524.807 | 573.827 | 8.50% / 8.54% |
| `xls_source_backed_open_one_cell` | 181.630 | 179.870 | 179.311 | 181.335 | 0.97% / 1.12% |
| `xls_owned_source_open` | 170.565 | 167.300 | 163.635 | 169.010 | 1.91% / 3.18% |
| `xls_owned_source_open_list_worksheets` | 170.285 | 167.475 | 164.631 | 168.651 | 1.65% / 2.38% |
| `xls_owned_source_open_one_cell` | 173.291 | 169.095 | 166.865 | 171.455 | 2.42% / 2.68% |

The few-large CFB open guard regresses 2.87%/2.76% in p50. Its original
p99 comparisons are +5.02%/+2.80%, triggering review. A separately declared
ABBA follow-up confirms a roughly 3% cost: p50 +3.04%/+3.08%, p95
+3.02%/+2.96%, and p99 +3.40%/+3.03%. The original evidence remains retained.
Tiny CFB p50 differences stay within 1.5%; its original control mean drift
exceeds the 5% ceiling, so that statistic is withheld. No CFB speedup is claimed.

This tradeoff is retained because the end-to-end XLS workflows improve in
both pairs, the change removes proven redundant work with no new machinery,
and the reproducible CFB cost remains below the 5% review threshold. The guard
remains a watch item for future shared CFB work. The profiles do not establish
a specific compiler/layout cause for its slowdown.

Original process peak RSS stays between 144,828 and 145,392 KiB for XLS
normal/allocator captures, and 82,604 to 82,744 KiB for CFB captures. All paired
RSS changes are below 0.4%. XLS allocation/deallocation/reallocation and byte
counts remain exact across roles. Plain and instrumented source open/list/cell
use 124/127/126 allocation calls and 223,742/223,807/223,774 allocated bytes.
CFB guard selectors do not expose operation-local allocation metrics, including
under the allocator binary; these values are unavailable, not zero.

Instrumented XLS open/list retains 334 logical reads / 138,459 bytes; one-cell
retains 362 reads / 138,593 bytes, including 134 selected-sheet bytes and no
opaque or unselected payload reads. These counters include defined observer
checks outside the timer; they are not physical I/O or timed-only source calls.
Source versions and output oracles remain equal. Allocation live/peak fields
are process snapshots and are reported descriptively.

## Residual profile and hardware evidence

The paired plain-source one-cell profiles retain 5,901 control and 5,806
candidate stack blocks, with zero lost, malformed, empty or unparsed blocks.
Blocks containing unknown frames account for 1.03%/0.98% of total event period.
Observed source-open ancestors account for 31.18%/30.63% of whole-process
period; their leaf `try_push<u32>` share is 26.59%/12.95%, while the exact
scratch collector remains 24.48%/26.66%. Other dynamic callers still require
fallible growth. Individual symbol shares can also change with compiler layout
and sampling; these profiles are diagnostic CPU attribution, not timed-phase
latencies or proof that all remaining helper calls are redundant.

Candidate source-open leaf samples also identify FAT loading (16.90%), stream
allocation validation (14.56%) and physical layout validation (8.16%). A future
experiment should first attribute the remaining helper callers and repeated
validation work while retaining ownership/cycle checks. Whole-process setup
and copies remain material; selected-cell-only sampling is sparse. No broad
rewrite or removal of required validation follows from these samples.

Four whole-process PMU captures use the same one-cell plain-source case, with
3,000 samples/10 warmups. Hardware events schedule at 83%, software events at
100%; counts are multiplex-scaled. IPC is approximately 1.88–1.91. Native L2
request/hit counters are retained; these are not exact L1/LLC miss rates or
operation-local IPC. No instruction, cache or scaling improvement is claimed.

## Remaining scope

The full non-iWork performance program remains open. This batch tests warm
synthetic XLS workflows and two substrate guards. It does not establish cold
or remote behavior, real-producer performance, concurrent scaling, or broad
CRUD completeness. The direct lifecycle implementations are serial; CPU
pinning yields one available CPU and no worker pool is used by these selectors.

## Evidence and tooling

The [retained bundle](../results/change-0413/README.md) contains all 26 reports
and 85,320 observations, separately scoped allocation and resource summaries,
three ABBA packages, paired profiles and replay scripts. Corpus/schema and
sample-order checks pass; seven report mutations and three profile mutations
are rejected with the expected diagnostics. The main profile parser is reused
from 0412. Every artifact and lossless compressed original is hashed.

The strict ABBA tool now understands the nine matched XLS lifecycle selectors,
their exact corpus/output oracles and v2 source contract, while preserving the
standalone semantic-open route and native numeric edit/save validation.
Sample-index permutation normalization is limited to validated lifecycle rows
on the exact corpus. Summary/package tests plus the updated seed-registry test
pass (100 tests); the other claim tests passed in the preceding full run.
The seed test previously omitted 0410 and now includes both 0410 and 0413.
Crate boundaries, CRUD coverage, strict eight-claim validation and report
classification are checked independently. No default benchmark selector or
coverage-index contract changes in this batch.
