# 0511: CFB FAT sector-batched entry reservation fast path

The CFB/OLE2 FAT loader's entry loop now extends one complete sector's decoded
entries at a time after its existing exact fallible reservation. The candidate
changes only the inner FAT-table insertion in `OleFile::load_fat` to
`fat.extend(sector_data.as_chunks::<4>().0.iter().copied().map(u32::from_le_bytes))`;
the generic `try_push` helper remains in place for data-dependent collections.
The exact count checks, source reads, entry decoding, marker validation,
ownership checks, limits and publication order remain unchanged.

This bounded optimization is retained after matched measurements and correctness
gates. The sector-sized extension relies on the local capacity proof below.
Larger mandatory validation loops remain unchanged and require independent
profiling and proof before any future optimization. No public API, dependency,
archive, source-ownership, or ODF/iWork behavior changes.

## Local capacity and error-order proof

Before the FAT entry loop, the loader checks that the DIFAT-derived
`fat_sectors` list has exactly the declared count. It then computes
`fat_entry_count` with checked arithmetic and performs the existing exact,
fallible `try_vec_with_capacity` reservation. CFB version and sector geometry
allow only 512- or 4,096-byte sectors, each contributing exactly
`sector_size / 4` complete `u32` entries. Every item yielded by
`as_chunks::<4>().0` is an entire `&[u8; 4]`, and `u32::from_le_bytes` decodes
it without a truncation branch. The outer and inner loops therefore extend
exactly `fat_entry_count` entries, and every sector-sized extension fits in
the reserved capacity without growth.

The reservation retains its `"FAT entries"` resource label and its existing
typed allocation failure at the same point, before the first FAT-sector read.
Each sector is still read in the same order, and its four-byte chunks are
decoded little-endian before extension. The old `read_u32_le` short-chunk
error is unreachable in this loop because `as_chunks::<4>().0` yields only
full chunks; no malformed-input acceptance boundary is removed. FAT/DIFAT
marker checks still run after decoding. Truncated final-sector zero-fill
follows the existing sector-read path. The FAT table is published only after
the same checks succeed.

This proof depends on the exact sector-list count check and reservation staying
adjacent to the loop. A future change that adds an insertion, changes sector
geometry, or weakens the count proof must restore a fallible growth check. The
MiniFAT loop, directory vectors, sector-location vectors and generic exact
chain helper remain on `try_push`; their call sites have different data-dependent
or error-order obligations and are outside this experiment.

## Fresh scoped profile

The fresh control profile totals **15,127,968 Callgrind instruction
references** over exactly five `SourceBackedWorkbook::from_read_at_with_limits`
calls. Collection is toggled only for that constructor body; selected-cell
query, setup, drop and report work are outside the toggle. `load_fat` accounts
for **9.30% exclusive** references in this fixed owned XLS source-open case.
This is operation attribution for selecting the candidate, not a before/after
latency or speedup result.

The same control shows the larger validation loops that remain mandatory:

| Constructor descendant | Exclusive share |
| --- | ---: |
| `SectorChainScratch::collect_exact` | 37.03% |
| `OleFile::claim_sector` | 16.46% |
| `validate_physical_sector_layout` | 13.17% |
| `validate_stream_allocations` | 13.08% |
| `load_fat` | 9.30% |

The chain collector, sector ownership checks, physical-layout reconciliation
and stream-allocation validation perform required cycle, marker, overlap,
length and limit checks. They are unchanged by this candidate. Their larger
shares do not authorize removing or combining those walks without a new
bounded representation, exact error-order proof and separate experiment.
Callgrind references are simulated instructions and exclude the selected
query's full timer region; they do not establish physical I/O cost or broad
XLS/OLE2 performance.

### Initial push-only rejection

The first candidate changed the per-entry `try_push` to `push` while retaining
the same reservation. Its scoped Callgrind profile increased from 15,127,968
to **15,462,114** instruction references (**+2.21%**), and `load_fat`
exclusive references increased from **1,407,020** to **1,740,600**. That
push-only candidate is rejected and archived under
`docs/performance/results/change-0511/initial-push/`; it has no formal native
measurements and is not a retained speedup or regression result. The archive
is relevant compiler-work evidence for selecting the final sector-batched
extension.

## Matched evidence protocol

The native capture uses the unchanged source-backed and owned-source harness
on CPU 2, with serial ABBA lanes (`before/r1`, `after/r1`, `after/r2`,
`before/r2`). The retained rows and sample counts are:

| Lane | Rows and samples | Boundary |
| --- | --- | --- |
| XLS lifecycle | 9 rows × 1,000 samples × 4 lanes | eager open/list/cell, source-backed open/list/cell, and owned-source open/list/cell |
| CFB guard | 2 incompressible shapes × 1,000 samples × 4 lanes | tiny and few-large `cfb_open`, matching the retained 0413 guard workload |
| Allocator diagnostic | 9 rows × 30 samples × 4 lanes | instrumented allocation deltas; excluded from native timing |

The normal XLS rows keep archive/layout construction, names/value oracles,
post-operation source checks and digest/report work outside their operation clocks according
to the runner family. The CFB guard clocks `OleFile::open` and keeps its
file-size oracle outside. The allocator lane is a separate instrumented
binary, so its timings cannot be pooled with the native rows. Exact corpus
identity, source bindings, output/oracle values and adverse thresholds remain
validated by the retained replay.

The measured allocation behavior is unchanged: the candidate removes the
per-entry reserve checks and decodes each complete sector through one
sector-sized extension after capacity has already been reserved. It does not
remove an allocation or change retained capacity. No allocation-count, allocation-bytes, RSS, or memory-reduction claim is made.

## Correctness and architecture boundaries

The source and capacity reviews find no source-level blocker. The direct
sector extension does not add an infallible growth path on the proven FAT
loop. Header/DIFAT count conversion, physical bounds checks, source version fences,
`ReadAt` boundaries, malformed metadata handling and typed errors remain in
their existing order. The candidate changes no CFB stream layout, FAT values,
directory topology, MiniFAT behavior or workbook consumer API.

The CFB/XLS test surface includes hostile count and DIFAT cases, marker and
overlap rejection, cycle and terminal-marker checks, fragmented FAT and
MiniFAT reads, exact boundary behavior, truncated-sector zero-fill, source
I/O/version fences, and writer-produced reopen cases. The reviews inspected
these requirements but did not run builds or tests. The completed CFB/XLS/DOC/PPT feature matrix and performance guardrails
are recorded below.

## Matched native results

All 44,000 native durations are retained with matching case/corpus and output
identities. The eager XLS medians improve 14.87–17.25%, tiny CFB improves
8.85–11.99%, and few-large CFB improves 45.78–46.30%. Plain positional-source
results are smaller and mixed, including R2 open +1.35% and one-cell +0.52%.
No latency, throughput or whole-child RSS comparison crosses the 5% adverse
review threshold, and no same-role repeat crosses its drift ceiling. This
supports the eager and container-opening improvement, not a blanket claim
about every source workflow or tail metric.

Each native row uses 1,000 samples after 20 warmups, on CPU 2 in serial ABBA
order. The shared host is not an isolated machine. Comparisons are within
one runner family, corpus and build protocol; there is no cross-era or
cross-family ranking. Instrumented source callbacks execute inside their
source-backed operation clocks; their post-operation counter checks are
outside. Plain OwnedSource omits that observer.

| Repeat | Case / CFB shape | Control p50 ns | Candidate p50 ns | p50 | Mean | p95 | p99 | Throughput |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| r1 | `xls_semantic_open` | 502,707 | 423,212 | -15.81% | -15.85% | -14.62% | -14.27% | +18.83% |
| r1 | `xls_eager_open_list_worksheets` | 510,322 | 422,292 | -17.25% | -17.34% | -16.95% | -17.16% | +20.97% |
| r1 | `xls_eager_open_one_cell` | 504,842 | 427,782 | -15.26% | -15.21% | -13.24% | -11.90% | +17.95% |
| r1 | `xls_source_backed_open` | 136,755 | 134,811 | -1.42% | -1.74% | -0.87% | +0.07% | +1.78% |
| r1 | `xls_source_backed_open_list_worksheets` | 137,326 | 131,090 | -4.54% | -4.76% | -3.28% | -2.36% | +5.00% |
| r1 | `xls_source_backed_open_one_cell` | 138,425 | 137,145 | -0.92% | -0.46% | +1.69% | +2.01% | +0.46% |
| r1 | `xls_owned_source_open` | 127,390 | 123,345 | -3.18% | -3.94% | -3.06% | -2.22% | +4.11% |
| r1 | `xls_owned_source_open_list_worksheets` | 124,970 | 124,650 | -0.26% | -0.59% | +1.09% | +1.83% | +0.59% |
| r1 | `xls_owned_source_open_one_cell` | 128,055 | 124,625 | -2.68% | -2.33% | -3.42% | -1.78% | +2.38% |
| r2 | `xls_semantic_open` | 501,107 | 423,022 | -15.58% | -15.54% | -14.52% | -14.22% | +18.39% |
| r2 | `xls_eager_open_list_worksheets` | 504,287 | 429,312 | -14.87% | -15.03% | -14.29% | -13.66% | +17.69% |
| r2 | `xls_eager_open_one_cell` | 503,212 | 423,481 | -15.84% | -16.18% | -16.70% | -16.87% | +19.30% |
| r2 | `xls_source_backed_open` | 136,276 | 132,065 | -3.09% | -3.57% | -3.12% | -2.57% | +3.70% |
| r2 | `xls_source_backed_open_list_worksheets` | 134,800 | 132,060 | -2.03% | -2.35% | -1.43% | -1.06% | +2.40% |
| r2 | `xls_source_backed_open_one_cell` | 138,830 | 137,790 | -0.75% | -1.33% | +0.32% | +0.77% | +1.34% |
| r2 | `xls_owned_source_open` | 124,115 | 125,786 | +1.35% | +1.43% | +1.16% | -0.79% | -1.41% |
| r2 | `xls_owned_source_open_list_worksheets` | 126,090 | 121,670 | -3.51% | -3.80% | -4.02% | -2.64% | +3.95% |
| r2 | `xls_owned_source_open_one_cell` | 126,960 | 127,615 | +0.52% | +0.32% | +0.55% | +0.49% | -0.32% |
| r1 | `cfb_open/tiny` | 2,670 | 2,350 | -11.99% | -12.57% | -12.09% | -11.96% | +14.37% |
| r1 | `cfb_open/few-large` | 169,101 | 90,810 | -46.30% | -46.27% | -44.73% | -43.80% | +86.12% |
| r2 | `cfb_open/tiny` | 2,600 | 2,370 | -8.85% | -14.65% | -8.30% | -8.89% | +17.17% |
| r2 | `cfb_open/few-large` | 168,761 | 91,510 | -45.78% | -45.75% | -44.07% | -43.74% | +84.32% |

Whole-child XLS RSS changes 145,512→145,508 KiB and 145,240→145,544 KiB;
CFB RSS changes 49,152→49,400 KiB and 49,160→49,216 KiB. These peaks include
fixture generation, copies, oracles and all selected cases. No peak-memory
improvement follows from these small differences.

## Final instruction attribution and hardware limits

The final source-open profile reduces 15,127,968→13,973,601 simulated
references (−7.63%). `load_fat` exclusive references fall
1,407,020→250,820 (−82.17%). Collection covers exactly five source constructor
calls and excludes their selected-cell queries, setup and drops.

A supplemental profile investigates the larger native CFB gain. It collects
one few-large guard runner with five constructor calls, excluding fixture
generation but including timers, file-size checks, drops and result building.
Runner references fall 22,681,105→13,081,195 (−42.33%); the five constructors'
inclusive references alone fall 22,672,032→13,072,122. The control legacy path
has 3,814,965 exclusive references in `try_push` and 2,682,000 in
`read_u32_le`, falling to 2,485 and 29,840. The profile therefore exposes
per-entry helper work that is much larger in this legacy path than in the
positional-source specialization. Larger chain and ownership checks retain
their exact reference counts. These are diagnostic instruction references,
not hardware cycles or a new source-latency percentage.

Four whole-child `perf stat` captures each run 1,000 plain OwnedSource
open/one-cell operations without warmups. The cycles/instructions/branches/
branch-misses group has identical event runtime and 100% running time in each
capture. Instructions decrease 4.22%/4.14% and branches 5.76%/5.69%, but cycles
increase 0.46%/3.64%. No cycle improvement is claimed. Setup, full input
clones, selected queries, oracles, drops and JSON reporting remain inside
this whole-child scope; these are not operation-local counters or IPC.
Cache events are omitted because the earlier broad probe was unreliable.

## Allocation, checks and custody

All 1,080 allocator samples retain identical operation allocation/reallocation/
deallocation counts, failures and allocated/deallocated bytes across all four
captures. Eager open and one-cell each allocate 1,015,869 bytes in 7,260 calls
with 307 reallocations; eager listing adds three calls and 113 bytes. Tracked
and plain source open each allocate 223,742 bytes in 124 calls with 25
reallocations; listing adds three calls and 65 bytes, and one-cell adds two
calls and 32 bytes. There are no failed allocations. All 18 paired incremental
region-peak ranges are unchanged. These are operation deltas and incremental
demand, not standalone document or process peaks. CFB guard allocation
metrics are unavailable, not zero.

The allocator captures ran from frozen binaries while the all-feature test
compilation ran. Their instrumented elapsed times and RSS are excluded;
native and hardware captures completed before that compilation started.
[Allocation replay](../results/change-0511/allocation-summary.json) retains the
full vectors' comparison and scope.

Tracked source open/list retain 334 logical reads and 138,459 bytes per sample;
one-cell retains 362 reads and 138,593 bytes. All ordinary opaque-payload read
counts and bytes remain zero. These are caller-observed logical ranges, not
physical device I/O.

The all-feature CFB/XLS/DOC/PPT all-target run passes 4,026 tests with 5
ignored; its doctests pass 41 with 22 ignored. CFB without default features
passes 305 with 1 ignored. That is **4,372 passing test executions and 28
ignored**, including repeated CFB coverage across feature configurations,
not a count of unique tests. Formatting, strict CFB clippy and rustdoc, crate
boundaries, ten strict registered claims and report classification all pass.
The unchanged harness is exercised by each captured case's correctness oracle;
no new harness-unit-test run is claimed. Source-bound commands, logs and
counts are retained in [gates.json](../results/change-0511/gates.json).
The [replay summary](../results/change-0511/summary.json) validates raw vectors,
corpus identities, source reads, build/artifact hashes and the negative short-
vector probe. The [acceptance decision](../results/change-0511/acceptance.json)
retains the native gains with the mixed source and hardware results explicit.

Control base: `b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7`.
Normal binary SHA-256 values are
`6e25e32d926756354de400bd089ea3ecdb214e338bbf8831d452ad7305407eab` (control) and
`575edf83f86566350faf206b7a0449f8b4950938f4bbed40be58d831c6beccc3` (candidate).
The final source manifest is
`4a0317050b0a28390a307b89321ad65d76d42a2a6d5619b6a5ba778db8378632`.
Only `crates/litchi-cfb/src/file.rs` changes in production. All 29 accepted
ADRs and their README remain hash-identical to the previously read versions.
Separate allocator builds bind the same source manifests. Raw samples,
source/binary receipts, profiles, rejected-variant evidence and replay tools
are retained under [change-0511](../results/change-0511/).

The [cleanup receipt](../results/change-0511/cleanup.json) records removal of
only `/tmp/litchi-goal-0511`: 5,878 files and
2,795,122,688 allocated bytes across unique file inodes.
No process referenced that directory before removal. Frozen executables and
target files were removed; raw samples, build/source identities, logs and
profiles remain. The replay succeeds without scratch artifacts;
[verification.json](../results/change-0511/verification.json) and
[SHA256SUMS](../results/change-0511/SHA256SUMS) seal the retained evidence.

## Priority after this batch

The requested priority after the preceding ODF work is OLE2 and OOXML. This
CFB/OLE2 FAT experiment is part of that priority. Further ODF optimization is
deferred until the OLE2/OOXML optimization goal is complete.

The retained [profile-scope review](../results/change-0511/profile-scope-review.md),
[CFB source review](../results/change-0511/cfb-source-review.md),
[FAT reservation review](../results/change-0511/fat-reservation-review.md),
[admission record](../results/change-0511/admission.json),
[profile plan](../results/change-0511/plan.json), and [control Callgrind
receipt](../results/change-0511/before/callgrind-receipt.json) provide the
investigation history and scope limits. The final
[sector capacity review](../results/change-0511/sector-batch-review.md) and
[final evidence review](../results/change-0511/final-review.md) cover the retained variant.
