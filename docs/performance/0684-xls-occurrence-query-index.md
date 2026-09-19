# 0684 — bounded XLS occurrence query indexes

Status: retained after correctness, paired measurement and independent source review. `performance_claim: none`;
no registry or CRUD coverage promotion. OLE2 and OOXML remain active; iWork is excluded.

## Mechanism and semantic boundary

This batch implements the worksheet locator cache proposed by
[0678](0678-xls-query-cache-design.md). A first selected query scans normally.
The second successful scan may publish an occurrence index; subsequent queries
replay the selected coordinate's occurrences in source order. The index retains
locators, not decoded cell values. Repeated coordinates cannot be collapsed:
an earlier SST or formula error must still refuse even when a later occurrence
would otherwise overwrite its value. Typed cell errors retain the existing
last-occurrence behavior.

Admission requires complete worksheet scanning and final source/execution
checks. Replay preserves selected SST decoding, formula STRING/CONTINUE grammar,
XF checks, and both leading and trailing fences, including missing targets.
Failed, stale, cancelled and partial candidates are discarded. Optional cache
allocation/admission failure falls back to the ordinary scan. Visitors and text
extraction keep their existing scan path. The mandatory SST catalog and CFB
chain-hint lifetime are unchanged.

`SourceBackedLimits::max_query_index_bytes` defaults to 2 MiB; zero disables
indexing. Cloned workbook handles share the snapshot-local cache; newly opened
owners do not. The ceiling covers the local cache tables, retained indexes and
concurrent candidate weights. Slot vector capacity, rather than just length,
is charged at 24 bytes per slot, with 128-byte entry overhead and reservation
vector capacity. Geometric growth can prevent admission even when live slots
alone would fit. This is a logical weight bound, not an allocator or RSS bound.

Managed candidate reservations follow the retained `Arc` through its last pin.
LRU eviction skips pinned indexes, and duplicate builders release the losing
candidate. Cache locks protect bookkeeping only. Incremental admission may evict
clean entries before a candidate later fails; no publication guarantee depends
on retaining those entries. Opening has no execution-context parameter; its
optional table storage is fallible and locally bounded. Ordinary small `Arc`
allocations retain Rust's existing allocation-failure behavior.

## Evidence and remaining scope

The [packet](results/change-0684/README.md) binds baseline
`5805d54a1d8435e1bee96cde1303310703440d16`, source/probes, constraints and all
126 checked-in XLS fixtures. Native timings use core owned/file providers;
logical I/O counts use separate wrappers. Allocation gauges retain the owner
and returned value through capture. External counters and RSS include process
setup; repeated-query differences are separately reported.

No physical cold-cache, remote latency, cross-platform, or concurrent throughput
claim is made. Snapshot chain hints and full SST catalog scan work remain
separate opportunities.

## Paired timing and measured regression review

Release builds used Rust 1.95.0 on AMD EPYC 9R45, pinned to CPU 12 on a shared
Linux host. Each leg uses 30 fresh-owner samples after three warmups. Baseline
A/A controls precede final A/B/B/A windows; each prepared owner performs first
query A, construction-triggering query B, then repeated A. Formatting/digests
are outside query timers. The sums below exclude probe diagnostic work and
are not whole-process elapsed time. p50/mean/p95/p99 and both paired deltas
are retained for every route; p99 of 30 samples is only the sample maximum.
Paired windows are a repeatability check, not a cross-machine confidence bound.

| Stored target, owned input | first query A→B, µs | build query A→B, µs | warm query A→B, µs | open + three queries, paired p50 delta |
| --- | ---: | ---: | ---: | ---: |
| 54016 | 369.027→348.612 | 368.541→704.558 | 369.277→2.765 | −4.20% / −4.69% |
| 45365-2 | 63.295→64.740 | 62.706→143.305 | 63.166→1.055 | +8.98% / +9.60% |
| WithCustomViews, Plan1 | 24.015→24.395 | 23.435→41.485 | 23.980→1.270 | −4.29% / −5.27% |
| 15228 | 3.800→3.780 | 3.760→5.110 | 3.620→0.850 | −2.61% / −1.45% |
| Simple | 0.775→0.810 | 0.720→0.845 | 0.700→0.530 | −0.05% / −0.96% |

A/B in this table means baseline/candidate, not the queried coordinate.
File-backed warm stored queries improve 98.83–98.85% on 54016,
96.09–96.12% on 45365-2, and 89.22–89.68% on Plan1. Those are prepared-query
observations, not equal gains for open/edit/save or a general 10× claim.

The initial candidate regressed first queries about 6–22% on larger fixtures.
Its full evidence and source remain in `initial-candidate/`. Final code uses
separate constant-specialized collecting and ordinary sinks, so the ordinary
scan eliminates per-occurrence optional-collector branches. No final open,
first-query or visitor metric has a p50 regression above 5% in both paired
windows. The largest control drift is 11.11% on the tiny missing indexed query;
the full report preserves control and single-window variation.

The following regressions are retained explicitly:

- Index construction increases the second query about 23–129% across the
  medium/large stored cases and makes two-query workflows slower. 45365-2
  remains 6.60–9.60% slower including open plus three queries across both
  providers; the measured phases predict payback at a fourth similar query,
  rather than proving all three-query workloads improve.
- Simple's file-backed warm hit grows 5.42–8.91% (about 0.11–0.18 µs).
  Its source reads increase from two to three even though bytes fall from
  323 to 26: separate indexed header/payload reads cost more than this tiny
  sheet's complete buffered scan. No universal tiny-file benefit is claimed.
- Repeated selected formula refusals cannot publish an index. Their second
  and third queries pay optional collection overhead, about 16–22% in these
  paired phases. The error identity remains exact; no error result is cached.
- Too-small nonzero budgets can retry and abandon collection on subsequent
  queries. The 1 MiB 54016 control still reads the complete sheet on all five
  calls; zero avoids this attempted-build cost entirely. Retry policy and
  tiny-sheet admission remain opportunities, not hidden amortized wins.

These are accepted tradeoffs for bounded reuse on repeatedly queried sheets,
with an explicit zero-budget opt-out. Retention is justified by representative
repeated-query savings and hardware work reduction, not by the three-query
aggregate alone.

## Allocation, logical I/O, hardware and retention

Three allocation runs per route agree exactly. The counting allocator observes
one operation with setup outside its interval and owner/result alive through
the gauges. Native second-query targets differ from the allocation probe's
same-coordinate preparation, so their phase boundaries are not interchangeable.

On owned 54016, first-query allocations are unchanged: 15 calls / 130,658
requested bytes / 65,560 peak logical live-byte delta. The second query rises
to 30 calls / 3,276,274 requested bytes / 1,638,424 peak delta, retaining
1,572,948 bytes including the four-byte returned string. The third query uses
eight calls / 102 requested bytes / 102 peak delta and retains only its
four-byte returned string *in addition to the already-retained index*.
Plan1's build retains 98,435 bytes including its 51-byte returned string;
Simple's build retains 281 including nine returned bytes. Reallocation request
bytes are counted in full; logical live deltas exclude internal copy overlap,
allocator headers and process RSS.

The counted 54016 stored query falls from 26 reads / 615,829 bytes to three
reads / 21 bytes after admission; its missing query falls from 25 reads /
615,822 bytes to zero reads. Plan1 stored falls from eight reads / 40,603 bytes
to three / 68. Freshness checks remain. Across 24 candidate-only controls
(three fixtures × stored/missing × 0/1 byte/1 MiB/2 MiB), all 120 outcomes agree;
1 MiB admits Plan1/Simple but not 54016, while 2 MiB admits all three.

Native repeat diagnostics subtract 10-query process counters from 1,010-query
counters, repeated three times. Median per-extra-query instructions fall
99.45% owned / 99.29% file on 54016 and 96.36% / 94.15% on Plan1.
Cycles fall 99.45% / 98.93% and 95.97% / 90.17%, respectively. Subtracted
low-frequency counters can be noisy or negative; no percentage is reported
for negative estimates. Simple's file cycles rise 4.56%, consistent with its
small I/O cost. These diagnostics include process setup before subtraction.

Process peak RSS **increases**: for N=1,010, 54016 owned 4,348→4,896 KiB,
file 3,288→4,000 KiB; Plan1 owned 2,756→3,028 KiB; Simple file
2,500→2,736 KiB. These exceed the 5% review threshold and are accepted as
bounded retention/host variation, not memory improvements. Three process
samples and an allocator delta do not establish a general RSS ceiling.

Fresh profiles use the same repeat probe but different iteration counts for
adequate samples: baseline 10,000 and candidate 1,000,000 queries. Baseline
`WorksheetScan::next_frame` accounts for 24.16% self samples; candidate
`litchi_cfb::shared::next_chain_sector` accounts for 71.31%. Raw profile text,
commands and bindings remain. Absolute profile event totals are not compared
across unequal iteration counts. This supports snapshot chain hints as the
next measured warm-query residue, not permission to bypass CFB validation.

## Validation

Final checks pass: workspace formatting, XLS all-target/all-feature check,
warning-denied Clippy and rustdoc, 1,455 XLS tests (one existing ignored), and
61 facade tests. The added coverage includes 13 private cache tests and 17
integration tests for duplicate/error order, formula continuation, local and
hierarchical budgets, pinned retention, concurrent publication, cancellation,
source change and missing-target fences.

Both source modes exactly match the baseline corpus JSON: 126 fixtures,
119 opened and seven refused, 364 sheet-walk attempts (349 successful,
15 refused), and 4,027 selected/repeated query checks per mode with no mismatch.
This is an admitted fixture differential, not support for every XLS file.
Claim, coverage, non-iWork and crate-boundary gates pass. No registry claim,
coverage promotion, executor, ambient source or production unsafe code is added.
