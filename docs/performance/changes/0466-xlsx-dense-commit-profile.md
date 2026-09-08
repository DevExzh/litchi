# Change 0466: current dense XLSX commit/save attribution

`performance_claim: none`

`claim_authorized: false`

This batch investigates the current default `xlsx_one_percent_commit_save`
latency lead before changing production code. It retains the exact normal
binary measured in 0465, verifies all 7,032 source manifest entries unchanged,
and adds a separately bound same-source build with frame pointers, unwind
tables and debug level 1 for CPU attribution. No production or harness Rust,
default benchmark identity, checked CRUD coverage, or iWork behavior changes.

Build provenance comes from authenticated build receipts and executable hashes:
the reused normal executable was compiled with Rust 1.98.1 during 0465 at
`161cf53b20d8bb65fe79d4567b9ea7768430de7b`, with the final source epoch
`0021cdee1035d1dc28ae1506fb991396d26b28e6c48ce4f1452c6be4bd5a263d`.
That source is identical at this batch's base `418edc42a`. The normal reports'
runtime environment probe says Rust 1.95.0 and revision `418edc42a`; those
describe the capture environment, not the compiler that produced the reused
executable. Raw metadata remains unchanged. The diagnostic build explicitly
uses Rust 1.98.1 at the current base with the same source epoch and different
profiling flags. Neither report set is a clean distinct-revision comparison.

The synthetic dense-wide corpus has two 256-by-256 worksheets, 131,072 cells,
1,311 updates, seven ZIP members and a 384,525-byte source archive. Source SHA
is `5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`.
Every successful iteration compares output with deterministic expected bytes
and reopens/verifies cells. The output is 388,095 bytes in 37 sequential writes,
with a 65,536-byte largest write. This legacy selector does not independently
exercise the complete patch apply/inverse or native-producer matrix.

The timer contains ordinary `Edit::commit` plus `Workbook::write_to`. Source
open, staged edit planning, sink reservation, expected-output construction,
reopen, full cell verification and teardown are outside. External perf and
Heaptrack diagnostics include the whole process, including those operations
and warmups. Library ancestors can occur in expected-output construction, so
sampled callchain fractions are not retained-sample phase timers.

## Normal observations and capture corrections

All runs use CPU 2, one worker, and a fresh process per lane on the recorded
AMD EPYC 9R45 KVM host. Each normal repeat has 30 samples and three warmups.
The shared guest is not an exclusive machine and these are warm in-process
owned-byte scenarios, not physical-cold or simulated-remote measurements.

| Normal lane | p50 ms | p95 ms | p99 ms | Mean ms | Process maximum RSS KiB |
|---|---:|---:|---:|---:|---:|
| R1, initial | 401.483917 | 403.883151 | 408.610090 | 401.695508 | 109,696 |
| R2, initial | 402.405519 | 405.796413 | 408.106871 | 402.625274 | 110,072 |
| R3, supplemental | 399.470314 | 401.643834 | 403.522390 | 399.073028 | 109,728 |
| R4, supplemental | 398.785007 | 400.513689 | 400.758971 | 399.029414 | 109,724 |

The first sequence overlapped two team-owned `perf report` postprocessors.
Their actual CPU contribution was not measured. Both were terminated and
verified absent before R3/R4, which use the same binary, workload and counts.
All initial artifacts remain retained; the additional runs do not replace
them or establish the cause of the timing difference. R3/R4 are the current
normal-latency observations. Their mean 95% Student-t intervals are
[398.400001, 399.746055] and [398.674752, 399.384077] ms.

The initial sampled workload succeeded, then its wrapper failed when
serializing `stat.st_size` as a callable. Recovery authenticated the existing
outputs without rerunning the workload. The original wrapper is retained;
the recovered receipt explicitly lacks an original start timestamp. The
standard release binary's short/incomplete DWARF callchains motivated the
separate frame-pointer build rather than treating missing ancestors as zero
work. Profile-instrumented latency and wrapper RSS are not normal baselines.

## Counters, allocations and next implementation

The frame-pointer capture contains 15,483 samples and 135,131,742,466
weighted event periods, with zero lost samples reported by perf. Exact
`Edit::commit` ancestors contain 74,497,523,419 periods (55.13% of the
whole-process denominator); the exact `PackageWriter::write_to_stream` with
the harness `CountingSink` contains 21,137,002,598 (15.64%). Within commit,
worksheet Parser ancestors contain 35,501,949,212 periods (47.66%), and
`unqualified_attribute_value` contains 11,458,403,980 (15.38%). These inclusive
rows overlap. They are method-context evidence across the whole process,
including warmups and expected-output construction, not elapsed phases or
operation-local hardware counters. The separate build changes compiler flags,
so its timing is not compared with the normal executable.

The separate 30-sample/three-warmup perf-stat run records 84,856,505,155
cycles, 307,742,221,394 instructions, 70,608,069,981 branches, 78,260,516
branch misses, 77,475,046 generic cache misses and 561,005 page faults.
These are whole-process totals, including the disclosed overlap sequence.
All listed counters report 100% scheduled time. Generic cache misses are not
an independently established LLC event on this host; no per-operation IPC,
cache-locality or instructions-per-byte claim is derived.

The separate five-sample/one-warmup Heaptrack print export reports 56,349,806
allocation calls, 26,103,884 temporary allocations and a rounded 104.38 MB
peak heap. Its raw end-of-recording temporary count differs by one from the
print export; both are retained. These are whole-process values, not a
per-commit allocation delta or peak. The strongest allocation lead is
quick-xml `RawVec` growth under repeated attribute checks, which accounts for
20,333,756 calls across all contexts in this trace. Complete stacks are
retained in the bundle; overlapping inclusive totals must not be added.

Source review identifies two original and two rewritten worksheet Store
parses per ordinary commit, plus XML scanning, rewriting, compaction and
publication audit. Cell-start parsing separately scans the same attributes
for `r`, `s`, `cm`, `vm` and `t`. The next implementation should use the
retained attribution to remove repeated work while preserving duplicate and
malformed-input errors, entity normalization, metadata/style bounds, full
publication validation and lexical/unknown-byte preservation. The previous
unrestricted Store-handoff experiment failed its memory gate; raising that
retention limit is not justified by this profile.

The [evidence bundle](../results/change-0466/) retains raw reports, catalogs,
profiles, commands, original and supplemental observations, source review,
analysis and verification helpers. The larger non-iWork goal remains open:
no speedup, bounded streaming, native acceptance, remote input, or parallel
scaling result is claimed by this investigation.

Final verification passes 11 parser tests and seven bundle-verifier tests,
including five semantic tamper probes. Both summaries replay exactly across
different Python hash seeds. Live precleanup, fresh-copy portable verification
with adjacent frozen 0465 provenance, and postcleanup verification pass. The
final seal covers 103 artifacts. Cleanup removed both authenticated temporary
executables (521,933,936 bytes); an initial finalization refusal also identified
three generated Python bytecode files, which were removed before retry. Raw
capture/profiling artifacts remain retained in losslessly compressed form.
No new workspace-wide Rust or native-application gate is claimed: production
and harness Rust are unchanged from the validated 0465 source epoch.
