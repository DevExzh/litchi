# 0513: XLSX operation allocation metrics

This batch adds a harness-only metric enabler for four existing XLSX commit
and commit/save cases. It records operation-scoped allocation observations
alongside the existing elapsed samples and adds a stable helper boundary for the
exact save profile. No XLSX format implementation, public API, dependency,
archive, or source-format behavior changes. The evidence record has
`performance_claim: none` and `claim_authorized: false`; the batch makes no
registered speedup or allocation-reduction claim.

## Harness change and preserved boundaries

For `xlsx_one_cell_commit`, `xlsx_one_percent_commit`,
`xlsx_one_cell_commit_save`, and `xlsx_one_percent_commit_save`, the harness
starts an allocation region immediately before the existing `Instant::now`
operation clock and finishes it after the elapsed interval. The existing
operation boundaries remain:

* archive cloning, workbook parsing, update preparation and expected-output
  construction occur before the timed region;
* commit-only rows time `edit.commit()`; and
* commit/save rows time the commit plus `workbook.write_to(&mut sink)` through
  `xlsx_commit_save_operation`, while sink reservation is already complete.

The elapsed clock still stops before patch-count checks, exact sink/output
oracles, reopen-and-verify work, report construction and teardown/drop. The
save helper is `#[inline(never)]` and returns the `Commit` before its caller's
post-clock checks and drop, giving Callgrind a stable operation boundary
without moving those existing lifecycle operations into the timer. A short
counting sink still receives partial output and propagates its typed failure.

Normal runs publish an explicit unavailable allocator status because they do
not install the counting allocator. The separate allocator binary measures
the same operation region; its instrumentation overhead, elapsed time and RSS
are excluded from native timing comparisons. Per-sample observations retain
the elapsed sample order and are promoted only after alignment checks.

## Evidence protocol and completed native results

The native protocol covers four cases across three shapes (`tiny`, `medium` and
`dense-wide`), with 100 measured samples and three warmups per row in serial
ABBA order (`before/r1`, `after/r1`, `after/r2`, `before/r2`) on CPU 2. That is
12 rows per lane and 4,800 measured samples in the paired capture. The
allocator-only protocol repeats the 12 rows after-only with 10 measured
samples, one warmup and two repeats. Its 240 instrumented samples are
separate from the native timing population.

The exact profile lane uses the `xlsx_commit_save_operation` helper for the
dense-wide one-percent commit/save case, with three samples and no warmups.
It collects three direct commit calls and three writer calls while excluding
fixture construction and expected-output commit/write work. This is a
diagnostic helper profile, not a whole operation latency or allocation result;
the collected boundary is verified below.

The retained [plan](../results/change-0513/plan.json) fixes these scopes and
requires drift thresholds of 5% for p50/mean, 10% for p95 and 15% for p99.
The native captures are now retained with their source and binary bindings.
The p50 values below are descriptive paired observations, with no latency or
repeat-drift flags. The allocator-only captures are also
retained below. The scoped allocator-feature test receipt records **486 passed, 1 ignored**;
the remaining completed gates are listed below.

The p50 columns use `B1 = before/r1`, `A1 = after/r1`, `B2 = before/r2`, and
`A2 = after/r2`; each delta is candidate minus its paired control. All values
are milliseconds.

| Case | Shape | B1 p50 | A1 p50 | B2 p50 | A2 p50 | A1−B1 | A2−B2 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xlsx_one_cell_commit` | tiny | 0.162515 | 0.162296 | 0.161811 | 0.161816 | -0.135% | +0.003% |
| `xlsx_one_percent_commit` | tiny | 0.303116 | 0.301646 | 0.301701 | 0.301931 | -0.485% | +0.076% |
| `xlsx_one_cell_commit_save` | tiny | 0.204601 | 0.205215 | 0.205410 | 0.205466 | +0.300% | +0.027% |
| `xlsx_one_percent_commit_save` | tiny | 0.369601 | 0.370121 | 0.372226 | 0.372072 | +0.141% | -0.041% |
| `xlsx_one_cell_commit` | medium | 1.783661 | 1.792121 | 1.765046 | 1.787507 | +0.474% | +1.273% |
| `xlsx_one_percent_commit` | medium | 7.137411 | 7.302747 | 7.043876 | 7.142466 | +2.316% | +1.400% |
| `xlsx_one_cell_commit_save` | medium | 2.206568 | 2.234743 | 2.216798 | 2.241563 | +1.277% | +1.117% |
| `xlsx_one_percent_commit_save` | medium | 8.740237 | 8.862068 | 8.809643 | 8.868118 | +1.394% | +0.664% |
| `xlsx_one_cell_commit` | dense-wide | 108.586955 | 110.143715 | 108.496740 | 110.006180 | +1.434% | +1.391% |
| `xlsx_one_percent_commit` | dense-wide | 220.704400 | 223.942449 | 219.741820 | 224.735909 | +1.467% | +2.273% |
| `xlsx_one_cell_commit_save` | dense-wide | 153.020744 | 154.858438 | 154.140029 | 155.342580 | +1.201% | +0.780% |
| `xlsx_one_percent_commit_save` | dense-wide | 307.153163 | 312.534515 | 309.664123 | 312.514798 | +1.752% | +0.921% |

The p50 pair deltas range from **-0.485% to +2.316%**. The recorded maximum
resident set sizes for the two ABBA pairs are **138,012 → 141,672 KiB**
(`B1 → A1`) and **136,532 → 137,028 KiB** (`B2 → A2`), or +2.652% and
+0.363% respectively. These are whole-child `/usr/bin/time -v` observations,
not operation allocation-region peaks. The raw reports retain p95/p99 and per-sample values; the replay also derives
throughput from each mean. No pair crosses the latency/throughput, RSS or
repeat-drift review threshold.

## Completed allocator observations

The after-only allocator lane repeats each of the 12 rows with ten measured
samples and one warmup in two repeats. For every dense-wide row below, all ten
samples in allocator r1 and allocator r2 agree exactly, including allocation
calls, allocated bytes, entry live bytes and region peak. The incremental peak
is a derived value, `region_peak_live_bytes - live_bytes_before`, shown only to
explain the absolute readings.

| Case | Allocation calls | Allocated bytes | Entry live bytes (absolute) | Region peak live bytes (absolute) | Incremental peak (derived) |
| --- | ---: | ---: | ---: | ---: | ---: |
| `xlsx_one_cell_commit` | 1,191,586 | 135,042,963 | 23,613,552 | 67,147,822 | 43,534,270 |
| `xlsx_one_percent_commit` | 2,400,578 | 273,128,176 | 38,821,852 | 97,318,672 | 58,496,820 |
| `xlsx_one_cell_commit_save` | 1,257,527 | 141,906,071 | 6,371,601 | 49,905,871 | 43,534,270 |
| `xlsx_one_percent_commit_save` | 2,532,326 | 286,872,324 | 6,665,149 | 65,161,969 | 58,496,820 |

The commit-only loop retains its prior `final_commit` while the next iteration
prepares and enters its operation region; that value is replaced after the
current timed operation. Save rows do not retain that cross-iteration
`final_commit`. Consequently, commit-only rows have larger absolute entry and
region peaks than their save counterparts, while the corresponding derived
incremental peaks are identical. Absolute peaks must therefore not be
compared as an optimization result. The allocator instrumentation remains
operation-scoped callback evidence and is excluded from native timing.

## Allocation fields and peak interpretation

The allocator observer keeps absolute process counters and takes checked
snapshots at the operation boundaries; it does not reset counters. Allocation,
deallocation and byte fields are published as differences between the two
snapshots for the operation. The live-byte and high-water fields retain their
absolute meaning:

| Field | Meaning |
| --- | --- |
| `live_bytes_before` / `live_bytes_after` | Absolute process live bytes at the two boundaries |
| `peak_live_bytes_before` / `peak_live_bytes_after` | Absolute process-lifetime high-water values at those boundaries |
| `region_peak_live_bytes` | Absolute live-byte high-water observed during the region, including the entry snapshot |

`region_peak_live_bytes` is therefore not incremental operation demand. A
subtraction from `live_bytes_before` can be a derived diagnostic, but it is
affected by other process threads, observer ordering and allocator behavior;
the metric envelope does not present that subtraction as an operation peak.
The region high-water value also excludes allocator-internal reallocation
overlap and physical RSS. The absolute fields must not be reported as a
memory reduction or a process RSS result without the matched evidence and
scope notes.

## Semantic and test boundaries

The helper returns the same successful `Commit` and writes through the same
`CountingSink`; it does not alter workbook serialization, update counts,
output bytes, sink write calls or reopen behavior. The focused test module
checks warmup/sample alignment, normal-binary allocator unavailability,
promotion of deterministic save-sink summaries, exact output and semantic
commit results, and partial output after a short sink failure. The current
tests receipt records 486 passed tests and one ignored test for the
allocator-feature library run.

This is an instrumentation change in `tools/perf-baseline` only. It does not
establish behavior for the complete XLSX CRUD surface, other OOXML formats,
provider/native producers, cold-cache runs or full-workspace execution.

## Exact post-oracle save profile

The normal candidate profile collects **14,458,431,895 Ir** across exactly
three `xlsx_commit_save_operation` calls. The direct helper children are
`Edit::commit`: **11,911,855,692 Ir (82.39%)**, three calls; and
`PackageWriter::write_to_stream`: **2,546,575,570 Ir (17.61%)**, three calls.
The remaining 633 Ir are helper-local work and a memcpy edge. These disjoint
children include their descendants; they are not exclusive leaf costs or
native wall-clock fractions.

Collection starts disabled and toggles only on the private helper. Fixture
construction and expected-output commit/write work use the original route,
so no counter reset at runner entry is needed. Successful Commit destruction
and all caller oracles remain outside the helper. Descendant call metadata
may still contain collection-off calls; only the exact direct runner/helper
and helper/commit/writer edges establish the invocation proof.

Valgrind reports `brk segment overflow` and then exits successfully. The raw
profile, annotations and warning are retained. This diagnostic environment
supports scoped simulated instruction attribution; its time, RSS and heap
behavior do not support native performance or allocation claims. There is one
save profile in this batch, with three operations; no repeat-stability claim
is made. See the [scope review](../results/change-0513/scope-review.md).

## Gates and retention

The allocator-feature harness library tests pass: **486 passed, one ignored**.
Formatting, warning-denied Clippy for the library and both binaries, and
warning-denied rustdoc pass. Boundary, all ten strict registered claim checks and report classification
also pass. The full evidence replay passes, including corrupted/short-vector rejection,
source/binary/corpus/sink identities and allocation alignment/bounds. Owned
scratch cleanup removed 3,137 files (1,669,324,800 allocated bytes), after
confirming no live process references. Raw evidence remains committed with
its checksummed inventory. No new format test,
fuzz, native Office, cold/provider or scaling campaign is claimed.

The measured-enabler decision is scoped to adding aligned allocation evidence
and an exact save boundary, with native overhead below the declared review
thresholds. There is no registered speedup or allocation-reduction claim;
`performance_claim: none` and `claim_authorized: false` remain unchanged.

## Priority and remaining scope

This metric enabler supports the requested OLE2/OOXML-first priority. ODF
optimization remains deferred until their full optimization goal is complete. iWork
remains excluded, the broader non-iWork goal remains active, and no statement
that the full goal is complete follows from this harness-only batch.

The retained [capture protocol](../results/change-0513/capture.py),
[source-bound checks](../results/change-0513/check.py),
[control build receipt](../results/change-0513/before/build-receipt.json),
[candidate build receipt](../results/change-0513/build-receipt.json),
[allocator build receipt](../results/change-0513/after/allocator-build-receipt.json),
[tests receipt](../results/change-0513/tests-receipt.json),
[candidate source manifest](../results/change-0513/source-manifest.json),
and [replay verifier](../results/change-0513/verify.py), together with the
[harness implementation](../../../tools/perf-baseline/src/lib.rs) and
[focused metric tests](../../../tools/perf-baseline/src/xlsx_commit_metrics_tests.rs),
define the current evidence boundary.
