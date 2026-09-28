# 0819 — real-file ordinary-save baseline

Fresh release measurements establish a preservation-admitted ordinary-save
baseline for three checked-in DOCX, XLSX, and PPTX files. Open/edit/save median
latencies are **5.220218 / 5.419553 / 7.422314 ms**, respectively, with default
full durability. Production, harness code, and dependency locks are unchanged
from `0c784f7aecda2a38bbe6cc98a684ce68da1f0dc8`. This is a current-source baseline,
not a before/after optimization or a historical speedup claim.

[Plan](results/change-0819/plan.json),
[protocol](results/change-0819/protocol-review.md),
[source review](results/change-0819/source-review.md),
[raw evidence](results/change-0819/),
[replayed analysis](results/change-0819/analysis.md), and
[independent review](results/change-0819/results-review.md).

## Inputs and preservation admission

| Format | Checked-in input | Source bytes | Published bytes |
| --- | --- | ---: | ---: |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | 23,535 |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | 8,521 |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | 68,284 |

DOCX appends one paragraph, XLSX edits the first sheet’s A1 cell, and PPTX edits
the first admitted slide/shape text. Exact source and output hashes are retained
in [admission](results/change-0819/artifact-admission.json). Three generated
controls supplement artifact qualification but do not enter the timing matrix.
Six corpora × five policies produce thirty byte-identical-within-case outputs.
The independent XML/relationship/closure audit and the independent ZIP payload,
metadata, member-order, and archive-comment check both pass. Untouched DOCX
relationship XML remains byte-exact after the 0818 repair. Default, full,
file-only, no-sync, and stream policies are artifact controls; timing uses the
default documented route.

The historical [provenance witness](results/change-0819/provenance.json) binds
the earlier Litchi/LibreOffice resave/readback record to these source files.
Current outputs were not resaved by an external Office application. This does
not establish a Microsoft producer identity or current external compatibility.
The three small files are not a representative corpus for all Office workloads.

## Capture and statistics

Fresh quality checks pass: formatting, all-feature/all-target compilation,
full all-feature tests (**641 passed, one ignored, zero failed**, 28 test-result
summaries, including the doctest invocation), warning-denied Clippy and rustdoc,
and crate boundaries (65 packages, 244 internal declarations, 11 existing debts).
Native, exporter, and observer release builds then run serially with opt-level 3,
thin LTO, one codegen unit, debug level 1, panic unwind, incremental disabled,
and two build jobs. Rust 1.95.0 and both exact lock graphs are retained. The
standalone harness lock differs from the workspace lock; neither was updated,
and no measurements using another graph are pooled here.

The shared Linux EPYC 9R45 host uses CPUs 12–19; the captured filesystem is ext4.
Cgroup ancestors report no CPU or memory quota, but this is not exclusive-host
evidence. Inputs and untimed probes warm filesystem/provider caches; no physical
cold-cache, remote-network, device-bandwidth, or hardware-counter claim is made.

| Lane | Processes | Samples/process | Warmups/process | Samples |
| --- | ---: | ---: | ---: | ---: |
| Qualification, instrumented | 12 | 1 | 0 | 12 |
| Native | 72 | 30 | 3 | 2,160 |
| Observer, instrumented | 24 | 3 | 0 | 72 |
| Total | 108 | — | — | 2,244 |

Native blocks run in forward/reverse/forward/reverse/reverse/forward order.
Observer blocks run forward/reverse. These counts are reports and samples,
not saved-file counts: edit samples publish nothing, and counting samples use
an in-memory sink. Native timing has no observer features. Observer elapsed
values are diagnostic and are never pooled with native latency.

The table uses nearest-rank quantiles within each 30-sample process, then the
median across six processes. The absolute p50 bootstrap interval uses 10,000
resamples, seed 819819, and sorted endpoints 250/9749. Raw harness p50 uses the
integer midpoint of the middle two samples; both definitions are preserved.
Raw means replay the producer’s sorted-sample Welford calculation. Full means,
block values, logical workload rates, flags, and samples remain in
[analysis.json](results/change-0819/analysis.json) and
[native.csv](results/change-0819/native.csv).

## Native latency

All values below are milliseconds.

| Format | Phase | p50 | p95 | p99 | p50 bootstrap 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| DOCX | Open/edit/save | 5.220218 | 5.346438 | 5.380229 | [5.197042, 5.230577] |
| DOCX | Edit | 0.055670 | 0.064020 | 0.068666 | [0.055530, 0.055945] |
| DOCX | Save to path | 5.021311 | 5.123737 | 5.143362 | [5.014162, 5.025476] |
| DOCX | Counting sink | 0.051236 | 0.063446 | 0.067035 | [0.050960, 0.052226] |
| XLSX | Open/edit/save | 5.419553 | 5.585354 | 5.683445 | [5.409298, 5.436379] |
| XLSX | Edit | 0.262591 | 0.414977 | 0.437247 | [0.261942, 0.268231] |
| XLSX | Save to path | 4.930676 | 5.059077 | 5.108102 | [4.911730, 4.938136] |
| XLSX | Counting sink | 0.065325 | 0.080265 | 0.109866 | [0.065036, 0.066010] |
| PPTX | Open/edit/save | 7.422314 | 7.528644 | 7.586289 | [7.401553, 7.449039] |
| PPTX | Edit | 1.407427 | 1.421712 | 1.425837 | [1.404372, 1.411978] |
| PPTX | Save to path | 5.620500 | 5.733260 | 5.752925 | [5.593029, 5.645875] |
| PPTX | Counting sink | 0.304592 | 0.317317 | 0.319481 | [0.301556, 0.305677] |

Eight cases carry fourteen max/min >1.05 spread flags across p95, p99, or mean;
none has a p50 spread flag. DOCX and XLSX edit/counting cases carry the four
p99/p50 >1.05 tail flags. All flags remain visible; no outlier removal or sample
replacement was used. Six process blocks do not establish stability on another
host or workload.

Open/edit/save includes opening, one edit, and a default full-durability path
save. Edit opens outside its clock and performs no timed publication. Save to
path opens and edits outside its clock, then includes temporary-file sync and
parent-directory sync. Counting opens and edits outside its clock, reserves the
sink before timing, and measures serialization. Destruction, digests, readback,
and cleanup are outside all brackets. The phases use independently prepared
owners and are not an additive decomposition; subtracting their medians would
not measure sync or parsing cost.

PPTX counting calls `to_bytes()` and then writes that complete buffer once to
the sink. It is not streaming or bounded-memory serialization evidence. DOCX
and XLSX use their sequential sink routes. Logical source/output bytes per
second describe workload sizes divided by phase p50, not physical I/O or memory
bandwidth; edit’s output descriptor does not imply a timed save.

## Resource diagnostics

The following are descriptive medians of six instrumented samples per case
(two processes × three samples), excluding qualification. Peak above entry is
`region_peak_live_bytes - live_bytes_before`, not whole-process RSS. Allocation
scope is the operation’s global system allocator region; destruction occurs
outside the measured region.

| Format | Phase | Allocation calls | Allocated bytes | Peak live bytes above entry |
| --- | --- | ---: | ---: | ---: |
| DOCX | Open/edit/save | 2,617 | 1,258,172 | 634,214 |
| DOCX | Edit | 1,375 | 121,273 | 10,337 |
| DOCX | Save to path | 380 | 590,203 | 568,921 |
| DOCX | Counting sink | 375 | 524,290 | 503,281 |
| XLSX | Open/edit/save | 3,542 | 4,561,313 | 1,103,104 |
| XLSX | Edit | 2,510 | 2,898,254 | 1,066,957 |
| XLSX | Save to path | 126 | 983,559 | 561,522 |
| XLSX | Counting sink | 121 | 917,646 | 495,882 |
| PPTX | Open/edit/save | 11,402 | 5,814,055 | 939,897 |
| PPTX | Edit | 7,682 | 2,719,494 | 305,970 |
| PPTX | Save to path | 674 | 641,339 | 597,823 |
| PPTX | Counting sink | 677 | 845,216 | 532,183 |

[Resource summary](results/change-0819/resource-summary.json) retains ranges,
process-counter values, and external whole-child RSS. All observer operation
`read_bytes` values are zero on these warmed inputs; this is a procfs observation,
not proof of zero physical reads. Lifecycle/path-save `write_bytes` are
24,576 / 12,288 / 69,632 for DOCX/XLSX/PPTX; edit and counting record zero.
Their `syscw` deltas are 1 / 1 / 2 for path publication and zero for edit/counting.
These are same-process counters including probe activity, not device attribution.
Each instrumented child retains 32 empty adjacent probe controls without
subtraction. CPU tick granularity is 100 ticks/sec and is too coarse to resolve
many of these short operations. RSS includes process setup and corpus probes,
so it must not be interpreted as an operation-local allocation peak.

## Interpretation and next experiment

The full-durability save observations are around 5 ms for these small files,
while counting serialization is much shorter. This motivates a fresh controlled
durability-attribution experiment using explicit policies on the same workloads;
it does not identify sync cost by subtraction or justify changing the default.
PPTX edit is also a useful profile target at 1.407427 ms and 7,682 allocation
calls in this case. Profile the exact ordinary edit path before proposing a
production optimization. Preserve full durability and exact untouched members
as the reference contract.

## Replay and custody

All workload handles terminated before offline analysis. Reader-only schema
corrections and their failed attempts are retained in the packet; no raw capture,
frozen plan, driver, or runtime source was changed. The artifact audit and ZIP
preservation checks replay independently. Owned build and scratch trees are
removed after review (7,012 files, 7,452,891,288 logical bytes); [cleanup.json](results/change-0819/cleanup.json) retains
binary identities and deletion accounting. [seal.json](results/change-0819/seal.json)
binds the complete owned change set. Unrelated workspace edits remain untouched.

```sh
python3 -B docs/performance/results/change-0819/artifact_audit.py --check docs/performance/results/change-0819/artifact-audit.json
python3 -B docs/performance/results/change-0819/preservation.py --check
python3 -B docs/performance/results/change-0819/analyze.py --check
python3 -B docs/performance/results/change-0819/validate.py --final
python3 -B docs/performance/results/change-0819/seal.py --check-head
```

Recorded descriptors include absolute workspace paths; relocation requires
explicit rebasing. Fresh reproduction requires the captured inputs, toolchain,
locks, host/affinity conditions, quality/build/export/admission sequence, and
all three capture lanes; old timings are not a replacement for that run.
