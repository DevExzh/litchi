# 0827 — ordinary-save effect of the shipped PPTX compaction proof

The shipped 0824 optimization improves this real PPTX full-lifecycle p50 by
2.494% and edit p50 by 11.524% in a fresh matched comparison. No frozen p50
regression, allocation increase, or RSS review threshold is triggered. This covers twelve
ordinary-save scenarios: three real DOCX/XLSX/PPTX inputs, each measured for
lifecycle, edit, atomic path publication, and counting publication. It makes
no new adoption decision. The baseline restores the exact two pre-0824 source
archives; the after leg uses the committed source at `c5d375083c`. All other
production and harness source, locks, corpus bytes, and build settings are fixed.

| Format | Input | Input bytes | Admitted output bytes |
| --- | --- | ---: | ---: |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | 23,535 |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | 8,521 |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | 68,284 |

Input and output SHA-256 identities are retained in the frozen plan and
artifact-admission receipts. Edit-only scenarios publish no archive; their
semantic outcome is checked separately.

## Scope and method

The host is an AMD EPYC 9R45, x86_64, with 32 logical CPUs. All workload
processes run serially on CPU 12. Both source legs receive fresh release builds
of the native harness, artifact exporter, and allocation/process observer.
Release settings are opt-level 3, thin LTO, one codegen unit, debug level 1,
unwind panic handling, incremental compilation disabled, and two build jobs.

Before measurement, each leg exports six corpora under five output policies
and undergoes independent XML/OPC and ZIP preservation checks. Historical
artifacts supply exact byte oracles only; no historical timings or executables
enter the comparison. Each leg must then pass twelve one-sample qualification
reports and separate qualification admission.

Native timing uses six counterbalanced blocks, thirty samples per process,
and three warmups. The separate observer uses two blocks, three samples per
process, and no warmups. Ratios are after/before, paired by block; uncertainty
uses 10,000 bootstrap resamples with seed 827827. The frozen review rule flags
a native p50 ratio whose lower confidence bound exceeds 1.05, allocation
increases, or a paired RSS median ratio above 1.05. All twelve rows remain
visible regardless of direction.

Lifecycle includes opening, editing, and default full-durability save. Atomic
publication includes file synchronization, rename, and parent-directory
synchronization; opening and editing precede its clock. Counting publication
also prepares the owner outside the clock. PPTX counting publication materializes
an archive with `to_bytes`. It does not demonstrate streaming output.

Package-owner destruction, output hashing, readback, and destination removal
occur after the clock. The returned PPTX publication snapshot is dropped inside
the edit helper. Allocation regions end while the owner remains live, so signed
net-live bytes describe retained state at that boundary. Independent phase
medians cannot be added to reconstruct lifecycle latency.

## Results

All 216 reports and 4,488 samples completed. Independent arithmetic replay
agrees with the primary reader on every native row, paired confidence interval,
and observer allocation summary. No native p50 regression, allocation increase,
or paired RSS review threshold was triggered.

For this PPTX file, the paired p50 ratio is **0.975058** for full lifecycle
(95% interval **0.973040–0.982246**), a **2.494%** improvement. Edit alone has
ratio **0.884765** (**0.882642–0.887019**), an **11.524%** improvement. The
atomic-publication and counting-publication intervals include 1.0. The result
supports a modest complete-save benefit for this one admitted workload; it
does not extend the edit-only percentage to the full save or other documents.

Times below are milliseconds, summarized as the median of six process statistics.
Ratios are medians of the six paired ratios, so they need not equal the ratio
of the displayed aggregate medians. Each process p50 uses the integer midpoint;
p95/p99 use nearest rank. Bootstrap endpoints use sorted zero-based ranks 250
and 9,749, resetting the seed for each metric.

| Format / phase | p50 before | p50 after | Paired ratio | 95% interval |
| --- | ---: | ---: | ---: | --- |
| DOCX / lifecycle | 5.274724 | 5.272876 | 0.999294 | 0.993437–1.002358 |
| DOCX / edit | 0.056015 | 0.056168 | 1.002049 | 0.994976–1.007723 |
| DOCX / atomic_publish | 5.088365 | 5.098148 | 1.001989 | 0.996772–1.005880 |
| DOCX / counting_publish | 0.051980 | 0.052310 | 1.006355 | 0.989117–1.021903 |
| XLSX / lifecycle | 5.464522 | 5.479965 | 1.001050 | 0.994168–1.006286 |
| XLSX / edit | 0.263688 | 0.264493 | 0.999528 | 0.973970–1.019405 |
| XLSX / atomic_publish | 5.000233 | 5.017995 | 1.002887 | 1.000733–1.004710 |
| XLSX / counting_publish | 0.064937 | 0.066000 | 1.012539 | 0.996532–1.029812 |
| PPTX / lifecycle | 7.474617 | 7.304064 | 0.975058 | 0.973040–0.982246 |
| PPTX / edit | 1.412057 | 1.249406 | 0.884765 | 0.882642–0.887019 |
| PPTX / atomic_publish | 5.693326 | 5.689858 | 1.000518 | 0.996739–1.003876 |
| PPTX / counting_publish | 0.304329 | 0.302829 | 0.993335 | 0.989625–1.007070 |

P95, p99, means, and process peak RSS are retained for all rows below. Arrows
mean before → after. RSS is whole-process `/usr/bin/time` maximum resident set
size in KiB; it is not operation-local heap usage.

| Format / phase | p95 ms | p99 ms | Mean ms | Peak RSS KiB |
| --- | ---: | ---: | ---: | ---: |
| DOCX / lifecycle | 5.398627 → 5.395717 | 5.401192 → 5.428912 | 5.274782 → 5.280984 | 137,082 → 137,072 |
| DOCX / edit | 0.063415 → 0.064900 | 0.073790 → 0.068401 | 0.056960 → 0.056955 | 137,054 → 137,206 |
| DOCX / atomic_publish | 5.183516 → 5.181556 | 5.218991 → 5.248176 | 5.081486 → 5.098340 | 137,058 → 137,124 |
| DOCX / counting_publish | 0.065495 → 0.065751 | 0.069346 → 0.070835 | 0.054077 → 0.054430 | 137,092 → 137,058 |
| XLSX / lifecycle | 5.670094 → 5.680949 | 5.726699 → 5.823410 | 5.477018 → 5.499301 | 137,042 → 137,088 |
| XLSX / edit | 0.296061 → 0.398947 | 0.309136 → 0.459407 | 0.268492 → 0.282100 | 137,172 → 137,100 |
| XLSX / atomic_publish | 5.128681 → 5.111161 | 5.151311 → 5.165551 | 4.996622 → 5.008819 | 137,124 → 137,106 |
| XLSX / counting_publish | 0.083730 → 0.109296 | 0.119741 → 0.112196 | 0.068314 → 0.070148 | 137,082 → 137,100 |
| PPTX / lifecycle | 7.612579 → 7.462853 | 7.751665 → 7.502668 | 7.489178 → 7.335493 | 137,102 → 137,110 |
| PPTX / edit | 1.422708 → 1.264352 | 1.425232 → 1.266517 | 1.411175 → 1.250807 | 137,066 → 137,082 |
| PPTX / atomic_publish | 5.808409 → 5.810694 | 5.871565 → 5.847309 | 5.698297 → 5.695103 | 137,048 → 137,074 |
| PPTX / counting_publish | 0.319147 → 0.316376 | 0.323877 → 0.319451 | 0.306082 → 0.303517 | 137,102 → 137,138 |

### Allocation and memory

Observer medians below use two processes per leg with three measured samples
each. All failed-allocation vectors are zero. PPTX lifecycle and edit each save
**298 allocation calls and 21,688 allocated bytes**; signed net-live and peak
above entry remain unchanged. All other rows have unchanged medians for the four
metrics in the table. Absolute live-byte endpoints and peak counters are three
bytes lower in the after leg across all twelve cases; this common offset is
retained in the raw vectors and does not establish an operation memory saving.
All eleven raw counters, derived values, per-sample vectors, paired RSS, and
process diagnostics remain in [analysis.json](results/change-0827/analysis.json).

| Format / phase | Allocation calls | Allocated bytes | Net live bytes, both legs | Peak above entry bytes, both legs |
| --- | ---: | ---: | ---: | ---: |
| DOCX / lifecycle | 2,617 → 2,617 | 1,258,172 → 1,258,172 | 66,800 | 634,214 |
| DOCX / edit | 1,375 → 1,375 | 121,273 → 121,273 | 9,296 | 10,337 |
| DOCX / atomic_publish | 380 → 380 | 590,203 → 590,203 | 1,507 | 568,921 |
| DOCX / counting_publish | 375 → 375 | 524,290 → 524,290 | 1,507 | 503,281 |
| XLSX / lifecycle | 3,542 → 3,542 | 4,561,313 → 4,561,313 | 40,789 | 1,103,104 |
| XLSX / edit | 2,510 → 2,510 | 2,898,254 → 2,898,254 | 4,642 | 1,066,957 |
| XLSX / atomic_publish | 126 → 126 | 983,559 → 983,559 | 0 | 561,522 |
| XLSX / counting_publish | 121 → 121 | 917,646 → 917,646 | 0 | 495,882 |
| PPTX / lifecycle | 11,402 → 11,104 | 5,814,055 → 5,792,367 | 342,074 | 939,897 |
| PPTX / edit | 7,682 → 7,384 | 2,719,494 → 2,697,806 | 158,985 | 305,970 |
| PPTX / atomic_publish | 674 → 674 | 641,339 → 641,339 | 0 | 597,823 |
| PPTX / counting_publish | 677 → 677 | 845,216 → 845,216 | 0 | 532,183 |

### Spread and tail flags

Native block families produce **30 metric spread flags** (maximum/minimum
above 1.05); the observer produces **16**. The complete list and ratios are
in [metrics.md](results/change-0827/metrics.md). Native p99/p50 exceeds 1.05
in eleven case/leg summaries; nine observer summaries also exceed it. No RSS
or positive allocation-counter block spread exceeds 1.05. Zero or signed
counter spreads are undefined where a multiplicative ratio is inappropriate.

These flags limit tail claims. In particular, XLSX edit/counting publication
vary across blocks, and several p99 estimates spread by more than 5%. The
unchanged DOCX/XLSX formats serve as controls, not evidence of a cross-format
optimization. The slightly slower XLSX atomic-publication ratio is 1.002887
(1.000733–1.004710), below the frozen 5% regression threshold. No broad tail
or RSS benefit is claimed.

## Qualification and validation history

Six fresh harness quality gates pass: formatting, all-feature/all-target
checking, tests, warning-denied Clippy, warning-denied documentation, and the
crate-boundary audit. Fresh tests total 641 passed, zero failed, and one ignored.
The exact-source 0824 PPTX quality receipts are explicitly reused: 1,241 before
and 1,253 after tests passed, with three ignored per leg. Those are historical
checks, not fresh 0827 test runs.

Three preflight suites passed before protocol freeze. Driver and reader
preflights cover 325 historical reports; the allocation fixture census contains
6,733 sample envelopes. Qualification admission also checks thirteen retained
qualification fixtures. These are schema and admission checks, not reused
performance measurements.

Two failed preflight invocations are retained with complete source snapshots
and logs. The driver first expected one publication hash rather than one per
sample; the independent reader first omitted the eleven explicitly unavailable
native allocation metric objects. Both assumptions were corrected and their
preflights passed before freeze. The shared allocation checker also enforces
unsigned 64-bit vector values and exact signed live-byte conservation.

Four later primary-reader failures and one independent-reader failure are also
retained. Repairs reconcile the frozen provenance's keyed paths with normalized
corpus descriptors, compare the two-file transition subset, classify
qualification as allocator-instrumented, use consistent case identifiers in
grouped and paired results, and read the independent audit's `rows` container.
These corrections change offline readers only. Every attempt preserves its
source snapshot and log; frozen drivers, protocol, binaries, reports, and
admission evidence remain unchanged. All seven failed invocations are tooling
failures. No Cargo, exporter, qualification, or comparative workload failed.

The final primary reader, independent arithmetic reader, complete custody audit,
and validator pass. The capture census is 24 qualification reports/samples,
144 native reports/4,320 samples, and 48 observer reports/144 samples. The
independent reader rederives native quantiles, means, paired ratios and intervals,
and allocation summaries from the raw reports without importing the primary
reader. Source, binary, corpus, artifact, and exact command checks remain
separate from the statistical comparison.

## Replay and cleanup

```sh
python3 -B docs/performance/results/change-0827/analysis.py --check
python3 -B docs/performance/results/change-0827/raw_audit.py --check
python3 -B docs/performance/results/change-0827/validate.py --final
python3 -B docs/performance/results/change-0827/seal.py --check-head
```

The [packet](results/change-0827/README.md) retains the frozen protocol, six
build receipts, source transitions, artifact and qualification admission,
raw samples, statistical outputs, and fourteen reader-attempt snapshots.
Verified cleanup removed only the owned target and filesystem scratch roots:
7,015 files and 8,056,907,786 logical bytes. Binary hashes were checked before
removal and remain bound by the cleanup witness. Offline replay uses retained
evidence and does not rebuild or rerun workloads.

The committed production source is restored exactly; all 35 normative inputs,
benchmark runtime, locks, and the three unrelated workspace files retain their
recorded hashes. The commit adds measurement evidence and documentation only.

## Limits

Claims are limited to this host, three small admitted files, fixed operations,
and warm provider/filesystem state. Process counters include adjacent procfs
probes and cannot establish exact production syscall counts or physical-device
I/O. Identical compressed payloads identify potential passthrough bytes; they
do not prove runtime copying behavior. Observer elapsed time is kept separate
from native latency.

This trial does not establish cold-cache behavior, broader producer coverage,
external Office resaves, cross-platform results, remote I/O, concurrency, or
completion of the broader performance goal. No iWork work is included.
