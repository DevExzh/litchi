# 0832 — defer dense XLSX column-map storage until the second record

The completed experiment adopts this change. It keeps zero or one complete
column record inline and promotes to the existing bounded dense map only when a
second valid record is assigned. The experiment removes the two fixed dense maps
observed for the pinned one-record worksheet while preserving the parser,
validator, writer, overlap, and default-save contracts. The root decision is
`adopt: true` with no regression flags.

## Change and preservation boundary

The baseline is `7eeaab48c0281b06f53527d4c4f4ea79050d27e7`. The candidate
changes four XLSX source files: `column.rs`, the worksheet codec, the edit
validator, and the snapshot column writer. `Assignments<T>` now starts with an
empty vector and an optional complete `Assigned<T>` value. The first valid
record is stored inline. A second valid record fallibly allocates the original
`2 * COLUMNS` dense tree, replays both records into a local value, and publishes
the tree only after both assignments succeed. Later assignments retain the
bounded tree algorithm.

The first-record lookup and range conversion are constant-time paths. Promotion
is a one-time `O(COLUMNS)` initialization; later tree assignments remain
`O(log COLUMNS)`. The inline range conversion reserves its one final range
directly, avoiding the dense traversal's intermediate raw-range vector. The
ordered traversal still compacts adjacent equal ranges. Replaying complete
records preserves last-matching-record replacement semantics: omitted
properties in a later record are reset rather than inherited.

Construction is now infallible, so allocation failure moves from map
construction to the second valid assignment when promotion is needed. Invalid
records and existing semantic refusals retain their validation order. The
promotion allocation is local; a failure leaves the prior inline record
available and publishes no partial map. The source review found no correctness
blocker. Its remaining limitation is that the dense reservation has no injected
fault test; the atomicity argument is explicit in the implementation and the
existing parser, writer, transaction, allocation, and full-output gates remain
required.

A declaration-only layout diagnostic reports `Assignments<Properties>` at
24 bytes in the baseline and 64 bytes in the candidate. Dense
`Node<Properties>` remains 32 bytes. These are Rust value-object sizes from
source-extracted declarations, not a production runtime allocation trace.

## Corpus and experiment

The structural census checks all 180 checked-in `test-data/**/*.xlsx` files and
389 worksheet members. It finds 43 worksheets with one direct physical column
record, 159 with multiple records, and 187 with none; all 202 worksheets with
a direct `<cols>` element contain at least one record. The census demonstrates
that the inline case occurs beyond the pinned fixture, while also showing that
most column-bearing worksheets promote. It does not establish producer share,
MCE admission, or a general latency distribution.

The parent experiment retains nine cases, 180 reports, and 20,304
samples, including qualification. Native elapsed measurements and allocator
observer measurements are separate. The real edit uses the pinned
LibreOffice `dateAutofilter.xlsx` fixture and its full-output oracle. The
matrix also includes lifecycle, one-cell and one-percent scale guards, a
dense-wide guard, and an exact no-op control. The requested-byte adoption gate
is at least 50% lower for the real edit; native p50, RSS, allocation-call,
requested-byte, and entry-adjusted region-peak regressions are reviewed
individually. The fixed matrix makes no nonempty-column-action latency claim.

The before quality witness reuses the completed 0831 after run: its eight
terminal quality receipts, logs, source hash, and normalized full-input census
are retained and independently checked. The 0832 after leg receives a fresh
eight-gate quality run. Both source legs build native and observer binaries and
run their qualification lanes before comparative capture is admitted.

Native capture uses six counterbalanced A/B and B/A blocks per source leg, with
three warmups. Real edit, real lifecycle, and the synthetic no-op control
use 500 measured samples per process; the six synthetic edit scale cases use
30. The observer capture uses two counterbalanced blocks, three samples, and
zero warmups. All workloads are pinned to CPU 12. Within-process p50/p95/p99
use nearest-rank order statistics; six-process summaries use midpoint medians,
with the harness-reported p50 checked separately. Paired ratios use 10,000
bootstrap resamples with seed `832832` and sorted endpoints 250 and 9,749.
Qualification uses one sample per source/binary/case and is not a performance
sample. No aggregate across cases or formats is reported.

The [host receipt](results/change-0832/host.json) records an AMD EPYC 9R45,
Linux x86-64, Rust 1.95.0, and Python 3.14.4. Release builds use optimization
level 3, thin LTO, one codegen unit, debug level 1, unwind panics, and no
incremental compilation; the [plan](results/change-0832/plan.json) and command
receipts retain the exact environment and build flags.

The timed real-edit operation measures edit and commit/adopt outcomes without
serializing every timed result. A separate pinned probe checks complete
reference output bytes on both source legs. The lifecycle case includes open,
edit, and default atomic save and checks every published output hash; that
publication-equality check is distinct from the pinned full-byte oracle.

The supplemental promotion guard exercised two generated derivatives of the
pinned fixture: a disjoint two-record worksheet and a 128-record nested,
complete-property worksheet. Its protocol produced 48 reports and 12,040
samples: 16 one-sample qualification reports, 24 native edit reports with 500
samples each, and eight observer edit reports with three samples each. It is a
separate guard for promoted parser paths; it does not change the frozen
nine-case parent matrix or create a broad multi-record performance claim.

## Accepted ADR mapping

| Accepted ADR | Application to this slice |
| --- | --- |
| [0003 — snapshots, edits, patches, and concurrency](../adr/0003-snapshots-edits-and-patches.md) | Promotion stages the new dense map locally and publishes it only after replay succeeds; the source snapshot and existing commit boundary remain unchanged. Lifecycle capture checks publication hashes. |
| [0005 — I/O, memory, and measured performance](../adr/0005-io-memory-and-performance.md) | The experiment measures requested bytes, allocation calls, region peaks, RSS, and native elapsed time separately, with fixed affinity and statistical rules. It makes no cache-miss, throughput, or cross-format claim. |
| [0006 — validation, security, and compatibility](../adr/0006-validation-security-and-compatibility.md) | Complete last-record replacement, reset of omitted properties, existing validation order, and untouched package preservation remain explicit gates. The change does not repair or normalize input. |
| [0008 — migration and verification](../adr/0008-migration-and-verification.md) | This is an incremental XLSX owner change with focused source review, full quality gates, parser/writer tests, qualification, and byte-oracle verification before adoption. It does not certify every Office producer or build. |
| [0010 — archive ownership below the facade](../adr/0010-facade-archive-ownership.md) | No facade detector or archive dependency changes; XLSX continues to use its existing format/package owners. |
| [0011 — physical package ownership below OOXML](../adr/0011-ooxml-physical-package-ownership.md) | No production ZIP/OPC ownership or error boundary changes. Raw ZIP copying is confined to the supplemental test-fixture generator and is not a production package path. |
| [0024 — current workspace topology](../adr/0024-current-topology.md) | The four edits stay inside the standalone `litchi-xlsx` owner; no package membership, facade, or dependency topology is changed. |

## Packet evidence

The packet retains the [experiment design](results/change-0832/design.md),
[source and scope review](results/change-0832/source-review.md), [structural
census](results/change-0832/census.json), and [declaration-only layout
receipt](results/change-0832/layout/compile.json). The inherited before leg is
bound by the [baseline witness](results/change-0832/baseline-witness.json),
[quality receipt](results/change-0832/quality-before.json), [build
receipt](results/change-0832/build-before.json), [qualification
receipt](results/change-0832/qualification-before.json), and [source
installation receipt](results/change-0832/install.json). The supplemental
protocol is described in its [guard README](results/change-0832/promotion-guard/README.md).
These links identify the evidence available to the later readers; they do not
turn the checked-in source fixtures into packet archives.

Final result evidence is retained in the [main analysis](results/change-0832/analysis.json),
[independent audit](results/change-0832/audit.json), [root decision](results/change-0832/decision.json),
[promotion analysis](results/change-0832/promotion-guard/analysis.json), [corrected
promotion receipt](results/change-0832/promotion-guard/corrected-analysis.json),
[independent guard check](results/change-0832/guard-root-check.json), and
[cleanup receipt](results/change-0832/cleanup.json).

## Results and disposition

The inherited before quality witness is the exact completed 0831 after run: all
eight quality receipts were reused and checked. Fresh 0832 after quality passed
all eight gates with 2,090 XLSX tests, 555 harness tests (one ignored), and
three pinned oracle tests. All four release binaries and all 36 one-sample
qualification reports passed. The parent analysis and independent audit both
pass and match for 180 reports and 20,304 samples. The corrected promotion
analysis passes for 48 reports and 12,040 samples. Together the parent and
guard retain 228 reports and 32,344 samples. The pinned full-output oracle and
lifecycle publication checks pass.

The root decision is `adopt: true` with an empty `regression_flags` list. On
the pinned real edit, requested observer bytes fall from 2,373,966 to 276,494,
a reduction of 2,097,472 bytes (88.353%). Allocation calls fall from 2,509 to
2,505. The component-only expectation is the two dense trees: 2,097,152 bytes
and two calls. The remaining 320 bytes and two calls are consistent with
removing the two intermediate raw-range vectors. The decision input keeps
the dense-tree estimate component-only rather than predicting an exact total.

### Native p50

Values are midpoint medians of six native processes per leg, in milliseconds.
The interval is the paired bootstrap interval for the after/before ratio.

| Case | Before → after p50 (ms) | Paired ratio [95% bootstrap interval] |
| --- | ---: | ---: |
| real-edit | 0.248627 → 0.222521 | 0.895998 [0.887658, 0.898114] |
| real-lifecycle | 5.398110 → 5.342501 | 0.990366 [0.981989, 0.996211] |
| one-cell-tiny | 0.101645 → 0.101046 | 0.991685 [0.989030, 1.004339] |
| one-cell-medium | 0.881204 → 0.879649 | 0.998534 [0.993177, 1.001959] |
| one-cell-dense-wide | 62.029293 → 61.949488 | 0.998260 [0.996921, 0.999939] |
| one-percent-tiny | 0.171276 → 0.169666 | 0.993354 [0.987780, 1.002298] |
| one-percent-medium | 3.688327 → 3.701007 | 1.002474 [1.000201, 1.007438] |
| one-percent-dense-wide | 125.725891 → 125.437475 | 0.996840 [0.994864, 1.000012] |
| noop-medium | 0.000565 → 0.000570 | 1.017549 [0.991379, 1.053896] |

### Allocator observer

These are midpoint medians from the two observer blocks. Requested bytes and
calls are measured inside the scoped operation; observer elapsed time is not a
latency result.

| Case | Allocation calls before → after | Requested bytes before → after |
| --- | ---: | ---: |
| real-edit | 2,509 → 2,505 | 2,373,966 → 276,494 |
| real-lifecycle | 3,541 → 3,537 | 4,037,025 → 1,939,553 |
| one-cell-tiny | 1,136 → 1,136 | 620,268 → 620,268 |
| one-cell-medium | 5,007 → 5,007 | 1,528,036 → 1,528,036 |
| one-cell-dense-wide | 142,182 → 142,182 | 42,470,598 → 42,470,598 |
| one-percent-tiny | 1,773 → 1,773 | 1,121,089 → 1,121,089 |
| one-percent-medium | 18,889 → 18,889 | 5,881,397 → 5,881,397 |
| one-percent-dense-wide | 303,078 → 303,078 | 88,521,416 → 88,521,416 |
| noop-medium | 0 → 0 | 0 → 0 |

### RSS and scoped peaks

Native RSS remains a diagnostic whole-process high-water value and includes
harness setup. The entry-adjusted peak is the decision's adoption diagnostic:
within each observer process it subtracts `live_bytes_before` from the raw
region peak, then takes the midpoint median across the two processes. Both
forms are shown because a raw region peak and an entry-adjusted peak answer
different questions.

| Case | Native RSS before → after (KiB), paired ratio [interval] | Raw region peak before → after (bytes) | Entry-adjusted peak before → after (bytes) | Net live before → after (bytes) |
| --- | ---: | ---: | ---: | ---: |
| real-edit | 137,592 → 137,600; 1.000174 [0.999346, 1.000596] | 1,130,221 → 89,604 | 1,066,957 → 26,343 | 4,642 → 4,642 |
| real-lifecycle | 137,606 → 137,588; 0.999753 [0.999419, 1.000364] | 1,130,300 → 629,504 | 1,103,104 → 602,311 | 40,789 → 40,789 |

The seven synthetic entry-adjusted peaks are unchanged: one-cell-tiny
518,759 → 518,759; one-cell-medium 941,014 → 941,014; one-cell-dense-wide
21,210,673 → 21,210,673; one-percent-tiny 552,204 → 552,204;
one-percent-medium 2,340,425 → 2,340,425; one-percent-dense-wide
36,443,008 → 36,443,008; and noop-medium 0 → 0. The decision therefore
claims no general RSS reduction or synthetic peak reduction.

### Promotion guard

The guard covers the disjoint two-record and nested 128-record complete-property
fixtures. Lifecycle qualification produced one shared output hash per variant
across both source legs and both binaries: `785c924300549de8a72a9bbc08cc54d63ec26ca4a637325a06e5d2c91c9782fc`
for `disjoint-two-records`, and
`adfc359c926c6faf3a9f8bdeb83c5170e701cf0b7b87dfcbb412778606b5109b` for
`overlap-128-wide-complete`.

| Variant | Native p50 before → after (ms) | Native p50 ratio [interval] | Native RSS ratio [interval] |
| --- | ---: | ---: | ---: |
| disjoint-two-records | 0.256182 → 0.256166 | 1.001557 [0.978550, 1.010032] | 0.999419 [0.999070, 1.000829] |
| overlap-128-wide-complete | 1.992034 → 1.998495 | 1.001599 [0.998202, 1.005419] | 1.000160 [0.999942, 1.000698] |

The guard observer reports no change in calls, requested bytes,
entry-adjusted peaks, or net live bytes. Disjoint records remain at 2,564
calls and 2,376,852 bytes with entry-adjusted peak 1,067,184 bytes and net
live 4,731 bytes. The 128-record fixture remains at 20,300 calls and
4,655,684 bytes with entry-adjusted peak 1,150,714 bytes and net live 26,491
bytes. Raw region peaks are three bytes lower in each variant (1,130,805 →
1,130,802 and 1,233,596 → 1,233,593). No guard regression flag was produced;
the independent root guard check also recomputed all 32 comparative reports,
quantiles, bootstrap intervals, and counters successfully.

### Diagnostics, custody, and limits

The parent retained 20 descriptive spread flags across the nine cases and nine
tail flags: real-edit 1/2, real-lifecycle 3/1, one-cell-tiny 3/2,
one-cell-medium 0/0, one-cell-dense-wide 0/0, one-percent-tiny 3/2,
one-percent-medium 2/0, one-percent-dense-wide 0/0, and noop-medium 8/2.
The largest cross-block spread ratio is 1.3008 for the real-lifecycle after
p99; the largest p99/p50 tail ratio is 2.068 for real-edit before. These are
descriptive variance diagnostics. The parent and promotion guard have no
adoption regression flags.

The frozen guard reader initially omitted `SOURCES`, then the first additive
replay reader also omitted `write_once`; attempts [06](results/change-0832/promotion-guard/attempts/06-reader/receipt.json)
and [08](results/change-0832/promotion-guard/attempts/08-replay_reader/receipt.json)
retain those failures. `replay_reader_v2.py` supplied only those two globals,
preserved the three frozen guard scripts and all parser/statistics/workload
logic, and reran no workloads. Its [bind receipt](results/change-0832/promotion-guard/reader-correction-v2.json),
[corrected admission](results/change-0832/promotion-guard/corrected-admission.json),
and [corrected analysis](results/change-0832/promotion-guard/corrected-analysis.json)
bind the correction and pass; the successful [v2 admission attempt](results/change-0832/promotion-guard/attempts/10-replay_reader_v2/receipt.json)
and [replay check](results/change-0832/promotion-guard/attempts/13-replay_reader_v2/receipt.json)
are retained.

Five retained offline failures are preserved in total: parent preflight 00,
parent custody preflight 03, guard parser preflight 06, and guard reader
attempts 06 and 08. None launched a workload or changed a measured report.

The evidence supports the named pinned real edit and the two controlled
promotion fixtures. It does not establish a general producer distribution,
nonempty-column-action latency result, cross-format speedup, cache-miss result,
observer-lane latency result, or process-RSS gain. The parent timed edit checks
admitted outcomes while the separate pinned probe checks complete reference
bytes; guard lifecycle outputs are equality-checked and then deleted. The
source review's missing injected dense-reservation fault test remains a
documented limitation, not an adoption failure.

All measurement, analysis, audit, correction, decision, cleanup, and
post-cleanup replay gates passed. The [packet README](results/change-0832/README.md)
lists offline replay commands. Its seal binds the exact owned source, report,
and evidence files; `seal.py verify` checks their committed blobs and parent
revision while confirming unrelated work remains unchanged.
