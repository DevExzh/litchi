# 0840 — fresh CFB emission avoids large payload gap initialization

**Adopted.** The fresh CFB writer now emits large streams before the ministream,
matching their existing allocation order. On the deterministic 128-paragraph,
2.56 MB-text DOC case, paired median latency falls **9.86%**, with ratio
**0.901444 [0.892770, 0.903134]**. All successful output bytes remain identical.
The change removes unnecessary zero initialization by a fresh `Cursor<Vec<u8>>`;
it does not change sector allocation, padding, compression or publication policy.

The frozen gates pass, but not every metric improves. The 512-paragraph DOC
case's whole-process RSS increases **4.80%**, just below the predeclared 5%
point-estimate threshold. Tiny DOC allocation calls rise from 165 to 166, and
mini-only CFB p50 rises 1.57%. These observations remain visible below.

## Mechanism and contract

Base is `f43040e1da`. Change 0753 identified CFB destination zero-fill as a
material remaining fresh DOC writer cost. Current source confirmed the cause:
`OleWriter::write_to` allocates large streams first, then the ministream, but
previously wrote the later ministream first. On an empty growable cursor that
write initialized the gap covering large stream payloads, which were then
written over the initialized bytes.

The production diff moves the ministream block after the large-stream loop.
It leaves the earlier source-layout reuse return, allocation and directory
order, absolute sector offsets, FAT/MiniFAT/DIFAT, reserved-hole clearing,
stream padding, and final flush intact. The independent untimed sink trace
replays cursor positions and the output high-water length from every event:

| Case | Gap bytes before | Gap bytes after |
| --- | ---: | ---: |
| DOC tiny | 12,288 | 0 |
| DOC 512 paragraphs | 94,208 | 0 |
| DOC payload | 5,136,384 | 0 |
| CFB tiny mixed | 8,192 | 0 |
| CFB 4 MiB mixed, v3 | 4,194,304 | 0 |
| CFB 4 MiB mixed, v4 | 4,194,304 | 0 |
| CFB large-only / mini-only | 0 | 0 |
| CFB 8 MiB DIFAT control | 512 | 512 |

Write-call counts, written-byte totals and seek-call counts stay equal. The
mixed cases lose one backward seek. DIFAT is still emitted after its later-
allocated FAT; that distinct 512-byte gap remains. Gap counts identify avoided
cursor initialization, not total copies, allocator zeroing, physical page
faults, or a wall-time fraction.

Five focused tests cover v3/v4 gap behavior and round trips, nonzero reused
sinks, short writes and `Interrupted` retries, typed partial sink failure, and
zero-length/4095/4096-byte boundaries. The baseline fails only the new gap
assertion; all five pass after the change. A failed direct sink may observe a
different partial-byte order, as permitted by `write_to`; the test preserves
typed I/O failure behavior, not an identical failed-output prefix.

ADR 0001 correctness priorities, 0002/0024 ownership, 0005 caller-owned I/O and
performance evidence, 0006 deterministic preservation, and 0026 directory
metadata remain intact. No public API, dependency, unsafe production code,
execution policy, durable wire or ordinary filesystem durability changes.
The v4 fixture does not cross the 2 GiB range-lock boundary; existing owner tests
remain the evidence there. ODF remains deferred and iWork is excluded.

## Protocol and results

Two fresh native builds and two separate allocator builds use identical probe
sources and root-aligned dependency versions. Rust 1.95.0, opt3, fat LTO, one
codegen unit and CPU 12 of the recorded EPYC 9R45 host are used. Equal-length
executable/output paths avoid the previously observed argument-length confound.
This is a shared host, not exclusive CPU isolation.

Each formal process runs three warmups and thirty measured operations. Six
paired blocks alternate before/after and case order. Nine cases cover fresh
DOC construction plus save, mixed CFB, unchanged-route controls, v4 sectors and
DIFAT. Native timing includes writer construction, registration of prepared
text/payloads, `write_to` and writer destruction; input generation, verification
and returned output destruction are outside. This differs from 0753's older
generator-inclusive boundary, so its timings are motivation only, not pooled
baseline data. Allocator regions cover the same operation in a separate binary;
they do not measure wall or CPU time.

All **270 reports / 6,534 samples** validate: 108 native reports / 3,240 samples,
108 allocation reports / 3,240 samples, eighteen native qualifications,
eighteen untimed seek observers and eighteen allocator preflights. Every child
finished successfully; no measured run was discarded or replaced. Formal
workloads run serially without overlapping a root Cargo command. Qualification
can overlap the read-only boundary check; it is not performance evidence.

The reducer uses per-process nearest-rank quantiles and six paired block
ratios, with 10,000 bootstrap resamples, seed 840084 and sorted endpoints
249/9749. The target requires at least 3% DOC-payload p50 improvement with an
upper interval below one. A p50, whole-child RSS or operation-allocation ratio
above 1.05 with lower interval above one triggers review. Zero/nonpositive
baselines use absolute increases. Thresholds and cases did not change.

| Case | Before p50 ms | After p50 ms | Paired ratio [95% interval] |
| --- | ---: | ---: | --- |
| DOC tiny | 0.006620 | 0.006640 | 1.003018 [0.987263, 1.011412] |
| DOC 512 paragraphs | 0.130186 | 0.126526 | 0.967539 [0.961656, 0.979714] |
| DOC payload | 0.696438 | 0.623442 | 0.901444 [0.892770, 0.903134] |
| CFB tiny mixed | 0.001635 | 0.001590 | 0.975535 [0.951515, 0.990683] |
| CFB 4 MiB mixed | 0.643088 | 0.620873 | 0.951122 [0.938176, 0.969352] |
| CFB large-only | 0.610317 | 0.612748 | 1.001106 [0.993861, 1.008835] |
| CFB mini-only | 0.002225 | 0.002275 | 1.015725 [1.011192, 1.026779] |
| CFB v4 mixed | 0.649913 | 0.607487 | 0.952860 [0.933729, 0.970989] |
| CFB DIFAT control | 2.489796 | 2.489290 | 1.002937 [0.984037, 1.007982] |

Absolute columns are medians of process p50 values; their quotient need not
match the median of paired ratios. Equal-weight geometric means of individual
p50 ratios are 0.956398 for the three DOC shapes and 0.982893 for the six CFB
shapes. These are named synthetic groups, not a whole-Office aggregate.
No p95/p99 point estimate exceeds a 5% increase. Complete individual metrics,
intervals and raw block vectors are in the [paired tables](results/change-0840/paired.md),
[CSV](results/change-0840/paired.csv) and [analysis](results/change-0840/analysis.json).

## Memory and observed increases

DOC-payload operation allocation calls stay at 915; requested bytes fall from
23,697,590 to 23,672,438, region peak above entry from 18,182,936 to 18,166,168,
and retained live-byte delta from 10,274,176 to 10,257,408. All nine allocation
peak/retained point estimates stay equal or improve. Tiny DOC uses one more
allocation call (165 → 166, about 0.61%) while requested bytes fall
80,025 → 73,305. The no-ministream CFB controls keep allocation metrics equal.

DOC 512-paragraph RSS is **7,822 → 8,138 KiB**, paired ratio
**1.048042 [1.012574, 1.053363]**. The lower interval exceeds one, and the upper
interval crosses 1.05; the point estimate remains below the frozen 1.05 trigger.
This observed increase is retained without attributing it to a particular heap
or kernel mechanism. Whole-process RSS includes corpus preparation, semantic
verification and allocator history, while the operation allocation metrics
measure a narrower region. Their differing directions do not establish a
physical-memory explanation. DOC-payload RSS is 80,068 → 80,174 KiB, paired
ratio 1.001074 [1.000076, 1.004552].

Mini-only CFB p50 rises 1.57%, with interval above one, despite its unchanged
emission branch. No cause is inferred from that control. Zero frozen flags
means the stated gates passed; it does not mean all metrics were unchanged.
Independent final review accepted the bounded change with these disclosures.

## Validation and custody

Fresh affected-owner gates pass on both legs: formatting, all-feature/all-target
compilation, warning-denied library Clippy, rustdoc and the full boundary check.
CFB/DOC tests pass **1,863 before and 1,868 after**, with fifteen ignored on each
leg. Each standalone native probe passes five tests and each allocator companion
passes 36. Fourteen rejected reader mutations, quantile/bootstrap checks,
independent CFB sector/stream parsing, exact cross-process/cross-leg artifact
comparison and 27 independently recomputed native quantile rows support replay.
DOC readback checks every generated paragraph and the complete text through
public APIs outside the timer.

The evidence retains initial probe Clippy failures, a test-only formatting
failure, allocator schema and CFB sequence-order reader failures, and the
expected baseline gap-test failure. The frozen analyzer then failed on duplicate
`before`/`after` keyword fields returned by the statistics helper. Its unchanged
source and failed receipt remain; additive `analyze-v2.py` removes only those
duplicate fields and binds its own identity. No measurement, reducer, interval,
threshold or case changed, and no workload was rerun for that repair.

The [packet](results/change-0840/README.md) contains exact source/build bindings,
raw artifacts, command receipts, mutation checks and repair provenance. Cleanup
hash-verifies the four retained binaries and removes only marker-owned roots.
Unrelated work is preserved. This closes one measured work-elimination target;
opened-package workflows, other Office formats, physical I/O, producer coverage,
concurrency and the wider OLE2/OOXML goal remain open.
