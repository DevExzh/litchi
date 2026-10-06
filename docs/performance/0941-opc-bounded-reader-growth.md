# 0941: reject bounded reader growth after paired latency regression

The candidate reduces DOCX reader-open gross allocator requests by **75.70%**,
but fails the prospective all-pair latency safety rule. Reader block 5 increases
operation p50 from 3,132.833 to 3,377.569 µs (**7.812%**). The production change
is rejected and `read_limited` is restored to the parent implementation. Four
reader behavior unit tests, six public ingress tests, a numerically unchanged
test-seed lint correction, and the measurement tooling remain.

## Hypothesis and boundaries

[0940](0940-docx-owned-ingress-attribution.md) measured 256 additional
reader-open reallocations versus vector ingress. The candidate retains 8 KiB
reads and fallible reservation but, when capacity is exhausted, requests
`min(maximum, max(new_len, capacity + min(capacity / 32, 64 KiB)))` with
saturating arithmetic. Full 8 KiB reads request exact capacity through 256 KiB;
a short final read can still leave spare capacity. This is a requested-capacity
policy, not a portable guarantee about allocator physical capacity.

The [0779 rejection](0779-opc-bounded-input-growth.md) already showed that a
larger growth fraction can reduce gross requests while worsening retained memory
and RSS. This experiment therefore includes the same four XLSX fixtures and six
open/lifecycle cells, including the 8,224-byte boundary fixture. It does not
revive the rejected 0779 policy or change any earlier rejection threshold.

Maximum-plus-one probing, short and interrupted reads, refusal of overreported
read lengths, I/O-error ordering, input limits and exact payloads are preserved.
Independent review found no semantic or ADR 0005/0006/0011/0030 blocker. No API,
unsafe code, dependency, concurrency or iWork changes are introduced.

## Capture and identity

The sole DOCX capture contains 48 normal probes-off children (12 paired blocks,
two routes, two versions) and 12 allocator children (three pairs per route).
Each child discards five warmups and retains 50 samples. Route and version order
alternate by block. The sole XLSX capture contains 72 native children (six pairs
per cell, three warmups and 30 samples) followed by 24 allocator children (two
pairs per cell, zero warmups and three samples). All children run sequentially on
CPU 12 under `/usr/bin/time` with a 900-second deadline. Eight real DOCX and 24
real XLSX preflight children are separate from these formal measurements. There
are no retries or replacement samples.

These are fresh-process, warm-cache measurements on the same AMD EPYC 9R45 host
as 0940, not exclusive-host, cold-storage or concurrent measurements. The XLSX
lifecycle probe uses explicit `NoSync`; ordinary save durability is unchanged.
Normal/native timing and allocator-instrumented timing are never interchanged.

The baseline reuses the frozen 0940 source closure and release executables.
Both DOCX reports identify historical base
`3ae856d76424f2f9b58d26dcd20fd074b8e97541`, **not this batch's parent HEAD**.
The candidate closure adds the production change, reader integration tests and
an existing ZIP64 fixture required by its test closure. A 7,928-file comparison
verified the baseline runtime source against the current checkout before the
candidate change; binary SHA-256 values distinguish the measured versions.
Candidate source v3 has 10,782 files / 260,028,642 bytes. Quality v4 adds only the
lint correction; final quality v5 restores baseline production and retains the
behavior tests. Build receipts record Rust 1.95.0, release, offline, locked,
`-j2`, unset Rust flag variables, and the exact source/binary bindings.

## DOCX results

Operation and RSS values below are medians of twelve child medians/observations.
The paired-ratio gate is computed separately; a ratio of these aggregate medians
is not the gate statistic.

| Route | Baseline operation µs | Candidate operation µs | Baseline max RSS KiB | Candidate max RSS KiB |
|---|---:|---:|---:|---:|
| Vector | 3,055.525 | 3,049.408 | 33,898 | 33,836 |
| Reader | 3,411.304 | 3,388.632 | 35,152 | 35,084 |

The median of twelve paired reader baseline/candidate p50 ratios is **1.038011**,
with descriptive bootstrap 95% interval **1.006246–1.070276** (10,000 draws,
seed 941, midpoint medians, sorted ranks 250 and 9,750). It also misses the
prospective 1.05 primary improvement threshold. More decisively, reader block 5
fails the independent requirement that every candidate/baseline p50 be at most
1.05. Neither aggregate improvement nor the allocation alternative waives it.

| Reader open metric | Baseline | Candidate |
|---|---:|---:|
| Gross requested bytes | 271,870,860 | 66,071,775 |
| Allocation calls, including reallocations | 645 | 488 |
| Reallocations | 295 | 138 |
| Live bytes after open | 10,999,368 | 11,022,513 |
| Absolute peak live bytes after open | 22,155,064 | 22,178,209 |
| Region peak live bytes | 13,091,170 | 13,114,315 |

These child medians agree across all three allocator pairs. Vector open remains
285,401 requested bytes, 388 calls and 39 reallocations. Gross requests include
the full requested new size at each reallocation; they do not measure physical
copies, memory traffic or RSS savings. Absolute memory gauges are not summed.

The reducer retains all **21 review flags**: 17 tail/spread flags, two normal
RSS flags, the one p50 safety failure, and one allocator RSS flag. Normal reader
block 4 RSS rises 8.440%, normal vector block 2 rises 5.611%, and allocator reader
pair 3 rises 7.246%. No flag is removed or resolved by assumption. All individual
pairs remain in [DOCX reduced evidence](results/change-0941/docx-reduced-evidence-v1.json).
Both original and moved packets independently reproduce the same reduction.

## XLSX controls

All 96 children pass terminal, source, output and allocator-accounting validation.
The following operation values are descriptive medians of six child medians.
The last column counts failures among the six **individual paired** p50 ratios.

| Cell | Baseline µs | Candidate µs | Pairs above 1.05 |
|---|---:|---:|---:|
| Generated workbook open | 375.999 | 369.494 | 2 |
| Generated workbook lifecycle | 4,595.099 | 4,595.449 | 0 |
| Conditional-formatting open | 1,078.807 | 1,072.859 | 0 |
| Conditional-formatting lifecycle | 3,141.458 | 3,135.808 | 0 |
| Small workbook open | 54.960 | 54.616 | 0 |
| 8,224-byte boundary open | 63.228 | 62.595 | 0 |

Generated-workbook open pairs 2 and 6 (zero-based blocks 1 and 5 in JSON) have
candidate/baseline ratios **1.130724** and **1.141847**. These independently fail
the frozen safety rule. The 36 native pairs retain 35 timing/spread review flags.
None of the 12 allocator pairs exceeds 1.05 for retained live bytes, region peak
or whole-child RSS. This does not resolve the DOCX RSS flags or the latency failures.

Generated open requests fall from 1,093,821,155 to 173,786,566 bytes, calls from
1,573 to 1,190, and reallocations from 635 to 252. Retained live bytes rise from
4,531,971 to 4,592,803 (+1.342%), and region peak from 4,616,928 to 4,677,760
(+1.318%). The boundary fixture requests 224 extra bytes (585,699 to 585,923),
with unchanged call counts; retained live rises from 36,043 to 36,270 and region
peak from 116,793 to 117,020. Small-fixture requests and call counts are unchanged.
These are measured process metrics, not the earlier exact-capacity simulation.

Post-capture review caught an aggregation error in the v4 XLSX reducer: it applied
the p50 gate to per-leg aggregate medians rather than each pair. The original
reducer and output are retained as descriptive history. The separate v5
[pairwise supplement](results/change-0941/xlsx-pairwise-evidence-v5.json) implements
the original frozen rule using the same validated 96 rows; it changes no threshold
and launches no workload. Aggregate “hard reject = false” fields in the v4 output
are not safety conclusions.

## Correctness and retained preparation history

The measured candidate passes OPC default/all-features all-target tests
(1,014/1,036 passed, one ignored in each), OPC doctests, warning-denied rustdoc,
DOCX all-target tests (1,936 passed, one ignored), DOCX doctests, and the crate
boundary check. Candidate v3 Clippy exposed a pre-existing hexadecimal grouping
warning; quality v4 fixes only that literal's grouping and passes both Clippy
feature variants. Its numeric value remains 357,695,443,540. Full workspace fmt
still reports unrelated baseline drift; it is not claimed to pass.

After rejection, final v5 OPC default/all-features all-target suites pass
**1,013/1,035 tests**, with one ignored in each. Both warning-denied Clippy
variants and scoped formatting of `phys_pkg.rs` and `reader_ingress.rs` pass.
The removed capacity-helper test accounts for the one-test difference from the
candidate suites. [Final quality evidence](results/change-0941/final-quality.json)
records commands, source hashes, terminal receipts and counts separately from
[measured-candidate quality](results/change-0941/new-quality.json).

Preparation failures remain external: an initial unused test import and missing
fixture, the inherited Clippy warning, stale XLSX lockfile failures, and a stale
baseline probe Cargo.toml hash in early receipts. The XLSX baseline was rebuilt
with corrected source bindings and produced identical binaries. All measured
probe legs use the same resolved lockfile. Native probe dead-code warnings are
retained, not described as warning-denied Clippy success.

Runner review fixed moved-packet absolute-path reads before measurement. Synthetic
controls use real-shaped time/taskset commands and remove the original packet
before moved replay; they are not Rust workload evidence. XLSX retained-report
controls reject 19 malformed cases, including a boolean sample index accepted
by the earlier validator. The v4 XLSX runner corrects only the completed-child
count receipt and version references; it retains v3 child schemas. The real
XLSX preflight's post-run `binary_before` field is supplemented by an explicit
24-child audit against each true prelaunch start receipt. No failed preparation
run is converted into a formal pass.

## Evidence and reproduction

The [decision](results/change-0941/decision-v1.json), full small reductions,
[DOCX audit](results/change-0941/docx-formal-independent-audit-v1.json) and
[XLSX audit](results/change-0941/xlsx-formal-independent-audit-v1.json) are committed.
The [custody summary](results/change-0941/custody-v1.json) binds source listings,
plans, builds, preflights, formal terminals, replay, archive verification and cleanup.
The measured patch is retained externally as `candidate-v1.patch` (8,554 bytes,
SHA-256 `94845a71626398d07665217185abcadabbff42520cecbad0d23e843199e651e4`).
It is archived for reproduction and is not applied to production.

The verified archive is
`/home/zhuhe/litchi-performance-0941-archive/litchi-performance-0941-v2.tar.zst`
(**573,880,217 bytes**, SHA-256
`6f92bcc3549fae309164908b6c004eb483d2371a3bb005481dbdfc511d1a09c8`).
Extraction verified **83,158 files / 2,689,623,533 bytes**, including raw captures,
source closures, binaries, failed preparation history and moved replay assets.
The archive v1 invocation omitted `--create`; the first v2 preparation also
refused an incorrect target path before archiving. Both failures are retained;
corrected v2 archived and verified successfully. The exact tar command and
per-file inventory remain beside the archive. The bound baseline DOCX binaries
and baseline source closure remain in the separately verified
[0940 archive](0940-docx-owned-ingress-attribution.md#evidence-and-reproduction).

The candidate source listing is `source-candidate-manifest-v3.jsonl`
(1,665,867 bytes, SHA-256
`f0628c6b23bd02968b9948f29a3f2787a6bc4736b1f2a5c791da819157c2cb29`).
`regenerate-source-listing-v1.py source-candidate-v3 OUTPUT.jsonl` recreates it
byte-for-byte from the retained closure; this was verified. Large raw evidence,
source listings and compiled executables remain outside Git. Every committed
result file is smaller than 1 MiB.

Run the following from `/home/zhuhe/litchi-performance-0941` using the retained
canonical paths. These commands replay existing measurements and write new
output files; they do not launch benchmark children. The last command uses the
retained validated v4 input. Its timestamp and input-path descriptors may differ
when reproducing, while the pair statistics and decisions remain identical.

```sh
python3 -B repository-artifacts/comparison-runner.py --replay --packet formal-v1/packet --output /tmp/litchi-0941-docx-reproduced.json
python3 -B xlsx-probes/xlsx-reduce-v4.py --plan xlsx-probes/xlsx-formal-authorized-v4.json --packet-root xlsx-probes/captures --output /tmp/litchi-0941-xlsx-reproduced.json
python3 -B xlsx-probes/xlsx-pairwise-v5.py --plan xlsx-probes/xlsx-formal-authorized-v4.json --reduction reductions/xlsx-formal-v1.json --json-output /tmp/litchi-0941-xlsx-pairs-reproduced.json --text-output /tmp/litchi-0941-xlsx-pairs-reproduced.txt
```

The DOCX original/moved reductions are byte-identical. XLSX original/moved
reductions differ only in `packet_root`; their numerical and validation fields
are identical. The audit asset manifest maps earlier temporary paths to retained
copies before those six owned scratch paths were deleted. Archive cleanup removes
only this batch's build caches, synthetic scratch and verification extraction;
the bound DOCX release binaries and canonical measured packets remain. Existing
0940 artifacts, checkout build files and unrelated user changes are untouched.

Build start receipts give exact regeneration commands and executable SHA-256
values. Frozen schedules and capture wrappers describe how to collect new
measurements; new timing observations are not expected to be byte-identical.
The archive script is committed alongside the evidence. Its archive operation
requires fresh destination paths and a restored full source set after cleanup;
its receipt records the exact original tar command.
