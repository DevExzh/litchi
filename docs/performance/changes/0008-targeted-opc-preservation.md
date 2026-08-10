# Change 0008: targeted OPC raw-member preservation

Date: 2026-08-10

## Decision

Accept targeted raw-member publication for mutation-touched, owned OPC
packages when the source ZIP layout and the modeled OPC member topology are
both provably unchanged.

The package retains private source bytes and semantic provenance independently
from clone-local authorization for exact whole-source publication. Save first
builds and audits the normal OPC `PublicationPlan`. It then raw-copies members
whose payload and semantic relationship/content-type state are unchanged and
regenerates only the changed closure. The sequential sink remains bounded to
64 KiB writes and preserves accepted-byte error accounting.

Before the sink sees a byte, the path falls back to the established full
rewrite for absent provenance, add/remove/rename or `.rels` topology changes,
ambiguous member mapping, and unsupported ZIP layouts (including ZIP64,
projected ZIP64, prefixes, multiple disks, overlap, truncation and
non-contiguous records).

## Correctness evidence

The focused corpus proves that unchanged local spans and central records retain
comments, local and central extras, data descriptors, physical order, central
order, and unknown non-part members. Separate tests cover relationship and
content-type closure regeneration, copy-all after a semantic no-op mutation,
clone isolation, topology fallback, prefixed/ZIP64 fallback, reopen/readback,
and partial-sink byte counts.

Final gates passed: 125 unit tests, 13 integration tests, five doctests,
all-target/all-feature check, warning-denied Clippy, rustdoc without
dependencies, formatting, and diff checks.

## Matched latency result

The release binaries were frozen independently of the worktree:

- before: `raw_opc_before_d6bd13c90`, SHA-256
  `f0ecff68ad8a63a2f11f91360579ce0008170c0a6d16890cdd12ff7adebcbcb5`
- after: `raw_opc_after_9a1562920f`, SHA-256
  `9a1562920f35d856ad3399b46f0f12f4f1c723ae046ebca8009135f995457320`

The fixed-CPU ABBA order was before-A, after-A, after-B, before-B. Each leg
used five warmups and 30 samples; the table pools the two 30-sample legs for
each state. Mean intervals are two-sided Student's-t 95% intervals over the 60
pooled samples. Times are milliseconds.

| Corpus | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| many-small, compressible | 1.577 / 1.687 / 1.963 | 0.189 / 0.237 / 0.324 | **-87.99%** | 1.571 (1.545-1.598) | 0.198 (0.191-0.204) | **-87.43%** |
| many-small, incompressible | 5.728 / 6.077 / 6.428 | 0.206 / 0.243 / 0.359 | **-96.41%** | 5.720 (5.658-5.781) | 0.213 (0.206-0.221) | **-96.27%** |
| few-large, compressible | 3.253 / 3.798 / 3.960 | 1.358 / 1.508 / 1.629 | **-58.24%** | 3.291 (3.237-3.345) | 1.379 (1.358-1.399) | **-58.11%** |
| few-large, incompressible | 216.299 / 223.294 / 227.751 | 61.206 / 64.065 / 64.771 | **-71.70%** | 216.361 (215.444-217.278) | 61.374 (61.085-61.664) | **-71.63%** |

The four-cell p50 geometric-mean delta is **-84.98%** and the mean
geometric-mean delta is **-84.64%**. Output byte counts stayed identical in
every matching corpus. Deterministic after-leg sink summaries also matched:
517 writes for each many-small archive, 13 for few-large compressible, and 461
for few-large incompressible; the largest accepted write was 65,536 bytes.

Raw samples:
[`before A`](../results/abba-raw-opc-before-a.json),
[`after A`](../results/abba-raw-opc-after-a.json),
[`after B`](../results/abba-raw-opc-after-b.json), and
[`before B`](../results/abba-raw-opc-before-b.json).

## CPU and memory result

A matched `perf stat` process run on the few-large/incompressible cell used
five warmups and 15 samples. It includes deterministic corpus/setup work, so
only matched process deltas are used. The timed median moved from 243.946 ms to
59.637 ms (**-75.55%**). Task clock fell 69.28%, cycles 69.21%, instructions
69.85%, branches 70.16%, branch misses 71.11%, cache references 65.20%, and
absolute cache misses 44.16%. IPC moved from 2.48 to 2.43 and the cache-miss
ratio from 1.23% to 1.98%; page faults rose 1.77%.

Raw counters:
[`before CSV`](../results/perf-raw-opc-before.csv),
[`before run`](../results/perf-raw-opc-before.json),
[`after CSV`](../results/perf-raw-opc-after.csv), and
[`after run`](../results/perf-raw-opc-after.json).

The latency win has a measured memory cost. Retaining the 16.78 MB compressed
source beside eager Parts increased the one-shot maximum RSS from 94,528 to
115,572 KiB (**+22.26%**). Heaptrack increased from 737 to 878 allocation calls
(**+19.13%**), from 89.29 to 122.49 MB peak heap (**+37.18%**), and from 100.13
to 124.89 MB profiler RSS (**+24.73%**). The profile attributes a 4.19 MB
payload copy to the regenerated action in addition to the deliberately
retained source. This is accepted as a scoped latency/CPU improvement, not as
a memory improvement; source-backed editable packages and shared regenerated
payload ownership remain follow-up work.

Raw one-shot records:
[`before time`](../results/time-raw-opc-before.txt),
[`before run`](../results/time-raw-opc-before.json),
[`after time`](../results/time-raw-opc-after.txt), and
[`after run`](../results/time-raw-opc-after.json). Heaptrack summaries:
[`before`](../results/heaptrack-raw-opc-before.txt) and
[`after`](../results/heaptrack-raw-opc-after.txt).
