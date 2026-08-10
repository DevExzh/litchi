# Change 0009: deterministic range source and explicit scaling evidence

Date: 2026-08-10

## Harness change

The standalone harness now has 44 selectable cases: the existing 36-case,
198-record default matrix plus six opt-in simulated-range cases and two opt-in
execution-scaling cases.

The range source has no network dependency. It deterministically splits each
nonempty logical `ReadAt` call into bounded physical requests and applies the
configured fixed latency, per-request overhead, bandwidth, and maximum range.
Every measured sample records logical calls/bytes, physical requests/bytes,
the sorted physical request sizes, and fixed size buckets. The XLSX cases also
enforce zero timed requests for listing and zero physical overlap with every
unselected worksheet range.

The scaling cases use only caller-created local pools. They record the resolved
worker count and exact logical task/byte totals, prewarm pool construction
outside timing, and verify every OPC Part or CFB stream after every sample.
Worker selections are capped to visible parallelism and deduplicated, preventing
a pinned process from silently oversubscribing its CPU set.

Validation passed: formatting, 18/18 tests, warning-denied all-target Clippy,
six range release smokes, and OPC/CFB scaling release smokes.

## High-latency range-source record

The fixed-CPU capture used two warmups, ten samples, 2,000 us fixed latency,
250 us request overhead, 16 MiB/s bandwidth, and a 65,536-byte maximum physical
range. OPC used the many-small/incompressible corpus; XLSX used dense-wide.

| Case | p50 / p95 / p99 | Median physical requests / bytes | Proven property |
|---|---:|---:|---|
| OPC structural open | 27.275 / 27.396 / 27.396 ms | 11 / 19,627 | Ordinary main payload remained deferred |
| OPC open + main Part | 36.711 / 36.863 / 36.863 ms | 15 / 20,709 | Selected main payload loaded and verified |
| XLSX structural open | 44.382 / 44.578 / 44.578 ms | 19 / 1,639 | Workbook metadata read; all worksheets deferred |
| XLSX list sheets | 0.001 / 0.001 / 0.001 ms | **0 / 0** | Listing issued no timed logical or physical request |
| XLSX first cell | 102.248 / 106.790 / 106.790 ms | 9 / 190,574 | Selected worksheet only |
| XLSX narrow column | 102.228 / 103.432 / 103.432 ms | 9 / 190,574 | Selected worksheet only |

Every physical request was at most the configured 65,536 bytes; the observed
maximum was 32,768 bytes. Logical and physical byte totals match. All ten
samples for every XLSX case recorded zero unselected-worksheet read calls and
bytes. The first-cell and narrow-range physical I/O totals are intentionally
identical because the current source-backed worksheet boundary loads the
selected compressed worksheet before applying the row index; the row index
reduces semantic traversal, not ZIP-member fetch bytes.

Raw record: [`range-source-high-latency.json`](../results/range-source-high-latency.json).

## Explicit worker scaling

The unpinned host exposed 12 CPUs. The capture used workers 1, 2, 4, 8 and 12,
three warmups and 20 samples on incompressible many-small and few-large
corpora. Times below are p50 milliseconds; speedups use the same case's
one-worker p50.

| Case / exact work | w1 | w2 | w4 | w8 | w12 | w12 speedup / efficiency |
|---|---:|---:|---:|---:|---:|---:|
| OPC few-large, 6 ZIP tasks / 16,778,178 B | 5.678 | 4.015 | 1.304 | 1.260 | 1.255 | **4.52x / 37.7%** |
| CFB few-large, 4 streams / 16,777,216 B | 3.988 | 0.887 | 0.696 | 0.742 | 0.672 | **5.93x / 49.4%** |
| OPC many-small, 258 ZIP tasks / 285,282 B | 0.573 | 0.759 | 0.674 | 0.693 | 0.783 | 0.73x / 6.1% |
| CFB many-small, 256 streams / 262,144 B | 0.201 | 0.182 | 0.151 | 0.210 | 0.386 | 0.52x / 4.3% |

Large-task scaling saturates at the available independent work: six OPC ZIP
members and four CFB streams. At 12 workers the p50 Amdahl-like serial-fraction
estimate `s = (1/S - 1/N) / (1 - 1/N)` is about 15.0% for OPC and 9.3% for
CFB. CFB w2/w4 and OPC w4 are superlinear in this warm-memory measurement, so
their negative formula outputs are not meaningful Amdahl estimates. The
many-small cases are scheduling/coordination dominated; their `S < 1` formula
outputs exceed one and likewise violate the model. The measured policy lesson
is to keep small batches serial or thresholded and reserve wider pools for
coarse independent payloads.

Tail behavior is preserved in the raw record. At 12 workers OPC few-large was
1.255 / 2.674 / 2.889 ms p50/p95/p99 and CFB few-large was
0.672 / 1.453 / 1.934 ms. The high-worker medians therefore do not justify an
unbounded default, especially for only four to six tasks.

Raw record: [`execution-scaling.json`](../results/execution-scaling.json).
