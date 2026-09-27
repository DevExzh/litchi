# 0786 independent results review

## Verdict

The retained 0786 evidence is internally consistent and suitable as a bounded
finite-budget scaling baseline for the three public low-level read routes. I
found no correctness, output-parity, timing-boundary, source-counter, or
permit-release failure. The evidence does not support a production optimization
claim or a whole-program CRUD conclusion.

This review was independent of the benchmark executable. It read the retained
receipts and JSON reports only; it did not run Cargo, native children,
profilers, or hardware-counter tools. The production source remains byte
identical to origin revision
`76dde712f53f007d51be6b29a64b575c3a4c3eb9`. The tool-source census is frozen
across the build and capture lanes.

## Contract and fixture checks

The standalone generator produces 32 distinct deterministic payloads for each
shape. The logical payload totals are 131,072 bytes for `small`, 8,388,608 for
`large`, and 8,130,560 for `mixed` (31 large members followed by one 4 KiB
member). I independently checked all 96 reconstructed member digests and the
ordered sequence digest for every shape. Every report verifies 32 members,
ordered output, and matching member SHA-256 values.

The OPC archive has two structural ZIP entries and 32 typed payload entries.
The production eager loader excludes the content-types and relationship entries
from the payload callback, so the 64 KiB task floor sees the intended 32
payload sizes. The small and mixed cases therefore intentionally serialize at
that floor because one selected member is below the floor; the large case is
eligible for parallel admission. This is an admission-policy result, not a
scheduler failure. The CFB 4 KiB streams are normal FAT streams because the
format's MiniFAT condition is strictly below 4096 bytes; this batch does not
exercise the multi-MiniFAT convergence path.

Container identities, member manifests, sequence digests, and deterministic
outputs agree across the native, observer, and qualification lanes for all
nine route/shape groups. No report has a failed verification marker.

## Timing and budget boundaries

For fresh cases, session construction is outside the timed public operation.
CFB/Parts also prepare package metadata before timing; OPC times the full eager
`from_bytes` open, including its metadata work. Primed cases verify one preload
outside the clock, drop its
result, and time the next operation on the same session. Primed Parts is an
explicit source-backed cache-hit control. The source is immutable in-memory
data, so these are warm-source measurements and do not represent physical cold
cache, filesystem, network, or remote range latency.

Native timing is isolated in the normal binary. Observer timing uses the
separate `source-metrics` binary and is never pooled with native latency. OPC
uses `OpenSession::from_bytes`, so external `ReadAt` counters are correctly
reported as not applicable. Process CPU time is read around the operation and
can be slightly wider than the wall interval. RSS comes from the whole child
process and includes corpus construction, warmups, verification, and report
writing; it is not an operation-only peak measurement.

Each sample creates a fresh finite root budget. Workers and positional I/O are
bounded by the requested width, cumulative CPU tasks by 1,000,000, in-flight
tasks by 32, in-flight bytes by 16 MiB, and aggregate parallel admission by 64
KiB. Across all 22,200 samples, fresh CPU-task usage follows 0 -> 32 -> 32
and primed usage follows 32 -> 64 -> 64 at before-operation, after-operation,
and after-drop snapshots. Every value stays within its configured limit, and
Workers and IoConcurrency are zero after the session and output are dropped.
Observer active-read counts also return to zero and no short read was recorded.

The report snapshots only Workers, IoConcurrency, and CpuTasks. Memory,
InputBytes, OutputBytes, Objects, Depth, Work, and cache-byte usage are finite
configured ceilings, not measured usage. This is an evidence limitation and is
not interpreted as a peak-memory or cache-accounting claim.

The initial quality attempt retained in `quality-0` failed only because three
unused enum helpers triggered `clippy -D warnings`. The corrected `quality-1`
run passed formatting, all-feature check, tests, clippy, rustdoc warnings, and
the crate-boundary checker. Qualification passed all 120 cases before the
primary capture.

## Independent native replay

I independently replayed nearest-rank p50 values and the frozen 10,000-resample
paired-block bootstrap (`seed=786078`, six blocks, width-one reference). The
following values are median speedups over the six paired blocks; each bracket is
the independent 95% bootstrap interval. Values below one are negative scaling.

| Route | State | Floor | w2 | w4 | w8 | w32 |
|---|---|---:|---:|---:|---:|---:|
| OPC | fresh | 0 | 0.5441 [0.5207, 0.5675] | 0.8024 [0.7580, 0.8790] | 1.0265 [1.0132, 1.0606] | 0.6627 [0.6538, 0.6725] |
| OPC | fresh | 65536 | 0.5511 [0.5218, 0.5745] | 0.7885 [0.7818, 0.8087] | 1.0191 [0.9843, 1.0675] | 0.6637 [0.6573, 0.6696] |
| OPC | primed | 0 | 0.5636 [0.5372, 0.6009] | 0.8650 [0.8490, 0.8865] | 1.2377 [1.2062, 1.2629] | 1.5677 [1.5463, 1.5798] |
| OPC | primed | 65536 | 0.5329 [0.5211, 0.5708] | 0.8501 [0.8307, 0.8651] | 1.3100 [1.2874, 1.3398] | 1.5197 [1.4942, 1.5556] |
| CFB | fresh | 0 | 0.9898 [0.9441, 1.0275] | 1.4938 [1.4046, 1.6466] | 1.9194 [1.8178, 2.0989] | 1.4972 [1.4424, 1.5291] |
| CFB | fresh | 65536 | 1.0065 [0.9629, 1.0511] | 1.5510 [1.4469, 1.5746] | 2.1056 [1.9730, 2.2006] | 1.5297 [1.4691, 1.5394] |
| CFB | primed | 0 | 0.9731 [0.9351, 1.0567] | 1.5205 [1.4392, 1.7060] | 2.3949 [2.3038, 2.5686] | 3.1642 [3.0674, 3.2992] |
| CFB | primed | 65536 | 1.0002 [0.9517, 1.0341] | 1.5904 [1.4836, 1.6697] | 2.3556 [2.2080, 2.5262] | 3.1913 [3.0839, 3.3203] |
| Parts | fresh | 0 | 0.5405 [0.5090, 0.5483] | 0.8617 [0.8417, 0.9053] | 1.1977 [1.1436, 1.2296] | 0.9773 [0.9462, 1.0112] |
| Parts | fresh | 65536 | 0.5213 [0.5044, 0.5334] | 0.8575 [0.8236, 0.8812] | 1.1704 [1.1508, 1.2048] | 0.9517 [0.9409, 0.9816] |
| Parts | primed | 0 | 0.0132 [0.0129, 0.0138] | 0.0131 [0.0130, 0.0132] | 0.0101 [0.0099, 0.0107] | 0.0052 [0.0051, 0.0054] |
| Parts | primed | 65536 | 0.0132 [0.0131, 0.0136] | 0.0130 [0.0126, 0.0135] | 0.0100 [0.0099, 0.0102] | 0.0052 [0.0051, 0.0054] |

The full native matrix contains 60 negative-scaling rows; 41 have intervals
entirely below one. The dominant anomaly is primed Parts: its p50 is about 3.0
microseconds at width 1 and 225--301 microseconds at widths 2--8, rising to
about 574--578 microseconds at width 32. The observer lane shows zero source
calls, requested bytes, and returned bytes for all primed Parts cases, so this
large penalty is a cache-hit scheduling/admission observation rather than a
source-read cost. It should not be generalized to cold or remote sources.

CFB large fresh and primed cases show useful scaling through width 8 and, in
the primed state, through width 32. OPC fresh large has only modest width-8
benefit and regresses at width 32, while primed OPC improves through width 32.
Parts fresh improves through width 8 and is neutral or negative at width 32.
These requested-width curves do not infer active worker counts. Process CPU
ratios and `max_simultaneous_reads` are supporting diagnostics, not worker
occupancy measurements. Whole-child RSS spread flags, especially at width 32,
are retained in the scaling output and should not be described as operation
peak-memory regressions.

## Observer evidence

The independent observer and qualification checks cover all 360 observer
reports. CFB reports consistently show 32 logical source calls with requested
bytes equal to returned bytes and zero short reads. Fresh Parts reports show 64
logical source calls with requested bytes equal to returned bytes; the request
volume is the compressed archive range volume (about 73,590 bytes for small,
146,041 for large, and 143,782 for mixed), not the 8 MiB logical payload size.
Primed Parts reports show zero calls and zero requested/returned bytes at every
width and floor. OPC source counters remain not applicable. Histograms account
for every observer call, and active-read snapshots are zero after each timed
operation.

The observer's maximum active-read value is often below the requested width on
the in-memory source. That is expected for this fast source and does not prove
that the scheduler used fewer worker permits; the report correctly exposes
requested-width efficiency and treats active-read observations separately.
The histogram is retained as eight bins; interpreting the bin boundaries still
requires the frozen tool source, so future packets should make those labels
self-describing in the report.

## Next measured candidate

The next candidate should be a focused correctness-and-measurement experiment
for an all-cache-hit `SourceBackedPackage::read_parts_ordered` fast path. The
experiment can test whether an already-resident ordered batch can avoid worker
spawning and admission overhead while preserving caller-selected worker, CPU,
memory, I/O, cancellation, source-version, ordering, and output/object
accounting. It must pair fresh and primed cases and retain the zero-source-call
witness. No production change is adopted from this baseline.

Other required follow-up evidence remains outside this packet: physical cold
and filesystem/network range sources, cross-session contention, CFB MiniFAT
corpora, native-format CRUD workflows, allocator and hardware counters, and
incompressible or producer-diverse corpora. The 0786 result is therefore a
well-formed bounded scaling baseline, not completion of the program-level
performance goal.
