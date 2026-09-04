# Change 0408: attributable OPC materialization evidence

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

## Measured motivation and implementation

0406 attributed most whole-process CPU samples to repeated deterministic
payload generation and verification. This batch prepares compact Part names,
URIs, lengths, SHA-256 digests, and relationship expectations once before the
sample loop, generating and dropping each expected payload individually.
Every resulting package is still fully verified after its timed conversion.
The configuration marker `prepared-part-digest-v1` separates this harness
from historical reports that regenerated their oracle per iteration.

The additive consuming and borrowed OPC materialization accounting methods
reuse the existing serial ZIP session and cold-read counter seam. Cache hits
add no ZIP work; managed conversion refuses before report mutation or payload
reads; partial progress survives a later CRC or source-change error. The new
opt-in `opc_source_materialize_accounted` selector carries dedicated accounting
claim and counter scopes. Existing default selectors stay at 36; the full
selectable registry now contains 422 entries.

## ADR scope

| Constraint | Implementation |
| --- | --- |
| 0001, 0004: layered explicit API | Caller-owned low-level OPC report; ordinary methods retain their signatures. |
| 0002, 0010, 0011, 0024: package ownership | ZIP counters remain in the ZIP/OPC owners; no format or facade archive dependency added. |
| 0003, 0006: immutable source and preservation | Source freshness, CRC/framing/size checks, cache sharing and signed-package policy remain in existing paths. |
| 0005: bounded explicit execution and evidence | Managed refusal retained, no hidden worker pool, accounting is opt-in, profiler totals retain whole-process scope. |
| 0008: verification | Mixed Stored/Deflate, warm-cache, managed refusal and partial-error tests accompany the APIs. |

## Capture

Production API commit `562d1b979` and harness/validation commit
`eba1f302eb1e04519925d1791a2f9d299e908d89` identify this capture. Rust 1.98.1
release binaries were built with debug level 1, forced frame pointers and unwind
tables. CPU affinity is 2 on AMD EPYC 9R45. Each of six sequential runs uses
20 warmups and 500 retained samples on the four-4-MiB-Part `few-large`,
`incompressible` corpus, archive SHA-256
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`.
The worktree contains only the user's untracked `docs/GOAL.md`, so the reports
correctly record a dirty worktree and explicitly exclude a clean ABBA claim.

| Observation | Value | Scope |
| --- | ---: | --- |
| Normal p50 / p95 / p99 | 2,183,708 / 2,213,238 / 2,222,408 ns | Timed materialization, 500 samples |
| Normal mean; 95% mean interval | 2,184,756.712; [2,183,305.513, 2,186,207.911] ns | Student's t interval, in-process samples |
| Source calls / bytes | 532 / 16,782,540 per sample | Timed logical ReadAt returned bytes |
| Allocations / deallocations | 55 / 71 per sample | Separate counting-allocator operation deltas |
| Allocated / deallocated bytes | 16,861,253 / 86,122 per sample | Same allocator run |
| Reallocations / failed allocations | 0 / 0 per sample | Same allocator run |
| Compressed Deflate bytes read | 16,782,356 per sample | Separate accounted selector |
| Decoded bytes produced / accepted | 16,777,216 / 16,777,216 per sample | Same accounted selector |
| Output bytes / recompressed bytes | 0 / 0 per sample | Same accounted selector; no output sink |
| Process wall time / peak RSS | 7.03 s / 82,624 KiB | Normal process, `/usr/bin/time -v` |
| Cycles / instructions | 31,552,257,515 / 40,209,793,690 | Separate whole-process perf-stat run |
| Branches / branch misses | 1,354,765,076 / 15,233,750 | Same perf-stat run |
| Generic cache misses / page faults | 83,887,031 / 31,362 | Same perf-stat run |

All allocation and ZIP counter vectors above are constant across their 500
retained samples. The normal operation allocates 84,037 bytes more than its
16 MiB decoded payload; this measures allocation volume, not memory-copy volume
or total live memory. Decoder workspace is reused within a materialization;
no new decoder/session optimization is claimed in this batch.

The CPU sample report has zero lost samples. Self samples attribute 74.49% to
SHA-256, 10.00% to memory movement, 7.30% to CRC32, and 0.66% to deterministic
payload generation. The inclusive report resolves the verifier at 71.82% and
the serial materializer at 16.34% of whole-process samples. It therefore
separates surrounding verification from the library path. These percentages
are sampled CPU attribution, not elapsed phase timers or a complete partition;
inclusive rows overlap and must not be added together. The raw script contains
7,124 sample records. A 0.61% unresolved kernel frame and two `addr2line`
warnings per postprocessor remain in the retained outputs, so complete unwinding
is not claimed. Binary identity hashing accounts for 2.95% inclusive samples
and corpus construction for 3.83%, both outside the materialization timer.

The optional cache command exits successfully, but LLC loads/misses explicitly
report `<not supported>`. Both L1 aliases return zero despite 100% scheduling;
they are unvalidated and are not accepted as evidence of zero L1 activity.
Native PMU event validation remains open. The six basic counters report 100%
scheduling. No instructions-per-logical-byte library claim is inferred from
whole-process totals.

The [retained bundle](../results/change-0408/) contains the capture script,
exact command journal, build/test logs, binary hashes/build IDs, schema-2
catalogs, all samples, raw perf data, symbolized caller reports, and a
[CPU flamegraph](../results/change-0408/cpu-flamegraph.svg). The large symbolized
perf script is retained losslessly as Zstandard; its original hash and size
are recorded. Rebuilt ELF binaries are not archived. The script, reports and
flamegraph retain symbolization independently of the temporary build directory.

Each iteration opens a fresh in-memory source-backed view before timing, so
logical payload caches start cold. This is not filesystem cold-cache evidence;
the generic filesystem configuration fields do not change this selector's
in-process semantics. Process snapshots and full semantic verification bracket
the operation. External CPU counters and profiles include that surrounding work.
Allocator calls/bytes are operation deltas; live/peak allocator fields and RSS
retain their absolute process/lifetime scope.

This is descriptive current-state evidence. The changed oracle and frame-pointer
build prohibit interpreting differences from 0406 as a production speedup.
The broad CRUD, native-producer, cold-cache and scaling matrix remains open.

## Validation and limitations

All 425 OPC tests pass, including mixed compression, cache-only conversion,
managed refusal and partial-error accounting. Scoped warning-denied OPC Clippy
and rustdoc pass. A broader Clippy attempt remains blocked by existing
`chunks_exact` warnings in `litchi-core`; no full-workspace lint success is
claimed. Crate boundaries pass for 64 packages and 240 internal declarations,
with 14 explicitly recorded pre-existing migration debt items.

The final release harness library suite passes 257 tests with one ignored.
The first run's stale selector-count assertion was corrected from 421 to 422.
The allocator test binary passes all five tests after making the maximal
failure-test request sizes opaque to optimization. The allocator implementation
is unchanged; real null-return checks, failure counters, retained allocation
contents and deallocation assertions remain. Initial failing logs are retained.
The full Python suite runs 860 tests: 840 pass and 20 skip. Its focused
comparator suite passes 84 tests. Report validation binds materialization claims
to their selector names, requires the exact oracle marker for the new accounted
case, and permits omission only for historical normal reports. Complete
configuration matching prevents old/new oracle comparison.

Every capture report passes corpus binding, parallel metrics, operation schema,
selector/claim, oracle, binary and source identity checks. Compressed artifacts
pass Zstandard integrity checks and decompressed SHA-256 equality. All required
capture and postprocessing commands succeed. Cache-event viability remains
qualified above. Strict verification passes all six registered performance claims; the report
classification checker validates 167 historical/current rows. The initial
structural-only registry invocation correctly required strict mode for landed
claims; both invocation logs are retained. The broader goal and non-iWork
scenario matrix remain open.
