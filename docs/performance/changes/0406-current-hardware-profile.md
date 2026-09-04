# Change 0406: current hardware profile of OPC materialization

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

## Capture and scope

Source revision `9eff696b6421e6182e5a96418268a2d8016f0af9`, stable Rust
1.98.1, release optimization, system allocator, and CPU affinity 2 identify this
current-state capture on an AMD EPYC 9R45 host. The harness was freshly rebuilt
with `CARGO_PROFILE_RELEASE_DEBUG=0`; function symbols remain available, but
source-line debug information is not claimed. The binary hash, machine/tool
metadata, exact commands, retained samples, capture script, and resumed
postprocessing script are bound by the
[artifact manifest](../results/change-0406/artifact-manifest.json).

The selector is `opc_source_materialize`, shape `few-large`, payload
`incompressible`: four 4 MiB logical Parts in a six-member, 16,783,632-byte
Deflate OPC/ZIP archive. Its archive SHA-256 is
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`.
Each normal/counter/CPU-profile run uses 20 warmups and 500 retained samples;
Heaptrack uses 20 warmups and 100 samples. Runs are sequential.

Only `into_opc_package()` is inside the operation timer. Source construction,
source-backed open, deterministic payload regeneration, relationship/Part
verification, and process metadata probes surround that interval. External
profilers observe the whole process. Their totals must not be assigned to the
materialization interval or divided by one document's bytes to make a library
instructions-per-byte or allocation claim.

## Observations

| Measurement | Observed value | Scope |
| --- | ---: | --- |
| p50 elapsed | 2,215,288 ns | Uninstrumented materialization, 500 samples |
| p95 elapsed | 2,266,958 ns | Same operation and run |
| p99 elapsed | 2,318,899 ns | Same operation and run |
| Mean elapsed | 2,219,236.874 ns | Same operation and run |
| Source calls per sample | 532 | In-process `ReadAt`, all 500 samples |
| Source bytes per sample | 16,782,540 | In-process returned bytes, all 500 samples |
| Materialized Parts | 4 | All 500 samples |
| Maximum concurrent reads | 1 | All 500 samples |
| Process wall time | 26.38 s | Separate `/usr/bin/time -v` run |
| Process peak RSS | 82,596 KiB | Same process lifetime, not operation peak |
| Cycles | 118,636,236,819 | Separate whole-process `perf stat` run |
| Instructions | 274,141,602,384 | Same counter run |
| Branches / branch misses | 23,350,087,395 / 15,290,075 | Same counter run |
| Cache misses / page faults | 86,407,324 / 50,778 | Same counter run |

The counter file records 100% scheduling for all six events. The user-cycle
profile reports no lost samples. Self samples attribute 65.50% to the harness's
`payload_bytes` and 26.20% to SHA-256 compression. The verifier regenerates and
hashes each deterministic Part after the timed operation, then repeats the
main-document payload check. Thus this process profile is dominated by fixture
and oracle work. It is evidence for improving profiling efficiency; it does
not justify a SIMD or algorithm change in a production owner.

The remaining visible self samples include CRC32 (2.63%), memory movement
(2.55%), and memory initialization (0.91%). Those are process-wide symbols;
the inclusive report contains substantial unresolved caller frames. Complete
phase attribution is therefore unavailable from this build. The
[CPU flamegraph](../results/change-0406/cpu-flamegraph.svg) retains that
limitation and is descriptive only.

Heaptrack records 41,209 allocation calls, 8,513 temporary allocations,
59.16 MB peak heap consumption, and 76.39 MB peak RSS including instrumentation
(as rounded by Heaptrack). These are process totals for the separate 20/100
run. Its stacks distinguish corpus ZIP-output growth (17.09 MB reported peak
contribution), materialized payload buffers (16.78 MB over 480 read-entry
allocations), and deterministic payload generation (4.19 MB peak over 604
allocations). These stack observations are not operation-local aggregate
allocation or RSS counters.

## Reproducibility and limits

The exact source is committed. The user's intentionally untracked
`docs/GOAL.md` remains untouched and hash-bound in the capture. Reports retain
`git_worktree_dirty: true`; no clean ABBA claim is made. This is one synthetic
current-state baseline, with no control revision, speedup, regression, native
producer, physical-cold-cache, scaling, or general Office CRUD conclusion.
`fincore` describes executable residency, not corpus cache state. `strace`
observes process read/write syscalls; it is not a physical package-I/O measure.
Compressed/decompressed/recompressed byte counters and operation-local
allocator counters remain explicitly unavailable in the normal harness report.

The initial inclusive-symbol report stalled on the inherited Ubuntu debuginfod
URL. Only that postprocessing subprocess was terminated. Its failure record
is retained, and analysis resumes with `DEBUGINFOD_URLS` empty. Completed
baseline, time, counter, and CPU-sampling workloads are reused without reruns.
The initial driver had not flushed its command journal: their argv is
reconstructed from the retained script and outputs, and original command
timing/exit metadata is explicitly marked unavailable. Resumed commands
retain directly captured timing and exit records.

## Next evidence

Prepare reusable verification expectations outside repeated iterations while
preserving complete Part, relationship, and main-document checks. Version the
harness identity and re-establish a baseline, because changed verification can
affect cache and allocator state even though it sits outside the timer. Then
use an explicitly identified build with reliable caller unwinding,
operation-scoped allocation evidence, and call-chain attribution before
selecting a production optimization. The wider non-iWork CRUD and native
producer matrix remains open in the [goal audit](../GOAL_AUDIT.md).
