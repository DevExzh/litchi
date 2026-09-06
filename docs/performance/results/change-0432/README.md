# XLSX streaming creation: operation memory across row counts

This batch extends the existing `xlsx_streaming_create` measurement with the
same operation process/allocation envelope already used by RTF streaming.
The production writer, deterministic four-cell row generator, exhaustive
owned-workbook oracle and timed hashing discard sink remain unchanged.

The frozen protocol runs 64, 8,192 and 131,072 rows in fresh normal and
allocator processes, twice in reversed order. Each process uses three warmups
and thirty samples on CPU 2. Normal timing and allocator diagnostics remain
separate; instrumentation latency is not an optimization result.

The configured row scratch is 4 KiB. That measures one authoring buffer and
does not include the compressor, package metadata, context or caller text.
The allocator region measures callback-order logical heap peak during the
writer call, independently of historical setup highwater. The full-process
RSS includes artifact construction and exhaustive materializing reopen.
These are distinct observations, not interchangeable memory bounds.

This is streaming creation of one new inline-scalar worksheet. It does not
measure logical append to an existing workbook, adding an existing package
Part, or arbitrary modification and repackaging. Those append semantics remain
separate in the required CRUD taxonomy.

## Results

All twelve formal processes passed, retaining 360 samples. Two tiny pilot
reports separately checked the report schema before formal capture; their
60 samples are not part of this matrix. Source revision is
`4c9cfd1fc9dd9df1a0f53a086bce274a8cf687ae`.

| Rows | Cells | Output bytes | Normal p50 ms, R1 / R2 | Allocator peak above entry |
|---:|---:|---:|---:|---:|
| 64 | 256 | 3,451 | 0.1160 / 0.1161 | 420,110 bytes |
| 8,192 | 32,768 | 167,418 | 13.4558 / 12.5297 | 420,110 bytes |
| 131,072 | 524,288 | 2,563,433 | 188.6571 / 188.1209 | 420,110 bytes |

Every one of the 180 allocator samples has the same 420,110-byte incremental
region peak, zero live-byte change at exit and zero failed allocation calls.
This supports stable observed incremental requested heap across these three
row counts. It is not a universal heap, allocator-internal, RSS or all-feature
bound. The independently modeled row scratch remains 4,096 bytes.

Per-operation requested allocation bytes are 2,491,428 / 3,157,924 / 13,234,084
for tiny/medium/large. Request counts are 144 / 8,272 / 131,152, each including
six reallocations. Cumulative allocation work grows even though the measured
incremental live peak stays fixed. No allocation-stack or physical-copy
attribution is claimed.

Normal medium p50 differs 6.88% between repeats, with corresponding mean/tail
and throughput flags. Those values remain descriptive; a stable medium
latency baseline is not accepted. Other normal p50 repeat differences are
below 0.3%. No normal-versus-allocator latency comparison is made. Full raw
samples, p95/p99, means, Student-t mean intervals and repeat flags remain in
`summary.json` and its source reports.

Process maximum RSS is approximately 80.7 MiB for tiny/medium and 279.5–279.9
MiB for large. That includes the untimed materializing oracle and process
setup. It is not the streaming writer's operation-local memory peak.

## Validation and reproducibility

The release harness library and allocator-target suites pass 310 tests with
one ignored;
31 streaming writer tests and two focused debug oracle/resource tests pass.
The oracle checks the exact six members and all stored cells, including
coordinate-specific E1 and member-set mutation regressions. Documentation with
warnings denied and scoped formatting pass. Strict Clippy retains the same
29 preexisting findings across 17 message/file groups, with none in changed
streaming/oracle regions. Five unchanged files retain prior formatting debt.

`checks/development-notes.md` retains the intentionally interrupted full debug
suite, initial test import/formatting correction and initial analysis import
failure. `verify.py --portable-check` validates source/driver/artifact custody,
rederives results and lint comparison, and exercises report/verifier mutations
in an isolated copy. `seal.py` records lossless compressed logs and complete
hash inventory. Replay requires no original executable or repository; rebuilding
still requires the repository and compile-time assets.

See [measurement-clarifications.md](measurement-clarifications.md) for absolute
allocator counters and honest untracked-file status, [next-work.md](next-work.md)
for the ODS authoring gap, and [native-gap.md](native-gap.md) for the native PPTX
name/identity blocker. The broader non-iWork goal remains open.

The `precleanup-portable` replay passed isolated export and all three mutation
checks. `task-cleanup` then removed only the two copied executables; exact
hashes and preserved paths are retained in `cleanup.json`.

`aftercleanup-portable` also passed after both copied binaries were removed.
The final policy requires both portable receipts and the cleanup receipt.
