# 0523: current CFB/OLE2 attribution and allocation measurement

This batch retains an operation-metrics enabler in the standalone benchmark
harness and a current baseline. It makes no production optimization or
before/after speedup claim. The raw CFB open runner now brackets its existing
constructor timer with the canonical allocator region and publishes aligned
observations; the normal executable publishes explicit unavailable allocation
status. CFB/XLS production, providers, corpus generators and default selectors
are unchanged.

See the [change record](../../changes/0523-cfb-open-allocation-attribution.md)
for numerical results, reviewed variation and the next optimization target.

## Frozen setup and reproduction

The plan starts at `0e6379307c5174babfc832edc6113a16c8e9233a`. Apply
`baseline/source.patch` to that revision to reproduce the retained harness
source. Both the full tracked source manifest and that patch are replayed
through a private Git index by the verifier. All 30 previously read ADR/index
hashes remain unchanged. The standalone release build uses Rust 1.95.0, two
jobs, no incremental compilation and its ordinary release defaults; root
workspace LTO settings do not apply.

The focused preflight test runs before freezing. Its first build overlapped
the final addition of two test assertions. That original source-custody
record is retained separately; the focused test is repeated against the final
unchanging file before freezing and before any benchmark capture.

Use a fresh evidence directory and owned target/scratch paths for a rerun;
`run.py` refuses existing receipts and logs. With the frozen plan and source,
execute this serial sequence:

```sh
python3 -B docs/performance/results/change-0523/run.py freeze
python3 -B docs/performance/results/change-0523/run.py build-normal
python3 -B docs/performance/results/change-0523/run.py native
python3 -B docs/performance/results/change-0523/run.py profile
python3 -B docs/performance/results/change-0523/run.py hardware
python3 -B docs/performance/results/change-0523/run.py build-alloc
python3 -B docs/performance/results/change-0523/run.py alloc
python3 -B docs/performance/results/change-0523/checks.py
```

Each child binds the source manifest, plan, driver and executable hash before
and after execution. Native blocks run XLS/CFB/CFB/XLS. Every raw result,
corpus catalog, log, RSS record, numbered profile dump and hardware CSV is
retained. Scheduling is serial for this campaign; the host is shared.

After the owned binaries have been removed, replay the reports and seal:

```sh
python3 -B docs/performance/results/change-0523/analyze.py /tmp/0523-numeric.json
python3 -B docs/performance/results/change-0523/analyze_profiles.py /tmp/0523-profiles.json
python3 -B docs/performance/results/change-0523/analyze_hardware.py /tmp/0523-hardware.json
python3 -B docs/performance/results/change-0523/verify.py --sealed
```

The numerical analyzer reuses retained 0511 row/metric checks and the common
corpus/elapsed validators, with helper hashes in its output. Direct CFB now
has an operation-metrics envelope, so its current validation replaces the
older CFB assertion that the envelope is absent. Existing report files are
not edited or reclassified. The verifier also rejects four independent
in-memory semantic corruptions without modifying captured artifacts.

## Measurement boundaries

The native matrix has 24,000 durations: nine fixed XLS workflows and three
CFB shapes, each with two fresh child observations of 1,000 samples after 20
warmups. Samples within each child share a process and already materialized
input. The printed filesystem-cache/process-isolation defaults do not turn
these selectors into filesystem or per-sample fresh-child measurements.

Direct `cfb_open` measures `OleFile::open(Cursor<&[u8]>)`. Fixture generation,
the file-size oracle, observation/report construction and object drop are
outside its timer. The XLS source-backed and owned-source workflows include
the constructor and selected open/list/one-cell operation; input construction
and cloning, version probes and post-operation oracles follow their existing
runner boundaries. Compare only within a named runner/corpus. Eager XLS,
source-backed XLS and raw CFB have different semantics and costs.

The separate allocator binary retains 720 operation samples, 30 per row per
repeat after three warmups. Its region surrounds the existing constructor or
operation clock, includes timer calls, and excludes subsequent correctness
oracles and object drop. Live-byte balance must reconcile exactly; incremental
peak is region peak minus entry live bytes. Absolute live bytes and whole-child
RSS are separate metrics. Instrumented elapsed times are excluded from native
latency summaries.

CFB's borrowed in-memory Cursor has no logical source observer; source metrics
are not applicable rather than zero. Instrumented XLS reports classify caller
ReadAt ranges and prove zero ordinary opaque-payload reads; these are logical
ranges, not physical I/O. OwnedSource and eager XLS do not fabricate source
read counters. Current source/corpus/output identities are matched across
repeats and instrumentation lanes.

Eight profile children each execute five measured constructor calls. The
exact CFB constructor toggle also captures the fixture generator's setup open;
its separate numbered dump remains retained. Only positive incoming
benchmark-runner dumps enter operation attribution. XLS profiling stops at
the source-backed constructor and excludes the later selected-cell query.
Raw summary, constructor incoming Ir, and self plus direct-callee Ir must
reconcile. Inner call metadata can include collection-off work and is never
used as an allocation or timed-event count. Valgrind warnings remain in the
raw logs.

Two grouped `perf stat` children cover whole-process XLS owned-source
open/one-cell work, including setup, clones, queries, oracles, drops and JSON
reporting. Grouped cycles/instructions/branches/branch-misses require matching
event runtime and 100% running time before IPC or grouped metrics are usable.
Unsupported/failed events remain explicit. These 2,000 instrumented durations
do not enter the native matrix or establish operation-local hardware costs.

Within-child standard deviation and the producer's Student-t mean interval
are recomputed from raw samples. Every same-build timing or RSS variation
above 5% is retained and reviewed; two children do not establish cross-host
confidence or a causal explanation for variation.

## Limits and next work

The synthetic CFB matrix covers tiny MiniFAT, many-small MiniFAT/directory and
few-large regular FAT storage. The fixed XLS corpus includes understood
workbook data and large opaque streams. Existing generator/reopen and
operation oracles remain. No native Office producer, fuzz, physical-provider,
cold-cache, range-latency, concurrency scaling or broad CRUD coverage completion
is asserted. The benchmark enabler is a necessary measurement capability; it
is not an adopted speedup.

The chain-owner review proposes a private checked visited-bit lookup/set
experiment if the fresh profiles confirm material work. Required sector
ownership and physical reconciliation remain separate until an equivalent
validation/error-order proof exists. OLE2/OOXML is the active priority; ODF
is deferred until that optimization goal completes, and iWork is excluded.

## Final validation

All ten quality gates pass, with 610 CFB test executions across two feature
configurations and one valid focused harness preflight. The four semantic
negative vectors are rejected. After removing both owned build paths,
`verify.py` passes replay of 8,584 source files, 32 serial intervals, all three
analysis reports and 80 profile annotations. No owned Python cache remains.
The first preflight's source-custody mismatch remains explicitly separate.
The recursive `SHA256SUMS` inventory is verified after sealing.
