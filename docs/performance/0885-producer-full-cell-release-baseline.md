# 0885 Producer XLSX full-cell release baseline

This batch measures the explicit `xlsx_producer_medium_source_full_cell_scan`
workload from committed source `335a40350bf06f49691745c14d310fd7a3ca9117`.
It changes no production code or benchmark workload. It makes no candidate
comparison or speedup claim, and does not reuse deleted 0882 measurements.

All eighteen reports and 450 measured iterations pass semantic and artifact
replay, including [independent review](results/change-0885/capture-review.json).
Native process-level p50 has a median of **3.574367 ms**, with a range
of **3.544454–3.638469 ms**. The predeclared variation flag does not fire.

## Results

The [verification record](results/change-0885/verification.json) retains every
process-level statistic and the underlying reports. Native measurements cover
360 iterations; the separate allocator population covers ninety. The 54 warmup
iterations are excluded from these summaries.

| Native metric | Median across 12 children | Child range |
| --- | ---: | ---: |
| p50 elapsed | 3.574367 ms | 3.544454–3.638469 ms |
| p95 elapsed | 3.613800 ms | 3.589814–3.663595 ms |
| p99 elapsed | 3.630855 ms | 3.596424–3.675354 ms |
| Cells/second from child p50 | 1,145,938 | 1,125,748–1,155,608 |
| Process-lifetime peak RSS | 21,106,688 bytes | 19,505,152–21,237,760 bytes |

The native p50 max/min ratio is 1.0265, below the frozen 1.20 spread threshold.
Ranges describe this population's observed variation; they are not confidence
intervals or predictions for other machines.

The six allocator children have identical per-child medians: 25,933 allocation
calls, 366 reallocation calls, 37,618,900 allocated bytes, and 9,252,094 region
peak live bytes. All ninety samples report zero failed allocations. The separate
instrumented p50 median is 3.988001 ms, with a child range of
3.976626–4.016426 ms; it is not substituted for native latency.

## Protocol and scope

The fixed protocol runs twelve native processes, each with three warmups and
thirty measured iterations, followed by six allocator processes with three
warmups and fifteen measured iterations. Each process is pinned to CPU 12.
The host reports an AMD EPYC 9R45 under KVM, with 32 logical CPUs. It is shared;
affinity is not an exclusive CPU or host guarantee.
The [environment](results/change-0885/environment.json) records hardware and
software observations. Both binaries use Rust 1.95.0, release optimization
level 3, fat LTO, one codegen unit, debug information, and no incremental
compilation. Build commands are offline and locked.

Each iteration opens a fresh source-backed workbook and visits all stored cells
in four worksheets: 4,096 cells, including 3,686 numbers and 410 strings.
The timer includes workbook opening, sheet enumeration, production freshness
checks, ordered callback validation, and digest updates. Corpus and source
preparation, the harness's extra source-version comparison, aggregate tolerance
report check, endpoint observations, digest finalization, and handle destruction
are outside the timer. See [0884](0884-producer-full-cell-scan-harness.md) for
the pinned archive, canonical oracle, and workload qualification.

The generated XLSX is a warm in-memory owned positional source. Its marker
census makes all four worksheets ineligible for source streaming. Eligibility
is not an observed runtime profile. No physical-cold, native Office, hardware
counter, scalability, or general Office performance claim follows.

## Interpretation

Raw elapsed vectors are retained per process. The verifier recomputes Rust's
floored-midpoint p50 and nearest-rank p95/p99, then reports medians and ranges
across process-level statistics; it does not pool the children. With thirty
samples per native process, p99 is that process's maximum, so it is a limited
tail observation. The predeclared native p50 max/min threshold of 1.20 flags
spread without discarding or replacing any result.

Throughput divides 4,096 by each native process's p50 workload duration; it
includes the opening and validation work described above. Native RSS is the
maximum reported process-lifetime high-water value per child, not the operation's
incremental memory use. Allocator results come from a separate instrumented
binary and bracket the operation; their elapsed times are not native estimates.
Region peak live bytes include allocations already live at entry.

## Reproduction and custody

The packet retains the [build script](results/change-0885/build.py), full
[source manifest](results/change-0885/source.json), build logs and receipts,
[freeze script](results/change-0885/freeze.py),
[capture driver](results/change-0885/capture.py), and
[verifier](results/change-0885/verify.py). The frozen plan binds source, build,
helpers, passing helper checks, methodology review, host observations, and
`taskset` before capture. Every child retains its exact command, environment,
stdout, stderr, report, producer evidence, and exit status. Invalid semantic
evidence stops capture before the next child; no retries are allowed.

Six focused helper tests and AST parsing pass. An additional offline schema
rehearsal used copies of 0884 debug reports with an in-memory profile override;
it launched no workload and contributes no release measurements. Its first
attempt used an incorrect sidecar filename, corrected before validation.

The exact source archive and both captured binaries are retained locally under
the packet's ignored `source.tar.gz` and `binaries/` paths. They are intentionally
not duplicated into Git. Tracked source revision, file hashes, toolchain, locked
dependencies, and build commands permit source reconstruction, but do not
promise byte-identical rebuilt executables. Archived replay on this host uses
the retained executable hashes. A new capture needs a separate evidence packet
and newly frozen identities; the scripts refuse to overwrite this population.

Replay the retained population into a new output file:

```sh
python3 -B docs/performance/results/change-0885/verify.py \
  --plan docs/performance/results/change-0885/plan.json \
  --capture-root docs/performance/results/change-0885/capture \
  --output /tmp/litchi-0885-replay.json
```

## Cleanup and next work

The [cleanup receipt](results/change-0885/cleanup.json) confirms removal of
`/tmp/litchi-0885-source` and `target/litchi-0885-build` after retained artifact
checks and independent review. The source archive replay verifies all 10,765
files; 3,846 native and 3,847 allocator source dependencies belong to that
manifest. Post-cleanup replay passes using retained artifacts. This batch
created no registered Git worktrees and left no Python cache files. Historical
recovery material and other teams' worktrees remain preserved.

The [read-only source audit](results/change-0885/next-source-audit.json)
identifies namespace resolution in x14ac capture as a hypothesis to profile.
No optimization is adopted from that observation. The broader non-iWork
performance program remains open.
