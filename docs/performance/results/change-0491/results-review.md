# 0491 DOCX provider and cache-state baseline

This is a descriptive baseline on the retained machine and source build, with no before/after optimization claim. Each formal process has three warmups and 30 retained samples; two repeats are reported separately. Quantiles describe these small samples, and repeat variation is not a confidence interval. The host is shared; CPU affinity is not exclusive CPU reservation. The filesystem raw reports also retain Student-t mean intervals. Tail estimates require more samples before a regression threshold can be calibrated.

## Source-provider lifecycle

The timer covers typed package open, document preparation, extraction and package/document destruction. Source construction and file staging/open are outside. The returned 10,000-byte text is verified after timing. Counters are logical calls; the range adapter cap/pacing is simulated. All available media-range proofs report zero returned compressed media overlap.

| Role | Repeat | Provider | p50 ms | p95 ms | p99 ms | Calls | Inner short reads | Heap peak increment KiB | Whole-child RSS MiB |
|---|---:|---|---:|---:|---:|---:|---:|---:|---:|
| normal | 1 | bytes | 0.273 | 0.285 | 0.286 | unavailable | — | — | 92.52 |
| normal | 1 | file | 0.280 | 0.289 | 0.293 | 19 | — | — | 92.06 |
| normal | 1 | instrumented-bytes | 0.281 | 0.288 | 0.291 | 19 | — | — | 92.31 |
| normal | 1 | range-64-0us | 0.284 | 0.294 | 0.296 | 74 | 55 | — | 92.46 |
| normal | 1 | range-65536-1000us-104857600bps-minimum-service | 20.343 | 25.351 | 34.356 | 19 | 0 | — | 92.81 |
| allocator | 1 | bytes | 0.288 | 0.296 | 0.297 | unavailable | — | 135.85 | 92.57 |
| allocator | 1 | file | 0.295 | 0.308 | 0.313 | 19 | — | 135.94 | 92.42 |
| allocator | 1 | instrumented-bytes | 0.288 | 0.299 | 0.300 | 19 | — | 135.94 | 92.34 |
| allocator | 1 | range-64-0us | 0.297 | 0.307 | 0.311 | 74 | 55 | 136.60 | 92.57 |
| allocator | 1 | range-65536-1000us-104857600bps-minimum-service | 20.485 | 25.362 | 26.434 | 19 | 0 | 135.94 | 92.51 |
| allocator | 2 | range-65536-1000us-104857600bps-minimum-service | 20.808 | 23.476 | 24.967 | 19 | 0 | 135.94 | 92.31 |
| allocator | 2 | range-64-0us | 0.301 | 0.310 | 0.316 | 74 | 55 | 136.60 | 92.51 |
| allocator | 2 | instrumented-bytes | 0.291 | 0.302 | 0.303 | 19 | — | 135.94 | 92.36 |
| allocator | 2 | file | 0.297 | 0.308 | 0.314 | 19 | — | 135.94 | 92.55 |
| allocator | 2 | bytes | 0.289 | 0.300 | 0.302 | unavailable | — | 135.85 | 92.61 |
| normal | 2 | range-65536-1000us-104857600bps-minimum-service | 20.322 | 20.365 | 21.139 | 19 | 0 | — | 92.27 |
| normal | 2 | range-64-0us | 0.286 | 0.297 | 0.298 | 74 | 55 | — | 92.32 |
| normal | 2 | instrumented-bytes | 0.278 | 0.286 | 0.290 | 19 | — | — | 92.29 |
| normal | 2 | file | 0.283 | 0.296 | 0.296 | 19 | — | — | 92.05 |
| normal | 2 | bytes | 0.274 | 0.287 | 0.288 | unavailable | — | — | 92.17 |

Heap peak increment subtracts region-start live bytes from the callback-order region peak. Allocated bytes and allocation/reallocation counts remain in JSON/raw reports. Absolute region peaks include pre-existing live memory; whole-child RSS includes corpus generation and diagnostics. Neither is a total of operation allocations. The short-read arm reports inner-adapter short reads separately from the outer wrapper.

## Filesystem lifecycle

This timer includes facade path open, full-text extraction and document destruction. It differs from the typed provider setup scope and is not an equal-scope comparison. Each sample runs in a fresh child. Verified-cold means sync, accepted DONTNEED, zero fincore residency/dirty/writeback, and positive process read_bytes during the lifecycle; it does not identify physical-device cache state.

| Role | Repeat | State | p50 ms | p95 ms | p99 ms | Median read_bytes delta | Heap peak increment KiB | Child high-water RSS MiB |
|---|---:|---|---:|---:|---:|---:|---:|---:|
| normal | 1 | warm | 0.582 | 0.594 | 0.604 | — | — | 10.50 |
| normal | 1 | cold-requested | 1.934 | 3.026 | 3.700 | — | — | 10.49 |
| normal | 1 | cold-verified | 2.753 | 4.348 | 4.435 | 81920.0 | — | 22.99 |
| allocator | 1 | warm | 0.589 | 0.603 | 0.633 | — | 135.96 | 10.57 |
| allocator | 1 | cold-requested | 2.031 | 2.481 | 2.798 | — | 135.96 | 10.50 |
| allocator | 1 | cold-verified | 2.724 | 3.302 | 3.638 | 81920.0 | 135.96 | 23.09 |
| allocator | 2 | warm | 0.603 | 0.627 | 0.979 | — | 135.96 | 10.56 |
| allocator | 2 | cold-requested | 2.008 | 5.547 | 7.452 | — | 135.96 | 10.41 |
| allocator | 2 | cold-verified | 2.797 | 3.515 | 5.824 | 81920.0 | 135.96 | 22.99 |
| normal | 2 | warm | 0.584 | 0.617 | 0.621 | — | — | 10.43 |
| normal | 2 | cold-requested | 2.164 | 3.712 | 4.766 | — | — | 10.38 |
| normal | 2 | cold-verified | 4.361 | 6.031 | 8.197 | 81920.0 | — | 22.96 |

## Repeat variance and limits

15 latency quantiles vary by more than 5% between repeats. These are review flags, not optimization regressions:

| Matrix | Role | Arm/state | Metric | Repeat 2 vs 1 |
|---|---|---|---|---:|
| provider | normal | range-65536-1000us-104857600bps-minimum-service | p95 | -19.67% |
| provider | normal | range-65536-1000us-104857600bps-minimum-service | p99 | -38.47% |
| provider | allocator | range-65536-1000us-104857600bps-minimum-service | p95 | -7.43% |
| provider | allocator | range-65536-1000us-104857600bps-minimum-service | p99 | -5.55% |
| filesystem | normal | cold-requested | p50 | +11.93% |
| filesystem | normal | cold-requested | p95 | +22.68% |
| filesystem | normal | cold-requested | p99 | +28.82% |
| filesystem | normal | cold-verified | p50 | +58.43% |
| filesystem | normal | cold-verified | p95 | +38.72% |
| filesystem | normal | cold-verified | p99 | +84.81% |
| filesystem | allocator | warm | p99 | +54.72% |
| filesystem | allocator | cold-requested | p95 | +123.58% |
| filesystem | allocator | cold-requested | p99 | +166.36% |
| filesystem | allocator | cold-verified | p95 | +6.45% |
| filesystem | allocator | cold-verified | p99 | +60.10% |

Provider analysis also retains RSS/absolute-region-peak repeat flags; filesystem analysis retains individual tail-spread diagnostics. No adverse row is averaged away. Prepared-query controls are explicitly cold-ineligible and produce no timed result.

Aligned cold archives contain 564 zero comment bytes. Their one 65,536-byte EOCD search overlaps 58,808 compressed media bytes and 4,500 other unselected payload bytes, with zero successful cache loads at open. The replay preserves these raw overlaps. Main preparation materializes one part; the prepared query performs zero I/O. The aligned source is structurally proven and text-identical, but is not byte-identical to the original warm source.

Profiles are whole-child diagnostics, including corpus/setup/report work. They cannot attribute CPU/syscalls to the timed operation. The stable delayed source has 19 paced calls; bounded range coalescing is a measured follow-up opportunity, subject to identity, budget and preservation tests. Genuine non-static borrowed input, concurrency/scaling, native producers, publication and the remaining CRUD intersections are still open.

Reproduce with the commands in README.md. The final seal rechecks build, report, protocol, profile and cleanup custody. Failed development attempts remain separate from the formal evidence.
