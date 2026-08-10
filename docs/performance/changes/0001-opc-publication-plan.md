# OPC publication-plan reuse

Status: accepted after measurement
Production base: `2665d572b78f0b3efd9ecfc4bd1fda09f8786ae3`

## Mechanism

`PackageWriter` previously constructed and serialized content types, package
relationships, every Part relationship set, and the sorted Part list once for
publication validation and again for output. A private `PublicationPlan` now
performs those fallible operations once, audits all generated/authored XML
before output, and emits the same retained bytes and order.

This does not skip validation, change Deflate behavior, retain source ZIP
records, or introduce a cache/executor. It strengthens the sequential-sink
failure boundary: plan errors occur before the sink accepts a byte.

## Before/after protocol

Two release binaries were built from identical harness sources. The before
binary used the production base; the after binary differed only by the
`PublicationPlan` patch. Runs were interleaved before/after/after/before. Tiny,
many-small, and wide-root cells contain 200 samples per variant after 10
warm-ups per replicate. Few-large cells contain 60 samples per variant after
three warm-ups per replicate. Inputs and outputs had identical SHA-256 corpus
hashes and sink byte/write summaries.

Raw machine-readable reports:

- `results/abba-publication-plan-before-a.json`
- `results/abba-publication-plan-after-a.json`
- `results/abba-publication-plan-after-b.json`
- `results/abba-publication-plan-before-b.json`
- the corresponding four `abba-publication-plan-few-large-*.json` reports

| OPC no-op save corpus | Before mean | After mean | Mean change | Before p50 | After p50 |
|---|---:|---:|---:|---:|---:|
| tiny, compressible | 35.7 us | 34.5 us | -3.42% | 34.3 us | 33.3 us |
| tiny, incompressible | 64.4 us | 63.2 us | -1.96% | 61.4 us | 60.3 us |
| 256 Parts, compressible | 1.569 ms | 1.601 ms | +2.08% | 1.544 ms | 1.572 ms |
| 256 Parts, incompressible | 5.623 ms | 5.462 ms | -2.88% | 5.594 ms | 5.445 ms |
| 2,048 Parts, compressible | 12.607 ms | 11.915 ms | -5.49% | 12.566 ms | 11.995 ms |
| 2,048 Parts, incompressible | 17.356 ms | 16.975 ms | -2.19% | 17.379 ms | 16.992 ms |
| four 4 MiB Parts, compressible | 3.246 ms | 3.224 ms | -0.66% | 3.240 ms | 3.188 ms |
| four 4 MiB Parts, incompressible | 208.861 ms | 212.159 ms | +1.58% | 208.237 ms | 211.499 ms |

The geometric-mean latency change across these explicitly named cells is
-1.65% (improvement). The largest benefit is the intended many-Part case:
5.49%. The two regressions are below the initial 5% review trigger; the
4 MiB-entry result confirms the expected Amdahl limit because Deflate, not
metadata planning, dominates that scenario. No claim is made that every save
is faster.

## Allocations and memory

Heaptrack used the identical 100-iteration 256-Part incompressible command.

| Process total | Before | After | Change |
|---|---:|---:|---:|
| Allocation calls | 356,632 | 224,828 | -37.0% |
| Total allocated bytes | 10,825,833,116 | 10,813,414,883 | -0.11% |
| Temporary allocations | 79,136 | 53,280 | -32.7% |
| Peak heap | 1.73 MB | 1.69 MB | -2.3% |
| Peak RSS including profiler | 12.39 MB | 12.53 MB | +1.1% |

Process totals include deterministic corpus generation, one package open,
runtime startup, and report generation; they are comparable because the
command and corpus are identical. Total allocated bytes are dominated by the
unchanged per-entry compression workspace, which explains why removing many
small metadata allocations barely moves that byte total. The small profiled RSS increase is below the
5% trigger and conflicts with the lower peak-heap result, so it is treated as
profiler/process noise rather than a retained-memory claim. The duplicate
validation/emission `ContentTypesItem` stacks disappeared from the after
profile.

## Correctness, concurrency, and limits

- `cargo test -p litchi-opc --all-features`: 108 tests passed (95 unit,
  13 integration) plus 5 doctests.
- Warning-denied all-feature/all-target Clippy and format checks passed.
- Added a test proving an authored-XML planning failure leaves a sequential
  sink untouched.
- Existing exact-source XML, deterministic enumeration, signature fixture,
  bounded chunk sink, partial-output, atomic replacement, limits, and malformed
  relationship tests pass.
- No parallel work, locks, cache entries, source reads, decompressed bytes,
  recompressed bytes, output bytes, or write-call distribution changed.

## Decision and remaining limit

Keep the change. The latency gain is deliberately modest, but the 37%
allocation-call reduction is material, the 2,048-Part scenario clears the 5%
useful-work threshold, and planning-before-output is a measured enabler for
future raw-copy publication. Full recompression remains the dominant
borrowed-source and mutation-touched save bottleneck and is not hidden by this
result; exact unchanged owned publication is addressed separately in change
0004.
