# Change 0183: ODS one-percent source-backed release evidence

Date: 2026-08-18

## Decision

Accept the existing source-backed ODS path for the fixed 21-existing-cell
workload (21 of 2,048 cells, approximately 1%). On one clean current-HEAD
release binary, its complete open, stage, commit, and sequential-publication
lifecycle is 72.07% and 72.61% lower at p50 than the eager owned-snapshot
control in the two A/B/B/A pair directions. Mean, p95, and p99 improve by
68.20%-72.33%, and same-implementation drift remains inside the predeclared
5%/5%/10%/15% p50/mean/p95/p99 thresholds.

This is an evidence closure, not a new production optimization. Change 0177
withheld the same 1% workload because its mean and tail stability gates failed.
The current clean rerun changes no production or harness code and makes no
claim beyond the already implemented, bounded existing-cell lifecycle. The raw
selector metadata therefore retains its conservative pre-acceptance text
`none; correctness and phase evidence only`; this reviewed change record and
summary carry the later evidence decision.

## Measurement contract

All four legs use revision
`63088f66332068cf57f6faf0afe68d0618ba6d8d` and release binary SHA-256
`9937ec3e5f9b09286ff728cddb02f8b45fe59cef66b5525b81dd50c7e1e038e0`.
Each fresh process is pinned to CPU 2 and retains 500 samples after 20 warmups.
A1/A2 run `ods_source_eager_one_percent_edit_save`; B1/B2 run
`ods_source_backed_one_percent_edit_save`.

The timer covers open, staging 21 bounded existing-cell replacements, commit,
and sequential publication through the fixed zero-retention hashing sink. The
source-backed logical `ReadAt` counters come from a separate untimed complete
lifecycle replay. They are not physical-I/O, decompression, or device-traffic
evidence.

| Metric | A1 eager | B1 source | B2 source | A2 eager | A1 -> B1 | B2 -> A2 | Eager drift | Source drift |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| p50 | 72.673 ms | 20.297 ms | 20.301 ms | 74.123 ms | -72.07% | -72.61% | 2.00% | 0.02% |
| mean | 73.911 ms | 20.940 ms | 20.870 ms | 75.424 ms | -71.67% | -72.33% | 2.05% | 0.34% |
| p95 | 83.793 ms | 25.401 ms | 24.666 ms | 86.488 ms | -69.69% | -71.48% | 3.22% | 2.89% |
| p99 | 89.889 ms | 26.369 ms | 28.699 ms | 90.236 ms | -70.66% | -68.20% | 0.39% | 8.83% |

The 16,790,689-byte deterministic source has two sheets, 2,048 cells, and
eight opaque media resources. Every leg produces the same 16,790,961-byte
output SHA-256
`5d9fc848e830ccea59ca2632715d1eaf54d8aeafc3683c92241c22404285db06`
and semantic SHA-256
`70712048d45fd0d6e5066f86b32d4d17df31ff4cf3ae6063947c417fd7df47de`.
Within each implementation, deleting timing vectors yields identical A1/A2
and B1/B2 correctness projections. The eager and source-backed raw schemas
intentionally contain different backend-specific gates, so no cross-backend
raw-projection identity is claimed.

The source-backed replay is fixed at 617 logical reads, 16,801,025 logical
bytes, and 10,336 bytes of classified range overlap per sample. The sink
retains no output and at most a 16 KiB authoring window. Exact output, semantic
reopen, media preservation, raw untouched-member identity, patch
forward/inverse, foreign-source refusal, exact no-op, replacement limit,
partial-sink progress, source immutability, and phase-within-lifecycle gates
all pass.

## Scope retained

No allocation/RSS, physical-I/O, decompression, cache-temperature,
real-producer, durable ZIP patch, atomic-save, formula, merge, structural-row,
insert/delete, or broad ODS CRUD claim is made. The larger next production
seams remain touched-sheet cloning and complete rewritten-worksheet
validation/materialization; this result does not imply either is solved.

Artifacts:

- [summary](../results/ods-one-percent-release-0183-summary.json)
- [manifest](../results/ods-one-percent-release-0183-manifest.json)
- compressed raw A1/B1/B2/A2 reports listed in the manifest
