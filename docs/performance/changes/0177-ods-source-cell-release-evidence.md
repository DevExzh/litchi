# Change 0177: ODS source-backed existing-cell release evidence

Date: 2026-08-17

## Decision

Accept the matched source-backed existing-cell publication path for the fixed
one-cell workload. On the same clean release binary, its complete lifecycle is
74.27%-75.03% lower at p50 than the eager owned-snapshot control in both
A/B/B/A directions. Mean, p95, and p99 also improve in both directions, and
same-implementation drift remains inside the predeclared 5%/5%/10%/15%
p50/mean/p95/p99 thresholds.

Retain the 21-cell deterministic 1% selector as correctness and phase evidence
only. Its apparent p50 reduction is 73.16%-73.59%, but source-backed mean drift
and both implementations' tail drift exceed the relevant stability thresholds.
No 1% latency claim is accepted.

## Measurement contract

Harness commit `2d2c54685` records an aligned lifecycle plus open, staging,
commit, and sequential-publication phase vectors. Their checked sum never
exceeds the corresponding lifecycle sample. The timed path remains
uninstrumented. A separate untimed `InstrumentedSource` replay records logical
`ReadAt` range overlap and re-runs the complete source-backed lifecycle; it is
not physical-I/O evidence.

Source-backed-only gates are nullable and omitted from eager JSON records. They
cover source-bound patch forward/inverse, foreign-source refusal, exact no-op,
replacement limit, exact typed partial-sink progress, raw untouched-member
identity, and source immutability. Stale/version and cancellation refusal,
signed/protected packages, formulas, unknown or repeated rows, and transaction
bounds remain production-test evidence and are explicitly outside this
selector.

## Clean release A/B/B/A

All four legs use clean revision `2d2c546856936dd643619957d16b1962f1ddaad2`
and release binary SHA-256 `94927d403afca7fa77f9404113531b82850b2df3211e7e410172b8bdb0854c8f`.
The host exposes CPU 2 only, with 20 warmups and 500 retained samples per
workload and leg. A1/A2 are eager owned snapshots; B1/B2 are source-backed.
Positive values mean lower source-backed latency.

| Workload | Metric | A1 -> B1 | B2 -> A2 | Eager drift | Source drift | Decision |
|---|---|---:|---:|---:|---:|---|
| one existing cell | p50 | 75.03% | 74.27% | 4.13% | 1.19% | accepted |
| one existing cell | mean | 75.14% | 74.33% | 4.35% | 1.24% | accepted |
| one existing cell | p95 | 76.68% | 74.67% | 9.24% | 1.40% | accepted |
| one existing cell | p99 | 76.37% | 74.76% | 7.16% | 0.83% | accepted |
| 21 existing cells (1%) | p50 | 73.59% | 73.16% | 2.69% | 4.38% | withheld with workload |
| 21 existing cells (1%) | mean | 73.57% | 73.10% | 4.03% | **5.86%** | withheld |
| 21 existing cells (1%) | p95 | 73.50% | 72.83% | **11.27%** | **14.06%** | withheld |
| 21 existing cells (1%) | p99 | 73.47% | 73.09% | **16.75%** | **18.41%** | withheld |

Both implementations emit the same deterministic hashes for each workload.
The 16,790,689-byte source contains two sheets, 2,048 cells, and eight opaque
media resources. The fixed sink retains zero output bytes and at most a 16 KiB
authoring window. The source replay is deterministic at 617 logical reads and
16,801,025 logical bytes per sample, but those values do not describe physical
device traffic or decompression.

Descriptive phase p50 values locate the one-cell difference in staging and
commit: eager A1/A2 stage at 16.82/15.96 ms and commit at 37.02/35.57 ms,
while source-backed B1/B2 stage at 1.72/1.70 ms and commit at 2.74/2.65 ms.
Source-backed sequential publication is slightly higher at 8.48/8.41 ms versus
7.05/6.97 ms eager. These phase medians are non-additive descriptive
attribution; the accepted claim remains the aligned complete lifecycle.

## Verification and scope

- focused four-selector test, including a two-sample source-backed record and
  serialized backend-specific field presence/absence;
- strict all-target harness Clippy with warnings and deprecations denied;
- exact revision, affinity, sample-cardinality, phase-sum, hash, semantic/media,
  sink, patch, refusal, and source-counter gates across all four raw reports;
- independent current-tree evidence-contract review.

No allocation/RSS, physical-I/O, decompression, cache-temperature,
real-producer, durable ZIP patch, atomic-save, formula, merge, structural-row,
insert/delete, or broad ODS CRUD claim is made.

Artifacts:

- [summary](../results/ods-source-cell-0177-summary.json)
- [manifest](../results/ods-source-cell-0177-manifest.json)
- raw A1/B1/B2/A2 reports listed in the manifest
