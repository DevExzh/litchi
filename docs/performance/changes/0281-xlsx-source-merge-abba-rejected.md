# Change 0281: bounded XLSX source-backed merge ABBA rejected

**Date:** 2026-08-25
**Status:** Rejected after evidence run
**Performance claim:** none
**Retained samples:** 1,000 per operation and leg (8,000 successful rows)

## Decision

Change 0281 retained the bounded XLSX source-backed merge/unmerge ABBA run, but
the predeclared all-operation authorization failed. Merge passed every strict
same-side and directional gate. Unmerge failed only the eager A1/A2 p99
same-side gate at `9.983807185381298%` and the source B1/B2 p99 same-side gate
at `5.578546712802768%`. The run therefore has no partial merge claim, even
though its merge result passed independently.

All 8,000 retained rows passed child-schema validation, semantic reopen and
oracle checks, complete package and one-part preservation checks, bounded sink
checks, and the source forward/inverse, stale, and foreign-source checks. The
retained failure file contains the single predeclared gate failure row.

The production revision was `df7261f21`. The runner and clean evidence
revision was `30d6c74c2026b5df6527582773f8ad28837bcfa8` (`30d6c74c2`). The run
used a clean detached worktree, `taskset` CPU 2, the release profile, 20
warmups, and 1,000 samples per cell.

## Frozen protocol

The retained protocol is schema `litchi.xlsx.source-merge-abba.v1`, produced by
`xlsx_merge_source_abba` package version `0.1.0`. It used evidence mode with a
20-sample evidence minimum, fresh child processes for every leg, and this fixed
order:

| Leg | Implementation |
|---|---|
| A1 | eager |
| B1 | source-backed |
| B2 | source-backed |
| A2 | eager |

Both `merge` and `unmerge` ran at every leg. The primary interval was commit
plus publication into a bounded retaining 64 KiB sink. Preparation, lifecycle,
semantic reopen, and logical `ReadAt`/length/version counters were recorded as
diagnostics. The same-side gate compared symmetric deltas of aggregate p50,
mean, p95, and p99 values and required at most 5%. Directional A1->B1 and
A2->B2 comparisons were claim-bearing only inside this same-revision
eager-vs-source path scope: p50 and mean required at least 1% or 50 us source
improvement, while p95 and p99 allowed at most 5% adverse change.

`cross_revision_evidence` was false. The protocol's initial and final
`performance_claim` was `none` because all-operation authorization failed.

## Pinned corpus and provenance

The deterministic fixture used `Sheet1`, populated cells A1 and C1, and merge
range `A1:B2`.

| Fixture | Bytes | Archive SHA-256 | Worksheet target SHA-256 |
|---|---:|---|---|
| Merge | 2,421 | `151fed9651e6f88a1e7e17183c8dac1f4885b6a922756214295b0d7c828a589e` | `692d1a1b71bd6af7bffc8d008d28ba75c5068dc913e0452363950fbe09d5b605` |
| Unmerge | 2,443 | `6329afb234f9f1ea073e37baa4ca9ab6a0bb559fd40ff46a94320d416296a03f` | `467a33b4a4635f43d4d1a582ed8f31390a0eb6b71c94fc13059de9e5c436798e` |

The protocol provenance was:

| Field | Value |
|---|---|
| Executable SHA-256 | `6636c541a2b32fd04409a00c7f2e91ae5fe793020ada444dbeb62968c40434dd` |
| Git revision | `30d6c74c2026b5df6527582773f8ad28837bcfa8` |
| Git dirty | `false` |
| Rust compiler | `rustc 1.95.0 (59807616e 2026-04-14)` |
| Target OS/arch | `linux` / `x86_64` |
| Memory | `MemTotal: 32812712 kB` |
| CPU environment | `null` in protocol; execution was pinned to CPU 2 with `taskset` |
| `RUSTFLAGS` | `null` |

## Aggregate timing

The exact per-leg aggregate values used by the gate calculations were:

| Operation | Leg | p50 (ns) | Mean floor (ns) | p95 (ns) | p99 (ns) |
|---|---|---:|---:|---:|---:|
| Merge | A1 eager | 190838 | 193113 | 224510 | 248883 |
| Merge | A2 eager | 191879 | 194898 | 228115 | 253689 |
| Merge | B1 source | 112404 | 115697 | 138131 | 155524 |
| Merge | B2 source | 111649 | 114420 | 134456 | 151969 |
| Unmerge | A1 eager | 165343 | 166520 | 193887 | 208117 |
| Unmerge | A2 eager | 165213 | 167614 | 197362 | 228895 |
| Unmerge | B1 source | 108234 | 110943 | 135828 | 152561 |
| Unmerge | B2 source | 107072 | 109419 | 132712 | 144500 |

The exact symmetric deltas of those aggregate values, which were the
same-side gate inputs, were:

| Operation | Side | p50 (%) | Mean (%) | p95 (%) | p99 (%) | Result |
|---|---|---:|---:|---:|---:|---|
| Merge | A1/A2 eager | 0.5454888439409342 | 0.9243292787124636 | 1.6057191216426885 | 1.9310278323549621 | pass |
| Merge | B1/B2 source | 0.6762263880554237 | 1.1160636252403426 | 2.7332361516034984 | 2.339292882100955 | pass |
| Unmerge | A1/A2 eager | 0.07868630192539328 | 0.6569781407638723 | 1.7922810709330692 | 9.983807185381298 | fail at p99 |
| Unmerge | B1/B2 source | 1.0852510460251046 | 1.392811120554931 | 2.3479414069564166 | 5.578546712802768 | fail at p99 |

The pooled eager-vs-source deltas were descriptive diagnostics, not a pooled
authorization rule:

| Operation | p50 (%) | Mean (%) | p95 (%) | p99 (%) |
|---|---:|---:|---:|---:|
| Merge | -41.401340485377006 | -40.692765650369836 | -39.822529743433996 | -38.60813958078396 |
| Unmerge | -34.828925399859656 | -34.04981235073353 | -31.547339393320396 | -31.400612128251936 |

## Claim boundary and exclusions

This package is retained as a rejected evidence result. It does not authorize
a merge-only or pooled performance claim. The 20/100 exploratory runs were not
retained.

The result makes no claim about:

- Cross-revision behavior or any result outside the pinned same-revision path.
- Materialization, resource usage, or a broader resource-bound claim.
- Eager durable patch apply/inverse, which the runner explicitly marked unavailable.
- Broad XLSX CRUD behavior beyond this bounded merge/unmerge fixture and path.

## Retained evidence

Evidence:
`docs/performance/results/0281-xlsx-source-merge-abba-rejected-20260825/`

The package contains exactly these seven artifacts:

- `artifact-manifest.json`
- `failures.jsonl`
- `process-time.txt`
- `protocol.json`
- `samples.jsonl`
- `sha256.txt`
- `summary.json`

Manifest SHA-256:
`c2c867e39204503074cdb8541543f1a24b79b3c184420686c4664989af047b12`

Manifest bytes: 981
