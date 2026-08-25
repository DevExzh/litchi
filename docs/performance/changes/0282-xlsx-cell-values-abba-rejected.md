# Change 0282: XLSX scalar cell-values ABBA rejected

**Date:** 2026-08-25
**Status:** Rejected after evidence run
**Performance claim:** none
**Retained samples:** 500 per cell and leg (12,000 successful rows)

## Decision

Change 0282 retained the clean release ABBA run for the existing eager and
source-backed XLSX scalar cell-value publishers, but the predeclared
all-cell authorization failed. `medium/one_edit` passed all 16 strict gates.
The other five cells failed their directional gates because source-backed
total latency was generally about 7-8% adverse. `medium/batch` additionally
failed the eager A1/A2 p99 same-side gate at `8.309521161168686%` against the
5% ceiling.

The protocol does not authorize partial claims, so the independently passing
`medium/one_edit` cell is diagnostic only. Across all cells, 56 of 96
individual timing gates passed and 40 failed. `failures.jsonl` contains the
single aggregate gate-failure row; it contains no child correctness failure.

All 12,000 retained rows passed child-schema, corpus, provenance, sink, and
same-side identity validation. Eager A1/A2 output hashes were identical for
every cell. Source B1/B2 output and semantic hashes, untouched-member evidence,
logical counters, cache diagnostics, and budget diagnostics were neutral for
every cell. Independent validation recomputed every aggregate and reported
12,190 of 12,190 structural/statistical checks passing.

The implementation and initial runner revision was `1da3e4141`. The corrected
smoke-oracle and clean evidence revision was
`3b826423bb073cf4e2edc854c5a1bdd09dbebbcc` (`3b826423b`). The run used a clean
detached worktree, release binaries, `taskset` CPU 2, one serial orchestrator,
fresh children for every leg, 20 warmups, and 500 retained samples per cell.

## Frozen protocol

The retained protocol is schema `litchi.xlsx.cell-values-abba.v1`, produced by
`xlsx_cell_values_abba` package version `0.1.0`. It measured these exact cells
for both `medium` and `dense-sparse` deterministic workbooks:

- One scalar cell edit.
- One-percent scalar cell edits.
- The exact bounded batch of 256 scalar cell edits.

Every sample used this fixed fresh-child order:

| Leg | Implementation |
|---|---|
| A1 | eager |
| B1 | source-backed |
| B2 | source-backed |
| A2 | eager |

The primary value was exactly the existing child `elapsed_ns.samples[0]`
total open/edit/save interval. Phase vectors were not compared because eager
and source staging are assigned to different child phase intervals. The
same-side gate compared symmetric deltas of aggregate p50, mean, p95, and p99
and required at most 5%. Directional A1-to-B1 and A2-to-B2 p50 and mean each
required at least 1% or 50 us source improvement; p95 and p99 allowed at most
5% adverse change.

Eager and source publication intentionally produce byte-distinct valid ZIP
packages. Cross-implementation ZIP identity was therefore excluded before the
authoritative run. Each child validated its own result. The runner required
same-side eager output identity and same-side source output, semantic,
untouched-member, and logical-counter identity.

## Pinned corpora and provenance

Both corpora contain four worksheets, 17 archive members, eight deterministic
512 KiB media parts, and `Sheet1!A1` as the target entry.

| Shape | Stored cells | Archive bytes | Archive SHA-256 |
|---|---:|---:|---|
| Medium | 9,216 | 4,226,429 | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` |
| Dense-sparse | 17,792 | 4,251,863 | `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a` |

The protocol provenance was:

| Field | Value |
|---|---|
| Benchmark SHA-256 | `fb095d2f55d5c272865b8e3d98903cfee8cc8c85425b4d79c5ba731d87e64816` |
| Benchmark bytes/profile | 51,073,272 / `release` |
| Runner SHA-256 | `14ed12cf8050eef6caf98911bc90fe2213b68ee18656a3c448afe3be48e099ea` |
| Git revision | `3b826423bb073cf4e2edc854c5a1bdd09dbebbcc` |
| Git dirty | `false` |
| Rust compiler | `rustc 1.95.0 (59807616e 2026-04-14)` |
| Target OS/arch | `linux` / `x86_64` |
| Memory | `MemTotal: 32812712 kB` |
| CPU environment | `null` in protocol; execution was pinned to CPU 2 with `taskset` |
| `RUSTFLAGS` | `null` |

The retained process interval was `1787647465188` through `1787657283419`
Unix milliseconds, or 9,818.231 seconds.

## Aggregate timing

The exact per-leg aggregates used by the gates were:

| Cell | Leg | p50 (ns) | Mean floor (ns) | p95 (ns) | p99 (ns) |
|---|---|---:|---:|---:|---:|
| Medium one edit | A1 eager | 7,888,014 | 7,972,717 | 8,595,818 | 9,117,555 |
| Medium one edit | B1 source | 7,711,040 | 7,803,777 | 8,469,170 | 9,051,370 |
| Medium one edit | B2 source | 7,701,902 | 7,784,841 | 8,335,828 | 9,228,530 |
| Medium one edit | A2 eager | 7,867,886 | 7,947,712 | 8,591,239 | 9,173,420 |
| Medium one percent | A1 eager | 27,111,523 | 27,202,870 | 28,981,757 | 30,133,335 |
| Medium one percent | B1 source | 29,266,366 | 29,352,447 | 30,879,444 | 32,158,049 |
| Medium one percent | B2 source | 29,234,808 | 29,372,496 | 30,924,771 | 32,314,891 |
| Medium one percent | A2 eager | 27,049,465 | 27,184,354 | 28,790,954 | 30,596,329 |
| Medium batch | A1 eager | 27,247,386 | 27,418,224 | 29,065,461 | 32,201,442 |
| Medium batch | B1 source | 29,409,772 | 29,547,966 | 31,220,566 | 33,175,325 |
| Medium batch | B2 source | 29,310,213 | 29,502,833 | 31,129,010 | 32,276,245 |
| Medium batch | A2 eager | 27,263,952 | 27,344,860 | 28,802,908 | 29,730,943 |
| Dense-sparse one edit | A1 eager | 48,489,888 | 48,616,726 | 51,296,681 | 53,990,978 |
| Dense-sparse one edit | B1 source | 52,104,479 | 52,424,356 | 55,598,624 | 58,087,614 |
| Dense-sparse one edit | B2 source | 52,073,509 | 52,295,341 | 54,781,910 | 56,777,640 |
| Dense-sparse one edit | A2 eager | 48,592,185 | 48,725,461 | 51,075,731 | 53,020,386 |
| Dense-sparse one percent | A1 eager | 53,002,543 | 53,119,435 | 55,892,100 | 58,081,291 |
| Dense-sparse one percent | B1 source | 57,192,839 | 57,391,417 | 60,233,104 | 65,371,293 |
| Dense-sparse one percent | B2 source | 57,058,780 | 57,353,615 | 60,542,689 | 62,896,988 |
| Dense-sparse one percent | A2 eager | 53,172,865 | 53,288,222 | 55,981,106 | 58,081,681 |
| Dense-sparse batch | A1 eager | 52,974,213 | 53,197,616 | 56,072,621 | 58,260,114 |
| Dense-sparse batch | B1 source | 57,043,664 | 57,230,526 | 59,884,427 | 61,571,965 |
| Dense-sparse batch | B2 source | 57,011,808 | 57,375,812 | 60,271,272 | 63,074,020 |
| Dense-sparse batch | A2 eager | 52,949,024 | 53,116,630 | 55,795,647 | 58,049,969 |

The all-cell gate outcomes were:

| Cell | Same-side gates | Directional gates | Result |
|---|---:|---:|---|
| `medium/one_edit` | 8/8 | 8/8 | pass, diagnostic only |
| `medium/one_percent` | 8/8 | 0/8 | fail |
| `medium/batch` | 7/8 | 1/8 | fail |
| `dense-sparse/one_edit` | 8/8 | 0/8 | fail |
| `dense-sparse/one_percent` | 8/8 | 0/8 | fail |
| `dense-sparse/batch` | 8/8 | 0/8 | fail |

The adverse source p50 observations were 7.16-8.08% across the five failing
cells and both directional pairings. Tail failures were also widespread; the
largest was `dense-sparse/one_percent` A1-to-B1 p99 at 12.551377344556624%
adverse. These values are diagnostic observations, not accepted claims.

## Claim boundary and exclusions

This package is retained as rejected evidence. It authorizes no source-backed
performance win, no `medium/one_edit` partial claim, and no pooled or
selector-wide claim. The smoke runs are not retained as evidence.

The result makes no claim about:

- Phase timing or causal attribution.
- Cross-implementation ZIP byte identity.
- Cross-revision behavior.
- Physical I/O, allocation, RSS, CPU utilization, or cache warmth.
- Managed-budget behavior or total resource bounds.
- Formula, date, structural, unsupported, or general XLSX CRUD behavior.
- Real-producer interoperability.

## Retained evidence

Evidence:
`docs/performance/results/0282-xlsx-cell-values-abba-rejected-20260825/`

The package contains exactly these seven artifacts:

- `artifact-manifest.json`
- `failures.jsonl`
- `process-time.txt`
- `protocol.json`
- `samples.jsonl`
- `sha256.txt`
- `summary.json`

Manifest SHA-256:
`3e9a0c42c42b0a5de76b5b1c27fb155e8f4bb1fb456ef0def27d39336c84bdd6`

Manifest bytes: 980
