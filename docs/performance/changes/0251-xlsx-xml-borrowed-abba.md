# Change 0251: XLSX borrowed XML worksheet parsing ABBA evidence

Date: 2026-08-21

Status: latency-supported on the listed cells, but HELD pending the allocator/
resource guardrail; the candidate is not landed and no broad XLSX claim is
made

## Evidence-package identity

The independently audited package is retained outside the repository at
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0251-xlsx-xml-borrowed-20260821/`.
Its manifest is `0251-xlsx-xml-borrowed-abba-manifest.json` with SHA-256
`06fbd58b2d9ae8fe460c656bc7316b1493ed78f47ec5d5c91e91d4c088e8c00b`.
The `summary.json` file SHA-256 is
`43cb20b1c68273cf42492b78f4c66a780b54e898bba9949ce082fc7d33291283`; its
canonical SHA-256 is
`c6045e6bd17d9181346692aa0463eda4b13c3b78f42f8e65b13990d17b3a04f9`.
The package is schema version 1, has change ID
`0251-xlsx-xml-borrowed-abba`, and contains four result rows from
`litchi-perf-abba-summary` 0.1.0.

The retained package members have these identities:

| Member | File SHA-256 | Canonical SHA-256 |
|---|---|---|
| `0251-xlsx-xml-borrowed-abba-manifest.json` | `06fbd58b2d9ae8fe460c656bc7316b1493ed78f47ec5d5c91e91d4c088e8c00b` | — |
| `summary.json` | `43cb20b1c68273cf42492b78f4c66a780b54e898bba9949ce082fc7d33291283` | `c6045e6bd17d9181346692aa0463eda4b13c3b78f42f8e65b13990d17b3a04f9` |
| `a1.json.zst` | `5a3a10bbdc9874de03d7914892890c483c1fa5be2c7019ca423d417bc78b2d42` | `4e586a1767e66ac3838ee8c4f5bfeb7f782ee5ba3ddbbeb66851bdefc46ce5bd` |
| `a2.json.zst` | `dd48891bae30664cc49d484480ab0a45c7490360cabf4033ae3ff8392f2a3654` | `f717916fb6d66504f19c75ad0363047861496d211ba06220ca3cd41e578e4826` |
| `b1.json.zst` | `25a63e377e538bf2b52d8e5883902841a426032d8819d6afd6864f7e67ad1e1e` | `2d288ef20afa077eca1cd52a17ce33af38d57213f46189f98675c50d28b3aaf4` |
| `b2.json.zst` | `8b6aa7aac6a97e2269f96b32583a1f6247448cd9dcc71f87d469bf969ea3d8ed` | `ee55db9cb1660e2b3b4e4e925ba2a6a6ea99bad3a8bb9b568d85cee65f90a5c6` |

The matched control revision is `1ac7d8d8b54695354d93220a1be2cafe912802b8`
(A1/A2); the candidate revision is
`2c5e46eba488271c722c4e7ec69ac4ca9615d9da` (B1/B2), measured over the
candidate range `1ac7d8d8b..2c5e46eba`. The harness is the release
`litchi-perf-baseline` profile, schema 1, x86_64 Linux; the summarizer is
`litchi-perf-abba-summary` 0.1.0. Both implementations report clean
worktrees, CPU affinity 2, one visible logical CPU, Rust 1.95.0, and the
recorded AMD EPYC host.

Case/corpus, configuration, stable-environment, and statistic-recomputation
checks pass. Per-row identity fields retain their applicable status: values
that are produced compare equal, while inapplicable source, sink, output, or
operation metrics are consistently absent. This package therefore does not
turn absent metrics into identity claims.

## Candidate and protocol

The candidate borrows `quick-xml` worksheet events through the worksheet codec
and bounds worksheet scans by XML depth and event-count limits, while retaining
namespace/error behavior. The focused worksheet codec, namespace, malformed
markup, and scan-bound suites passed for the candidate.

The order is A1 control, B1 candidate, B2 candidate, A2 control on CPU 2,
with 30 warmups and 500 retained samples per selector and leg. The matched
selectors are:

| Selector | Corpus |
|---|---|
| `xlsx_eager_cell_values_one_edit_save` | `xlsx-cell-values-medium`, four 48×48 sheets with media |
| `xlsx_source_backed_cell_values_one_edit_save` | `xlsx-cell-values-medium`, four 48×48 sheets with media |
| `xlsx_first_cell` | `xlsx-tiny`, three 8×8 sheets |
| `xlsx_source_first_cell` | `xlsx-tiny`, three 8×8 sheets |

The medium corpus is a 4,226,429-byte, 17-member XLSX archive with archive
SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`,
4,231,168 uncompressed bytes, and target `Sheet1!A1`. The tiny corpus is a
3,561-byte, 8-member XLSX archive with archive SHA-256
`69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428`,
768 uncompressed bytes, and the same target. Drift ceilings are
5%/5%/10%/15% for p50/mean/p95/p99. Positive readings mean lower candidate
elapsed time.

## Exact paired latency readings

Three selectors accept all four statistics: both one-edit-save selectors and
`xlsx_source_first_cell`. `xlsx_first_cell` accepts p50, mean, and p99; its
p95 is rejected solely because candidate drift is
`+10.941931063360043%` (10.941931% at six decimals), above the 10% p95
ceiling. No selector has an adverse-both statistic.

### `xlsx_eager_cell_values_one_edit_save`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +7.358527311127923% | +6.389063984374453% | accepted |
| mean | +7.45121206455559% | +7.072223312335558% | accepted |
| p95 | +8.04860381035411% | +10.173420108318425% | accepted |
| p99 | +7.738146427228101% | +10.969448261274541% | accepted |

### `xlsx_source_backed_cell_values_one_edit_save`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +7.779092885678692% | +8.206212271370346% | accepted |
| mean | +7.492230083534052% | +8.252032222989628% | accepted |
| p95 | +7.643348251929688% | +8.120220230248181% | accepted |
| p99 | +6.693922794402211% | +8.068135540769896% | accepted |

### `xlsx_first_cell`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +13.104303555110745% | +11.356584065381758% | accepted |
| mean | +10.980406522477562% | +11.567382619595335% | accepted |
| p95 | +10.64455782312925% | +4.976932983388488% | rejected: candidate drift |
| p99 | +8.464539693959434% | +11.234168064310914% | accepted |

### `xlsx_source_first_cell`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +10.799519025006624% | +13.273386394475162% | accepted |
| mean | +13.258675748547644% | +14.407737136930605% | accepted |
| p95 | +21.40482388733808% | +16.922898127241336% | accepted |
| p99 | +22.48132628894152% | +28.688545755009297% | accepted |

These are 15 accepted cells and zero adverse-both cells. The three fully
accepted selectors and the accepted cells of `xlsx_first_cell` support the
latency direction for this matched run, but they do not clear the separate
allocator/resource guardrail required for production landing.

## Validation and claim boundary

The durable baseline test failure reproduces on the unmodified control base
`1ac7d8d8b`; it is unrelated to this candidate. The targeted worksheet/parser
suites passed, as did the matched evidence identity and statistic checks. The
durable baseline failure is not reclassified as a regression from this change.

Production remains HELD pending allocator/resource guardrail evidence. The
candidate is not landed. This record is limited to the listed matched XLSX
worksheet parsing selectors and their fixed corpora; it makes no broad XLSX
claim and no claim about allocations, RSS, physical I/O, filesystem or
cold-cache behavior, decompression, durable preservation, real producers, or
other CRUD paths. Identity fields that are consistently absent remain absent;
they are not inferred from the latency result.

Verification for this documentation-only update is Python package/hash/count
and link checking plus `git diff --check`; no Cargo command or benchmark is
part of this record.
