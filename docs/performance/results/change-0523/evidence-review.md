# 0523 numerical evidence review

`scope: independent read-only review of the final 0523 evidence`

`performance_claim: none`

I reviewed the retained native, allocation, profile, hardware, variation and
verification artifacts, the 0523 change record, and the 0523 entries in the
global performance ledgers. This review only read artifacts and recomputed
their arithmetic; it did not run a build, test, capture, or source edit.

## Receipt and scope result

The retained reports are internally consistent and carry the declared scope:

- `analysis.json` is `pass` with 24,000 native samples, 720 separate
  allocation samples, 48 rows, and seven retained same-build flags.
- `profile-analysis.json` is `pass` with eight profile children, 40 timed
  constructor dumps, six CFB setup dumps, and zero-instruction final process
  dumps. Its validation retains the positive owner edges, one selected
  constructor call per timed dump, native identity, and all requested CFB
  targets.
- `hardware-analysis.json` is `pass` for two whole-child diagnostic groups.
  Both groups have matching event runtime and 100% running time; their IPC is
  1.670903288 and 1.680879869. The 2,000 hardware-lane durations are excluded
  from the native matrix and no operation-local hardware claim is made.
- Post-cleanup `verification.json` is `pass`: source replay covers 8,584
  files and 32 serial intervals, all ten quality gates pass, the four negative
  vectors are rejected, and owned paths are absent. The retained first
  preflight source-custody mismatch remains separate from the final passing
  preflight.

The evidence identities reviewed here are:

| Artifact | SHA-256 |
| --- | --- |
| `analysis.json` | `dd134240f54106ba562def30697051387524b951552233564fc40cacb8137bfe` |
| `profile-analysis.json` | `07b939d028d3cdbf1d970552219dd876ddd12bb35260108de04f6f3ee4b962e3` |
| `hardware-analysis.json` | `22e6dbae30387fa15b3340f0fd9a91e757b02c325f11762a66641853583563c3` |
| `variation-review.json` | `4041ffefb79fb147fbdc37d7e9d1eb1811a700f7dadeaab88131f391cce26407` |

## Native and allocation arithmetic

The representative rows in the change record match the raw reports. The
allocation columns below are stable for every retained sample of the named
row in both repeats; allocation elapsed times remain separate from native
latency.

| Workflow | Native p50, repeat 1 / 2 | Allocation calls | Reallocation calls | Allocated bytes | Incremental peak bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| CFB tiny | 2.270 / 2.260 us | 32 | 1 | 4,456 | 3,384 |
| CFB many-small | 136.070 / 135.560 us | 544 | 7 | 214,186 | 193,026 |
| CFB few-large | 102.470 / 101.920 us | 29 | 1 | 213,571 | 205,129 |
| XLS owned-source open-one-cell | 130.495 / 131.086 us | 126 | 25 | 223,774 | 191,084 |

I checked all 24 allocation rows and 30 samples per row. Every vector has
the expected elapsed/sample permutation and alignment, every allocation
status is measured under `operation_global_system_allocator`, failed
allocation calls are zero, and every sample satisfies
`live_before + allocated - deallocated = live_after`. Region peaks cover both
live endpoints, and post-operation peaks never fall below pre-operation peaks.
There are no allocation balance or vector-length violations. The normal
reports retain explicit `allocation: unavailable` and CFB source metrics retain
`not_applicable`; neither is converted into a zero observation.

The four corpus identities (one XLS corpus and three CFB shapes), output
identities, and cross-lane/repeat bindings also match. These checks support a
current harness baseline and allocation attribution only; they do not create a
production before/after speedup result.

## Variation review

All seven flags in `variation-review.json` were independently recomputed as
`(repeat_2 - repeat_1) / repeat_1 * 100`. The results are:

| Case and metric | Repeat 1 -> repeat 2 | Change |
| --- | ---: | ---: |
| Eager one-cell p50 | 424,616 -> 450,197 ns | +6.024502% |
| Eager one-cell p95 | 440,812 -> 468,822 ns | +6.354183% |
| Eager one-cell p99 | 446,172 -> 476,642 ns | +6.829205% |
| Eager one-cell mean | 425,230.497 -> 450,668.031 ns | +5.982058% |
| Owned-source open p95 | 136,300 -> 144,341 ns | +5.899486% |
| Owned-source open p99 | 139,591 -> 147,131 ns | +5.401494% |
| Source-backed open p95 | 141,951 -> 149,681 ns | +5.445541% |

The flags are all same-build second-child increases on a shared host. They are
retained, with no sample deletion or causal/stable-tail claim. No old 0511
control or final profile is used as a current timing comparison, and the old
9.30% `load_fat` share is not reused for current ranking.

## Profile attribution

The fresh profile percentages use the summed inclusive instruction references
of the selected constructor as the denominator and the named function's
exclusive `self_ir` as the numerator. The aggregate arithmetic is:

| Workload | Constructor Ir | `collect_exact` | `claim_sector` | Stream allocation validation | Physical reconciliation | `load_fat` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| XLS owned source | 27,947,884 | 11,202,280 (40.0827%) | 4,979,100 (17.8157%) | 3,958,820 (14.1650%) | 3,983,420 (14.2530%) | 501,640 (1.7949%) |
| CFB tiny | 490,284 | 10,400 (2.1212%) | 900 (0.1836%) | 8,570 (1.7480%) | 860 (0.1754%) | 3,490 (0.7118%) |
| CFB many-small | 28,061,254 | 1,528,970 (5.4487%) | 92,100 (0.3282%) | 1,074,490 (3.8291%) | 73,820 (0.2631%) | 11,130 (0.0397%) |
| CFB few-large | 26,145,238 | 11,143,890 (42.6230%) | 4,954,650 (18.9505%) | 3,935,640 (15.0530%) | 3,963,860 (15.1609%) | 533,630 (2.0410%) |

The per-repeat XLS `collect_exact` shares are 40.0883% and 40.0772%, which
matches the documented 40.08–40.09% range. The documented CFB shares of
42.62%, 2.12%, and 5.45% match the fresh aggregate values above. These are
exclusive descendant diagnostics with disjoint self-cost numerators; inclusive
parent/child costs must not be added. The CFB setup dump is excluded from the
five timed dumps per child, and the retained Valgrind overflow warning is
carried as a warning rather than a claim.

The current change record and global `REPORT.md`, `ADR_COMPLIANCE.md`,
`GOAL_AUDIT.md`, `HOTSPOTS.md`, and `BASELINE.md` entries agree with these
counts, percentages, boundaries, and limitations. They correctly identify the
checked visited-bit fusion as a next private CFB experiment, preserve sector
ownership and physical validation, keep OLE2/OOXML first, and defer ODF while
excluding iWork.

## Finding

There is no substantive numerical or scope error in the final evidence. One
minor clarity issue remains in the change record and its matching baseline
summary: `XLS owned-source open/one-cell` is shorthand for the exact
`xls_owned_source_open_one_cell` row, whose 126 allocation calls and
130.495/131.086-us native p50 values are correct. The plain
`xls_owned_source_open` row has 124 allocation calls and different p50 values.
Using the exact case name in the table would prevent those two workflows from
being conflated; it does not affect any calculation or claim.

Root resolved the clarity note before sealing: the representative change-record
table and baseline summary now use `xls_owned_source_open_one_cell` exactly.
