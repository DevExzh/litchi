# 0734 PPT owned-stream results review

Disposition: **retain the bounded optimization** for the ordinary public PPT
slide-index-1 removal workflow covered by `45543.ppt` and `41246-1.ppt`.
The sealed evidence shows a concrete allocation reduction, no central latency
regression above the review threshold, and one retained primary tail limitation.
This disposition is scoped to the measured source-owned CFB payload path.

## Evidence custody and independent replay

The packet is complete and internally consistent. [`analysis.json`](analysis.json)
and [`audit.json`](audit.json) both report `passed`; the audit reports
`analysis_match: true` and 48 processes. The capture manifest is complete with
36 native processes (nine paired process repeats per case) and 12 allocation
processes (three paired repeats per case). The synthetic preflight also passed;
its values are schema evidence and are not included as performance evidence.

| receipt | SHA-256 |
| --- | --- |
| `analysis.json` | `6d8f01ca477f0b5d8c2dee88152c851b0d5f44424f8e0730030e4e6c0ada3a4d` |
| captured process packet | `85989a5caa0c0c281c095e9097806aef7d6ce9f0cae5fe7a576135d8b11f650e` |
| frozen inputs and scripts | `43abb88e5ec1d986482c8128e482845b6bd93f14cb65a05bd65dd1383bb57980` |

I independently replayed the 48 raw capture JSONs without running Cargo,
native probes, or a profiler. The replay verified every manifest SHA and plan
row, zero exit codes, 50 samples and three warmups for each native process,
one sample and no warmup for each allocation process, and exact raw output
SHA/inventory equality within every process. All 1,812 measured owner outputs across the 48 processes had a
passing semantic oracle, and all 384 retained negative-control records (eight
per process) were rejected. Recomputed native p50/mean/p95/p99/maximum values,
all 18 paired differences and flags, and all allocation comparisons agree with
the sealed reports. [`qualification.json`](qualification.json) passed. The
quality summary has 1,211 target tests, 14 doctests, and four baseline plus
four candidate probe tests passed, with zero failures.

The 19 actual corruption controls in [`negative-checks.json`](negative-checks.json)
were all rejected. Eighteen reached both analyzer and audit; the retained
`quality command` control is analyzer-only and is recorded as a
coverage limitation rather than a measurement result. The post-cleanup replay
also passed all 19 controls.

## Timing result

The paired unit is a process, not an individual sample. The analysis uses the
midpoint p50 and a 10,000-resample percentile bootstrap over the nine process
pairs per case. A flag means an absolute paired change above 5%; its sign is
still reviewed.

| case | paired p50 median | bootstrap percentile interval for paired p50 median | paired mean median | flags and interpretation |
| --- | ---: | ---: | ---: | --- |
| primary | -5.548% | [-6.213%, -4.963%] | -7.570% | All nine p50 and mean pairs improve. Seven p50 pairs and all nine mean pairs exceed the 5% threshold. |
| secondary | +0.349% | [-0.412%, +0.873%] | -0.348% | No >5% flag. The interval crosses zero, so this fixture supplies no speedup claim. |

The primary p50 bootstrap interval reaches `-4.963%` at its upper endpoint,
so the aggregate p50 result is near the 5% boundary even though its paired
median is `-5.548%`; this is retained as context rather than rounded away.
The bootstrap intervals for the **mean of paired mean changes** are
`[-7.713%, -7.233%]` for primary and `[-0.413%, +0.151%]` for secondary.
These use the mean reducer; the main report separately reports the bootstrap
interval for the median of paired mean changes. Both secondary intervals
cross zero.

The only positive native flag above 5% is primary cycle 2, repeat 0:
nearest-rank p99 and maximum are both `+9.560%` (1,282,706 ns after versus
1,170,776 ns before). With 50 samples, p99 equals maximum in that process;
these are two reported fields for one tail event, not two independent
regressions. There is no primary p50 or mean regression and no secondary flag.
All native flags in [`analysis.json`](analysis.json) remain reported, including
the negative flags that represent improvements.

## Allocation result

All three allocation repeats for each case and variant are identical. The
change removes exactly five allocation calls in each fixture and removes the
selected payload byte total from both allocated and deallocated bytes.

| case | allocated bytes (before → after) | deallocated bytes (before → after) | allocation calls (before → after) | peak live bytes (before → after) | retained bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| primary | 11,776,674 → 11,391,768 (**-384,906; -3.268%**) | 11,386,530 → 11,001,624 (**-384,906; -3.380%**) | 5,663 → 5,658 (**-5; -0.088%**) | 2,686,521 → 1,992,885 (**-693,636; -25.819%**) | 390,144 → 390,144 |
| secondary | 10,724,448 → 10,444,631 (**-279,817; -2.609%**) | 10,439,264 → 10,159,447 (**-279,817; -2.680%**) | 18,667 → 18,662 (**-5; -0.027%**) | 1,976,223 → 1,976,223 (**0%**) | 285,184 → 285,184 |

The only allocation review flag is primary `peak_live_bytes`, and it is a
25.819% reduction. No peak-live growth occurred. `retained_bytes` is unchanged
in both cases. These counters are boundary-relative allocator measurements;
they are not RSS or a resident-memory claim.

## ADR limits and retention scope

The result stays within the accepted constraints in [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md), [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), [ADR 0008](../../../adr/0008-migration-and-verification.md), [ADR 0024](../../../adr/0024-current-topology.md), and [ADR 0026](../../../adr/0026-ole-directory-metadata-binding.md):

- The [source review](source-review.md) found no ownership, duplicate-path fallback, validation,
  typed-failure, public-API, or publication-boundary change that blocks the
  measured candidate. Output identity, stream inventory, directory policy,
  semantic reopen, and negative controls pass for both fixtures.
- The concrete benefit is limited to the existing `litchi-ppt`/CFB finish path:
  384,906 fewer allocated bytes and five fewer calls on the primary, and
  279,817 fewer bytes and five fewer calls on the secondary. No cache, policy,
  executor, or ambient-I/O behavior is inferred.
- This matrix covers two fixtures, one public operation, one host, and
  serialized CPU-pinned processes. It makes no RSS, cold-I/O, instruction,
  concurrency, throughput, or broad producer/CRUD claim. Samples within a
  process are not independent process observations.

Retain the candidate with the primary `+9.560%` p99/maximum event and the
secondary inconclusive timing result recorded as explicit limitations. Further
PPT optimization work should keep the same validation, custody, and flag
reporting requirements and should measure any different producer or workflow
separately.
