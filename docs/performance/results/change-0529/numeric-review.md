# 0529 XML-minifier pilot numeric review

Status: pilot rejected; conditional lanes remain unmeasured.

Frozen plan SHA-256: `17988e94a8d0993a6b453ca8c3ec3c46689bc89f2f823cace045a47f92b06a33`.
Comparison SHA-256: `70b58d3e35525883f8f122d27cc429cb291e00a0e6c2e3fde3a1c9b6107f1722`.

The canonical comparison contains 4 primary pairs, 4 publication allocation pairs, 18 native timing pairs, 46 matched adverse flags, and 76 same-build drift flags. Every flag is retained individually in `adverse-review.json` with all phase statistics and whole-child RSS context.

## Pilot result

| repeat | shape | total p50 | total mean | publication p50 | commit p50 diagnostic | row |
| ---: | --- | ---: | ---: | ---: | ---: | --- |
| 1 | dense-sparse | +0.537344% (need 2.0%) | +0.498459% (need 2.0%) | -0.082301% (need 5.0%) | +2.773597% | fail |
| 1 | medium | -0.766658% (need 2.0%) | -0.910607% (need 2.0%) | -0.016301% (need 5.0%) | -2.038794% | fail |
| 2 | dense-sparse | +1.589595% (need 2.0%) | +1.495296% (need 2.0%) | +5.157368% (need 5.0%) | +0.994648% | fail |
| 2 | medium | +2.130057% (need 2.0%) | +1.371701% (need 2.0%) | +8.532514% (need 5.0%) | +0.204601% | fail |

All four primary total-mean rows fail the 2% gate. Both repeat-1 publication p50 rows are effectively flat negative reductions (-0.082301% and -0.016301%); repeat 2 passes that gate at +5.157368% and +8.532514%, but the total gates still reject the pilot. Commit p50 is diagnostic only under the frozen plan.

| repeat | shape | publication allocation-call p50 reduction | publication realloc p50 delta | row |
| ---: | --- | ---: | ---: | --- |
| 1 | dense-sparse | +98.869532% (need 8.0%) | +0.0 | pass |
| 1 | medium | +97.848904% (need 8.0%) | +0.0 | pass |
| 2 | dense-sparse | +98.869532% (need 8.0%) | +0.0 | pass |
| 2 | medium | +97.848904% (need 8.0%) | +0.0 | pass |

All four publication allocation-call rows pass, with medium at +97.848904% and dense-sparse at +98.869532%; repeats are identical. Publication `incremental_region_peak_live_bytes` p50 is unchanged in all four pairs (0% change). Publication reallocation calls are also unchanged (40 to 40) and remain diagnostic.

The separate `commit_allocation_metrics` vectors are retained as diagnostics. Their allocation, deallocation, reallocation, byte, and incremental peak vectors are identical between baseline and candidate for all four pairs, so those counters do not explain or replace the publication allocation result.

## Flag review

The 46 matched flags are all native timing flags; no RSS flag was emitted. They cover elapsed, open, plan, publication, and reopen statistics. Publication flags are retained as possible candidate-facing audit-path observations. Open, plan, and reopen flags are retained as unresolved outside-path diagnostics; elapsed flags retain their component phase context. The 76 same-build flags are retained repeat-drift observations with the same full phase and RSS context. No flag is dismissed as noise.

Because the pilot is rejected, publication-Ir profiling, hardware, and eager lanes are unmeasured and cannot support a performance claim.
