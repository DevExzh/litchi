# 0806 results and disposition review

## Disposition

The candidate is rejected for production retention. The independent raw audit and
the workflow analysis agree that the frozen useful-workflow benefit gate is not
met. The final disposition is recorded in
[`disposition.json`](disposition.json): `status` is `rejected` and
`production_change_retained` is `false`. The source has been restored; the
crate tree has no remaining diff from the restored source.

This review covers the terminal reports and the frozen policy. It does not rerun
timing or bootstrap calculations.

## Evidence and policy application

The terminal packet contains 306 reports and 6,714 samples: 216 native reports,
72 allocation reports, and 18 qualification reports. The independent
[`root-audit.json`](root-audit.json) reports the same counts, with no benefit
rows, no latency violations, and no resource violations. Its
`production_adoption` result is `false`.

The frozen [`adoption-policy.json`](adoption-policy.json) requires at least one
capture or lifecycle case with a minimum 3% improvement and a 95% bootstrap
upper ratio below 1.0. Allocation reductions alone cannot satisfy that gate.
The audited paired p50 ratios contain no eligible case. Representative results
are:

| Case | After/before p50 | Change |
| --- | ---: | ---: |
| large/capture | 1.018526 | +1.853% |
| large/lifecycle | 1.031050 | +3.105% |
| valid-4attr/capture | 1.021855 | +2.186% |
| large/commit | 0.989316 | -1.068% |

The first three eligible-workflow rows are slower, and the only listed
improvement is below the 3% threshold and is a commit row, which is outside the
benefit modes. Across all 18 native rows, `benefit` is false. The latency guard
also passes: there are zero >5% vetoes. Passing the veto guard does not supply
the missing benefit.

The allocation evidence is real but insufficient under the policy. In both
large/capture allocation blocks, calls fall from 72,106 to 10,788 and allocated
bytes fall from 4,767,939 to 853,507. Net live bytes remain 278,201 and peak
live bytes remain 338,955. All 36 audited allocation rows pass the guarded
resource comparison, so this evidence establishes no guarded resource increase;
it cannot replace the required capture or lifecycle timing benefit.

The cross preview has no veto, and profile analysis qualifies 4/4 rows with
owner-counter conservation. Those checks support custody and counter
semantics, but they do not alter the frozen adoption rule. The source,
qualification, build, and quality gates therefore remain useful verification
evidence even though the production disposition is rejection.

## Review conclusion

The rejection is the policy result required by the terminal measurements:
`benefit_satisfied` and `adoption_eligible` are false, while the latency and
resource guards pass. Retaining the candidate because of allocator call or byte
reductions would violate `allocation_count_alone_sufficient: false`. No timing
recapture is justified for this disposition.
