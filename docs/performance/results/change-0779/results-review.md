# 0779 paired results review — reject bounded input growth

Disposition: **reject the candidate for production retention**. This is a
performance and resource tradeoff rejection, not a correctness failure. The
candidate remains archived at `candidate/applied-phys_pkg.rs`; the working
`crates/litchi-opc/src/phys_pkg.rs` was restored to the base source witness.
The retained disposition records `production_change_retained: false`.

## Evidence and attribution

The paired matrix uses the generated 4,226,429-byte workbook and the real
654,688-byte conditional-formatting workbook. It covers `open`, `edit`,
`save`, and `lifecycle` with six alternating native blocks and 30 samples per
phase, plus two allocator blocks with three samples. The 3,555-byte and
8,224-byte supplementary open controls are separate. Source hashes, output
hashes, marker/reopen checks, and the 0778 NoSync published-output digest all
match. The retained quality report has six successful gates and the final
analysis contains 32 primary allocation children, 96 primary native children,
and separate small-control lanes.

The direct Heaptrack attribution confirms the suspected site. For one
generated `read_limited` call, the baseline requested 1,092,697,469 bytes over
516 allocation events, while the candidate requested 40,650,096 bytes over 44
events. The interpreted two-open whole-process trace reports target requested
bytes falling from 2,185,394,938 to 81,300,192. These are allocator request
sizes. The trace explicitly does not measure bytes physically copied by a
reallocation, RSS, or latency, so the 96% reduction cannot be presented as a
96% memory or copy reduction.

The table below uses the nearest-rank 50th percentile across the six native
process p50s (the third sorted value), with percent changes calculated from
those displayed values. These are not median paired-block ratios. The main
report instead averages the middle two process p50s and reports paired ratios.
Allocator values are repeated process p50s.
`elapsed` values are converted from the reports' nanoseconds to milliseconds;
`live after` is the allocator live-byte snapshot at the end of the measured
region.

| Corpus / phase | Requested bytes before → after | Allocation / realloc calls before → after | Live after before → after | Elapsed p50 before → after |
| --- | ---: | ---: | ---: | ---: |
| Generated / open | 1,093,823,950 → 41,776,577 (-96.18%) | 1,615 / 635 → 1,143 / 163 | 4,531,873 → 4,854,874 (+7.13%) | 0.405252 → 0.398471 ms (-1.67%) |
| Generated / lifecycle | 1,100,160,814 → 48,113,441 (-95.63%) | 29,843 / 5,558 → 29,371 / 5,086 | 5,033,625 → 5,356,626 (+6.42%) | 4.805340 → 4.776620 ms (-0.60%) |
| Real / open | 33,020,061 → 12,403,384 (-62.44%) | 9,193 / 1,014 → 9,141 / 962 | 1,266,687 → 1,303,063 (+2.87%) | 1.155205 → 1.147935 ms (-0.63%) |
| Real / lifecycle | 39,278,222 → 18,661,545 (-52.49%) | 27,513 / 2,497 → 27,461 / 2,445 | 1,289,962 → 1,326,338 (+2.82%) | 3.255704 → 3.243774 ms (-0.37%) |
| 8,224-byte boundary / open | 587,981 → 596,141 (+1.39%) | 784 / 76 → 784 / 76 | 36,051 → 44,210 (+22.63%) | 0.064550 → 0.064820 ms (+0.42%) |

The generated open region also raises allocator `peak_above_entry` from
4,615,014 to 4,938,016 bytes (+7.00%); the generated lifecycle rises from
6,466,390 to 6,789,392 (+4.99%). The boundary control's corresponding peak
rises 7.10%. The 3,555-byte small control has identical allocation and
reallocation counts and bytes, so the candidate does not provide a general
small-input win.

The whole-process native RSS diagnostic rises by about 20.2% in the generated
save and lifecycle paired comparisons. The plan includes setup and post-clock
verification in this RSS value, so it is not an operation-local memory claim.
It is nevertheless adverse corroborating evidence alongside the direct
allocator live-byte increase. The allocator-region values are the stronger
causal observation for the retained input buffer.

## Decision against the goal rules

The candidate passes the semantic side of GOAL non-negotiable semantic rule 6: validation did not
mutate, NoSync save did not repair or normalize the package, output digests
remain bound to the existing result, and the reopen/marker checks pass. The
change only altered capacity reservation. That semantic rule therefore supplies no reason
to reject the bytes or typed behavior, but it also does not justify retaining
an allocation policy whose resource tradeoff is unfavorable.

The practical decision follows GOAL rule 1's priority for bounded resource use,
while rule 12 confirms that the unchanged limit and malformed-input defenses
remain admissible and rule 13 bars treating request counters as physical-memory
evidence. The candidate's
request-count and request-byte reductions are real observations in the
allocator instrument, but they do not establish less copying or less physical
memory. The end-to-end lifecycle gains are below one percent on both corpora,
while retained allocator live bytes rise 6–7% for the generated workbook and
about 3% for the real workbook. The boundary control is worse by about 23%
in retained live bytes without reducing any allocation or reallocation count.
Keeping this candidate would exchange a diagnostic request-volume reduction
for a measured retained-memory increase without a material lifecycle benefit.
That fails the practical bar for this shared owned-ingress policy under those
goal rules.

No data challenge is warranted. The source and fixture identities are bound,
the allocator and native lanes are separate, Heaptrack independently attributes
the growth site, and the final validation records the expected regression and
spread flags rather than hiding them. The base source is restored; any future
attempt would need a separately measured policy with an explicit retained-memory
budget and a material end-to-end benefit.
