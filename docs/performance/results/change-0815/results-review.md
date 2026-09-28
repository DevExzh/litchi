# 0815 numerical results and independent raw review

The terminal 0815 packet contains 306 main reports and 6,714 measured
samples: 216 native reports, 72 allocation reports, and 18 before-only
qualification reports. The three readers used here are the retained
analysis.json, root-audit.json, and quality-summary.json artifacts. Their
numerical result is adoption-ineligible; this is a policy result, and the
root decision and disposition records own source retention.

Both the reused baseline quality summary and the fresh candidate production
gates record 1,241 passed tests, zero failures, three ignored tests, and 85 suites. Both fresh probe legs record 36
passed tests and zero failures. The qualification lane has one sample per
case and imports no timing pool. Native timing uses six paired blocks with
seed 815815 and 10,000 bootstrap resamples; allocation uses two paired blocks.

## Frozen policy result

The frozen benefit rule considers capture and lifecycle p50 rows, requires at
least 3% improvement, and requires the bootstrap high endpoint below one.
Positive change below means that the after candidate is slower. No row meets
the benefit rule. The only latency veto is large/lifecycle: its ratio is above
1.05 and its interval is entirely above one. The numerical audit therefore
has zero eligible benefits, one latency veto, no resource violations, and
adoption_eligible=false.

The table is copied from the raw paired block audit. The interval is the 95%
bootstrap interval for the median of the six after/before block ratios; the
ratio is not calculated by dividing the two displayed p50 medians.

| Shape/workflow | Before p50 ns | After p50 ns | Ratio | Change | 95% interval | Policy |
| --- | ---: | ---: | ---: | ---: | --- | --- |
| tiny/capture | 228871.0 | 228036.0 | 0.996898273 | -0.310173% | [0.992916493, 0.998887707] | none |
| tiny/commit | 206986.0 | 206816.0 | 1.000625375 | +0.062537% | [0.993812546, 1.004552697] | outside benefit lane |
| tiny/lifecycle | 1403907.5 | 1413532.5 | 1.005987631 | +0.598763% | [1.003831638, 1.009265330] | none |
| medium/capture | 425122.0 | 433557.0 | 1.018741751 | +1.874175% | [1.010940192, 1.024773829] | none |
| medium/commit | 291241.5 | 289701.0 | 0.994114632 | -0.588537% | [0.991716027, 0.998095278] | outside benefit lane |
| medium/lifecycle | 1980175.0 | 1998045.5 | 1.010177271 | +1.017727% | [1.007406329, 1.014868097] | none |
| large/capture | 16031204.5 | 16724164.0 | 1.042702887 | +4.270289% | [1.037801272, 1.056069325] | none |
| large/commit | 1264437.0 | 1263136.5 | 1.000375655 | +0.037565% | [0.996839549, 1.003156047] | outside benefit lane |
| large/lifecycle | 25804666.5 | 27168548.0 | 1.052994372 | +5.299437% | [1.049080959, 1.057174796] | latency veto |
| vendor/capture | 509672.5 | 518318.0 | 1.016888698 | +1.688870% | [1.009502110, 1.022071005] | none |
| vendor/commit | 319211.5 | 318221.5 | 0.997337136 | -0.266286% | [0.991784445, 1.000614173] | outside benefit lane |
| vendor/lifecycle | 2129891.0 | 2158641.5 | 1.013712987 | +1.371299% | [1.010133551, 1.016159655] | none |
| unicode-vendor/capture | 513792.0 | 521048.0 | 1.015523658 | +1.552366% | [1.011320501, 1.018577269] | none |
| unicode-vendor/commit | 319661.5 | 319117.0 | 0.997842146 | -0.215785% | [0.994625434, 1.001512434] | outside benefit lane |
| unicode-vendor/lifecycle | 2140771.5 | 2169841.0 | 1.012083958 | +1.208396% | [1.010576913, 1.016337615] | none |
| valid-4attr/capture | 490362.5 | 494108.0 | 1.007168566 | +0.716857% | [1.002278028, 1.016080914] | none |
| valid-4attr/commit | 311521.0 | 311247.0 | 0.999099939 | -0.090006% | [0.994632997, 1.001668851] | outside benefit lane |
| valid-4attr/lifecycle | 2093296.0 | 2116496.0 | 1.012858440 | +1.285844% | [1.006683711, 1.014424213] | none |

The large/lifecycle veto is backed by six raw paired ratios:
1.057993888, 1.056355704, 1.051675304, 1.049336639, 1.054313441, and
1.048825280. Its median is 1.052994372, with interval
[1.049080959, 1.057174796]. Large capture is slower by 4.270289% but does
not cross the 1.05 veto threshold. The static codegen gate separately records
two arm-local copy sequences before and zero after; that observation does not
establish a workflow benefit or override the timing veto.

## Allocation equality

The allocation lane is separate from elapsed-time measurements. The raw audit
contains 18 cases times two blocks times four guarded metrics. Every after
value equals its before value, so all 144 comparisons are equal and there are
no decreases or increases.

| Guarded metric | Comparisons | Equal | Decreased | Increased |
| --- | ---: | ---: | ---: | ---: |
| allocation_calls | 36 | 36 | 0 | 0 |
| allocated_bytes | 36 | 36 | 0 | 0 |
| net_live | 36 | 36 | 0 | 0 |
| peak_above_entry | 36 | 36 | 0 | 0 |
| Total | 144 | 144 | 0 | 0 |

Allocation has no regression or spread flags. Equality is a resource guard
result; it is not an allocation reduction claim.

## Tail, RSS, and spread diagnostics

The following are diagnostics rather than p50 adoption guards. All individual
p95 and p99 paired block changes above 5% remain in the review:

| Shape/workflow | Metric | Block | Change |
| --- | --- | ---: | ---: |
| large/capture | p95 | 1 | +6.248% |
| large/capture | p95 | 2 | +6.537% |
| large/capture | p95 | 4 | +5.453% |
| large/capture | p99 | 1 | +8.233% |
| large/capture | p99 | 2 | +7.776% |
| large/capture | p99 | 4 | +6.193% |
| large/lifecycle | p95 | 0 | +5.730% |
| large/lifecycle | p95 | 1 | +5.277% |
| large/lifecycle | p95 | 2 | +7.192% |
| large/lifecycle | p95 | 3 | +5.196% |
| large/lifecycle | p95 | 4 | +5.826% |
| large/lifecycle | p95 | 5 | +6.711% |
| large/lifecycle | p99 | 0 | +5.696% |
| large/lifecycle | p99 | 1 | +5.453% |
| large/lifecycle | p99 | 2 | +5.690% |
| large/lifecycle | p99 | 3 | +5.114% |
| large/lifecycle | p99 | 4 | +6.066% |
| large/lifecycle | p99 | 5 | +6.733% |
| medium/lifecycle | p99 | 0 | +6.202% |
| medium/lifecycle | p99 | 2 | +5.312% |
| tiny/capture | p99 | 4 | +12.337% |
| tiny/commit | p99 | 5 | +32.672% |
| valid-4attr/capture | p99 | 5 | +7.245% |
| valid-4attr/commit | p99 | 4 | +6.918% |
| valid-4attr/lifecycle | p99 | 2 | +7.764% |
| valid-4attr/lifecycle | p99 | 3 | +14.463% |

There are 23 native spread flags: one p95 flag, eleven p99 flags, and eleven
RSS flags. No native p50 spread exceeds 5%; the maximum native p50 spread is
2.324% for large/capture/after. The largest native p95 spread is 5.422% for
valid-4attr/commit/after, the largest native p99 spread is 35.431% for
tiny/commit/after, and the largest native RSS spread is 10.676% for
tiny/commit/after. Allocation has no spread flags.

Process RSS is review-only evidence and does not imply an allocation result.
The six-repeat medians are:

| Shape/workflow | Before RSS KiB | After RSS KiB |
| --- | ---: | ---: |
| tiny/capture | 5170 | 5130 |
| tiny/commit | 5080 | 5146 |
| tiny/lifecycle | 5050 | 5200 |
| medium/capture | 5260 | 5202 |
| medium/commit | 5480 | 5516 |
| medium/lifecycle | 5356 | 5324 |
| large/capture | 18662 | 18732 |
| large/commit | 18668 | 18668 |
| large/lifecycle | 18668 | 18636 |
| vendor/capture | 5516 | 5454 |
| vendor/commit | 5454 | 5436 |
| vendor/lifecycle | 5388 | 5428 |
| unicode-vendor/capture | 5548 | 5452 |
| unicode-vendor/commit | 5418 | 5516 |
| unicode-vendor/lifecycle | 5406 | 5420 |
| valid-4attr/capture | 5482 | 5484 |
| valid-4attr/commit | 5484 | 5420 |
| valid-4attr/lifecycle | 5516 | 5482 |

Three paired RSS increases exceed 5%:

* tiny/commit, block 3: 5048 to 5308 KiB (+5.151%).
* tiny/commit, block 4: 4940 to 5240 KiB (+6.073%).
* medium/lifecycle, block 3: 5132 to 5452 KiB (+6.235%).

The RSS medians and these block observations are host-process diagnostics.
They do not establish a memory saving or a general RSS bound.

## Reader checks and limits

The completed offline writes were:

* analysis.py --write: PASS, 306 reports and 6,714 samples.
* root_audit.py --write: PASS, one latency violation, zero resource violations.
* quality_summary.py --write: PASS, schema
  litchi.performance.0815.quality-summary.v1.

The review has no historical timing pool, cross-format lane, heaptrack, perf,
latency-profile, causal-cycle, cold-cache, remote-source, non-seek-output, or
concurrency claim. Profile evidence is a separate reader and is not folded
into this numerical decision. The retained measurements describe this
deterministic PPTX matrix only.
