# Explicit review of the omitted peak-memory criterion

The original frozen practical gate is **not met**. Both medium/large repeats
miss the 5% normal-p50 and allocation-call/requested-byte thresholds. The
protocol and derived failed gate remain unchanged.

The root copied that gate from 0440, where namespace allocations were the
target, but failed to include peak memory despite explicitly naming peak
reduction in this batch's hypothesis. That was a protocol-design omission.
The program in user-owned GOAL.md requires improving peak live bytes and
allows keeping statistically and practically useful changes; it does not
require the root's particular calls/requested-byte threshold.

A separate **post-hoc** memory review therefore considers the complete
retained matrix. All 60 allocator observations per role and shape agree:

| Source slides | Peak above entry before | Candidate | Reduction |
|---:|---:|---:|---:|
| 64 | 812,062 B | 781,342 B | 3.783% |
| 4,096 | 19,993,648 B | 18,027,568 B | 9.834% |
| 8,192 | 39,890,548 B | 35,958,388 B | 9.857% |

The reduction equals the cumulative requested-byte saving at each size:
30,720 / 1,966,080 / 3,932,160 bytes. Source and commit remain live at the
endpoint, whose retained live delta is unchanged. This is the measured effect
of avoiding one temporary deep copy, not a bounded-memory append result.

Medium/large normal p50 is 0.664–1.279% slower across the four comparisons.
There are no matched >5% adverse flags; baseline tiny normal p99 has a
5.807% repeat flag. These observations support no normal latency improvement.
All observations remain in summary.json, including instrumented timings and
whole-process RSS. No RSS improvement is claimed.

Any retention decision must identify this post-hoc basis. It must not report
that the original frozen practical gate passed, convert this into a causal
latency claim, or imply a program-wide memory result. Correctness, source
isolation, preservation, complete profiles and portable evidence verification
remain required before the batch is finalized.
