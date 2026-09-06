# 0448: Calibrate minimum-service-time range pacing

Previous turn: progress, committed 0447. Worktree starts with only user GOAL.md
untracked and no owned CPU process live. Accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49, unchanged from the prior complete read.

Evidence: 0447 plain calls request 2.767330 ms transfer sleep yet grow from
126–131 ms unpaced to 154–156 ms paced, with a second OS sleep per returned chunk.
Add an explicit minimum-service policy beside the existing separate-sleeps policy.
It preserves fixed delay before the wrapped read, then waits only the remaining
fixed-plus-nominal-transfer target after elapsed request work. Fixed-delay
oversleep and source work count toward this lower-bound service time. This is a
different simulation policy, not a production optimization or network claim.

Existing default/with_limits behavior stays separate sleeps. The optional CLI
policy requires range provider plus an explicit transfer rate. Both policies
retain nominal transfer counters; these counters are targets, not actual sleep.
Preserve empty/EOF/zero-cap/error behavior and checked duration/counter overflow.
Validate pure deadline boundaries and exact real PPTX publication equivalence.

After correctness gates and pilots, freeze a same-build comparison: plain and
media-rich, separate/minimum-service, 64 KiB max range, 200 us fixed request delay,
25 MiB/s transfer target, 30 samples/3 warmups, CPU 2 and one worker, two reversed
repeats (8 reports/240 samples), plus four media-rich profiles. Verify the combined
nominal service floor against every enclosing serial API clock in both policies.
Require identical underlying reads and nominal transfer counters in every sample.

Retain the new policy only if plain API p50 is at least 5% lower in both repeats,
all service-floor/correctness gates pass, and every absolute 5% paired/repeat
latency/RSS trigger is reviewed. No production speedup, ideal physical bandwidth,
shared-link scaling, cold I/O, native or operation-allocation claim. Record the
comparison as timer-model calibration. The full non-iWork goal remains active.
