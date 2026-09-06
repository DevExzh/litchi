# 0447: Explicit range transfer pacing

The standalone PPTX range provider could simulate request delay and short
reads, but had no transfer-rate control. It now accepts optional
`--transfer-bytes-per-second` and records requested transfer delay separately
from fixed request delay. Existing unconfigured behavior, production ReadAt,
versions, budgets and publication contracts remain. This is a measured harness
capability, not a production optimization.

The frozen single-build matrix uses plain/media-rich managed cross-slide-copy
lifecycles, 64 KiB maximum reads, 200 us fixed delay, and either no pacing or
25 MiB/s requested transfer pacing. Eight processes retain 240 samples with
30 samples/3 warmups, CPU 2, one worker and reversed repeats. Four profiles
and all source/cache/budget ownership boundaries remain retained.

| Corpus | Unpaced API p50 ms, R1/R2 | Paced API p50 ms, R1/R2 | Requested transfer delay ms |
| --- | ---: | ---: | ---: |
| Plain | 130.597 / 126.092 | 156.006 / 154.093 | 2.767330 |
| Media-rich | 549.966 / 539.105 | 2,613.293 / 2,600.987 | 1,923.768399 |

Every sample matches underlying read work and output identity across both
configurations/repeats. Media-rich open/plan/publication return 19,226 /
16,794,014 / 33,617,184 bytes through 162 / 580 / 1,258 calls. Requested pacing
is 0.734 / 640.641 / 1,282.394 ms respectively. Publication therefore accounts
for about two-thirds of the requested transfer delay; that is a model-derived
quantity, not a measured wall-time share or Amdahl serial fraction.

Twelve absolute 5% repeat flags remain, all tails. Three increase: paced plain
publication p95/p99 (+5.233%/+5.272%) and paced media-rich open p99 (+10.694%).
Nine decrease. No API median or process-RSS repeat crosses 5%. See
[complete measurements](../results/change-0447/measurements.md) and
[raw/derived evidence](../results/change-0447/summary.json).

Each successful nonempty read requests a second sleep, rounded up to a whole
nanosecond. OS sleep granularity can dominate small requests: the plain observed
increase exceeds its requested 2.767 ms transfer delay. These observations do not
prove actual bandwidth. Concurrent reads do not share a link budget. Future
calibration should assess a combined deadline model before claiming ideal link
behavior. Managed phase-boundary memory and approximately 784 MiB media-rich
whole-process RSS include different ownership/setup costs; operation allocation
metrics are unavailable and the sink retains full output.

The 378 harness tests pass, including five new pacing/CLI tests and expanded
PPTX lifecycle equivalence. Strict harness lint adds zero diagnostics relative
to 29 inherited diagnostics. Workspace/feature checks, warning-denied rustdoc,
formatting and boundaries pass. Four 1/0 pilots and 17 report corruption probes
pass before freeze. The initial 3/1 pilot oracle rejection is retained.

Profiles retain unknown symbols and four unpaced samples without callchains
(0.024307% of period). SHA-256 accounts for 60.2–60.9% of self period within the
lifecycle-frame subset, followed by memory moves. That subset includes untimed
work; blocked sleep is absent from cycles. No production CPU attribution follows.
The initial profile-parser and pilot-count replay corrections remain disclosed.

Reproduce portable validation with
`python3 -B docs/performance/results/change-0447/verify.py --sealed --cleanup`.
The [validation record](../results/change-0447/validation-notes.md) documents
source custody, oracles, failures and cleanup. The full non-iWork goal remains
active: native breadth, cold I/O, bounded existing append, repackaging and scaling
are still incomplete. Registry and representative-index statuses are unchanged.
