# PPTX minimum-service calibration

Both policies use 64 KiB ranges, 200 us fixed request delay and 25 MiB/s nominal
transfer targets. One build, CPU 2, one worker, two reversed repeats; 8 reports
and 240 retained samples (30 samples/3 warmups per process). These are delay-model
observations, not a production speedup or measured network bandwidth.

| Corpus | Policy | Repeat | API p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process RSS MiB |
| --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| plain | separate-sleeps | R1 | 152.814823 | 153.161050 | 153.164661 | 152.812658–152.815789 | 19.387 |
| media-rich | separate-sleeps | R1 | 2578.538702 | 2578.905309 | 2581.266538 | 2578.419619–2578.627854 | 785.297 |
| plain | minimum-service | R1 | 123.322546 | 123.668612 | 123.813771 | 123.316706–123.333035 | 19.695 |
| media-rich | minimum-service | R1 | 2454.029005 | 2456.061126 | 2459.440518 | 2453.893789–2454.225695 | 784.301 |
| media-rich | minimum-service | R2 | 2460.901378 | 2462.251059 | 2465.502466 | 2460.681221–2461.111359 | 785.281 |
| plain | minimum-service | R2 | 123.346862 | 123.666453 | 123.803233 | 123.322911–123.464162 | 19.406 |
| media-rich | separate-sleeps | R2 | 2572.652557 | 2573.441965 | 2575.392194 | 2572.514754–2572.744496 | 784.059 |
| plain | separate-sleeps | R2 | 152.816473 | 152.819372 | 152.820022 | 152.815938–152.817112 | 19.691 |

p50 uses the midpoint; p95/p99 use nearest rank. The 2,000 bootstrap resamples
describe within-process uncertainty, not machine/day uncertainty. All individual
phase timers and vectors remain in summary.json.

| Corpus | Repeat | Minimum-service API p50 relative to separate sleeps |
| --- | --- | ---: |
| plain | R1 | -19.299% |
| plain | R2 | -19.284% |
| media-rich | R1 | -4.829% |
| media-rich | R2 | -4.344% |

| Corpus | Phase | Logical reads | Returned bytes | Nominal fixed + transfer floor ms |
| --- | --- | ---: | ---: | ---: |
| plain | opened | 162.0 | 17,572.0 | 33.070414 |
| plain | planned | 36.0 | 10,756.0 | 7.610328 |
| plain | published | 281.0 | 44,209.0 | 57.886588 |
| media-rich | opened | 162.0 | 19,226.0 | 33.133509 |
| media-rich | planned | 580.0 | 16,794,014.0 | 756.640830 |
| media-rich | published | 1,258.0 | 33,617,184.0 | 1533.994060 |

All underlying counters and nominal targets are equal across configurations,
repeats and samples. Every enclosing serial API clock satisfies the independently
checked combined service floor. The target is not a measured sleep duration.
Minimum-service credits time already spent, including wrapped-source work and
fixed-wait overshoot; separate sleeps request the full extra transfer wait.

Absolute 5% repeat review: 0 flags.


Every paired difference remains in summary.json; all absolute 5% paired and
repeat triggers are reviewed in decision.json. No observations are discarded.
The sink retains full output; operation allocator attribution is unavailable.
Managed boundary gauges and whole-process RSS do not establish allocation peaks
or bounded total memory. Profiles include untimed work and omit blocked sleep.
No physical bandwidth, cold I/O, shared-link concurrency, native compatibility
or scaling claim follows.
