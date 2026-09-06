# PPTX range transfer-pacing baseline

One build; 8 reports/240 retained samples; CPU 2, one worker, 30 samples and
three warmups. Both configurations use 64 KiB maximum reads and 200 us fixed
request delay. Paced reads additionally request transfer sleeps at 25 MiB/s.
This measures simulation cost, not a production speedup or physical network rate.

| Corpus | Paced | Repeat | API p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Whole-process RSS MiB |
| --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| plain | False | R1 | 130.597022 | 139.397943 | 147.616270 | 129.957709–133.633895 | 19.863 |
| media-rich | False | R1 | 549.966205 | 568.957388 | 580.894658 | 545.780491–556.199853 | 783.816 |
| plain | True | R1 | 156.006251 | 161.816545 | 162.862841 | 153.392074–156.895523 | 20.152 |
| media-rich | True | R1 | 2613.293088 | 2646.565440 | 2652.880113 | 2606.864741–2618.502976 | 785.258 |
| media-rich | True | R2 | 2600.987479 | 2643.139564 | 2656.345011 | 2588.314506–2613.892035 | 783.824 |
| plain | True | R2 | 154.093050 | 160.303767 | 160.840290 | 153.074335–156.768293 | 19.402 |
| media-rich | False | R2 | 539.104992 | 556.439666 | 557.900545 | 535.943540–542.070793 | 783.547 |
| plain | False | R2 | 126.092147 | 133.713851 | 136.388666 | 123.936460–127.917631 | 19.473 |

All API vectors, per-phase timing and counter observations are in summary.json.
p50 uses the midpoint; p95/p99 use nearest rank. Bootstrap intervals use 2,000
deterministic within-process resamples, not machine/day uncertainty.

| Corpus | Paced | Repeat | Phase | Logical reads | Returned bytes | Requested transfer delay ms |
| --- | --- | --- | --- | ---: | ---: | ---: |
| plain | False | R1 | opened | 162.0 | 17,572.0 | 0.000000 |
| plain | False | R1 | planned | 36.0 | 10,756.0 | 0.000000 |
| plain | False | R1 | published | 281.0 | 44,209.0 | 0.000000 |
| media-rich | False | R1 | opened | 162.0 | 19,226.0 | 0.000000 |
| media-rich | False | R1 | planned | 580.0 | 16,794,014.0 | 0.000000 |
| media-rich | False | R1 | published | 1,258.0 | 33,617,184.0 | 0.000000 |
| plain | True | R1 | opened | 162.0 | 17,572.0 | 0.670414 |
| plain | True | R1 | planned | 36.0 | 10,756.0 | 0.410328 |
| plain | True | R1 | published | 281.0 | 44,209.0 | 1.686588 |
| media-rich | True | R1 | opened | 162.0 | 19,226.0 | 0.733509 |
| media-rich | True | R1 | planned | 580.0 | 16,794,014.0 | 640.640830 |
| media-rich | True | R1 | published | 1,258.0 | 33,617,184.0 | 1282.394060 |
| media-rich | True | R2 | opened | 162.0 | 19,226.0 | 0.733509 |
| media-rich | True | R2 | planned | 580.0 | 16,794,014.0 | 640.640830 |
| media-rich | True | R2 | published | 1,258.0 | 33,617,184.0 | 1282.394060 |
| plain | True | R2 | opened | 162.0 | 17,572.0 | 0.670414 |
| plain | True | R2 | planned | 36.0 | 10,756.0 | 0.410328 |
| plain | True | R2 | published | 281.0 | 44,209.0 | 1.686588 |
| media-rich | False | R2 | opened | 162.0 | 19,226.0 | 0.000000 |
| media-rich | False | R2 | planned | 580.0 | 16,794,014.0 | 0.000000 |
| media-rich | False | R2 | published | 1,258.0 | 33,617,184.0 | 0.000000 |
| plain | False | R2 | opened | 162.0 | 17,572.0 | 0.000000 |
| plain | False | R2 | planned | 36.0 | 10,756.0 | 0.000000 |
| plain | False | R2 | published | 281.0 | 44,209.0 | 0.000000 |

| Corpus | Repeat | Phase | Requested transfer delay / paced API median |
| --- | --- | --- | ---: |
| plain | R1 | opened | 1.276% |
| plain | R1 | planned | 3.267% |
| plain | R1 | published | 1.869% |
| media-rich | R1 | opened | 1.422% |
| media-rich | R1 | planned | 75.740% |
| media-rich | R1 | published | 74.801% |
| media-rich | R2 | opened | 1.422% |
| media-rich | R2 | planned | 76.291% |
| media-rich | R2 | published | 75.133% |
| plain | R2 | opened | 1.306% |
| plain | R2 | planned | 3.307% |
| plain | R2 | published | 1.884% |

These are requested-sleep/median ratios, not Amdahl serial fractions or
measured sleep attribution. Fixed delay, OS oversleep and actual API work
remain in the denominator; no causal subtraction is performed.

Absolute 5% repeat review: 12 flags.

- plain, paced=False, plan_ns.p95: -14.480% R2 versus R1.
- plain, paced=False, plan_ns.p99: -25.655% R2 versus R1.
- plain, paced=False, publication_ns.p95: -8.227% R2 versus R1.
- plain, paced=False, publication_ns.p99: -7.599% R2 versus R1.
- plain, paced=False, api_sum_ns.p99: -7.606% R2 versus R1.
- plain, paced=True, open_ns.p95: -7.459% R2 versus R1.
- plain, paced=True, publication_ns.p95: +5.233% R2 versus R1.
- plain, paced=True, publication_ns.p99: +5.272% R2 versus R1.
- media-rich, paced=False, open_ns.p95: -15.000% R2 versus R1.
- media-rich, paced=False, open_ns.p99: -5.750% R2 versus R1.
- media-rich, paced=False, publication_ns.p99: -5.510% R2 versus R1.
- media-rich, paced=True, open_ns.p99: +10.694% R2 versus R1.

Every paired API difference remains visible in summary.json. Added pacing cost
is expected model behavior and is not a production regression. Counter equality
is checked across configurations, repeats and samples; output/source identities
are frozen from the pilots. No observation is silently recaptured or discarded.

The full-retaining CountingSink and fixed fixtures contribute to process memory.
Managed phase-boundary memory is not an allocation high-water mark. Operation
allocator attribution is unavailable. Final memory/object/depth budgets return
to zero under the existing lifecycle oracle. No bounded-total-memory, cold I/O,
shared-link concurrency, native compatibility or scaling claim follows.

Profiles include the whole fresh process and untimed work; cycles omit blocked
sleep. The zero L1 miss event on this guest must not be interpreted as a cache
benefit. Fixed and transfer sleeps are separate; OS oversleep affects observed
latency. The nominal pacing rate is not a measured achieved bandwidth.
