# Measured baseline

Normal elapsed time is in milliseconds. Each row contains 30 retained samples; the p50 interval is a deterministic 2,000-resample 95% bootstrap interval.

| Source slides | Repeat | p50 ms | p50 95% interval ms | p95 ms | p99 ms |
|---:|---|---:|---|---:|---:|
| 64 | R1 | 2.075 | 2.072–2.079 | 2.093 | 2.128 |
| 4,096 | R1 | 85.178 | 84.833–85.356 | 86.485 | 86.514 |
| 8,192 | R1 | 170.742 | 170.529–171.093 | 171.804 | 171.818 |
| 8,192 | R2 | 175.416 | 175.090–175.988 | 177.434 | 177.976 |
| 4,096 | R2 | 85.979 | 85.760–86.087 | 86.426 | 86.449 |
| 64 | R2 | 2.081 | 2.076–2.105 | 2.130 | 2.130 |

Allocator observations below are identical across all 60 retained observations per shape, including both repeats. Requested bytes are cumulative requests within the operation; peak above entry and retained live delta are separate measurements.

| Source slides | Allocation calls | Requested bytes | Peak above entry bytes | Retained live delta bytes |
|---:|---:|---:|---:|---:|---:|
| 64 | 15,933 | 10,824,868 | 812,062 | 73,324 |
| 4,096 | 733,682 | 122,864,214 | 19,993,648 | 4,041,249 |
| 8,192 | 1,462,779 | 236,704,188 | 39,890,548 | 8,074,341 |

Whole-process peak RSS spans 84,590,592–85,782,528 bytes. One of 45 repeat checks crosses the 5% review trigger: allocator large p99 −6.229%. Allocator elapsed time is descriptive and is not a normal latency claim. No normal or allocation-counter repeat check crosses the trigger.

The successful stat profile records 30,079,967,633 user cycles and 125,266,533,586 user instructions. Whole-process self samples include memcmp (8.38%), namespace-event processing (6.91%), attribute iteration (5.21%), ODP ElementAttrs::get (4.71%), and memmove (4.19%). These include setup and oracle work; they do not isolate the timed operation. The zero L1 counter is retained without a cache-miss conclusion.
