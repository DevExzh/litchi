# 0499 matched before/after measurements

Sixty children contain 3,600 measured samples and 360 warmups. All byte, source-counter, work-counter, and released-budget oracles match. Percentiles use nearest rank. Whole-child RSS includes setup and verification and is counted once per capture. Positive latency/RSS deltas and negative throughput deltas greater than five percent are retained as adverse flags.

| Workload | p50 before µs | p50 after µs | p50 change | p99 change | RSS change | Adverse flags |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| few-large-owned-serial-w1 | 587.73 | 585.85 | -0.32% | -9.28% | +0.59% | — |
| few-large-owned-batch-w1 | 585.22 | 593.02 | +1.33% | -20.61% | -0.39% | — |
| few-large-owned-batch-w2 | 141.13 | 131.29 | -6.97% | +0.92% | +0.97% | — |
| few-large-owned-batch-w4 | 89.32 | 87.69 | -1.82% | -11.18% | -0.06% | — |
| few-large-owned-batch-w8 | 85.26 | 88.53 | +3.84% | -20.48% | +0.44% | p95_us |
| few-large-file-serial-w1 | 577.34 | 578.13 | +0.14% | -4.88% | +0.14% | — |
| few-large-file-batch-w1 | 578.80 | 582.94 | +0.72% | -1.64% | +0.36% | — |
| few-large-file-batch-w2 | 152.40 | 142.29 | -6.63% | -64.88% | -0.09% | — |
| few-large-file-batch-w4 | 95.04 | 95.17 | +0.14% | -8.17% | +0.46% | — |
| few-large-file-batch-w8 | 92.96 | 93.46 | +0.54% | +6.14% | -0.31% | p95_us, p99_us |
| few-large-instrumented-serial-w1 | 82467.92 | 82312.08 | -0.19% | +0.25% | -0.04% | — |
| few-large-instrumented-batch-w1 | 82439.00 | 82228.95 | -0.25% | +0.54% | +0.83% | — |
| few-large-instrumented-batch-w2 | 39932.50 | 39907.46 | -0.06% | -8.40% | +1.04% | — |
| few-large-instrumented-batch-w4 | 19996.37 | 20023.90 | +0.14% | -0.26% | +0.43% | — |
| few-large-instrumented-batch-w8 | 20008.76 | 19979.46 | -0.15% | -0.40% | -1.60% | — |
| many-small-owned-serial-w1 | 95.72 | 95.86 | +0.15% | -4.42% | -0.07% | — |
| many-small-owned-batch-w1 | 99.64 | 94.41 | -5.25% | -5.49% | +0.30% | — |
| many-small-owned-batch-w2 | 669.84 | 186.51 | -72.16% | -67.16% | -0.81% | — |
| many-small-owned-batch-w4 | 542.70 | 186.04 | -65.72% | -67.61% | -3.17% | — |
| many-small-owned-batch-w8 | 475.25 | 228.74 | -51.87% | -50.27% | -12.08% | — |
| many-small-file-serial-w1 | 183.64 | 182.04 | -0.87% | -0.36% | +0.30% | — |
| many-small-file-batch-w1 | 197.71 | 200.50 | +1.41% | -14.81% | +0.65% | — |
| many-small-file-batch-w2 | 780.59 | 296.04 | -62.07% | -76.24% | +1.15% | — |
| many-small-file-batch-w4 | 562.99 | 248.77 | -55.81% | -61.07% | -4.43% | — |
| many-small-file-batch-w8 | 496.17 | 273.70 | -44.84% | -32.86% | -10.27% | — |
| many-small-instrumented-serial-w1 | 29586.75 | 29726.19 | +0.47% | +0.93% | -0.12% | — |
| many-small-instrumented-batch-w1 | 29591.38 | 29574.47 | -0.06% | +1.38% | +0.63% | — |
| many-small-instrumented-batch-w2 | 16012.90 | 15467.24 | -3.41% | -8.09% | +2.69% | — |
| many-small-instrumented-batch-w4 | 8037.23 | 7821.47 | -2.68% | -2.01% | +1.71% | — |
| many-small-instrumented-batch-w8 | 4223.06 | 4068.20 | -3.67% | +30.79% | -0.09% | p95_us, p99_us |

Aggregate flag counts: {"p95_us": 3, "p99_us": 2}.
Per-repeat latency/throughput flag counts: {"mean_us": 1, "p50_us": 1, "p95_us": 3, "p99_us": 4, "throughput_bytes_s": 1}.

Per-repeat metrics and throughput deltas are retained in comparison.json. Shared-host variation and sparse tail samples limit causal interpretation; no broad end-to-end or isolated-host result is implied.
