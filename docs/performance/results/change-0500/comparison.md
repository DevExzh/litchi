# 0500 managed DOCX comparisons

Thirty-six children contain 2,160 measurements and 216 warmups. Every output identity agrees across matched routes. Work charges may differ as passes are removed. Percentiles use nearest rank; RSS includes whole-child setup and verification.

## after_batch_vs_scalar

| Workload | Lifecycle p50 before → after µs | Edit p50 before → after µs | Lifecycle p99 change | RSS change | Flags |
| --- | ---: | ---: | ---: | ---: | --- |
| p128-k1-owned | 596.82 → 640.63 | 270.59 → 269.65 | +4.95% | +1.71% | elapsed_p50_us, elapsed_p95_us, elapsed_mean_us, throughput_output_bytes_s |
| p128-k1-file | 639.63 → 643.54 | 273.67 → 274.17 | +0.93% | -3.70% | — |
| p128-k8-owned | 1483.37 → 632.24 | 1148.78 → 299.91 | -55.95% | +2.42% | — |
| p128-k8-file | 1595.61 → 686.39 | 1181.21 → 311.87 | -56.23% | +1.84% | — |
| p128-k32-owned | 5483.62 → 783.68 | 5103.10 → 403.74 | -86.08% | -0.82% | — |
| p128-k32-file | 5926.34 → 853.13 | 5507.08 → 432.38 | -85.34% | +0.63% | — |
| p512-k1-owned | 2245.88 → 2235.11 | 1031.97 → 1033.61 | +0.30% | +0.05% | — |
| p512-k1-file | 2276.56 → 2274.17 | 1032.18 → 1036.86 | -0.57% | -0.86% | — |
| p512-k8-owned | 5323.19 → 2273.93 | 4103.84 → 1061.87 | -58.52% | +2.34% | — |
| p512-k8-file | 5380.68 → 2298.04 | 4132.39 → 1089.63 | -62.41% | +5.77% | rss_kib |
| p512-k32-owned | 17165.06 → 2421.56 | 15922.74 → 1191.54 | -85.40% | +1.66% | — |
| p512-k32-file | 17271.56 → 2419.75 | 16033.39 → 1191.58 | -85.15% | +1.83% | — |

Aggregate flags: {"elapsed_mean_us": 1, "elapsed_p50_us": 1, "elapsed_p95_us": 1, "rss_kib": 1, "throughput_output_bytes_s": 1}

## batch_after_vs_scalar_before

| Workload | Lifecycle p50 before → after µs | Edit p50 before → after µs | Lifecycle p99 change | RSS change | Flags |
| --- | ---: | ---: | ---: | ---: | --- |
| p128-k1-owned | 650.08 → 640.63 | 274.49 → 269.65 | -3.25% | +2.89% | — |
| p128-k1-file | 638.72 → 643.54 | 271.80 → 274.17 | +1.23% | -2.53% | — |
| p128-k8-owned | 1518.32 → 632.24 | 1139.27 → 299.91 | -56.91% | +4.22% | — |
| p128-k8-file | 1530.17 → 686.39 | 1166.86 → 311.87 | -54.98% | +0.97% | — |
| p128-k32-owned | 5516.12 → 783.68 | 5138.42 → 403.74 | -85.55% | -1.47% | — |
| p128-k32-file | 5919.41 → 853.13 | 5518.19 → 432.38 | -86.01% | -0.40% | — |
| p512-k1-owned | 2214.50 → 2235.11 | 1014.68 → 1033.61 | -0.29% | +1.05% | — |
| p512-k1-file | 2256.75 → 2274.17 | 1021.65 → 1036.86 | +0.72% | +3.11% | — |
| p512-k8-owned | 5268.44 → 2273.93 | 4066.09 → 1061.87 | -55.73% | +2.56% | — |
| p512-k8-file | 5367.89 → 2298.04 | 4119.37 → 1089.63 | -56.66% | +7.89% | rss_kib |
| p512-k32-owned | 16696.95 → 2421.56 | 15490.13 → 1191.54 | -85.07% | -0.32% | — |
| p512-k32-file | 17268.22 → 2419.75 | 16042.28 → 1191.58 | -85.20% | +5.83% | rss_kib |

Aggregate flags: {"rss_kib": 2}

## scalar_before_after

| Workload | Lifecycle p50 before → after µs | Edit p50 before → after µs | Lifecycle p99 change | RSS change | Flags |
| --- | ---: | ---: | ---: | ---: | --- |
| p128-k1-owned | 650.08 → 596.82 | 274.49 → 270.59 | -7.81% | +1.17% | — |
| p128-k1-file | 638.72 → 639.63 | 271.80 → 273.67 | +0.30% | +1.21% | — |
| p128-k8-owned | 1518.32 → 1483.37 | 1139.27 → 1148.78 | -2.18% | +1.75% | — |
| p128-k8-file | 1530.17 → 1595.61 | 1166.86 → 1181.21 | +2.85% | -0.85% | — |
| p128-k32-owned | 5516.12 → 5483.62 | 5138.42 → 5103.10 | +3.85% | -0.66% | — |
| p128-k32-file | 5919.41 → 5926.34 | 5518.19 → 5507.08 | -4.57% | -1.03% | — |
| p512-k1-owned | 2214.50 → 2245.88 | 1014.68 → 1031.97 | -0.58% | +0.99% | — |
| p512-k1-file | 2256.75 → 2276.56 | 1021.65 → 1032.18 | +1.30% | +4.01% | — |
| p512-k8-owned | 5268.44 → 5323.19 | 4066.09 → 4103.84 | +6.72% | +0.22% | elapsed_p99_us |
| p512-k8-file | 5367.89 → 5380.68 | 4119.37 → 4132.39 | +15.31% | +2.00% | elapsed_p99_us |
| p512-k32-owned | 16696.95 → 17165.06 | 15490.13 → 15922.74 | +2.27% | -1.94% | — |
| p512-k32-file | 17268.22 → 17271.56 | 16042.28 → 16033.39 | -0.37% | +3.92% | — |

Aggregate flags: {"elapsed_p99_us": 2}

Per-repeat values, all named phases, source counters, charged Work, and every retained flag appear in comparison.json. Two repeats on a shared host do not establish isolated causal costs or native/cold behavior.
