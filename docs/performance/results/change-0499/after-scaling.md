# Change 0499 after: serial versus batch scaling

Generated: `2026-09-10T20:21:25+00:00`

The analysis uses measured sample indexes 3–32 from each of two repeats; indexes 0–2 are warmups. Percentiles use the nearest-rank rule used by the Rust harness. RSS is the whole child maximum from `/usr/bin/time -v`. The measured process was pinned to the configured benchmark CPU set on a shared host.

## Serial versus batch

Positive latency or RSS deltas and negative throughput deltas are adverse. The flag threshold is greater than 5 percent.

| corpus | source | workers | p50 batch/serial (us) | p95 batch/serial (us) | p99 batch/serial (us) | mean batch/serial (us) | throughput delta | RSS delta | flags |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| few-large | owned | 1 | 593.02/585.85 | 609.85/607.72 | 636.47/663.69 | 583.67/577.87 | -0.99% | -0.30% | — |
| few-large | owned | 2 | 131.29/585.85 | 146.37/607.72 | 163.10/663.69 | 133.71/577.87 | 332.19% | 14.84% | rss_regression_gt_5pct |
| few-large | owned | 4 | 87.69/585.85 | 106.72/607.72 | 241.74/663.69 | 96.62/577.87 | 498.06% | 24.93% | rss_regression_gt_5pct |
| few-large | owned | 8 | 88.53/585.85 | 130.18/607.72 | 240.78/663.69 | 97.53/577.87 | 492.53% | 24.93% | rss_regression_gt_5pct |
| few-large | file | 1 | 582.94/578.13 | 596.29/586.33 | 646.94/588.53 | 571.34/562.73 | -1.51% | 0.12% | latency_p99_us_regression_gt_5pct |
| few-large | file | 2 | 142.29/578.13 | 178.16/586.33 | 199.39/588.53 | 146.75/562.73 | 283.45% | 15.16% | rss_regression_gt_5pct |
| few-large | file | 4 | 95.17/578.13 | 111.45/586.33 | 273.30/588.53 | 103.49/562.73 | 443.75% | 25.56% | rss_regression_gt_5pct |
| few-large | file | 8 | 93.46/578.13 | 129.09/586.33 | 258.43/588.53 | 101.65/562.73 | 453.59% | 24.93% | rss_regression_gt_5pct |
| few-large | instrumented | 1 | 82228.95/82312.08 | 82560.84/82722.67 | 83180.72/82856.78 | 82165.34/82246.05 | 0.10% | 0.61% | — |
| few-large | instrumented | 2 | 39907.46/82312.08 | 40173.94/82722.67 | 40850.41/82856.78 | 39947.02/82246.05 | 105.89% | 15.21% | rss_regression_gt_5pct |
| few-large | instrumented | 4 | 20023.90/82312.08 | 20258.47/82722.67 | 20678.93/82856.78 | 20079.96/82246.05 | 309.59% | 25.33% | rss_regression_gt_5pct |
| few-large | instrumented | 8 | 19979.46/82312.08 | 20238.64/82722.67 | 20660.77/82856.78 | 20026.39/82246.05 | 310.69% | 23.71% | rss_regression_gt_5pct |
| many-small | owned | 1 | 94.41/95.86 | 97.32/101.40 | 106.54/103.42 | 94.80/96.57 | 1.87% | 0.00% | — |
| many-small | owned | 2 | 186.51/95.86 | 210.66/101.40 | 257.93/103.42 | 189.68/96.57 | -49.09% | 5.57% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | owned | 4 | 186.04/95.86 | 207.45/101.40 | 214.94/103.42 | 186.73/96.57 | -48.28% | 5.94% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | owned | 8 | 228.74/95.86 | 255.74/101.40 | 261.35/103.42 | 230.40/96.57 | -58.08% | 1.44% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct |
| many-small | file | 1 | 200.50/182.04 | 207.84/191.14 | 211.26/193.96 | 201.41/183.35 | -8.96% | 0.35% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct |
| many-small | file | 2 | 296.04/182.04 | 331.34/191.14 | 356.41/193.96 | 298.70/183.35 | -38.62% | 5.92% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | file | 4 | 248.77/182.04 | 262.89/191.14 | 266.63/193.96 | 248.43/183.35 | -26.19% | 5.25% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | file | 8 | 273.70/182.04 | 307.19/191.14 | 404.18/193.96 | 277.02/183.35 | -33.81% | 1.46% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct |
| many-small | instrumented | 1 | 29574.47/29726.19 | 29895.05/30006.88 | 30136.58/30210.94 | 29635.31/29760.16 | 0.42% | 0.40% | — |
| many-small | instrumented | 2 | 15467.24/29726.19 | 15558.98/30006.88 | 15646.07/30210.94 | 15469.75/29760.16 | 92.38% | 6.27% | rss_regression_gt_5pct |
| many-small | instrumented | 4 | 7821.47/29726.19 | 7898.35/30006.88 | 7959.80/30210.94 | 7829.25/29760.16 | 280.11% | 7.14% | rss_regression_gt_5pct |
| many-small | instrumented | 8 | 4068.20/29726.19 | 4788.38/30006.88 | 5593.08/30210.94 | 4158.11/29760.16 | 615.71% | 1.90% | — |

## Flag counts

RSS flags are counted only at aggregate comparison scope because each capture has one whole-child receipt; per-repeat counts cover latency and throughput flags only.

| scope | flag | count |
| --- | --- | ---: |
| aggregate | latency_mean_us_regression_gt_5pct | 7 |
| aggregate | latency_p50_us_regression_gt_5pct | 7 |
| aggregate | latency_p95_us_regression_gt_5pct | 7 |
| aggregate | latency_p99_us_regression_gt_5pct | 8 |
| aggregate | rss_regression_gt_5pct | 15 |
| aggregate | throughput_regression_gt_5pct | 7 |
| per-repeat, RSS excluded | latency_mean_us_regression_gt_5pct | 14 |
| per-repeat, RSS excluded | latency_p50_us_regression_gt_5pct | 14 |
| per-repeat, RSS excluded | latency_p95_us_regression_gt_5pct | 14 |
| per-repeat, RSS excluded | latency_p99_us_regression_gt_5pct | 15 |
| per-repeat, RSS excluded | throughput_regression_gt_5pct | 14 |

The JSON contains the corresponding per-repeat latency, throughput, and RSS comparisons. The rows below use batch worker 1 as the scaling reference; requested efficiency is p50 latency speedup divided by requested worker count, and effective efficiency uses the selected-Part work cap.

## Batch scaling

| corpus | source | workers | p50 speedup | throughput speedup | efficiency requested/effective | simple Amdahl result |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| few-large | owned | 2 | 4.52 | 4.37 | 2.26/2.26 | invalid: superlinear |
| few-large | owned | 4 | 6.76 | 6.04 | 1.69/1.69 | invalid: superlinear |
| few-large | owned | 8 | 6.70 | 5.98 | 0.84/1.67 | invalid: 4-Part cap |
| few-large | file | 2 | 4.10 | 3.89 | 2.05/2.05 | invalid: superlinear |
| few-large | file | 4 | 6.13 | 5.52 | 1.53/1.53 | invalid: superlinear |
| few-large | file | 8 | 6.24 | 5.62 | 0.78/1.56 | invalid: 4-Part cap |
| few-large | instrumented | 2 | 2.06 | 2.06 | 1.03/1.03 | invalid: superlinear |
| few-large | instrumented | 4 | 4.11 | 4.09 | 1.03/1.03 | invalid: superlinear |
| few-large | instrumented | 8 | 4.12 | 4.10 | 0.51/1.03 | invalid: 4-Part cap |
| many-small | owned | 2 | 0.51 | 0.50 | 0.25/0.25 | invalid: no speedup |
| many-small | owned | 4 | 0.51 | 0.51 | 0.13/0.13 | invalid: no speedup |
| many-small | owned | 8 | 0.41 | 0.41 | 0.05/0.05 | invalid: no speedup |
| many-small | file | 2 | 0.68 | 0.67 | 0.34/0.34 | invalid: no speedup |
| many-small | file | 4 | 0.81 | 0.81 | 0.20/0.20 | invalid: no speedup |
| many-small | file | 8 | 0.73 | 0.73 | 0.09/0.09 | invalid: no speedup |
| many-small | instrumented | 2 | 1.91 | 1.92 | 0.96/0.96 | 0.05 |
| many-small | instrumented | 4 | 3.78 | 3.79 | 0.95/0.95 | 0.02 |
| many-small | instrumented | 8 | 7.27 | 7.13 | 0.91/0.91 | 0.01 |

Simple-Amdahl values are descriptive model outputs for matched captures. Superlinear observations and widths above the selected-Part work cap are labelled invalid for this model rather than being fitted or clamped; raw fractions remain in JSON for traceability. No scaling row establishes a causal mechanism.

## Matched baseline boundary

Both 0499 executables use the unchanged harness with plain data() on the ordinary serial path. This table compares serial and batch routes only within the after executable; comparison.json contains the separate matched before/after results. The earlier 0498 historical accounting-path confound does not apply to these fresh 0499 controls.

The final files are process-level measurements: setup, repeated package opens, and post-timer verification are included in the whole-child RSS receipt, while the timed CSV interval excludes corpus construction and post-timer digest verification. These results support descriptive matched comparisons and do not establish operation-local causal costs.
