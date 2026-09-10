# Change 0498 final benchmark analysis

Generated: `2026-09-10T19:46:36+00:00`

The analysis uses measured sample indexes 3–32 from each of two repeats; indexes 0–2 are warmups. Percentiles use the nearest-rank rule used by the Rust harness. RSS is the whole child maximum from `/usr/bin/time -v`. The measured process was pinned to the configured benchmark CPU set on a shared host.

## Serial versus batch

Positive latency or RSS deltas and negative throughput deltas are adverse. The flag threshold is greater than 5 percent.

| corpus | source | workers | p50 batch/serial (us) | p95 batch/serial (us) | p99 batch/serial (us) | mean batch/serial (us) | throughput delta | RSS delta | flags |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| few-large | owned | 1 | 581.66/575.85 | 589.30/586.71 | 665.27/592.94 | 577.44/563.33 | -2.44% | -0.06% | latency_p99_us_regression_gt_5pct |
| few-large | owned | 2 | 136.46/575.85 | 147.57/586.71 | 172.38/592.94 | 138.76/563.33 | 305.99% | 14.75% | rss_regression_gt_5pct |
| few-large | owned | 4 | 82.36/575.85 | 100.24/586.71 | 243.48/592.94 | 90.51/563.33 | 522.42% | 24.99% | rss_regression_gt_5pct |
| few-large | owned | 8 | 82.52/575.85 | 105.89/586.71 | 255.29/592.94 | 92.08/563.33 | 511.78% | 25.21% | rss_regression_gt_5pct |
| few-large | file | 1 | 577.20/578.63 | 588.87/586.32 | 597.11/600.76 | 566.56/567.25 | 0.12% | 0.04% | — |
| few-large | file | 2 | 146.22/578.63 | 190.22/586.32 | 534.71/600.76 | 157.99/567.25 | 259.04% | 15.42% | rss_regression_gt_5pct |
| few-large | file | 4 | 91.02/578.63 | 115.39/586.32 | 255.82/600.76 | 99.59/567.25 | 469.59% | 24.58% | rss_regression_gt_5pct |
| few-large | file | 8 | 91.70/578.63 | 107.93/586.32 | 260.04/600.76 | 98.03/567.25 | 478.67% | 20.15% | rss_regression_gt_5pct |
| few-large | instrumented | 1 | 82474.82/82119.57 | 82566.89/82520.22 | 82760.61/82646.33 | 82361.90/82116.12 | -0.30% | -0.30% | — |
| few-large | instrumented | 2 | 39924.87/82119.57 | 40228.21/82520.22 | 40303.18/82646.33 | 39980.81/82116.12 | 105.39% | 14.22% | rss_regression_gt_5pct |
| few-large | instrumented | 4 | 20028.23/82119.57 | 20332.73/82520.22 | 20715.60/82646.33 | 20090.43/82116.12 | 308.73% | 24.34% | rss_regression_gt_5pct |
| few-large | instrumented | 8 | 20016.39/82119.57 | 20260.52/82520.22 | 20707.01/82646.33 | 20076.83/82116.12 | 309.01% | 25.60% | rss_regression_gt_5pct |
| many-small | owned | 1 | 93.92/94.77 | 98.71/100.85 | 105.36/102.74 | 94.50/95.18 | 0.73% | -0.37% | — |
| many-small | owned | 2 | 677.78/94.77 | 747.87/100.85 | 768.79/102.74 | 686.70/95.18 | -86.14% | 4.08% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct |
| many-small | owned | 4 | 560.02/94.77 | 653.39/100.85 | 694.31/102.74 | 563.27/95.18 | -83.10% | 8.17% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | owned | 8 | 480.61/94.77 | 546.69/100.85 | 587.76/102.74 | 481.15/95.18 | -80.22% | 15.61% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | file | 1 | 201.81/182.26 | 235.98/191.48 | 238.36/195.40 | 205.01/183.76 | -10.36% | -0.72% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct |
| many-small | file | 2 | 727.71/182.26 | 823.75/191.48 | 863.97/195.40 | 735.10/183.76 | -75.00% | 5.59% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | file | 4 | 577.15/182.26 | 633.24/191.48 | 855.57/195.40 | 581.33/183.76 | -68.39% | 6.71% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | file | 8 | 512.80/182.26 | 578.92/191.48 | 754.17/195.40 | 519.21/183.76 | -64.61% | 12.99% | latency_p50_us_regression_gt_5pct, latency_p95_us_regression_gt_5pct, latency_p99_us_regression_gt_5pct, latency_mean_us_regression_gt_5pct, throughput_regression_gt_5pct, rss_regression_gt_5pct |
| many-small | instrumented | 1 | 29600.10/29572.76 | 29994.87/29662.56 | 30074.74/30144.34 | 29655.28/29594.74 | -0.20% | -0.65% | — |
| many-small | instrumented | 2 | 16040.22/29572.76 | 16056.85/29662.56 | 17043.52/30144.34 | 16052.97/29594.74 | 84.36% | -1.14% | — |
| many-small | instrumented | 4 | 8016.39/29572.76 | 8155.03/29662.56 | 8418.76/30144.34 | 8043.98/29594.74 | 267.91% | 2.55% | — |
| many-small | instrumented | 8 | 4190.62/29572.76 | 4250.52/29662.56 | 4290.40/30144.34 | 4192.60/29594.74 | 605.88% | 0.00% | — |

## Flag counts

RSS flags are counted only at aggregate comparison scope because each capture has one whole-child receipt; per-repeat counts cover latency and throughput flags only.

| scope | flag | count |
| --- | --- | ---: |
| aggregate | latency_mean_us_regression_gt_5pct | 7 |
| aggregate | latency_p50_us_regression_gt_5pct | 7 |
| aggregate | latency_p95_us_regression_gt_5pct | 7 |
| aggregate | latency_p99_us_regression_gt_5pct | 8 |
| aggregate | rss_regression_gt_5pct | 14 |
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
| few-large | owned | 2 | 4.26 | 4.16 | 2.13/2.13 | invalid: superlinear |
| few-large | owned | 4 | 7.06 | 6.38 | 1.77/1.77 | invalid: superlinear |
| few-large | owned | 8 | 7.05 | 6.27 | 0.88/1.76 | invalid: 4-Part cap |
| few-large | file | 2 | 3.95 | 3.59 | 1.97/1.97 | invalid: superlinear |
| few-large | file | 4 | 6.34 | 5.69 | 1.59/1.59 | invalid: superlinear |
| few-large | file | 8 | 6.29 | 5.78 | 0.79/1.57 | invalid: 4-Part cap |
| few-large | instrumented | 2 | 2.07 | 2.06 | 1.03/1.03 | invalid: superlinear |
| few-large | instrumented | 4 | 4.12 | 4.10 | 1.03/1.03 | invalid: superlinear |
| few-large | instrumented | 8 | 4.12 | 4.10 | 0.52/1.03 | invalid: 4-Part cap |
| many-small | owned | 2 | 0.14 | 0.14 | 0.07/0.07 | invalid: no speedup |
| many-small | owned | 4 | 0.17 | 0.17 | 0.04/0.04 | invalid: no speedup |
| many-small | owned | 8 | 0.20 | 0.20 | 0.02/0.02 | invalid: no speedup |
| many-small | file | 2 | 0.28 | 0.28 | 0.14/0.14 | invalid: no speedup |
| many-small | file | 4 | 0.35 | 0.35 | 0.09/0.09 | invalid: no speedup |
| many-small | file | 8 | 0.39 | 0.39 | 0.05/0.05 | invalid: no speedup |
| many-small | instrumented | 2 | 1.85 | 1.85 | 0.92/0.92 | 0.08 |
| many-small | instrumented | 4 | 3.69 | 3.69 | 0.92/0.92 | 0.03 |
| many-small | instrumented | 8 | 7.06 | 7.07 | 0.88/0.88 | 0.02 |

Simple-Amdahl values are descriptive model outputs for matched captures. Superlinear observations and widths above the selected-Part work cap are labelled invalid for this model rather than being fitted or clamped; raw fractions remain in JSON for traceability. No scaling row establishes a causal mechanism.

## Historical baseline boundary

The retained before executable used data_with_accounting on its primary serial path, while the final serial control uses plain data() by default. Any before-versus-final latency comparison is therefore descriptive and has an accounting-path confound; this script does not merge historical baseline numbers into the matched batch/scaling tables.

The final files are process-level measurements: setup, repeated package opens, and post-timer verification are included in the whole-child RSS receipt, while the timed CSV interval excludes corpus construction and post-timer digest verification. These results support descriptive matched comparisons and do not establish operation-local causal costs.
