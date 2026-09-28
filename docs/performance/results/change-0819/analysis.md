# 0819 real-file ordinary-save baseline

Offline replay of admitted real-file ordinary-save receipts.
Native elapsed values are a baseline only; no historical timing or speedup claim is made.
Observer allocation and procfs counters remain diagnostic and are never pooled with native latency.
Analysis p50/p95/p99 use nearest rank within each process block; raw harness p50 is its integer midpoint.

Native reports/samples: 72 / 2160
Observer reports/samples: 24 / 72
Qualification reports/samples: 12 / 12

| Selector | p50 ns | p95 ns | p99 ns | mean ns | p50 CI95 ns | source B/s | published B/s | spread | tail |
|---|---:|---:|---:|---:|---|---:|---:|---|---|
| docx_real_file_ordinary_save_lifecycle | 5.22022e+06 | 5.34644e+06 | 5.38023e+06 | 5.22523e+06 | [5.19704e+06, 5.23058e+06] | 4.5023e+06 | 4.50843e+06 | p99 | - |
| docx_real_file_ordinary_save_edit | 55670 | 64020.5 | 68666 | 56852.6 | [55530, 55945] | 4.22184e+08 | 4.22759e+08 | p95,p99 | yes |
| docx_real_file_ordinary_save_atomic_publish | 5.02131e+06 | 5.12374e+06 | 5.14336e+06 | 5.02928e+06 | [5.01416e+06, 5.02548e+06] | 4.68065e+06 | 4.68702e+06 | p99 | - |
| docx_real_file_ordinary_save_counting_publish | 51235.5 | 63445.5 | 67035.5 | 53514.6 | [50960, 52225.5] | 4.58725e+08 | 4.59349e+08 | p95,p99 | yes |
| xlsx_real_file_ordinary_save_lifecycle | 5.41955e+06 | 5.58535e+06 | 5.68344e+06 | 5.41904e+06 | [5.4093e+06, 5.43638e+06] | 1.5564e+06 | 1.57227e+06 | p99 | - |
| xlsx_real_file_ordinary_save_edit | 262591 | 414977 | 437247 | 286319 | [261942, 268231] | 3.21222e+07 | 3.24497e+07 | p95,p99,mean | yes |
| xlsx_real_file_ordinary_save_atomic_publish | 4.93068e+06 | 5.05908e+06 | 5.1081e+06 | 4.934e+06 | [4.91173e+06, 4.93814e+06] | 1.71072e+06 | 1.72816e+06 | - | - |
| xlsx_real_file_ordinary_save_counting_publish | 65325.5 | 80265 | 109866 | 68149.6 | [65036, 66010.5] | 1.29123e+08 | 1.30439e+08 | p95,p99,mean | yes |
| pptx_real_file_ordinary_save_lifecycle | 7.42231e+06 | 7.52864e+06 | 7.58629e+06 | 7.42004e+06 | [7.40155e+06, 7.44904e+06] | 9.27231e+06 | 9.19983e+06 | - | - |
| pptx_real_file_ordinary_save_edit | 1.40743e+06 | 1.42171e+06 | 1.42584e+06 | 1.40875e+06 | [1.40437e+06, 1.41198e+06] | 4.88992e+07 | 4.85169e+07 | - | - |
| pptx_real_file_ordinary_save_atomic_publish | 5.6205e+06 | 5.73326e+06 | 5.75292e+06 | 5.62306e+06 | [5.59303e+06, 5.64588e+06] | 1.22448e+07 | 1.21491e+07 | - | - |
| pptx_real_file_ordinary_save_counting_publish | 304592 | 317317 | 319482 | 306707 | [301556, 305677] | 2.25948e+08 | 2.24182e+08 | p99 | - |

No optimization, adoption, or historical comparison is inferred.
