# 0821 real-file save durability attribution

Offline replay of admitted real-file save-durability receipts.
Native elapsed values describe configuration attribution; no production optimization, historical timing, or default-weakening claim is made.
Default save and explicit full durability are matched controls. File-only and no-sync change durability semantics.
Observer allocation and procfs counters remain diagnostic and are never pooled with native latency.
Analysis p50/p95/p99 use nearest rank within each process block; raw harness p50 is its integer midpoint. Ratios pair each policy with default in the same block, format, and phase.

Native reports/samples: 144 / 4320
Observer reports/samples: 48 / 144
Qualification reports/samples: 24 / 24

| Selector | policy | p50 ns | p95 ns | p99 ns | p50 CI95 ns | ratio/default | ratio CI95 | source B/s | published B/s | spread | tail |
|---|---|---:|---:|---:|---|---:|---|---:|---:|---|---|
| docx_real_file_ordinary_save_lifecycle | default | 5.22294e+06 | 5.35774e+06 | 5.42715e+06 | [5.1904e+06, 5.234e+06] | 1 | [1, 1] | 4.49996e+06 | 4.50609e+06 | p95,p99 | - |
| docx_real_file_ordinary_save_lifecycle | full | 5.20502e+06 | 5.3488e+06 | 5.44843e+06 | [5.19414e+06, 5.21719e+06] | 0.998016 | [0.994463, 1.00161] | 4.51545e+06 | 4.5216e+06 | p99 | - |
| docx_real_file_ordinary_save_lifecycle | file-only | 3.37792e+06 | 3.49815e+06 | 3.53148e+06 | [3.35451e+06, 3.39106e+06] | 0.647016 | [0.64229, 0.651661] | 6.95784e+06 | 6.96731e+06 | p99 | - |
| docx_real_file_ordinary_save_lifecycle | no-sync | 242312 | 270827 | 277566 | [237571, 244776] | 0.0463158 | [0.0456941, 0.0469256] | 9.69948e+07 | 9.71268e+07 | p99 | yes |
| docx_real_file_ordinary_save_atomic_publish | default | 5.02451e+06 | 5.14704e+06 | 5.19516e+06 | [5.0125e+06, 5.03953e+06] | 1 | [1, 1] | 4.67767e+06 | 4.68404e+06 | - | - |
| docx_real_file_ordinary_save_atomic_publish | full | 5.02684e+06 | 5.15813e+06 | 5.2212e+06 | [5.01588e+06, 5.04656e+06] | 1.00087 | [0.998006, 1.00368] | 4.6755e+06 | 4.68187e+06 | p99 | - |
| docx_real_file_ordinary_save_atomic_publish | file-only | 3.19592e+06 | 3.30794e+06 | 3.32873e+06 | [3.17803e+06, 3.20845e+06] | 0.635722 | [0.631614, 0.639438] | 7.35406e+06 | 7.36407e+06 | - | - |
| docx_real_file_ordinary_save_atomic_publish | no-sync | 83995.5 | 102591 | 105336 | [83165, 85165] | 0.0167552 | [0.0165419, 0.0169114] | 2.79813e+08 | 2.80194e+08 | p95,p99 | yes |
| xlsx_real_file_ordinary_save_lifecycle | default | 5.41084e+06 | 5.54776e+06 | 5.60957e+06 | [5.38802e+06, 5.42023e+06] | 1 | [1, 1] | 1.55891e+06 | 1.5748e+06 | - | - |
| xlsx_real_file_ordinary_save_lifecycle | full | 5.41563e+06 | 5.53608e+06 | 5.59842e+06 | [5.37757e+06, 5.43469e+06] | 1.00087 | [0.993522, 1.00727] | 1.55753e+06 | 1.57341e+06 | - | - |
| xlsx_real_file_ordinary_save_lifecycle | file-only | 3.59399e+06 | 3.8049e+06 | 3.8686e+06 | [3.57169e+06, 3.6039e+06] | 0.663721 | [0.660877, 0.667434] | 2.34698e+06 | 2.3709e+06 | p99 | yes |
| xlsx_real_file_ordinary_save_lifecycle | no-sync | 534128 | 602153 | 735094 | [531463, 538373] | 0.0990247 | [0.0980517, 0.0996071] | 1.57921e+07 | 1.59531e+07 | p95,p99,mean | yes |
| xlsx_real_file_ordinary_save_atomic_publish | default | 4.94868e+06 | 5.05503e+06 | 5.11111e+06 | [4.93283e+06, 4.96126e+06] | 1 | [1, 1] | 1.70449e+06 | 1.72187e+06 | p99 | - |
| xlsx_real_file_ordinary_save_atomic_publish | full | 4.93782e+06 | 5.04041e+06 | 5.09277e+06 | [4.91692e+06, 4.96841e+06] | 0.998569 | [0.9933, 1.00418] | 1.70824e+06 | 1.72566e+06 | - | - |
| xlsx_real_file_ordinary_save_atomic_publish | file-only | 3.10477e+06 | 3.2014e+06 | 3.22382e+06 | [3.08817e+06, 3.12492e+06] | 0.628414 | [0.624003, 0.630893] | 2.71679e+06 | 2.74449e+06 | - | - |
| xlsx_real_file_ordinary_save_atomic_publish | no-sync | 99605 | 129920 | 154821 | [98875.5, 99985] | 0.0200907 | [0.0199802, 0.0202552] | 8.46845e+07 | 8.55479e+07 | p95,p99,mean | yes |
| pptx_real_file_ordinary_save_lifecycle | default | 7.4206e+06 | 7.56774e+06 | 7.58813e+06 | [7.40468e+06, 7.44222e+06] | 1 | [1, 1] | 9.27446e+06 | 9.20195e+06 | p99 | - |
| pptx_real_file_ordinary_save_lifecycle | full | 7.40384e+06 | 7.54115e+06 | 7.622e+06 | [7.38905e+06, 7.41923e+06] | 0.99839 | [0.993476, 1.0007] | 9.29544e+06 | 9.22278e+06 | p95,p99 | - |
| pptx_real_file_ordinary_save_lifecycle | file-only | 5.56978e+06 | 5.64884e+06 | 5.69114e+06 | [5.54826e+06, 5.60219e+06] | 0.750027 | [0.747717, 0.754906] | 1.23563e+07 | 1.22597e+07 | p95,p99 | - |
| pptx_real_file_ordinary_save_lifecycle | no-sync | 2.0782e+06 | 2.152e+06 | 2.154e+06 | [2.07081e+06, 2.11308e+06] | 0.280308 | [0.278623, 0.284738] | 3.31162e+07 | 3.28574e+07 | p99 | - |
| pptx_real_file_ordinary_save_atomic_publish | default | 5.61833e+06 | 5.75646e+06 | 5.80605e+06 | [5.60234e+06, 5.64737e+06] | 1 | [1, 1] | 1.22496e+07 | 1.21538e+07 | p99 | - |
| pptx_real_file_ordinary_save_atomic_publish | full | 5.60874e+06 | 5.76743e+06 | 5.79678e+06 | [5.60551e+06, 5.61491e+06] | 0.997913 | [0.993732, 1.00147] | 1.22705e+07 | 1.21746e+07 | p99 | - |
| pptx_real_file_ordinary_save_atomic_publish | file-only | 3.78064e+06 | 3.8688e+06 | 3.8996e+06 | [3.76538e+06, 3.79688e+06] | 0.673256 | [0.670654, 0.673441] | 1.82038e+07 | 1.80615e+07 | - | - |
| pptx_real_file_ordinary_save_atomic_publish | no-sync | 296616 | 351747 | 355367 | [293047, 301437] | 0.0527395 | [0.0522058, 0.053536] | 2.32024e+08 | 2.3021e+08 | p95,p99,mean | yes |

No optimization, adoption, or historical comparison is inferred.
