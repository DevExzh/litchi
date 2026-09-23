| case | before p50 (ms) | after p50 (ms) | paired p50 ratio | 95% CI | before p95 | after p95 | mean ratio |
| --- | ---: | ---: | ---: | --- | ---: | ---: | ---: |
| `xls_semantic_one_edit_save/large` | 3.3771 | 1.7998 | 0.533 | [0.527, 0.546] | 3.4572 | 1.8430 | 0.532 |
| `xls_semantic_one_edit_save/tiny` | 0.0802 | 0.0420 | 0.524 | [0.517, 0.526] | 0.0880 | 0.0467 | 0.525 |
| `xls_visibility_eager_edit_save` | 25.33 | 5.4705 | 0.216 | [0.216, 0.216] | 25.40 | 5.5077 | 0.216 |
| `xls_visibility_eager_batch_edit_save` | 25.34 | 5.4698 | 0.216 | [0.216, 0.217] | 25.43 | 5.4995 | 0.216 |
| `xls_numeric_eager_rk_mulrk_edit_save` | 2.3790 | 0.5068 | 0.213 | [0.213, 0.214] | 2.3907 | 0.5170 | 0.213 |
| `xls_visibility_source_backed_edit_save` | 10.15 | 2.2377 | 0.220 | [0.219, 0.221] | 10.17 | 2.2575 | 0.221 |
| `xls_visibility_source_backed_batch_edit_save` | 10.25 | 2.2924 | 0.224 | [0.223, 0.224] | 10.32 | 2.3226 | 0.224 |
| `xls_comments_source_backed_edit_save` | 33.49 | 20.26 | 0.605 | [0.491, 0.607] | 33.67 | 20.39 | 0.605 |
| `xls_comments_source_backed_batch_edit_save` | 36.76 | 20.90 | 0.568 | [0.502, 0.570] | 37.00 | 21.09 | 0.568 |
| `xls_numeric_source_backed_number_edit_save` | 78.40 | 30.61 | 0.391 | [0.389, 0.392] | 78.77 | 30.93 | 0.391 |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | 0.8240 | 0.2625 | 0.318 | [0.316, 0.319] | 0.8286 | 0.2760 | 0.321 |
| `xls_numeric_plan_only_number_edit_save` | 33.82 | 17.96 | 0.531 | [0.531, 0.532] | 33.93 | 18.08 | 0.532 |
| `xls_numeric_plan_only_rk_mulrk_edit_save` | 0.4028 | 0.2186 | 0.542 | [0.541, 0.544] | 0.4141 | 0.2282 | 0.543 |
| `cfb_file_owned_same_length_overlay_atomic_save` | 65.81 | 49.79 | 0.758 | [0.753, 0.758] | 69.40 | 52.67 | 0.759 |
| `cfb_file_same_length_overlay_atomic_save` | 98.17 | 97.94 | 0.998 | [0.992, 1.002] | 100.87 | 109.72 | 1.011 |
| `ppt_source_backed_one_shape_text/large` | 0.0287 | 0.0280 | 0.978 | [0.962, 0.987] | 0.0308 | 0.0307 | 0.973 |
| `xls_comments_eager_edit_save` | 21.42 | 21.00 | 0.980 | [0.972, 0.998] | 22.41 | 21.46 | 0.982 |
| `xls_numeric_eager_number_edit_save` | 28.08 | 28.00 | 0.995 | [0.985, 1.015] | 28.84 | 28.48 | 0.992 |
| `ole_common_one_edit_save` | 20.89 | 20.78 | 0.989 | [0.971, 1.024] | 21.61 | 21.59 | 0.988 |
| `doc_semantic_one_edit_save/large` | 0.6983 | 0.6953 | 0.994 | [0.974, 1.008] | 0.8037 | 0.8020 | 0.992 |
| `xls_semantic_open/large` | 1.4977 | 1.4813 | 0.982 | [0.962, 1.051] | 1.5564 | 1.5753 | 0.989 |
| `probe/54016/number-generic` | 23.68 | 14.26 | 0.600 | [0.591, 0.610] | 24.35 | 15.05 | 0.604 |
| `probe/54016/string-generic` | 24.47 | 14.97 | 0.619 | [0.577, 0.646] | 24.85 | 15.70 | 0.619 |
| `probe/54016/number-source-backed` | 15.78 | 14.05 | 0.891 | [0.869, 0.931] | 16.14 | 14.56 | 0.894 |
| `probe/54016/number-plan` | 11.88 | 11.72 | 0.995 | [0.940, 1.062] | 12.32 | 12.80 | 1.006 |
| `probe/54016/number-plan-publish` | 0.9497 | 0.0432 | 0.045 | [0.045, 0.048] | 0.9572 | 0.0477 | 0.046 |
| `probe/54016/open` | 11.75 | 11.88 | 1.010 | [0.979, 1.057] | 12.29 | 12.53 | 1.007 |
| `probe/xls-large/number-generic` | 3.4459 | 1.8662 | 0.546 | [0.530, 0.559] | 3.5238 | 1.9263 | 0.546 |
| `probe/xls-large/number-source-backed` | 2.0998 | 1.7512 | 0.839 | [0.819, 0.868] | 2.1652 | 1.8109 | 0.839 |

Regression flags (>5%):
   ('cfb_file_same_length_overlay_atomic_save', 'median', None, 'p95', 1.0877164275767868)
   ('xls_semantic_open/large', 'pair', 2, 'p50', 1.0507812096687044)
   ('xls_semantic_open/large', 'pair', 2, 'mean', 1.053801067425224)
   ('probe/54016/number-plan', 'pair', 2, 'p50', 1.0619957356849241)
   ('probe/54016/open', 'pair', 3, 'p50', 1.056523551051932)
   ('probe/54016/open', 'pair', 3, 'mean', 1.0541694558814314)

| case | before instructions (M) | after (M) | ratio | before cycles (M) | after (M) | ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `probe/54016/number-generic` | 213.622 | 162.427 | 0.760 | 108.422 | 67.809 | 0.625 |
| `probe/54016/number-source-backed` | 179.078 | 163.769 | 0.915 | 76.007 | 63.431 | 0.835 |
| `probe/54016/number-plan` | 148.272 | 143.132 | 0.965 | 61.265 | 56.868 | 0.928 |
| `probe/54016/open` | 148.500 | 148.407 | 0.999 | 51.091 | 52.125 | 1.020 |
| `probe/xls-large/number-generic` | 34.096 | 25.538 | 0.749 | 14.816 | 8.089 | 0.546 |
| `xls_semantic_one_edit_save/large` | 227.007 | 218.346 | 0.962 | 71.374 | 64.505 | 0.904 |
| `xls_visibility_eager_edit_save` | 180.159 | 69.619 | 0.386 | 135.362 | 46.581 | 0.344 |
| `xls_visibility_source_backed_edit_save` | 97.943 | 53.740 | 0.549 | 73.373 | 36.714 | 0.500 |
| `xls_numeric_eager_rk_mulrk_edit_save` | 17.066 | 6.538 | 0.383 | 12.309 | 3.467 | 0.282 |
| `xls_numeric_plan_only_rk_mulrk_edit_save` | 4.472 | 3.431 | 0.767 | 1.868 | 2.466 | 1.320 |
| `xls_numeric_source_backed_rk_mulrk_edit_save` | 7.986 | 4.851 | 0.607 | 6.368 | 3.364 | 0.528 |
| `xls_comments_source_backed_edit_save` | 346.826 | 260.137 | 0.750 | 282.694 | 212.133 | 0.750 |
| `xls_comments_eager_edit_save` | 204.387 | 204.349 | 1.000 | 187.161 | 184.084 | 0.984 |
| `ppt_source_backed_one_shape_text/large` | 6.844 | 6.844 | 1.000 | 1.627 | 1.629 | 1.001 |
| `doc_semantic_one_edit_save/large` | 41.391 | 41.386 | 1.000 | 10.870 | 10.962 | 1.008 |
| `cfb_file_same_length_overlay_atomic_save` | 9239.267 | 9239.263 | 1.000 | 3774.102 | 3768.773 | 0.999 |
| `cfb_file_owned_same_length_overlay_atomic_save` | 8811.525 | 8638.093 | 0.980 | 3455.183 | 3311.957 | 0.959 |
