# Fresh CFB emission paired results

Decision: **adopt**. 270 reports / 6534 samples; 0 frozen guard flags.

Ratios are medians of six matched process ratios. Absolute columns are medians of process summaries; they need not divide to the paired ratio.

| Case | Before p50 ms | After p50 ms | p50 ratio [95% CI] | RSS ratio [95% CI] | p95 ratio | p99 ratio |
| --- | ---: | ---: | --- | --- | ---: | ---: |
| doc-tiny | 0.006620 | 0.006640 | 1.003018 [0.987263, 1.011412] | 1.002063 [0.975607, 1.027688] | 1.005026 | 0.993799 |
| doc-large | 0.130186 | 0.126526 | 0.967539 [0.961656, 0.979714] | 1.048042 [1.012574, 1.053363] | 0.915221 | 0.903675 |
| doc-payload | 0.696438 | 0.623442 | 0.901444 [0.892770, 0.903134] | 1.001074 [1.000076, 1.004552] | 0.901401 | 0.887745 |
| cfb-tiny | 0.001635 | 0.001590 | 0.975535 [0.951515, 0.990683] | 1.010844 [0.967482, 1.069409] | 0.983130 | 0.978392 |
| cfb-large | 0.643088 | 0.620873 | 0.951122 [0.938176, 0.969352] | 0.998241 [0.996550, 1.001343] | 0.940599 | 0.955253 |
| cfb-large-only | 0.610317 | 0.612748 | 1.001106 [0.993861, 1.008835] | 1.008441 [0.995587, 1.012978] | 0.999588 | 0.992417 |
| cfb-mini-only | 0.002225 | 0.002275 | 1.015725 [1.011192, 1.026779] | 1.023050 [0.991814, 1.026760] | 0.987455 | 0.979974 |
| cfb-v4 | 0.649913 | 0.607487 | 0.952860 [0.933729, 0.970989] | 1.000921 [0.996898, 1.006444] | 0.964362 | 0.965396 |
| cfb-difat | 2.489796 | 2.489290 | 1.002937 [0.984037, 1.007982] | 1.001034 [0.997526, 1.006466] | 1.002672 | 1.007152 |

## Operation allocations

| Case | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta before → after |
| --- | ---: | ---: | ---: | ---: |
| doc-tiny | 165 → 166 | 80025 → 73305 | 47745 → 40193 | 25984 → 18432 |
| doc-large | 2798 → 2798 | 1.40482e+06 → 1.36738e+06 | 801246 → 776286 | 189824 → 164864 |
| doc-payload | 915 → 915 | 2.36976e+07 → 2.36724e+07 | 1.81829e+07 → 1.81662e+07 | 1.02742e+07 → 1.02574e+07 |
| cfb-tiny | 44 → 44 | 42130 → 40594 | 31734 → 30710 | 18432 → 17408 |
| cfb-large | 108 → 108 | 1.69932e+07 → 1.69871e+07 | 1.26644e+07 → 1.26603e+07 | 8.39373e+06 → 8.38963e+06 |
| cfb-large-only | 96 → 96 | 1.69165e+07 → 1.69165e+07 | 1.26553e+07 → 1.26553e+07 | 8.38963e+06 → 8.38963e+06 |
| cfb-mini-only | 50 → 50 | 33628 → 33628 | 25004 → 25004 | 10752 → 10752 |
| cfb-v4 | 45 → 45 | 1.68437e+07 → 1.68376e+07 | 1.26224e+07 → 1.26183e+07 | 8.4009e+06 → 8.3968e+06 |
| cfb-difat | 166 → 166 | 3.3894e+07 → 3.3894e+07 | 2.53073e+07 → 2.53073e+07 | 1.67782e+07 → 1.67782e+07 |

Full intervals and all fault/quantile rows are in paired.csv and analysis.json. RSS is whole-child GNU-time maximum, including setup and verification. Allocator regions exclude setup, verification and output destruction.
