# Separate 0844 tail confirmation

All times are ns; RSS KiB; faults counts. Six paired blocks,1000samples perchild,3warmups.

| Shape | Parts | Mode | Metric | Before median | After median | Paired ratio | CI low | CI high |
|---|---:|---|---|---:|---:|---:|---:|---:|
| shared | 4 | lifecycle | p50_ns | 37610.000 | 37360.000 | 0.996151 | 0.985086 | 1.003879 |
| shared | 4 | lifecycle | p95_ns | 42255.500 | 41905.500 | 0.991226 | 0.975105 | 1.015556 |
| shared | 4 | lifecycle | p99_ns | 47770.000 | 47955.000 | 1.009030 | 0.994376 | 1.021066 |
| shared | 4 | lifecycle | mean_ns | 38137.004 | 37921.865 | 0.996933 | 0.984895 | 1.004541 |
| shared | 4 | lifecycle | rss_kib | 5530.000 | 5536.000 | 0.999999 | 0.979748 | 1.019636 |
| shared | 4 | lifecycle | minor_faults | 641.500 | 631.000 | 0.986709 | 0.981354 | 0.991405 |
| shared | 4 | lifecycle | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
| shared | 64 | publish | p50_ns | 76430.000 | 75030.500 | 0.982037 | 0.980212 | 0.987035 |
| shared | 64 | publish | p95_ns | 83720.000 | 82200.500 | 0.982459 | 0.977118 | 0.992660 |
| shared | 64 | publish | p99_ns | 86725.000 | 85285.000 | 0.976626 | 0.971064 | 0.997787 |
| shared | 64 | publish | mean_ns | 77224.461 | 75779.558 | 0.981807 | 0.979540 | 0.988143 |
| shared | 64 | publish | rss_kib | 5790.000 | 5848.000 | 1.010061 | 0.975305 | 1.038361 |
| shared | 64 | publish | minor_faults | 687.500 | 688.500 | 1.001460 | 0.987091 | 1.013880 |
| shared | 64 | publish | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
| root | 4 | publish | p50_ns | 23735.000 | 23605.500 | 0.992594 | 0.988070 | 0.998737 |
| root | 4 | publish | p95_ns | 25180.000 | 25065.000 | 0.994819 | 0.983878 | 1.001811 |
| root | 4 | publish | p99_ns | 32515.000 | 32175.000 | 0.993342 | 0.955590 | 1.009949 |
| root | 4 | publish | mean_ns | 24160.493 | 23973.020 | 0.993363 | 0.988729 | 0.997150 |
| root | 4 | publish | rss_kib | 5558.000 | 5356.000 | 0.969240 | 0.958095 | 1.003277 |
| root | 4 | publish | minor_faults | 628.000 | 624.000 | 0.992834 | 0.990469 | 1.004785 |
| root | 4 | publish | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
| root | 4 | lifecycle | p50_ns | 36060.000 | 35880.500 | 0.996122 | 0.986443 | 1.000862 |
| root | 4 | lifecycle | p95_ns | 38980.000 | 38845.000 | 0.996788 | 0.961725 | 1.017580 |
| root | 4 | lifecycle | p99_ns | 46141.000 | 45825.500 | 0.991642 | 0.955446 | 1.009464 |
| root | 4 | lifecycle | mean_ns | 36593.378 | 36407.816 | 0.994696 | 0.986462 | 1.000704 |
| root | 4 | lifecycle | rss_kib | 5498.000 | 5362.000 | 0.976318 | 0.968430 | 1.023105 |
| root | 4 | lifecycle | minor_faults | 628.000 | 623.500 | 0.992817 | 0.988104 | 0.996820 |
| root | 4 | lifecycle | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
| deep-unique | 4 | publish | p50_ns | 26690.000 | 26800.000 | 1.001499 | 0.998135 | 1.015924 |
| deep-unique | 4 | publish | p95_ns | 28540.500 | 28565.000 | 0.997527 | 0.987186 | 1.018464 |
| deep-unique | 4 | publish | p99_ns | 35835.000 | 36185.000 | 1.015329 | 0.971853 | 1.057638 |
| deep-unique | 4 | publish | mean_ns | 27131.675 | 27200.732 | 1.001595 | 0.994386 | 1.014224 |
| deep-unique | 4 | publish | rss_kib | 5496.000 | 5552.000 | 1.005071 | 0.995245 | 1.021567 |
| deep-unique | 4 | publish | minor_faults | 628.500 | 626.000 | 0.995236 | 0.992056 | 1.002386 |
| deep-unique | 4 | publish | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
| deep-unique | 4 | lifecycle | p50_ns | 43530.500 | 43940.000 | 1.007559 | 1.003545 | 1.023540 |
| deep-unique | 4 | lifecycle | p95_ns | 47745.000 | 49305.000 | 1.023167 | 1.005742 | 1.035401 |
| deep-unique | 4 | lifecycle | p99_ns | 53740.000 | 54660.000 | 1.011310 | 1.007063 | 1.029067 |
| deep-unique | 4 | lifecycle | mean_ns | 44170.266 | 44652.094 | 1.008265 | 1.004650 | 1.025134 |
| deep-unique | 4 | lifecycle | rss_kib | 5482.000 | 5552.000 | 1.007283 | 0.978590 | 1.014622 |
| deep-unique | 4 | lifecycle | minor_faults | 628.500 | 625.500 | 0.992848 | 0.990476 | 1.005582 |
| deep-unique | 4 | lifecycle | major_faults | 0.000 | 0.000 | n/a | n/a | n/a |
