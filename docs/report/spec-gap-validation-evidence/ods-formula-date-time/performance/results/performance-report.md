# Date/time performance evidence

Validated 4,620 samples: 1,020 baseline and 3,600 candidate, with 15 samples and three warmups per group.

All ten allocation, work, read and output accounting sets match exactly across 68 control groups. Candidate-only date/time cases have no supported predecessor baseline.

Median threshold flags are retained below. The machine-readable report includes every group and all median/p95/p99 flags and uncertainty intervals.

| Group | Metric | Baseline | Candidate | Change | 95% bootstrap interval |
| --- | --- | ---: | ---: | ---: | --- |
| array-control-16x16-sin/parse-evaluate | rss_kib | 4348.0000 | 4584.0000 | +5.428% | +2.300% to +6.704% |
| array-control-4x4-sin/parse-evaluate | elapsed_ns_per_repeat | 11033.8125 | 11610.6875 | +5.228% | +0.000% to +10.909% |
| database-control-dstdev/evaluate | rss_kib | 3812.0000 | 4020.0000 | +5.456% | +3.455% to +6.322% |
| database-control-dstdev/parse-evaluate | rss_kib | 3804.0000 | 3996.0000 | +5.047% | +3.132% to +6.448% |
| database-control-dsum/evaluate | rss_kib | 3800.0000 | 4016.0000 | +5.684% | +2.914% to +7.090% |
| database-control-dvar/evaluate | rss_kib | 3820.0000 | 4012.0000 | +5.026% | +1.763% to +6.184% |
| database-control-dvar/parse-evaluate | rss_kib | 3848.0000 | 4048.0000 | +5.198% | +0.727% to +6.618% |
| lazy-if-cache-isblank/evaluate | rss_kib | 3828.0000 | 4028.0000 | +5.225% | +3.826% to +7.553% |
| lazy-if-cache-isblank/parse-evaluate | rss_kib | 3796.0000 | 3988.0000 | +5.058% | +3.413% to +7.384% |
| lazy-if-cache-isnumber/evaluate | rss_kib | 3800.0000 | 4028.0000 | +6.000% | +3.255% to +7.075% |
| lazy-if-cache-isnumber/parse-evaluate | rss_kib | 3800.0000 | 4040.0000 | +6.316% | +2.947% to +6.842% |
| lazy-if-cache-istext/evaluate | rss_kib | 3812.0000 | 4024.0000 | +5.561% | +1.675% to +7.053% |
| lazy-if-cache-istext/parse-evaluate | rss_kib | 3788.0000 | 4000.0000 | +5.597% | +4.536% to +7.097% |
| literal-aggregate-4x1-sum/evaluate | rss_kib | 3772.0000 | 4000.0000 | +6.045% | +4.215% to +7.991% |
| reference-control-average/evaluate | rss_kib | 3788.0000 | 3996.0000 | +5.491% | +2.746% to +7.286% |
| representative-percentrank/evaluate | rss_kib | 3584.0000 | 3772.0000 | +5.246% | +0.321% to +6.488% |
| representative-rank/evaluate | rss_kib | 3596.0000 | 3804.0000 | +5.784% | +1.824% to +8.072% |
| representative-rank/parse-evaluate | rss_kib | 3596.0000 | 3784.0000 | +5.228% | +0.000% to +6.637% |
| scalar-control-arithmetic/parse-evaluate | rss_kib | 3556.0000 | 3744.0000 | +5.287% | +0.000% to +7.119% |
| scalar-control-var/parse-evaluate | rss_kib | 3556.0000 | 3768.0000 | +5.962% | -0.320% to +6.757% |
| value-date-fraction-value-date/evaluate | rss_kib | 3576.0000 | 3772.0000 | +5.481% | -0.966% to +7.545% |
| value-date-fraction-value-mixed-fraction/parse-evaluate | rss_kib | 3596.0000 | 3784.0000 | +5.228% | +2.670% to +6.682% |
| value-date-fraction-value-time/parse-evaluate | rss_kib | 3580.0000 | 3764.0000 | +5.140% | -0.533% to +6.935% |

Accepted with disclosed threshold flags for this function-support batch; no overall speedup or causal regression claim.
Latency includes cached expected-value comparisons, complete checksumming and result drop; independent oracle generation is outside timing.
Baseline precedes candidate; CPU affinity is not pinned and system load is recorded without a rejection threshold.
Bootstrap intervals are descriptive independent resamples of fifteen child samples, not proof of a causal effect.
RSS is a process observation; equal allocation counters do not establish a cause for RSS differences.
Hardware-counter and concurrency scaling improvements are not claimed by this synchronous evaluator capture.

Units: elapsed time is nanoseconds per repeat; RSS is KiB. Full environment, raw child JSON, time receipts and source/harness/lock identities are retained beside this report.
