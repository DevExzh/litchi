# Independent date/time performance review

This review covers the completed before/after capture and is bound to:

- [`profile-analysis.json`](ods-formula-date-time/performance/results/profile-analysis.json), SHA-256 `6c73af987164a71f50757b40f302fb699dbd972b7a71f69a1a2b5f104d77fe88`;
- [`freeze.json`](ods-formula-date-time/gates/freeze.json), SHA-256 `461a76708e36b2a716cd622df45f014e009186223bd1db511c3e0ff0f7fa3561`.

The capture contains 1,020 baseline rows, 3,600 candidate rows, 68 matched groups, 240 candidate groups, three warmups, and fifteen samples per group. Preflight and correctness checks passed. The analysis reports 62 five-percent threshold flags across quantiles, including 46 positive flags. Bootstrap intervals use 10,000 independent resamples with seed `20260920`.

The only median latency trigger is the matched `array-control-4x4-sin/parse-evaluate` lane: `11,033.8125` to `11,610.6875` ns per repeat, or `+5.228%`. Its seeded 95% relative interval is `[0%, +10.909%]`. The same lane is `+7.324%` at p95, with interval `[+0.307%, +11.769%]`. Other latency positives are tail observations, including `reference-conditional-256x4-sumifs/parse-evaluate` at `+7.871%` p95 and `+13.762%` p99, and `reference-aggregate-64x4-sum/parse-evaluate` at `+9.970%` p95 and `+12.268%` p99. Several other tails move faster, so these results do not establish a broad latency slowdown.

RSS has 22 median triggers across 27 groups, ranging from `+184` to `+240 KiB` (`+5.026%` to `+6.316%`). The largest median increase is `lazy-if-cache-isnumber/parse-evaluate`, `3,800` to `4,040 KiB` (`+6.316%`, bootstrap interval `[+2.947%, +6.842%]`). RSS is an external high-water observation; allocation calls, requested/released bytes, peak live bytes, and retained memory remain unchanged.

All matched accounting sets are exact across the recorded quantiles: output bytes, allocation calls, requested/released bytes, peak live, retained memory, evaluator work, and resolver reads. The same toolchain, libc, kernel, and 32-CPU host were used. CPU affinity was not pinned. The recorded one-minute load before baseline was `[3.022, 2.595, 2.207]`; before candidate it was `[3.672, 2.830, 2.302]`. This shared-host context limits causal interpretation, and the load record is not treated as a reason to discard the triggers.

The disposition is descriptive acceptance of the capture with the latency and RSS triggers disclosed. It supports no causal regression or speedup claim and requires no production change or rerun. A focused follow-up is appropriate only if a causal performance claim is needed: repeat the 4×4 SIN parse-evaluate lane and RSS observation under controlled load.
