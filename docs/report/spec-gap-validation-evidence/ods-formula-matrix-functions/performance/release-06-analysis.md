# Matrix candidate 06 release measurements

All 99 processes succeeded: 33 cases, three rounds, CPU 6, two warmups and
31 measured iterations per process, using each case's explicit repeat count.
Full numeric/shape/error validation is untimed. Timings include evaluation,
checksum consumption, retained-memory observation, and result drop. RSS is
whole-process maximum RSS, including fixture/parser setup.

These are absolute measurements for the final candidate and an observational
comparison with the earlier candidate 05 window. The windows ran sequentially,
not interleaved AB/BA pairs, so their deltas do not isolate a causal regression.
The table reports all cases and flags positive changes over the GOAL 5% review
threshold in the receipt. Deltas use the median of three per-process summaries;
latency is divided by repeat, and RSS is not. No geometric mean hides a case.

| Case | p50 µs/evaluation | p50 delta | p95 delta | Work/evaluation | Allocations/evaluation | Retained bytes | RSS KiB | RSS delta |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| mdeterm-2 | 1.94–2.01 | +1.5% | +1.2% | 39 | 12 | 0 | 2900–2940 | +1.4% |
| minverse-2 | 2.35–2.36 | -0.6% | +0.2% | 60 | 13 | 352 | 2784–2896 | +1.2% |
| mmult-2 | 2.85–2.93 | -1.2% | -0.9% | 57 | 15 | 352 | 2656–2904 | -5.4% |
| munit-2 | 1.09–1.13 | -0.5% | -1.9% | 16 | 8 | 352 | 2820–2856 | +7.6% |
| transpose-2 | 1.87–1.89 | -1.2% | -1.0% | 32 | 12 | 352 | 2632–2876 | -1.9% |
| mdeterm-8 | 10.54–10.62 | -0.4% | -1.6% | 538 | 20 | 0 | 2632–2876 | -7.1% |
| minverse-8 | 20.82–21.11 | +0.3% | +0.7% | 1515 | 21 | 5632 | 2716–2888 | +8.4% |
| mmult-8 | 20.14–20.75 | -1.0% | -1.3% | 1096 | 23 | 5632 | 2928–3132 | +6.6% |
| munit-8 | 2.23–2.25 | -1.3% | -1.5% | 76 | 8 | 5632 | 2616–2908 | -0.7% |
| transpose-8 | 8.47–8.56 | +2.9% | +2.8% | 327 | 20 | 5632 | 2636–2868 | -2.1% |
| mdeterm-16 | 41.91–43.37 | -1.6% | -0.3% | 2790 | 24 | 0 | 3096–3164 | +6.9% |
| minverse-16 | 114.76–115.93 | -0.0% | +0.4% | 10119 | 25 | 22528 | 2908–3148 | -3.4% |
| mmult-16 | 103.23–109.24 | +0.5% | +1.7% | 6565 | 27 | 22528 | 2912–3184 | +1.4% |
| munit-16 | 5.72–5.74 | +0.0% | -0.8% | 269 | 8 | 22528 | 2632–2872 | -6.5% |
| transpose-16 | 27.07–27.48 | -0.8% | +0.2% | 1444 | 24 | 22528 | 2892–3172 | -4.1% |
| transpose-64 | 593.63–596.85 | +2.8% | +1.4% | 27581 | 32 | 360448 | 4896–5020 | +0.1% |
| munit-64 | 74.37–74.68 | -0.2% | -0.6% | 4109 | 8 | 360448 | 3172–3320 | -1.4% |
| mmult-rect-4x8x2 | 7.62–7.70 | -0.5% | -0.3% | 280 | 20 | 704 | 2844–2868 | +7.2% |
| transpose-rect-2x8 | 3.51–3.69 | -0.6% | -0.6% | 87 | 16 | 1408 | 2848–2912 | +8.3% |
| minverse-singular-2 | 2.10–2.12 | +2.1% | +2.0% | 55 | 12 | 0 | 2636–2784 | -5.4% |
| mdeterm-nonsquare-2x1 | 1.39–1.43 | -0.8% | -0.7% | 22 | 10 | 0 | 2672–2900 | +5.1% |
| mmult-incompatible-2x3-2x2 | 2.87–2.97 | -1.7% | +1.6% | 57 | 15 | 0 | 2656–2912 | +0.2% |
| munit-zero | 0.90–0.90 | +0.5% | +0.8% | 12 | 7 | 0 | 2784–2904 | -2.1% |
| if-selected-mdeterm-8 | 10.49–10.73 | -2.9% | -2.3% | 549 | 20 | 0 | 2664–2928 | +6.2% |
| if-unselected-mdeterm-8 | 0.60–0.61 | +0.8% | +0.8% | 14 | 4 | 0 | 2844–2900 | +1.1% |
| if-selected-minverse-8 | 21.00–21.06 | -0.3% | -2.6% | 1526 | 21 | 5632 | 2656–2928 | +5.5% |
| if-unselected-minverse-8 | 0.59–0.61 | +2.1% | +1.6% | 14 | 4 | 0 | 2876–2908 | +2.3% |
| if-selected-mmult-8 | 20.80–21.52 | +0.4% | +0.2% | 1107 | 23 | 5632 | 2912–3116 | -2.0% |
| if-unselected-mmult-8 | 0.60–0.61 | +0.0% | -0.4% | 14 | 4 | 0 | 2604–2660 | -2.3% |
| if-selected-munit-8 | 2.45–2.48 | -0.4% | -0.4% | 87 | 8 | 5632 | 2648–2672 | +0.6% |
| if-unselected-munit-8 | 0.60–0.64 | +1.2% | +1.2% | 14 | 4 | 0 | 2784–2900 | +0.8% |
| if-selected-transpose-8 | 8.43–8.74 | -4.4% | -4.0% | 338 | 20 | 5632 | 2632–2864 | +7.5% |
| if-unselected-transpose-8 | 0.60–0.61 | +2.5% | +1.6% | 14 | 4 | 0 | 2636–2876 | -4.5% |

There are 10 cases with at least one positive latency/RSS review flag,
0 within-capture deterministic mismatches, and 0 cases
with deterministic counter/result changes relative to candidate 05. Exact
fields and deltas are in the [receipt](release-06/receipt.json).
The [raw capture](release-06/capture.tar.gz) includes every command, status,
JSON row, stderr/time output, build log, frozen harness/lockfile, and runner.
Its source closure and all five gates are linked from the receipt. The exact
ELF is retained outside tmpfs. Loose capture files were removed after archive
verification.

Performance acceptance remains open: representative old scalar/reference
controls need replay against their retained baseline binaries. Reference-based
matrix branches remain uncached, and these short shared-host measurements do
not establish end-to-end recalculation throughput or the overall GOAL speedup.
The candidate is a measured functional addition, not a blanket no-regression
claim. The prior unsupported-function implementation is not a semantic matrix
performance baseline.
