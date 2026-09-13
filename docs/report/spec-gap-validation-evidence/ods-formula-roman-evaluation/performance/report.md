# ROMAN/ARABIC: final measured function-family extension

The candidate implements both remaining OpenFormula §6.19 functions and removes a per-byte cancellation-threshold comparison from string scanning. Checks remain bounded by 4096-byte windows, including a doubled quote crossing a window by one byte. This is a necessary measured feature extension, with explicit remaining performance qualifications; it is not an across-the-board speedup or no-regression claim.

Baseline: `d0c1ca700ceda177c27a820acffe0de534720390`. Candidate source is bound by [the exact patch](../candidate.patch), [gate hashes](../gates/results.json), and both build receipts. The scoped isolated gates pass 926 tests in 56 targets, two doctests, Clippy, rustdoc, formatting and crate boundaries. A separate main-worktree gate run also passed; unrelated root edits were preserved.

## Method and reproducibility

The final corpus contains 53 existing-function controls and 41 Roman/Arabic cases, each in parse, evaluate, and combined phases. All 441 initial rows pass exact result/refusal checks. Every phase uses three warmup batches and 15 timed batches. Repetition counts are fixed per case in [run.py](roman-harness/run.py); scaled inputs span 64, 256, 1024 and 4096 bytes/calls.

Final children are pinned to CPU 6 on the shared Linux host described in [environment.json](environment.json). No hard scheduling isolation is claimed. An unrelated benchmark occupied CPU 2 during exploration; the final complete comparison uses CPU 6. Baseline and candidate use the same four harness files, locked offline release builds and identical Cargo inputs. The runner affinity is the only change from the independently reviewed original harness, as verified by [final-harness-affinity.json](final-harness-affinity.json).

Timing excludes process launch, setup, preflight and evaluation-context construction. Evaluate mode reuses an immutable parsed expression; combined mode parses and evaluates per operation. Result destruction and allocator observation follow the harness contract. RSS is whole-process maximum RSS, including startup and preflight. Allocator counters include only the measured tracking interval; peak heap and RSS are not divided by repeat count. Hardware counters below cover the whole subprocess, including setup and warmups.

Displayed ns/op divides a recorded batch median by its repeat count. The four-repeat comparison is a ratio of per-revision medians across interleaved A/B pairs. With only 15 batches, p95/p99 both select the maximum batch; these are diagnostic tail flags, not estimates of production service percentiles. There is no confidence-interval, broad throughput, package-memory, cold-I/O or multicore-scaling claim.

To reproduce, materialize the base commit in the scratch path configured by [capture.py](capture.py), copy the retained [workspace lock](../gates/workspace-Cargo.lock), gate collector and four harness files, then run `python3 -B performance/capture.py baseline` from this evidence directory. Apply `candidate.patch` from the checkout root, run `gates/run.py` in the isolated candidate and retain its receipt here, then run `performance/capture.py candidate`. The exact release build commands, environment overrides and per-case commands are retained beside each capture. Run `performance/select_flags.py` to generate the >5% flag set from the four initial latency/RSS metrics, then run `performance/repeat_flags.py`, `performance/counters.py` and `performance/investigate_counters.py` serially. Recreate these disposable roots only for a new measurement run; final cleanup intentionally removes them.

## Comparable results and regression disposition

Every comparable result, checksum, refusal kind, allocation count, requested-byte count, peak tracked heap and output reservation matches the baseline. All individual results remain in [baseline CSV](baseline/comparable/raw.csv), [candidate CSV](candidate/comparable/raw.csv), and [four-pair CSV](abab/raw.csv). This exact allocation parity is scoped to these controls and does not establish package-wide RSS behavior.

All 34 initial scenarios exceeding 5% in p50/p95/p99 or RSS received four A/B pairs (272 rows). 2 combined-phase p50 regressions remain. They are accepted as explicit qualifications of this necessary function-family extension under the review-trigger rule in `docs/GOAL.md`; follow-up should target the coercion/error paths and code-layout sensitivity using representative formula workloads. No remaining regression is dismissed as noise. No repeated RSS median exceeds 5%.

| Phase/case | Baseline ns/op | Candidate ns/op | Repeated p50 change |
| --- | ---: | ---: | ---: |
| parse-evaluate / bitwise-coerce-text | 788.91 | 848.83 | +7.60% |
| parse-evaluate / radix-fraction-error | 621.56 | 653.36 | +5.12% |

All 8 remaining tail-flag lanes are disclosed below. Full repeated results, including improvements and flags that cleared, are in [summary.json](abab/summary.json).

| Phase/case | p95 change | p99 change |
| --- | ---: | ---: |
| parse / logical-and-64 | +9.25% | +9.25% |
| parse / radix-base-max | +10.09% | +10.09% |
| parse / radix-direct-negative | +21.89% | +21.89% |
| evaluate / failure-name | +18.58% | +18.58% |
| evaluate / lazy-if-false-reference | +26.29% | +26.29% |
| parse-evaluate / logical-false | +70.10% | +70.10% |
| parse-evaluate / bitwise-coerce-text | +31.05% | +31.05% |
| parse-evaluate / radix-fraction-error | +5.03% | +5.03% |

## New-function costs and scaling

These candidate-only timings have no supported baseline implementation. All five formats, error/refusal cases, text scanning and concatenated calls are retained in [Roman CSV](candidate/roman/raw.csv). The following evaluate-only rows show individual calls and every scaling point; repeat batches amortize timing overhead without dividing peak heap.

| Case | ns/op | Peak tracked heap bytes | Output reservation bytes |
| --- | ---: | ---: | ---: |
| roman-3888-format-0 | 3212.59 | 463 | 15 |
| roman-3888-format-1 | 4297.60 | 463 | 15 |
| roman-3888-format-2 | 3606.81 | 463 | 15 |
| roman-3888-format-3 | 4211.59 | 463 | 15 |
| roman-3888-format-4 | 4636.12 | 456 | 8 |
| arabic-indirect | 4763.06 | 456 | 0 |
| arabic-empty | 417.81 | 400 | 0 |
| arabic-input-64 | 510.78 | 400 | 0 |
| arabic-input-256 | 796.56 | 400 | 0 |
| arabic-input-1024 | 1780.00 | 400 | 0 |
| arabic-input-4096 | 5725.00 | 400 | 0 |
| roman-concat-64 | 33691.56 | 3489 | 64 |
| roman-concat-256 | 126873.72 | 12897 | 256 |
| roman-concat-1024 | 502562.38 | 50529 | 1024 |
| roman-concat-4096 | 2019849.50 | 201057 | 4096 |

ARABIC scans input linearly with a fixed i128 accumulator. ROMAN format 4 examines 64 fixed candidates and uses stack arrays rather than a 4000-entry heap table. Concatenation scaling includes the existing repeated-copy behavior of the scalar concatenation operator; it is not a claim of a linear rope or streaming implementation. The functions perform no workbook access, decompression, recompression or I/O and introduce no threads or caches.

## Hardware counters and scanner investigation

Five primary captures cover an existing text control before/after, classic ROMAN, shortest ROMAN, and long ARABIC input. Four additional captures investigate the prior UTF-8/coercion regressions with larger repeat batches. Counter totals are process totals and are not interchangeable with ns/op or production tail latency.

| Capture | Evaluate ns/op | Cycles | Instructions | Branches | Branch misses |
| --- | ---: | ---: | ---: | ---: | ---: |
| perf-stat/baseline-text-if-concat | 765.44 | 6192967178 | 11625066237 | 1964293131 | 297186 |
| perf-stat/candidate-arabic-input-4096 | 5746.72 | 468551930 | 2943629983 | 527662233 | 95115 |
| perf-stat/candidate-roman-3888-format-0 | 3205.18 | 2596093783 | 7575885068 | 1347173068 | 2412794 |
| perf-stat/candidate-roman-3888-format-4 | 4624.13 | 3740169423 | 11407754023 | 1743299887 | 781391 |
| perf-stat/candidate-text-if-concat | 774.49 | 6265426054 | 11599912894 | 1960701663 | 124457 |
| investigation-counters/baseline-bitwise-coerce-text | 584.35 | 4729136442 | 9480468340 | 1607463696 | 89935 |
| investigation-counters/baseline-control-utf8-left-4096 | 10867.78 | 117572380 | 273153268 | 81385130 | 39849 |
| investigation-counters/candidate-bitwise-coerce-text | 593.77 | 4812565196 | 9469705123 | 1611069899 | 233219 |
| investigation-counters/candidate-control-utf8-left-4096 | 10090.44 | 108637094 | 235345348 | 62506106 | 34701 |

The larger-batch UTF-8 lane measures about 10.87→10.09 µs/op and bitwise text coercion about 584→594 ns/op. These evaluate-only measurements do not erase the two combined-phase A/B regressions. The finite scan-window change removes real branch work; blanket dispatch/scanner outlining did not robustly improve the controls and was discarded. No cold-path annotation or padding was adopted.

Original captures, counter investigation, the CPU 6 confirmation probe and rejected experiments remain under [exploration](exploration/). Source and harness bytes for the original candidate were reconstructed and checked against the original receipts. Exploratory results are not substituted for the final source-bound CPU 6 captures. Shared-host interference and code placement limit causal attribution of small deltas.

## Verification and cleanup

The root verifier checks all raw stdout/status/time sidecars, command affinity, exact output/heap parity, all flag selection and repeated medians, source and harness custody, five primary and four investigative counter captures, and binary hashes. Both final ELFs and all three scratch roots were removed after review and measurement checks. [Cleanup](../gates/cleanup.json) records 1,555,263,488 allocated bytes reclaimed and hash-verified recovery of 22 unique files outside tmpfs. No broader formula or recalculation gap is declared closed.
