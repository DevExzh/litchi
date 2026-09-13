# Existing value workloads after matrix support

The retained pre-matrix const-specialized evaluator (`scalar-cell-specialized-03`) and the current candidate 06 evaluator ran the unchanged value harness in 81 AB/BA/AB pairs, 162 child processes, CPU 6, three warmups and 31 iterations. All processes succeeded with zero deterministic mismatches. Build source hashes match all 458 candidate-06 source inputs. The build archive retains compiler identity, command, environment, source hashes and exact harness files; the capture retains host metadata, commands, raw outputs and the frozen paired runner.

All captured allocation, requested/released bytes, live allocation and retained-budget counters agree across revisions. Ten cases trigger a positive latency or RSS change above 5%. This is a review trigger, not performance acceptance. In particular, text arithmetic has +6.57% median p50; repeated single-cell references and background-4096 have p95 flags. Three rounds cannot identify causality or establish statistical significance. Follow-up should reproduce these specific lanes and profile persistent latency changes before adding optimization complexity.

Values are percentage changes of the median of three process results, without dividing batch latencies by repeat. Positive means higher. RSS is whole-process peak.

| Case | Repeat | p50 Δ% | p95 Δ% | RSS Δ% |
|---|---:|---:|---:|---:|
| matrix-lazy-aggregate-4096 | 1 | -0.52 | -0.53 | +2.82 |
| matrix-lazy-inline-4096 | 1 | -2.60 | -2.58 | +2.46 |
| reference-background-4096 | 128 | +0.90 | +6.95 | +0.00 |
| reference-cancelled | 128 | -3.03 | -2.26 | -6.67 |
| reference-cell | 128 | -0.95 | +0.27 | -1.35 |
| reference-distinct-1 | 128 | +0.52 | -0.88 | -6.09 |
| reference-distinct-16 | 32 | +1.84 | +2.10 | +6.38 |
| reference-distinct-256 | 8 | +1.86 | +1.58 | +7.88 |
| reference-distinct-4096 | 1 | -1.65 | -1.63 | +3.35 |
| reference-empty | 128 | +0.80 | -3.39 | -0.54 |
| reference-empty-arithmetic | 128 | +3.27 | +2.15 | +5.67 |
| reference-error | 128 | +0.90 | +0.55 | +0.40 |
| reference-lazy | 128 | -0.00 | -4.26 | +9.93 |
| reference-limit-cells | 128 | +1.02 | +1.04 | -0.41 |
| reference-limit-memory | 128 | +0.15 | -0.10 | -1.50 |
| reference-limit-work | 128 | +0.49 | +0.91 | +1.80 |
| reference-matrix | 128 | -0.07 | +0.61 | -4.25 |
| reference-range-1 | 128 | -0.31 | -0.92 | +6.39 |
| reference-range-16 | 32 | -0.29 | +0.44 | +4.89 |
| reference-range-256 | 8 | -0.52 | +0.27 | +11.39 |
| reference-range-4096 | 1 | +0.53 | -0.15 | +5.60 |
| reference-repeat-1 | 128 | +1.17 | +6.13 | +1.90 |
| reference-repeat-16 | 32 | -2.88 | -1.42 | +2.36 |
| reference-repeat-256 | 8 | -0.74 | +0.52 | +3.55 |
| reference-repeat-4096 | 1 | -1.25 | -1.50 | +2.25 |
| reference-text | 128 | +2.52 | +2.58 | -2.12 |
| reference-text-arithmetic | 128 | +6.57 | +4.28 | +5.85 |

Reproduce with `python3 -B analyze_common.py common-01/capture.tar.gz`. Raw child files and pair validation remain authoritative; the analyzer summarizes the verified runner receipt. No runtime source changed in this batch. Broader OpenFormula implementation and end-to-end performance requirements remain open.

Independent read-only review by `ods_profiler` confirmed all 162 results, exact pairwise deterministic and memory counters, and the latency/RSS flags above. All loose build and capture files were deleted after byte-for-byte archive verification.
