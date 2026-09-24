# Paired matrix release comparison

The exact retained candidate 05 and 06 binaries ran in 99 adjacent pairs on CPU 6: before/after, after/before, before/after over three rounds, 33 cases each. Each process used two warmups and 31 measured iterations. All 198 processes passed their untimed numeric/shape/error oracle. Every non-timing JSON result field agrees across all six executions of each case, including checksums, work, allocation counts, requested/released bytes, peak live allocation delta, and retained budget bytes.

Fourteen cases exceed the 5% RSS review threshold; none exceeds 5% for the median of process p50 or p95 latencies. Five of the ten earlier sequential RSS flags recur: `if-selected-mdeterm-8`, `if-selected-transpose-8`, `mdeterm-16`, `mmult-8`, `transpose-rect-2x8`. These measurements do not establish a causal explanation or close RSS acceptance. Three pairs per case are limited evidence, and external time RSS includes process setup. GNU size totals differ by only four bytes (1,210,372 before; 1,210,376 after), which does not explain resident pages.

Values below are percentage changes of medians across three processes; positive means higher. RSS ranges are KiB.

| Case | p50 Δ% | p95 Δ% | RSS Δ% | Before RSS | After RSS |
|---|---:|---:|---:|---|---|
| mdeterm-2 | +0.84 | +0.67 | +5.36 | 2628–2704 | 2652–2844 |
| minverse-2 | +0.03 | +0.03 | +2.72 | 2696–2808 | 2684–2928 |
| mmult-2 | +1.65 | +2.69 | -1.48 | 2648–2908 | 2636–2912 |
| munit-2 | -1.62 | -1.27 | +1.02 | 2748–2808 | 2604–2836 |
| transpose-2 | -0.06 | -0.10 | +0.57 | 2788–2808 | 2628–2848 |
| mdeterm-8 | -0.17 | +0.78 | +6.55 | 2652–2708 | 2656–2912 |
| minverse-8 | +0.47 | -2.76 | -1.48 | 2660–2796 | 2656–2876 |
| mmult-8 | +0.57 | -2.10 | +8.38 | 2856–2948 | 2964–3164 |
| munit-8 | -0.44 | -1.31 | -1.92 | 2656–2900 | 2636–2848 |
| transpose-8 | +1.54 | +1.41 | +6.28 | 2672–2704 | 2804–2908 |
| mdeterm-16 | +1.86 | +1.53 | +5.71 | 2880–3000 | 2912–3144 |
| minverse-16 | -0.06 | +0.45 | +3.05 | 2944–3056 | 3076–3164 |
| mmult-16 | -1.45 | -1.29 | +4.86 | 2940–3096 | 3052–3164 |
| munit-16 | +0.44 | +1.48 | +6.43 | 2600–2856 | 2620–2888 |
| transpose-16 | -0.94 | +2.19 | +9.35 | 2856–3040 | 3164–3184 |
| transpose-64 | +3.29 | +3.02 | -0.88 | 4996–5088 | 4936–5004 |
| munit-64 | +0.04 | +3.23 | +2.86 | 3204–3280 | 3292–3376 |
| mmult-rect-4x8x2 | -1.07 | -0.42 | -0.30 | 2692–2896 | 2616–2848 |
| transpose-rect-2x8 | +1.00 | +1.56 | +7.76 | 2628–2696 | 2848–2908 |
| minverse-singular-2 | -0.09 | -0.29 | -1.63 | 2696–2796 | 2648–2656 |
| mdeterm-nonsquare-2x1 | -1.94 | -1.35 | -1.63 | 2684–2708 | 2652–2688 |
| mmult-incompatible-2x3-2x2 | +1.36 | +0.65 | +6.96 | 2648–2796 | 2864–2912 |
| munit-zero | +1.04 | -0.41 | +2.89 | 2700–2808 | 2656–2912 |
| if-selected-mdeterm-8 | -1.72 | -1.98 | +5.33 | 2652–2752 | 2656–2852 |
| if-unselected-mdeterm-8 | +1.70 | +2.92 | +5.48 | 2652–2800 | 2604–2888 |
| if-selected-minverse-8 | -0.86 | -3.77 | -1.34 | 2656–2704 | 2628–2908 |
| if-unselected-minverse-8 | +1.27 | +0.00 | +3.16 | 2732–2796 | 2652–2900 |
| if-selected-mmult-8 | -0.59 | +0.63 | +6.22 | 2952–3144 | 2912–3164 |
| if-unselected-mmult-8 | +1.69 | +0.41 | +3.74 | 2776–2864 | 2636–2928 |
| if-selected-munit-8 | +0.10 | +0.10 | -3.21 | 2612–2784 | 2620–2784 |
| if-unselected-munit-8 | +2.45 | +1.19 | +7.73 | 2672–2808 | 2708–2928 |
| if-selected-transpose-8 | -0.74 | -1.01 | +7.87 | 2632–2808 | 2664–2908 |
| if-unselected-transpose-8 | +2.14 | +0.82 | +3.34 | 2744–2784 | 2848–2908 |

Reproduce the analysis by extracting `pairs-01/capture.tar.gz` to a new disk-backed directory and running `python3 -B analyze_pairs.py DIRECTORY`. The analyzer checks raw artifact hashes, pair ordering, result parity and exit statuses. The archive retains the exact capture runner and raw time/stdout/stderr files. Source and gate provenance remains in candidate-05 and candidate-06.

Existing scalar/reference workload regression checks and broader end-to-end performance requirements remain open. No production source changed for this capture. Loose capture files were removed only after every archived member was compared byte-for-byte.

Independent source/evidence review by `ods_profiler` reconstructed all 198 archived rows and confirmed deterministic parity. The reviewer agreed that whole-process RSS and only three paired rounds do not support causal attribution or performance acceptance.
