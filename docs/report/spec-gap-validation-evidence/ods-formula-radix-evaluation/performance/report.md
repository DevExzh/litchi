# Radix evaluator performance

The final fourteen-function implementation passes the captured semantic preflights. Existing scalar, logical and bitwise controls retain identical allocation counts, requested bytes, peak tracked heap, result reservations and checksums. This is a measured feature addition; no overall speedup is claimed.

The initial screen flagged 34 of 141 comparable lanes. Four interleaved A/B pairs produced 272 repeat rows. No repeated median-latency delta exceeds 5%. Seven lanes retain tail-latency flags (5.09–55.10%); one evaluation lane retains a 5.48% process peak-RSS increase. These remain review limitations, not a blanket no-regression claim.

## Setup and provenance

The release harness runs serial processes pinned to CPU 2 on the shared AMD EPYC 9R45 x86_64 host. Each lane has three warmup batches and fifteen measured batches; each batch contains the recorded repeat count. No Cargo build overlaps the retained measurement runs. The host is not isolated from unrelated workloads. Environment and toolchain versions are in [environment.json](environment.json); exact build commands and before/after source hashes are retained under each revision.

| Artifact | SHA-256 |
| --- | --- |
| baseline release executable | `8437de6ad6b53d4883960f2b20fdb7e13b6af80822fd740934f42525ce015733` |
| candidate release executable | `5012e67ac4edbd037d07835c51eef6c0860c2953cad52abede7098d68e6e5e7a` |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `15ff312b132e6d3d0d97a2cd6ff897a8b33fcbd8b366b35e20737c6d7cfeaa84` |
| `crates/litchi-ods/src/codec/formula/evaluation/radix.rs` | `48cf638a72a4946f8f3c8284b34ff817e4c5ba77c962776ad6c2ae981991cf9d` |
| `crates/litchi-ods/tests/ods_formula_radix_evaluation.rs` | `c2d0ee542fdc4e465b6236b1e6234d1b7980791a82995d7c315239af48ba32bf` |
| harness `Cargo.toml` | `f337daad9e20c4b63c10fc84fc998f2ee16762170d513c59e1e7c4b471c9cdff` |
| harness `Cargo.lock` | `0bd5dc7cfe53d0a6625406040d86b10f3f257f9cc061747c4a77077e5f1c8df0` |
| harness `run.py` | `a96ad565525121c9f00135326c21798ffde574c39e4caf8dd34453ae32876c6f` |
| harness `src/main.rs` | `11c3d7df4156199ae15e4e9931551d2dcc4644601ce3dbfea0a273a13e576c61` |

Both revisions use exactly the same four harness files. The baseline is commit `e1976c59d5c4785e0c73a5d27e9349ba082cca2f`. The workspace lock is retained as an explicit pre-existing input. The final candidate includes the reviewed DECIMAL reservation-order fix: the temporary TextValue is dropped as a whole before rounding or pushing its numeric result. Earlier draft expectations and pre-fix captures were superseded before this final comparison.

## Corpus and measurement interpretation

The comparable corpus has 47 cases across parse, evaluate and parse-evaluate phases (141 rows per revision). The radix corpus has 53 cases across the same phases (159 rows): all fourteen functions, accepted lexical spellings, fractional and range errors, resource and cancellation refusals, large finite values, call-count scales and padding scales. Radix evaluation is candidate-only because the baseline does not implement this family; it is not used to claim before/after speedup.

The BASE/DECIMAL input-N lanes vary the number of calls, and DEC2HEX-padding-N repeats a seven-character result; these are not single-literal byte-length scales. BASE-padding-N varies the actual output length from 64 to 4096. Finite-maximum performance vectors use hexadecimal; independent Rust tests cover bases 2, 10 and 36 as well.

Latency below is the recorded batch p50 divided by repeat (ns/op). Requested heap bytes and allocation calls are measured by a counting allocator. Peak heap and RSS are batch/process peaks and are never divided by repeat. RSS includes executable and runtime pages, beyond tracked live heap. Every retained row returns live heap to its pre-sample level. Formula error values count as successful evaluations; typed refusals are reported separately.

With only fifteen measured batches, p95 and p99 both select the sample maximum. Their large deltas are finite-sample tail flags rather than reliable population quantiles. Four paired repeats improve the comparison, but do not prove the remaining flags are noise. No geometric mean hides individual results.

## Repeated review flags

[Initial screening deltas](initial-flags.json), [all repeat rows](abab/raw.csv), and [paired summaries](abab/summary.json) retain every selected lane.

| Phase | Case | Median delta % | Tail delta % | RSS delta % |
| --- | --- | ---: | ---: | ---: |
| parse | control-coerce-256 | 0.47 | 5.09 | 1.29 |
| parse | control-false | -2.32 | 46.47 | -0.90 |
| parse | failure-name | -0.04 | 47.29 | 1.69 |
| parse | failure-memory | 0.91 | 55.10 | -0.53 |
| parse | logical-and-64 | -0.62 | 5.88 | 0.99 |
| evaluate | control-utf8-left-4096 | -0.21 | 7.86 | -0.60 |
| evaluate | failure-reference | -0.16 | 32.41 | -0.15 |
| evaluate | bitwise-or-256 | 1.28 | 1.05 | 5.48 |

## Existing workloads: initial comparable capture

The table includes every initial comparable lane; flagged lanes must be interpreted with the repeats above. Exact allocations, requested bytes, output reservations, checksums, p95/p99 and commands are in [baseline raw data](baseline/comparable/raw.csv) and [candidate raw data](candidate/comparable/raw.csv).

| Phase | Case | Baseline ns/op | Candidate ns/op | Delta % | Baseline RSS KiB | Candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| parse | control-flat-64 | 1,451.11 | 1,463.44 | 0.85 | 2640 | 2664 |
| parse | control-flat-256 | 4,549.72 | 4,601.28 | 1.13 | 2512 | 2668 |
| parse | control-flat-1024 | 16,422.50 | 15,810.00 | -3.73 | 2908 | 2932 |
| parse | control-flat-4096 | 68,650.50 | 71,335.00 | 3.91 | 3736 | 3960 |
| parse | control-coerce-64 | 1,485.31 | 1,468.12 | -1.16 | 2708 | 2668 |
| parse | control-coerce-256 | 4,610.03 | 4,751.59 | 3.07 | 2692 | 2660 |
| parse | control-coerce-1024 | 16,123.75 | 15,835.00 | -1.79 | 2892 | 2888 |
| parse | control-coerce-4096 | 65,860.00 | 67,405.50 | 2.35 | 3808 | 3924 |
| parse | control-utf8-left-64 | 156.72 | 157.50 | 0.50 | 2640 | 2668 |
| parse | control-utf8-left-256 | 267.84 | 266.88 | -0.36 | 2616 | 2616 |
| parse | control-utf8-left-1024 | 697.50 | 701.25 | 0.54 | 2600 | 2552 |
| parse | control-utf8-left-4096 | 2,425.00 | 2,415.00 | -0.41 | 2644 | 2636 |
| parse | control-escaped-64 | 141.88 | 142.03 | 0.11 | 2600 | 2668 |
| parse | control-escaped-256 | 315.62 | 320.94 | 1.68 | 2716 | 2640 |
| parse | control-escaped-1024 | 1,026.25 | 1,025.00 | -0.12 | 2644 | 2668 |
| parse | control-escaped-4096 | 3,810.00 | 3,820.00 | 0.26 | 2644 | 2664 |
| parse | control-true | 106.25 | 101.17 | -4.78 | 2676 | 2668 |
| parse | control-false | 107.89 | 104.61 | -3.04 | 2452 | 2660 |
| parse | failure-reference | 196.48 | 200.39 | 1.99 | 2644 | 2628 |
| parse | failure-array | 205.94 | 203.75 | -1.06 | 2644 | 2456 |
| parse | failure-name | 102.58 | 100.94 | -1.60 | 2640 | 2660 |
| parse | failure-work | 128.44 | 126.88 | -1.22 | 2692 | 2656 |
| parse | failure-memory | 86.17 | 87.11 | 1.09 | 2512 | 2668 |
| parse | failure-stack | 178.59 | 175.62 | -1.66 | 2640 | 2656 |
| parse | failure-cancelled | 264.38 | 257.89 | -2.46 | 2644 | 2620 |
| parse | logical-true | 103.91 | 101.64 | -2.18 | 2516 | 2672 |
| parse | logical-false | 108.36 | 106.48 | -1.73 | 2644 | 2616 |
| parse | logical-and-64 | 2,590.47 | 2,581.11 | -0.36 | 2516 | 2668 |
| parse | logical-and-1024 | 39,172.75 | 38,776.50 | -1.01 | 2892 | 2888 |
| parse | logical-or-256 | 10,612.88 | 10,180.03 | -4.08 | 2624 | 2668 |
| parse | logical-xor-4096 | 173,606.00 | 174,771.00 | 0.67 | 3708 | 3652 |
| parse | logical-not-256 | 27,790.44 | 27,208.25 | -2.09 | 2728 | 2924 |
| parse | logical-if-1024 | 165,557.00 | 161,162.00 | -2.65 | 3680 | 3668 |
| parse | logical-iferror-4096 | 609,907.50 | 563,302.50 | -7.64 | 5072 | 4996 |
| parse | logical-ifna-4096 | 569,047.50 | 525,787.00 | -7.60 | 4920 | 4816 |
| parse | lazy-if-true-heavy-1024 | 540.00 | 531.38 | -1.60 | 2652 | 2668 |
| parse | lazy-if-false-reference | 345.70 | 342.89 | -0.81 | 2644 | 2668 |
| parse | text-if-escaped | 217.27 | 211.09 | -2.84 | 2644 | 2668 |
| parse | text-if-concat | 294.92 | 296.41 | 0.50 | 2516 | 2648 |
| parse | bitwise-and-64 | 6,609.09 | 6,354.72 | -3.85 | 2516 | 2604 |
| parse | bitwise-or-256 | 23,900.41 | 23,151.66 | -3.13 | 2752 | 2892 |
| parse | bitwise-xor-1024 | 99,646.75 | 95,072.88 | -4.59 | 3284 | 3328 |
| parse | bitwise-lshift-4096 | 425,292.00 | 414,507.00 | -2.54 | 4568 | 4492 |
| parse | bitwise-rshift-4096 | 425,277.00 | 413,647.00 | -2.73 | 4576 | 4544 |
| parse | bitwise-coerce-text | 190.55 | 189.38 | -0.62 | 2640 | 2652 |
| parse | bitwise-error-shift | 203.12 | 199.45 | -1.81 | 2628 | 2604 |
| parse | bitwise-lazy-selected | 440.39 | 425.47 | -3.39 | 2620 | 2656 |
| evaluate | control-flat-64 | 7,504.25 | 7,130.19 | -4.98 | 2580 | 2672 |
| evaluate | control-flat-256 | 26,020.12 | 26,755.75 | 2.83 | 2652 | 2668 |
| evaluate | control-flat-1024 | 104,754.12 | 101,181.75 | -3.41 | 2892 | 2888 |
| evaluate | control-flat-4096 | 411,637.00 | 426,987.00 | 3.73 | 3624 | 3816 |
| evaluate | control-coerce-64 | 7,650.19 | 7,387.53 | -3.43 | 2624 | 2672 |
| evaluate | control-coerce-256 | 27,231.38 | 26,180.75 | -3.86 | 2644 | 2620 |
| evaluate | control-coerce-1024 | 104,600.50 | 101,259.25 | -3.19 | 2836 | 2892 |
| evaluate | control-coerce-4096 | 402,581.50 | 412,077.00 | 2.36 | 3716 | 3948 |
| evaluate | control-utf8-left-64 | 983.30 | 980.16 | -0.32 | 2684 | 2668 |
| evaluate | control-utf8-left-256 | 1,457.50 | 1,459.06 | 0.11 | 2592 | 2624 |
| evaluate | control-utf8-left-1024 | 3,553.75 | 3,273.75 | -7.88 | 2644 | 2540 |
| evaluate | control-utf8-left-4096 | 10,640.00 | 10,645.00 | 0.05 | 2668 | 2540 |
| evaluate | control-escaped-64 | 2,002.05 | 1,990.80 | -0.56 | 2640 | 2644 |
| evaluate | control-escaped-256 | 6,652.84 | 6,681.88 | 0.44 | 2644 | 2588 |
| evaluate | control-escaped-1024 | 25,303.88 | 25,402.62 | 0.39 | 2580 | 2656 |
| evaluate | control-escaped-4096 | 99,850.50 | 100,310.50 | 0.46 | 2432 | 2640 |
| evaluate | control-true | 350.63 | 342.50 | -2.32 | 2516 | 2600 |
| evaluate | control-false | 345.16 | 345.47 | 0.09 | 2612 | 2652 |
| evaluate | failure-reference | 224.23 | 223.75 | -0.21 | 2516 | 2684 |
| evaluate | failure-array | 224.15 | 222.58 | -0.70 | 2644 | 2796 |
| evaluate | failure-name | 223.28 | 223.59 | 0.14 | 2580 | 2616 |
| evaluate | failure-work | 384.38 | 385.39 | 0.26 | 2616 | 2644 |
| evaluate | failure-memory | 280.16 | 280.55 | 0.14 | 2580 | 2680 |
| evaluate | failure-stack | 143.67 | 144.38 | 0.49 | 2644 | 2668 |
| evaluate | failure-cancelled | 6.41 | 7.50 | 17.07 | 2692 | 2628 |
| evaluate | logical-true | 347.66 | 357.04 | 2.70 | 2580 | 2640 |
| evaluate | logical-false | 346.02 | 345.47 | -0.16 | 2600 | 2664 |
| evaluate | logical-and-64 | 3,877.52 | 3,847.05 | -0.79 | 2452 | 2632 |
| evaluate | logical-and-1024 | 44,919.00 | 44,245.12 | -1.50 | 2856 | 2920 |
| evaluate | logical-or-256 | 11,627.53 | 11,486.62 | -1.21 | 2696 | 2620 |
| evaluate | logical-xor-4096 | 198,581.00 | 195,265.50 | -1.67 | 3628 | 3728 |
| evaluate | logical-not-256 | 27,895.75 | 27,088.88 | -2.89 | 2912 | 2616 |
| evaluate | logical-if-1024 | 126,228.00 | 124,923.00 | -1.03 | 3388 | 3392 |
| evaluate | logical-iferror-4096 | 448,847.00 | 446,752.00 | -0.47 | 4560 | 4584 |
| evaluate | logical-ifna-4096 | 455,437.00 | 435,132.00 | -4.46 | 4580 | 4544 |
| evaluate | lazy-if-true-heavy-1024 | 435.00 | 423.75 | -2.59 | 2644 | 2660 |
| evaluate | lazy-if-false-reference | 424.45 | 427.50 | 0.72 | 2644 | 2616 |
| evaluate | text-if-escaped | 582.34 | 568.67 | -2.35 | 2700 | 2644 |
| evaluate | text-if-concat | 785.62 | 784.30 | -0.17 | 2640 | 2536 |
| evaluate | bitwise-and-64 | 15,006.17 | 15,337.25 | 2.21 | 2708 | 2664 |
| evaluate | bitwise-or-256 | 56,626.84 | 57,658.09 | 1.82 | 2900 | 3052 |
| evaluate | bitwise-xor-1024 | 223,231.00 | 230,598.50 | 3.30 | 3404 | 3308 |
| evaluate | bitwise-lshift-4096 | 920,524.00 | 935,929.00 | 1.67 | 4944 | 4908 |
| evaluate | bitwise-rshift-4096 | 926,869.50 | 949,814.50 | 2.48 | 4756 | 4980 |
| evaluate | bitwise-coerce-text | 585.31 | 579.54 | -0.99 | 2624 | 2572 |
| evaluate | bitwise-error-shift | 580.23 | 571.25 | -1.55 | 2608 | 2628 |
| evaluate | bitwise-lazy-selected | 639.92 | 633.28 | -1.04 | 2580 | 2672 |
| parse-evaluate | control-flat-64 | 8,939.88 | 8,631.44 | -3.45 | 2716 | 2668 |
| parse-evaluate | control-flat-256 | 31,405.44 | 31,407.97 | 0.01 | 2676 | 2740 |
| parse-evaluate | control-flat-1024 | 118,421.75 | 118,356.75 | -0.05 | 2904 | 2900 |
| parse-evaluate | control-flat-4096 | 479,367.00 | 483,112.00 | 0.78 | 3724 | 3708 |
| parse-evaluate | control-coerce-64 | 9,316.61 | 8,969.27 | -3.73 | 2708 | 2604 |
| parse-evaluate | control-coerce-256 | 34,311.41 | 35,185.16 | 2.55 | 2728 | 2872 |
| parse-evaluate | control-coerce-1024 | 118,813.00 | 118,751.88 | -0.05 | 3092 | 3176 |
| parse-evaluate | control-coerce-4096 | 475,902.50 | 476,977.50 | 0.23 | 4244 | 4440 |
| parse-evaluate | control-utf8-left-64 | 1,171.72 | 1,154.53 | -1.47 | 2644 | 2636 |
| parse-evaluate | control-utf8-left-256 | 1,803.12 | 1,729.69 | -4.07 | 2616 | 2640 |
| parse-evaluate | control-utf8-left-1024 | 4,001.25 | 3,990.00 | -0.28 | 2644 | 2632 |
| parse-evaluate | control-utf8-left-4096 | 13,035.00 | 13,070.50 | 0.27 | 2640 | 2672 |
| parse-evaluate | control-escaped-64 | 2,150.94 | 2,134.55 | -0.76 | 2656 | 2668 |
| parse-evaluate | control-escaped-256 | 7,037.84 | 6,985.66 | -0.74 | 2676 | 2612 |
| parse-evaluate | control-escaped-1024 | 26,336.38 | 26,245.12 | -0.35 | 2576 | 2668 |
| parse-evaluate | control-escaped-4096 | 103,625.50 | 103,590.50 | -0.03 | 2644 | 2668 |
| parse-evaluate | control-true | 447.58 | 443.52 | -0.91 | 2640 | 2668 |
| parse-evaluate | control-false | 451.41 | 447.35 | -0.90 | 2676 | 2668 |
| parse-evaluate | failure-reference | 420.62 | 420.16 | -0.11 | 2624 | 2640 |
| parse-evaluate | failure-array | 427.89 | 423.99 | -0.91 | 2652 | 2600 |
| parse-evaluate | failure-name | 323.52 | 322.97 | -0.17 | 2640 | 2668 |
| parse-evaluate | failure-work | 507.35 | 508.36 | 0.20 | 2624 | 2624 |
| parse-evaluate | failure-memory | 365.31 | 373.59 | 2.27 | 2584 | 2668 |
| parse-evaluate | failure-stack | 327.97 | 321.25 | -2.05 | 2516 | 2636 |
| parse-evaluate | failure-cancelled | 270.08 | 267.59 | -0.92 | 2644 | 2616 |
| parse-evaluate | logical-true | 448.21 | 444.69 | -0.79 | 2644 | 2668 |
| parse-evaluate | logical-false | 450.87 | 454.45 | 0.80 | 2588 | 2668 |
| parse-evaluate | logical-and-64 | 6,736.75 | 6,570.66 | -2.47 | 2516 | 2540 |
| parse-evaluate | logical-and-1024 | 98,615.38 | 97,391.62 | -1.24 | 2708 | 2888 |
| parse-evaluate | logical-or-256 | 21,358.53 | 21,486.03 | 0.60 | 2652 | 2600 |
| parse-evaluate | logical-xor-4096 | 454,452.00 | 450,202.00 | -0.94 | 4072 | 4020 |
| parse-evaluate | logical-not-256 | 60,405.56 | 59,231.22 | -1.94 | 2884 | 2904 |
| parse-evaluate | logical-if-1024 | 294,100.00 | 281,741.25 | -4.20 | 3580 | 3568 |
| parse-evaluate | logical-iferror-4096 | 1,029,894.50 | 974,459.00 | -5.38 | 5252 | 5244 |
| parse-evaluate | logical-ifna-4096 | 979,964.50 | 922,614.00 | -5.85 | 5228 | 5256 |
| parse-evaluate | lazy-if-true-heavy-1024 | 993.75 | 965.00 | -2.89 | 2644 | 2676 |
| parse-evaluate | lazy-if-false-reference | 786.48 | 766.88 | -2.49 | 2516 | 2624 |
| parse-evaluate | text-if-escaped | 817.66 | 795.95 | -2.66 | 2644 | 2616 |
| parse-evaluate | text-if-concat | 1,101.09 | 1,099.62 | -0.13 | 2596 | 2668 |
| parse-evaluate | bitwise-and-64 | 21,598.06 | 21,539.00 | -0.27 | 2520 | 2560 |
| parse-evaluate | bitwise-or-256 | 80,558.19 | 79,693.47 | -1.07 | 2936 | 2872 |
| parse-evaluate | bitwise-xor-1024 | 325,530.25 | 318,021.38 | -2.31 | 3312 | 3324 |
| parse-evaluate | bitwise-lshift-4096 | 1,339,071.00 | 1,351,976.00 | 0.96 | 4872 | 4848 |
| parse-evaluate | bitwise-rshift-4096 | 1,368,816.00 | 1,367,091.00 | -0.13 | 5012 | 4912 |
| parse-evaluate | bitwise-coerce-text | 829.14 | 833.12 | 0.48 | 2620 | 2576 |
| parse-evaluate | bitwise-error-shift | 784.61 | 775.87 | -1.11 | 2704 | 2664 |
| parse-evaluate | bitwise-lazy-selected | 1,091.95 | 1,085.94 | -0.55 | 2632 | 2668 |

## New radix evaluation workloads

The following are the evaluate-phase rows; all three phases and all counters are retained in [radix raw data](candidate/radix/raw.csv). Refusal lanes measure rejection, not successful calculation. Input bytes count the complete generated formula.

| Case | Input bytes | ns/op | Peak tracked heap bytes | RSS KiB | Typed failure |
| --- | ---: | ---: | ---: | ---: | --- |
| radix-base | 13 | 926.41 | 450 | 2476 | none |
| radix-bin2dec | 16 | 454.93 | 400 | 2676 | none |
| radix-bin2hex | 16 | 707.59 | 401 | 2624 | none |
| radix-bin2oct | 16 | 766.33 | 402 | 2664 | none |
| radix-dec2bin | 12 | 892.97 | 404 | 2540 | none |
| radix-dec2hex | 13 | 762.27 | 402 | 2664 | none |
| radix-dec2oct | 12 | 761.02 | 402 | 2540 | none |
| radix-decimal | 17 | 633.59 | 448 | 2668 | none |
| radix-hex2bin | 13 | 903.83 | 404 | 2668 | none |
| radix-hex2dec | 14 | 455.63 | 400 | 2668 | none |
| radix-hex2oct | 13 | 757.42 | 402 | 2672 | none |
| radix-oct2bin | 14 | 904.38 | 404 | 2536 | none |
| radix-oct2dec | 14 | 457.27 | 400 | 2644 | none |
| radix-oct2hex | 14 | 706.09 | 401 | 2656 | none |
| radix-decimal-space | 19 | 632.11 | 448 | 2668 | none |
| radix-decimal-tab | 18 | 668.59 | 448 | 2632 | none |
| radix-decimal-prefix | 19 | 631.95 | 448 | 2668 | none |
| radix-decimal-h | 18 | 630.23 | 448 | 2720 | none |
| radix-decimal-b | 19 | 662.27 | 448 | 2668 | none |
| radix-base-truncate | 14 | 842.66 | 449 | 2664 | none |
| radix-error-fraction | 14 | 430.95 | 400 | 2664 | none |
| radix-error-invalid-digit | 17 | 623.05 | 448 | 2640 | none |
| radix-error-radix-low | 16 | 556.96 | 448 | 2640 | none |
| radix-error-radix-high | 17 | 567.89 | 448 | 2624 | none |
| radix-error-empty | 15 | 589.30 | 448 | 2668 | none |
| radix-error-arity | 8 | 418.36 | 400 | 2644 | none |
| radix-error-text-number | 13 | 443.36 | 400 | 2648 | none |
| radix-refusal-work | 13 | 272.89 | 264 | 2644 | resource-work |
| radix-refusal-memory | 13 | 820.94 | 488 | 2596 | resource-memory |
| radix-refusal-stack | 13 | 144.77 | 216 | 2644 | resource-objects |
| radix-refusal-cancelled | 13 | 7.50 | 0 | 2668 | cancelled |
| radix-base-input-64 | 832 | 38,499.39 | 6802 | 2536 | none |
| radix-base-input-256 | 3328 | 146,510.34 | 25618 | 2904 | none |
| radix-base-input-1024 | 13312 | 588,998.88 | 100882 | 3436 | none |
| radix-base-input-4096 | 53248 | 2,316,625.50 | 401938 | 4972 | none |
| radix-decimal-input-64 | 1088 | 19,649.16 | 6672 | 2668 | none |
| radix-decimal-input-256 | 4352 | 75,995.94 | 25104 | 2872 | none |
| radix-decimal-input-1024 | 17408 | 298,186.38 | 98832 | 3432 | none |
| radix-decimal-input-4096 | 69632 | 1,199,680.50 | 393744 | 4960 | none |
| radix-dec2hex-padding-64 | 960 | 38,672.20 | 7127 | 2632 | none |
| radix-dec2hex-padding-256 | 3840 | 148,434.12 | 26903 | 2924 | none |
| radix-dec2hex-padding-1024 | 15360 | 591,319.00 | 106007 | 3180 | none |
| radix-dec2hex-padding-4096 | 61440 | 2,373,405.50 | 422423 | 4952 | none |
| radix-base-padding-64 | 16 | 1,166.09 | 688 | 2668 | none |
| radix-base-padding-256 | 17 | 1,450.62 | 880 | 2476 | none |
| radix-base-padding-1024 | 18 | 2,506.38 | 1648 | 2668 | none |
| radix-base-padding-4096 | 18 | 6,780.00 | 4720 | 2616 | none |
| radix-max-base | 32 | 20,683.45 | 704 | 2624 | none |
| radix-max-decimal | 271 | 3,392.67 | 448 | 2668 | none |
| radix-big-base | 26 | 1,858.21 | 462 | 2628 | none |
| radix-error-large-dec2bin | 26 | 508.12 | 400 | 2668 | none |
| radix-error-large-dec2hex | 26 | 512.66 | 400 | 2648 | none |
| radix-big-decimal | 29 | 814.53 | 448 | 2668 | none |

## Hardware counters and mechanism

Five serial perf-stat captures cover the common text-concatenation control in both revisions and candidate finite-maximum BASE, finite-maximum DECIMAL and 4096-byte padding. Counters include whole-process setup, preflight and warmups; these totals are not isolated function costs. Exact commands and binary hashes are in the metadata sidecars.

| Capture | Cycles | Instructions | Branches | Branch misses |
| --- | ---: | ---: | ---: | ---: |
| baseline-text-if-concat | 6,269,421,800 | 11,592,684,758 | 1,970,011,276 | 113,656 |
| candidate-radix-base-padding-4096 | 557,343,909 | 1,812,547,945 | 405,846,532 | 82,011 |
| candidate-radix-max-base | 1,678,693,425 | 2,445,995,961 | 386,490,646 | 698,780 |
| candidate-radix-max-decimal | 275,558,388 | 1,001,654,190 | 68,102,445 | 59,948 |
| candidate-text-if-concat | 6,404,298,991 | 11,627,580,889 | 1,964,751,479 | 194,656 |

The new arithmetic uses a fixed 32-limb magnitude and fixed reverse-digit storage, with no heap big-integer dependency. Fallible text allocation retains its explicit reservation until the result is dropped. The outlined handler adds no evaluator frame variants or hidden parallel work. These are implementation mechanisms; the retained measurements support only the listed corpus, not an extrapolated throughput or cross-platform claim.

No Amdahl speedup estimate or scaling curve is applicable to this synchronous scalar family: the batch adds functionality and introduces no worker scheduling. Workbook/reference/array evaluation and Roman numerals remain outside this batch. Further profiling is needed before optimizing larger workbook workloads or treating shared-host tail/RSS thresholds as hard gates.
