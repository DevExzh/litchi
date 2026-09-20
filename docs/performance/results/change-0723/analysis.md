# Change 0723 worksheet replay checkpoint pilot

Status: **FAIL**

The packet binds the baseline revision, current candidate source, immutable probes,
raw fixture hashes, frozen binaries and per-phase build manifests before comparing
fresh-owner query timings. A/A runs precede A/B/B/A when the staged commands are used.

## Capture counts

- Native groups: 24; expected 24.
- Repeated-query groups: 16; expected 16, with nine 50,000-query processes per leg.
- Allocator groups: 96; expected 96 per phase, with three repeats.
- Primary counted-I/O groups: 12; expected 12.
- Budget-fence observations: 34; logical charge fences include 224 and 264 bytes.

## Hard gates

- Native semantic and p50/mean timing gates: FAIL.
- Repeated-query semantic and p50/mean timing gates: FAIL.
- Native 54016-late owned q8 benefit gate: PASS.
- Repeated 54016-late owned benefit gate: PASS.
- Allocator outcome and fixed-field bounds: PASS.
- Primary counted-I/O semantic parity: PASS.
- Budget-fence semantic parity: PASS.
- Binding and command audit: PASS.

## Selected timing rows

| Case | Mode | Metric | B1/A1 p50 | B2/A2 p50 | B1/A1 mean | B2/A2 mean |
|---|---|---|---:|---:|---:|---:|
| 54016-stored-0 | owned | q8 | +1.82% | +1.43% | +2.00% | +1.27% |
| 54016-stored-0 | owned | open-plus-eight | +1.59% | +1.30% | +1.60% | +1.36% |
| 54016-stored-0 | file | q8 | +0.74% | +2.01% | +0.32% | +1.62% |
| 54016-stored-0 | file | open-plus-eight | +0.36% | +1.55% | +0.33% | +1.54% |
| 54016-stored-2097152 | owned | q8 | +3.92% | +1.92% | +3.86% | +2.11% |
| 54016-stored-2097152 | owned | open-plus-eight | -1.70% | +0.02% | -2.41% | -0.93% |
| 54016-stored-2097152 | file | q8 | +0.96% | -0.96% | +0.62% | -0.24% |
| 54016-stored-2097152 | file | open-plus-eight | -0.61% | -0.59% | -1.80% | -1.81% |
| 54016-missing-1048576 | owned | q8 | +8.33% | +8.33% | +3.60% | +5.01% |
| 54016-missing-1048576 | owned | open-plus-eight | -0.17% | -0.57% | -0.63% | -0.97% |
| 54016-missing-1048576 | file | q8 | +0.00% | +1.59% | +0.14% | -10.24% |
| 54016-missing-1048576 | file | open-plus-eight | -1.04% | -0.85% | -1.02% | -0.73% |
| Plan1-stored-2097152 | owned | q8 | +4.26% | +13.04% | +4.36% | +12.06% |
| Plan1-stored-2097152 | owned | open-plus-eight | +1.96% | +3.72% | +1.90% | +4.14% |
| Plan1-stored-2097152 | file | q8 | +1.25% | +1.49% | +5.31% | +1.57% |
| Plan1-stored-2097152 | file | open-plus-eight | -0.29% | -0.30% | -0.28% | -0.32% |
| Simple-stored-2097152 | owned | q8 | +4.76% | +4.76% | +3.41% | +6.88% |
| Simple-stored-2097152 | owned | open-plus-eight | +1.58% | +2.10% | +1.10% | +3.83% |
| Simple-stored-2097152 | file | q8 | +0.50% | +1.50% | +4.03% | -1.31% |
| Simple-stored-2097152 | file | open-plus-eight | +0.14% | +1.14% | +1.10% | +1.17% |
| Simple-missing-2097152 | owned | q8 | +0.00% | -12.50% | -8.61% | -9.80% |
| Simple-missing-2097152 | owned | open-plus-eight | +0.07% | +0.01% | -3.02% | +0.89% |
| Simple-missing-2097152 | file | q8 | +0.88% | +0.00% | -20.25% | -12.97% |
| Simple-missing-2097152 | file | open-plus-eight | +1.20% | +0.62% | +0.74% | -0.18% |
| synthetic-70000-default | owned | q8 | +8.51% | +7.55% | +9.19% | +5.98% |
| synthetic-70000-default | owned | open-plus-eight | -3.25% | -3.48% | -3.22% | -3.53% |
| synthetic-70000-default | file | q8 | +2.24% | +0.99% | +1.79% | +4.06% |
| synthetic-70000-default | file | open-plus-eight | -3.17% | -3.30% | -3.20% | -3.34% |
| 54016-late | owned | q8 | -79.31% | -79.86% | -79.63% | -79.83% |
| 54016-late | owned | open-plus-eight | -1.15% | -0.66% | -3.27% | -2.73% |
| 54016-late | file | q8 | -42.91% | -43.73% | -39.23% | -43.11% |
| 54016-late | file | open-plus-eight | -0.41% | +6.13% | -0.40% | +6.19% |
| Plan1-late | owned | q8 | -22.58% | -25.81% | -22.77% | -26.11% |
| Plan1-late | owned | open-plus-eight | +3.63% | +0.19% | +3.22% | +0.46% |
| Plan1-late | file | q8 | -4.79% | -3.42% | +0.15% | -3.37% |
| Plan1-late | file | open-plus-eight | +0.16% | +2.31% | +0.34% | +1.82% |
| 45365-first | owned | q8 | +4.08% | +6.25% | +2.02% | +4.67% |
| 45365-first | owned | open-plus-eight | -0.67% | -0.94% | -0.87% | -0.85% |
| 45365-first | file | q8 | +1.23% | +3.00% | -1.83% | +2.30% |
| 45365-first | file | open-plus-eight | -1.24% | -0.66% | -1.16% | -0.71% |
| 45365-late | owned | q8 | -50.94% | -52.83% | -50.84% | -52.03% |
| 45365-late | owned | open-plus-eight | +0.04% | -0.61% | +0.03% | -0.88% |
| 45365-late | file | q8 | -13.53% | -15.20% | -11.90% | -20.29% |
| 45365-late | file | open-plus-eight | +1.36% | +1.20% | +1.45% | +0.81% |
| formula-refusal-2097152 | owned | q8 | +0.65% | +2.06% | +1.26% | +6.03% |
| formula-refusal-2097152 | owned | open-plus-eight | -0.41% | -0.62% | -0.31% | -0.25% |
| formula-refusal-2097152 | file | q8 | +1.11% | +0.19% | +1.09% | +1.40% |
| formula-refusal-2097152 | file | open-plus-eight | +0.43% | -1.16% | +0.18% | -0.75% |

## Budget fences

| Case | Budget | Semantic parity | Route metrics changed |
|---|---:|---|---|
| Simple-fence | 0 | PASS | no |
| Simple-fence | 223 | PASS | no |
| Simple-fence | 224 | PASS | no |
| Simple-fence | 263 | PASS | no |
| Simple-fence | 264 | PASS | no |
| Simple-fence | 450 | PASS | no |
| Simple-fence | 489 | PASS | yes |
| Simple-fence | 490 | PASS | yes |
| Simple-fence | 491 | PASS | no |
| Simple-fence | 2097152 | PASS | no |
| Simple-missing-fence | 0 | PASS | no |
| Simple-missing-fence | 223 | PASS | no |
| Simple-missing-fence | 224 | PASS | no |
| Simple-missing-fence | 263 | PASS | no |
| Simple-missing-fence | 264 | PASS | no |
| Simple-missing-fence | 450 | PASS | no |
| Simple-missing-fence | 489 | PASS | yes |
| Simple-missing-fence | 490 | PASS | yes |
| Simple-missing-fence | 491 | PASS | no |
| Simple-missing-fence | 2097152 | PASS | no |
| 54016-fence | 0 | PASS | no |
| 54016-fence | 223 | PASS | no |
| 54016-fence | 224 | PASS | no |
| 54016-fence | 263 | PASS | no |
| 54016-fence | 264 | PASS | no |
| 54016-fence | 1048576 | PASS | no |
| 54016-fence | 2097152 | PASS | no |
| 54016-missing-fence | 0 | PASS | no |
| 54016-missing-fence | 223 | PASS | no |
| 54016-missing-fence | 224 | PASS | no |
| 54016-missing-fence | 263 | PASS | no |
| 54016-missing-fence | 264 | PASS | no |
| 54016-missing-fence | 1048576 | PASS | no |
| 54016-missing-fence | 2097152 | PASS | no |

## Failures

None.
