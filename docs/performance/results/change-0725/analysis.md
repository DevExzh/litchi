# Change 0725 worksheet replay checkpoint pilot

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
- Repeated-query semantic and p50/mean timing gates: PASS.
- Native 54016-late owned q8 benefit gate: PASS.
- Repeated 54016-late owned benefit gate: PASS.
- Allocator outcome and fixed-field bounds: PASS.
- Primary counted-I/O semantic parity: PASS.
- Budget-fence semantic parity: PASS.
- Binding and command audit: PASS.

## Selected timing rows

| Case | Mode | Metric | B1/A1 p50 | B2/A2 p50 | B1/A1 mean | B2/A2 mean |
|---|---|---|---:|---:|---:|---:|
| 54016-stored-0 | owned | q8 | +9.08% | +9.71% | +9.07% | +9.53% |
| 54016-stored-0 | owned | open-plus-eight | +8.06% | +8.22% | +8.07% | +8.29% |
| 54016-stored-0 | file | q8 | +7.30% | +6.86% | +6.97% | +6.69% |
| 54016-stored-0 | file | open-plus-eight | +6.11% | +5.80% | +6.01% | +5.86% |
| 54016-stored-2097152 | owned | q8 | +1.92% | +1.96% | +1.56% | +0.93% |
| 54016-stored-2097152 | owned | open-plus-eight | +0.08% | +1.45% | -0.38% | +1.11% |
| 54016-stored-2097152 | file | q8 | -1.88% | -0.72% | -2.05% | +2.51% |
| 54016-stored-2097152 | file | open-plus-eight | +0.32% | +0.78% | -0.41% | -0.06% |
| 54016-missing-1048576 | owned | q8 | -25.00% | -25.00% | -24.30% | -24.69% |
| 54016-missing-1048576 | owned | open-plus-eight | +1.42% | +0.44% | +0.94% | +0.19% |
| 54016-missing-1048576 | file | q8 | -4.69% | -3.17% | -4.69% | -12.64% |
| 54016-missing-1048576 | file | open-plus-eight | +1.03% | +0.77% | +0.91% | +0.71% |
| Plan1-stored-2097152 | owned | q8 | +6.38% | +4.26% | +5.72% | +5.22% |
| Plan1-stored-2097152 | owned | open-plus-eight | +0.42% | -0.69% | +0.56% | -0.81% |
| Plan1-stored-2097152 | file | q8 | -1.20% | +1.73% | +3.55% | +1.58% |
| Plan1-stored-2097152 | file | open-plus-eight | +16.03% | +0.57% | +16.03% | +0.41% |
| Simple-stored-2097152 | owned | q8 | +0.00% | +4.76% | +2.58% | +3.40% |
| Simple-stored-2097152 | owned | open-plus-eight | -1.49% | -0.44% | -3.12% | -0.51% |
| Simple-stored-2097152 | file | q8 | -0.99% | +0.00% | -1.55% | -0.15% |
| Simple-stored-2097152 | file | open-plus-eight | -2.16% | -1.39% | -2.07% | -1.79% |
| Simple-missing-2097152 | owned | q8 | -14.29% | -14.29% | -20.51% | -18.75% |
| Simple-missing-2097152 | owned | open-plus-eight | -2.11% | -4.43% | -4.25% | -3.25% |
| Simple-missing-2097152 | file | q8 | -3.45% | -1.72% | +10.06% | -2.24% |
| Simple-missing-2097152 | file | open-plus-eight | -3.37% | -2.74% | -6.49% | -3.32% |
| synthetic-70000-default | owned | q8 | +8.33% | +6.25% | +6.22% | +6.73% |
| synthetic-70000-default | owned | open-plus-eight | -2.70% | -0.60% | -2.69% | -0.60% |
| synthetic-70000-default | file | q8 | +1.95% | +1.97% | +2.03% | -0.73% |
| synthetic-70000-default | file | open-plus-eight | -2.26% | -1.19% | -2.30% | -1.12% |
| 54016-late | owned | q8 | -80.00% | -80.82% | -79.85% | -80.51% |
| 54016-late | owned | open-plus-eight | +0.79% | +0.43% | -1.27% | -1.61% |
| 54016-late | file | q8 | -44.66% | -45.08% | -45.50% | -46.45% |
| 54016-late | file | open-plus-eight | +0.33% | +0.10% | +0.26% | +0.12% |
| Plan1-late | owned | q8 | -28.12% | -22.58% | -27.22% | -22.27% |
| Plan1-late | owned | open-plus-eight | -0.70% | -0.12% | -0.59% | -0.15% |
| Plan1-late | file | q8 | -5.48% | -6.21% | -5.15% | -6.07% |
| Plan1-late | file | open-plus-eight | -1.22% | -1.67% | -1.22% | -1.68% |
| 45365-first | owned | q8 | +5.10% | +10.42% | +4.34% | +3.44% |
| 45365-first | owned | open-plus-eight | -0.96% | -0.35% | -0.88% | -0.38% |
| 45365-first | file | q8 | +1.96% | +1.47% | +1.72% | +1.35% |
| 45365-first | file | open-plus-eight | -1.12% | -2.62% | -1.23% | -2.61% |
| 45365-late | owned | q8 | -52.83% | -52.83% | -53.62% | -52.32% |
| 45365-late | owned | open-plus-eight | -3.39% | -1.55% | -3.42% | -1.57% |
| 45365-late | file | q8 | -16.96% | -15.88% | -20.59% | -19.89% |
| 45365-late | file | open-plus-eight | -4.61% | +0.44% | -4.78% | -0.06% |
| formula-refusal-2097152 | owned | q8 | -1.28% | +0.26% | +0.17% | -0.17% |
| formula-refusal-2097152 | owned | open-plus-eight | -1.59% | -0.61% | -1.66% | -0.79% |
| formula-refusal-2097152 | file | q8 | -1.46% | -0.83% | -4.30% | -0.86% |
| formula-refusal-2097152 | file | open-plus-eight | -1.14% | -1.94% | -0.92% | -1.61% |

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
