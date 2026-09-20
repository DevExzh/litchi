# Change 0726 empty-slot setup pilot

Status: **FAIL**

The packet binds the baseline revision, current candidate source, immutable probes,
raw fixture hashes, frozen binaries and per-phase build manifests before comparing
fresh-owner query timings. A/A runs precede A/B/B/A when the staged commands are used.

## Capture counts

- Native groups: 24; expected 24.
- Repeated-query groups: 16; expected 16, with nine 50,000-query processes per leg.
- Allocator groups: 96; expected 96 per phase, with three repeats.
- Primary counted-I/O groups: 12; expected 12.
- Budget-fence observations: 34; unchanged224-byte index tested at all inherited boundaries.

## Hard gates

- Native semantic and p50/mean timing gates: FAIL.
- Repeated-query semantic and p50/mean timing gates: PASS.
- Native 54016-missing owned q8 benefit gate: PASS.
- Repeated 54016-missing owned benefit gate: PASS.
- Allocator outcome and fixed-field bounds: PASS.
- Primary counted-I/O semantic parity: PASS.
- Budget-fence semantic parity: PASS.
- Binding and command audit: PASS.

## Selected timing rows

| Case | Mode | Metric | B1/A1 p50 | B2/A2 p50 | B1/A1 mean | B2/A2 mean |
|---|---|---|---:|---:|---:|---:|
| 54016-stored-0 | owned | q8 | +1.31% | +3.33% | +2.15% | +2.93% |
| 54016-stored-0 | owned | open-plus-eight | +1.57% | +1.80% | +2.19% | +2.56% |
| 54016-stored-0 | file | q8 | +1.26% | +1.68% | +0.64% | +1.33% |
| 54016-stored-0 | file | open-plus-eight | +0.65% | +1.38% | +0.66% | +1.32% |
| 54016-stored-2097152 | owned | q8 | +0.00% | +1.92% | -11.21% | +2.49% |
| 54016-stored-2097152 | owned | open-plus-eight | -8.24% | -8.34% | -8.26% | -8.60% |
| 54016-stored-2097152 | file | q8 | -0.47% | +1.41% | -0.31% | +0.22% |
| 54016-stored-2097152 | file | open-plus-eight | -6.92% | -6.56% | -7.02% | -6.62% |
| 54016-missing-1048576 | owned | q8 | -25.00% | -25.00% | -27.68% | -26.03% |
| 54016-missing-1048576 | owned | open-plus-eight | -8.52% | -9.17% | -8.55% | -9.08% |
| 54016-missing-1048576 | file | q8 | -3.17% | -1.59% | -4.00% | -2.49% |
| 54016-missing-1048576 | file | open-plus-eight | -8.88% | -6.61% | -8.76% | -6.48% |
| Plan1-stored-2097152 | owned | q8 | +4.26% | +0.00% | +3.52% | -0.97% |
| Plan1-stored-2097152 | owned | open-plus-eight | -4.13% | -4.50% | -4.00% | -4.43% |
| Plan1-stored-2097152 | file | q8 | -3.35% | -0.98% | +0.09% | -3.34% |
| Plan1-stored-2097152 | file | open-plus-eight | -3.90% | -3.18% | -3.78% | -3.21% |
| Simple-stored-2097152 | owned | q8 | +1.16% | +2.33% | -6.64% | +0.51% |
| Simple-stored-2097152 | owned | open-plus-eight | +0.55% | -0.69% | +0.27% | -0.19% |
| Simple-stored-2097152 | file | q8 | +0.50% | -2.92% | +0.77% | -3.22% |
| Simple-stored-2097152 | file | open-plus-eight | +1.08% | -1.88% | +1.23% | -1.51% |
| Simple-missing-2097152 | owned | q8 | -25.00% | -25.00% | -23.44% | -23.83% |
| Simple-missing-2097152 | owned | open-plus-eight | -2.21% | -1.31% | -3.06% | -1.21% |
| Simple-missing-2097152 | file | q8 | -3.45% | -1.69% | -3.57% | -1.94% |
| Simple-missing-2097152 | file | open-plus-eight | -1.52% | +0.31% | -2.37% | +0.66% |
| synthetic-70000-default | owned | q8 | +4.17% | +4.17% | +3.08% | +3.22% |
| synthetic-70000-default | owned | open-plus-eight | -11.15% | -14.38% | -11.03% | -14.35% |
| synthetic-70000-default | file | q8 | -1.90% | -0.49% | -2.07% | -1.14% |
| synthetic-70000-default | file | open-plus-eight | -16.50% | -10.11% | -16.51% | -10.11% |
| 54016-late | owned | q8 | +0.00% | +1.38% | +1.66% | +1.13% |
| 54016-late | owned | open-plus-eight | -6.44% | -6.39% | -6.44% | -6.31% |
| 54016-late | file | q8 | -0.38% | +1.14% | -0.32% | +1.37% |
| 54016-late | file | open-plus-eight | -7.81% | -7.47% | -7.79% | -7.43% |
| Plan1-late | owned | q8 | +3.23% | +0.00% | -18.25% | +1.49% |
| Plan1-late | owned | open-plus-eight | -4.27% | -4.07% | -4.09% | -3.85% |
| Plan1-late | file | q8 | -2.03% | -1.35% | -1.61% | -1.47% |
| Plan1-late | file | open-plus-eight | -3.79% | -2.87% | -3.89% | -3.10% |
| 45365-first | owned | q8 | +2.04% | -2.91% | +1.87% | -1.67% |
| 45365-first | owned | open-plus-eight | +0.62% | +0.56% | +0.52% | +0.61% |
| 45365-first | file | q8 | +0.97% | -0.97% | +0.33% | -0.98% |
| 45365-first | file | open-plus-eight | -0.65% | -2.16% | -0.66% | -2.53% |
| 45365-late | owned | q8 | +0.00% | +0.00% | +0.44% | +0.35% |
| 45365-late | owned | open-plus-eight | +2.66% | +2.04% | +1.91% | +1.97% |
| 45365-late | file | q8 | -1.14% | +1.74% | -0.98% | +8.40% |
| 45365-late | file | open-plus-eight | +0.19% | +2.46% | +0.18% | +2.49% |
| formula-refusal-2097152 | owned | q8 | -3.08% | -3.46% | +3.22% | -8.00% |
| formula-refusal-2097152 | owned | open-plus-eight | -1.16% | -1.35% | -1.41% | -1.13% |
| formula-refusal-2097152 | file | q8 | -2.05% | -0.74% | -4.48% | -2.20% |
| formula-refusal-2097152 | file | open-plus-eight | -0.91% | -0.17% | -2.00% | +0.29% |

## Budget fences

| Case | Budget | Semantic parity | Route metrics changed |
|---|---:|---|---|
| Simple-fence | 0 | PASS | no |
| Simple-fence | 223 | PASS | no |
| Simple-fence | 224 | PASS | no |
| Simple-fence | 263 | PASS | no |
| Simple-fence | 264 | PASS | no |
| Simple-fence | 450 | PASS | no |
| Simple-fence | 489 | PASS | no |
| Simple-fence | 490 | PASS | no |
| Simple-fence | 491 | PASS | no |
| Simple-fence | 2097152 | PASS | no |
| Simple-missing-fence | 0 | PASS | no |
| Simple-missing-fence | 223 | PASS | no |
| Simple-missing-fence | 224 | PASS | no |
| Simple-missing-fence | 263 | PASS | no |
| Simple-missing-fence | 264 | PASS | no |
| Simple-missing-fence | 450 | PASS | no |
| Simple-missing-fence | 489 | PASS | no |
| Simple-missing-fence | 490 | PASS | no |
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
