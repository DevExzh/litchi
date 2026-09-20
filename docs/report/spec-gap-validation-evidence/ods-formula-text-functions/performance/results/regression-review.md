# Matched-control regression review

Disposition: accepted with explicit process-RSS cost for this new-function batch.

All 56 matched control/phase groups retain identical allocation calls, requested/released bytes, peak live heap, retained execution budget, work, reference reads and output bytes. Latency median shifts range from −3.88% to +3.69%; no latency median exceeds the +5% review threshold.

42 of 56 process-RSS comparisons exceed +5%; all RSS median shifts range from +1.14% to +10.00%. These are retained findings, not dismissed as noise. The table gives absolute deltas and independently bootstrapped intervals.

The candidate adds the Unicode tables and text/formatting implementation to the same linked evaluator. Increased resident code/data is a plausible contributor given unchanged measured heap accounting, but this capture does not attribute RSS causally or distinguish resident image pages from runtime overhead. The measurement is whole-child peak RSS, not per-call heap. Acceptance is for the bounded feature addition and these disclosed absolute costs; it is not a performance improvement claim. No threshold-based rerun was used.

| control / phase | baseline KiB | candidate KiB | delta KiB | delta % | bootstrap 95% delta % |
| --- | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic/evaluate | 3736 | 4044 | 308 | 8.24 | 5.64 to 9.29 |
| array-control-16x16-arithmetic/parse-evaluate | 3688 | 3964 | 276 | 7.48 | 6.22 to 9.32 |
| array-control-16x16-sin/evaluate | 4044 | 4264 | 220 | 5.44 | 5.24 to 6.97 |
| array-control-16x16-sin/parse-evaluate | 4032 | 4292 | 260 | 6.45 | 5.45 to 8.23 |
| array-control-4x4-arithmetic/evaluate | 3468 | 3752 | 284 | 8.19 | 6.20 to 9.43 |
| array-control-4x4-arithmetic/parse-evaluate | 3472 | 3704 | 232 | 6.68 | 5.07 to 7.50 |
| array-control-4x4-sin/evaluate | 3788 | 4068 | 280 | 7.39 | 5.87 to 8.81 |
| array-control-4x4-sin/parse-evaluate | 3784 | 4028 | 244 | 6.45 | 5.39 to 8.35 |
| concat-borrowed-literals/parse-evaluate | 3424 | 3664 | 240 | 7.01 | 4.57 to 14.09 |
| concat-growth-chain/evaluate | 3436 | 3668 | 232 | 6.75 | 2.11 to 12.33 |
| concat-growth-chain/parse-evaluate | 3448 | 3628 | 180 | 5.22 | 2.44 to 7.66 |
| concat-owned-left/parse-evaluate | 3432 | 3680 | 248 | 7.23 | 4.90 to 8.84 |
| concat-owned-right/parse-evaluate | 3376 | 3576 | 200 | 5.92 | 1.75 to 11.41 |
| database-control-dstdev/evaluate | 3484 | 3752 | 268 | 7.69 | 5.23 to 9.77 |
| database-control-dstdev/parse-evaluate | 3492 | 3712 | 220 | 6.30 | 4.59 to 10.05 |
| database-control-dsum/evaluate | 3480 | 3828 | 348 | 10.00 | 7.04 to 11.95 |
| database-control-dsum/parse-evaluate | 3496 | 3716 | 220 | 6.29 | 3.46 to 7.64 |
| database-control-dvar/evaluate | 3496 | 3752 | 256 | 7.32 | 3.79 to 9.50 |
| database-control-dvar/parse-evaluate | 3516 | 3736 | 220 | 6.26 | 4.25 to 8.66 |
| literal-aggregate-4x1-sum/evaluate | 3480 | 3720 | 240 | 6.90 | 5.55 to 8.35 |
| literal-aggregate-4x1-sum/parse-evaluate | 3496 | 3800 | 304 | 8.70 | 5.88 to 9.92 |
| reference-aggregate-64x4-sum/evaluate | 3476 | 3700 | 224 | 6.44 | 5.14 to 7.65 |
| reference-aggregate-64x4-sum/parse-evaluate | 3492 | 3764 | 272 | 7.79 | 6.33 to 10.38 |
| reference-array-16x4-arithmetic/evaluate | 3468 | 3728 | 260 | 7.50 | 5.06 to 11.23 |
| reference-array-16x4-arithmetic/parse-evaluate | 3464 | 3716 | 252 | 7.27 | 6.42 to 8.85 |
| reference-conditional-256x4-sumifs/evaluate | 3480 | 3712 | 232 | 6.67 | 5.96 to 10.44 |
| reference-conditional-256x4-sumifs/parse-evaluate | 3484 | 3728 | 244 | 7.00 | 2.06 to 9.16 |
| reference-control-average/evaluate | 3480 | 3724 | 244 | 7.01 | 4.99 to 8.90 |
| reference-control-average/parse-evaluate | 3500 | 3812 | 312 | 8.91 | 6.33 to 10.40 |
| reference-control-counta/evaluate | 3484 | 3832 | 348 | 9.99 | 6.33 to 11.56 |
| reference-control-counta/parse-evaluate | 3472 | 3752 | 280 | 8.06 | 5.97 to 9.12 |
| representative-median/parse-evaluate | 3420 | 3592 | 172 | 5.03 | 1.50 to 8.48 |
| scalar-aggregate-sum/evaluate | 3468 | 3684 | 216 | 6.23 | 3.10 to 7.56 |
| scalar-aggregate-sum/parse-evaluate | 3436 | 3684 | 248 | 7.22 | 5.17 to 7.85 |
| scalar-control-arithmetic/evaluate | 3436 | 3692 | 256 | 7.45 | 3.73 to 8.78 |
| scalar-control-arithmetic/parse-evaluate | 3456 | 3648 | 192 | 5.56 | 1.14 to 6.88 |
| scalar-control-counta/evaluate | 3468 | 3664 | 196 | 5.65 | 2.77 to 6.57 |
| scalar-control-imsum/parse-evaluate | 3444 | 3664 | 220 | 6.39 | 4.12 to 11.99 |
| scalar-control-sin/evaluate | 3732 | 3944 | 212 | 5.68 | 2.25 to 6.32 |
| scalar-control-sin/parse-evaluate | 3744 | 3940 | 196 | 5.24 | 3.72 to 7.48 |
| scalar-control-stdev/parse-evaluate | 3448 | 3680 | 232 | 6.73 | 2.57 to 7.94 |
| scalar-control-var/evaluate | 3448 | 3668 | 220 | 6.38 | 1.85 to 7.66 |

Intervals use the retained 15 samples per side, 10,000 independent median resamples, seed 20260920. They quantify this corpus/run only; they are not universal confidence bounds for deployment workloads.

Prior captures are diagnostic only: unsupported baseline operand producer, timed scalar-error consumer, and cancellation repeat-count mismatch. Their raw receipts and reconstructing harnesses remain in separate diagnostic directories. Final acceptance uses only the complete protocol-correct capture.
