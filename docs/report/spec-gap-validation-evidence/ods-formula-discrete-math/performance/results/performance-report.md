# ODS discrete-math evaluator performance profile

The baseline is committed `782339a2c`. It contributes only matched controls: arithmetic, SIN, IMSUM, DSUM, and SUM across scalar, literal-array, and streamed-reference entry points. The eleven new discrete-math functions have no valid baseline implementation, so their rows are candidate-only evidence. Each cell is the p50 across fifteen fresh child processes; time and work are normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-4x4-arithmetic | evaluate | 4,350 | 4,399 | +1.1% | 1,120 | 1,120 | 2,920 | 2,964 |
| array-control-4x4-arithmetic | parse-evaluate | 5,185 | 5,166 | -0.4% | 2,240 | 2,240 | 2,924 | 2,960 |
| array-control-4x4-sin | evaluate | 9,542 | 9,759 | +2.3% | 3,840 | 3,840 | 3,244 | 3,276 |
| array-control-4x4-sin | parse-evaluate | 10,530 | 10,560 | +0.3% | 5,040 | 5,040 | 3,224 | 3,268 |
| database-control-dsum | evaluate | 3,920 | 3,830 | -2.3% | 20 | 20 | 2,916 | 3,120 |
| database-control-dsum | parse-evaluate | 4,510 | 4,370 | -3.1% | 29 | 29 | 2,924 | 3,092 |
| literal-aggregate-4x4-sum | evaluate | 3,196 | 3,227 | +1.0% | 1,120 | 1,120 | 2,912 | 2,976 |
| literal-aggregate-4x4-sum | parse-evaluate | 4,046 | 4,103 | +1.4% | 2,320 | 2,320 | 2,932 | 2,972 |
| reference-aggregate-16x4-sum | evaluate | 3,443 | 3,531 | +2.6% | 640 | 640 | 2,940 | 2,988 |
| reference-aggregate-16x4-sum | parse-evaluate | 3,795 | 3,880 | +2.2% | 1,200 | 1,200 | 2,940 | 2,972 |
| reference-array-16x4-arithmetic | evaluate | 9,374 | 9,260 | -1.2% | 800 | 800 | 2,928 | 2,976 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,628 | 9,622 | -0.1% | 1,280 | 1,280 | 2,912 | 2,984 |
| reference-array-16x4-sin | evaluate | 29,268 | 29,490 | +0.8% | 11,120 | 11,120 | 3,228 | 3,272 |
| reference-array-16x4-sin | parse-evaluate | 29,852 | 29,922 | +0.2% | 11,680 | 11,680 | 3,228 | 3,272 |
| scalar-aggregate-sum | evaluate | 548 | 554 | +1.1% | 4,000 | 4,000 | 2,920 | 2,960 |
| scalar-aggregate-sum | parse-evaluate | 729 | 741 | +1.6% | 8,000 | 8,000 | 2,896 | 2,964 |
| scalar-control-arithmetic | evaluate | 555 | 554 | -0.2% | 5,000 | 5,000 | 2,884 | 2,956 |
| scalar-control-arithmetic | parse-evaluate | 700 | 704 | +0.6% | 8,000 | 8,000 | 2,900 | 2,960 |
| scalar-control-imsum | evaluate | 1,410 | 1,373 | -2.6% | 8,000 | 8,000 | 2,896 | 2,972 |
| scalar-control-imsum | parse-evaluate | 1,921 | 1,936 | +0.8% | 17,000 | 17,000 | 2,880 | 2,992 |
| scalar-control-sin | evaluate | 458 | 449 | -2.0% | 4,000 | 4,000 | 3,200 | 3,208 |
| scalar-control-sin | parse-evaluate | 628 | 637 | +1.4% | 8,000 | 8,000 | 3,204 | 3,240 |

## Candidate discrete-math workloads

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| literal-discrete-4x4-combin | evaluate | 35,069 | 467 | 0 | 4,080 | 2,232,960 | 12,696 | 1,408 | 2,952 |
| literal-discrete-4x4-combin | parse-evaluate | 36,123 | 467 | 0 | 5,840 | 3,318,080 | 19,668 | 1,408 | 2,972 |
| literal-discrete-4x4-combina | evaluate | 36,779 | 452 | 0 | 4,080 | 2,232,960 | 12,696 | 1,408 | 3,032 |
| literal-discrete-4x4-combina | parse-evaluate | 38,282 | 452 | 0 | 5,840 | 3,316,880 | 19,653 | 1,408 | 2,976 |
| literal-discrete-4x4-delta | evaluate | 12,770 | 397 | 0 | 4,080 | 2,232,960 | 12,696 | 1,408 | 2,964 |
| literal-discrete-4x4-delta | parse-evaluate | 14,271 | 397 | 0 | 5,840 | 3,316,720 | 19,651 | 1,408 | 2,976 |
| literal-discrete-4x4-even | evaluate | 11,765 | 245 | 0 | 3,840 | 1,624,960 | 8,216 | 1,408 | 3,116 |
| literal-discrete-4x4-even | parse-evaluate | 12,908 | 245 | 0 | 5,200 | 2,685,280 | 15,038 | 1,408 | 3,040 |
| literal-discrete-4x4-fact | evaluate | 22,241 | 594 | 0 | 3,840 | 1,624,960 | 8,216 | 1,408 | 2,972 |
| literal-discrete-4x4-fact | parse-evaluate | 23,042 | 594 | 0 | 5,040 | 2,152,960 | 11,712 | 1,408 | 2,960 |
| literal-discrete-4x4-factdouble | evaluate | 17,805 | 456 | 0 | 3,840 | 1,624,960 | 8,216 | 1,408 | 2,972 |
| literal-discrete-4x4-factdouble | parse-evaluate | 18,727 | 456 | 0 | 5,040 | 2,153,440 | 11,718 | 1,408 | 3,060 |
| literal-discrete-4x4-gcd | evaluate | 5,844 | 399 | 0 | 1,360 | 1,389,440 | 10,552 | 0 | 2,964 |
| literal-discrete-4x4-gcd | parse-evaluate | 7,091 | 399 | 0 | 3,120 | 2,473,200 | 17,507 | 0 | 2,992 |
| literal-discrete-4x4-gestep | evaluate | 13,123 | 406 | 0 | 4,080 | 2,232,960 | 12,696 | 1,408 | 3,108 |
| literal-discrete-4x4-gestep | parse-evaluate | 14,597 | 406 | 0 | 5,840 | 3,317,120 | 19,656 | 1,408 | 3,108 |
| literal-discrete-4x4-lcm | evaluate | 5,728 | 399 | 0 | 1,360 | 1,389,440 | 10,552 | 0 | 3,016 |
| literal-discrete-4x4-lcm | parse-evaluate | 6,968 | 399 | 0 | 3,120 | 2,473,200 | 17,507 | 0 | 3,016 |
| literal-discrete-4x4-multinomial | evaluate | 20,268 | 2,135 | 0 | 1,360 | 1,389,440 | 10,552 | 0 | 3,092 |
| literal-discrete-4x4-multinomial | parse-evaluate | 21,674 | 2,135 | 0 | 3,120 | 2,473,840 | 17,515 | 0 | 3,012 |
| literal-discrete-4x4-odd | evaluate | 11,582 | 244 | 0 | 3,840 | 1,624,960 | 8,216 | 1,408 | 3,088 |
| literal-discrete-4x4-odd | parse-evaluate | 12,627 | 244 | 0 | 5,200 | 2,685,200 | 15,037 | 1,408 | 3,020 |
| nested-discrete-1024-gcd | evaluate | 2,230,261 | 49,211 | 12,288 | 4,156 | 2,136,168 | 1,445,416 | 0 | 4,596 |
| nested-discrete-1024-gcd | parse-evaluate | 2,223,611 | 49,211 | 12,288 | 4,173 | 2,140,137 | 1,448,041 | 0 | 4,596 |
| nested-discrete-1024-lcm | evaluate | 2,167,271 | 49,211 | 12,288 | 4,156 | 2,136,168 | 1,445,416 | 0 | 4,592 |
| nested-discrete-1024-lcm | parse-evaluate | 2,172,651 | 49,211 | 12,288 | 4,173 | 2,140,137 | 1,448,041 | 0 | 4,616 |
| nested-discrete-1024-multinomial | evaluate | 2,878,675 | 483,427 | 12,288 | 4,156 | 2,136,168 | 1,445,416 | 0 | 4,588 |
| nested-discrete-1024-multinomial | parse-evaluate | 2,856,014 | 483,427 | 12,288 | 4,173 | 2,140,145 | 1,448,049 | 0 | 4,584 |
| nested-discrete-256-gcd | evaluate | 553,512 | 12,347 | 3,072 | 2,160 | 1,077,456 | 364,072 | 0 | 3,628 |
| nested-discrete-256-gcd | parse-evaluate | 556,008 | 12,347 | 3,072 | 2,194 | 1,085,388 | 366,694 | 0 | 3,512 |
| nested-discrete-256-lcm | evaluate | 536,972 | 12,347 | 3,072 | 2,160 | 1,077,456 | 364,072 | 0 | 3,624 |
| nested-discrete-256-lcm | parse-evaluate | 545,653 | 12,347 | 3,072 | 2,194 | 1,085,388 | 366,694 | 0 | 3,580 |
| nested-discrete-256-multinomial | evaluate | 720,798 | 120,931 | 3,072 | 2,160 | 1,077,456 | 364,072 | 0 | 3,472 |
| nested-discrete-256-multinomial | parse-evaluate | 722,188 | 120,931 | 3,072 | 2,194 | 1,085,404 | 366,702 | 0 | 3,644 |
| nested-discrete-64-gcd | evaluate | 136,710 | 3,131 | 768 | 1,232 | 557,472 | 93,736 | 0 | 3,292 |
| nested-discrete-64-gcd | parse-evaluate | 138,163 | 3,131 | 768 | 1,300 | 573,324 | 96,355 | 0 | 3,356 |
| nested-discrete-64-lcm | evaluate | 132,898 | 3,131 | 768 | 1,232 | 557,472 | 93,736 | 0 | 3,364 |
| nested-discrete-64-lcm | parse-evaluate | 134,100 | 3,131 | 768 | 1,300 | 573,324 | 96,355 | 0 | 3,288 |
| nested-discrete-64-multinomial | evaluate | 188,465 | 30,307 | 768 | 1,232 | 557,472 | 93,736 | 0 | 3,356 |
| nested-discrete-64-multinomial | parse-evaluate | 188,883 | 30,307 | 768 | 1,300 | 573,356 | 96,363 | 0 | 3,388 |
| reference-discrete-1024x4-gcd | evaluate | 447,032 | 16,399 | 8,192 | 12 | 1,944 | 1,880 | 0 | 3,288 |
| reference-discrete-1024x4-gcd | parse-evaluate | 446,252 | 16,399 | 8,192 | 21 | 3,323 | 3,227 | 0 | 3,284 |
| reference-discrete-1024x4-lcm | evaluate | 386,292 | 16,399 | 8,192 | 12 | 1,944 | 1,880 | 0 | 3,212 |
| reference-discrete-1024x4-lcm | parse-evaluate | 389,192 | 16,399 | 8,192 | 21 | 3,323 | 3,227 | 0 | 3,284 |
| reference-discrete-1024x4-multinomial | evaluate | 1,075,885 | 450,615 | 8,192 | 12 | 1,944 | 1,880 | 0 | 3,248 |
| reference-discrete-1024x4-multinomial | parse-evaluate | 1,084,905 | 450,615 | 8,192 | 21 | 3,331 | 3,235 | 0 | 3,256 |
| reference-discrete-256x4-gcd | evaluate | 111,716 | 4,111 | 2,048 | 24 | 3,888 | 1,880 | 0 | 3,108 |
| reference-discrete-256x4-gcd | parse-evaluate | 112,770 | 4,111 | 2,048 | 42 | 6,642 | 3,225 | 0 | 3,032 |
| reference-discrete-256x4-lcm | evaluate | 98,860 | 4,111 | 2,048 | 24 | 3,888 | 1,880 | 0 | 3,016 |
| reference-discrete-256x4-lcm | parse-evaluate | 99,465 | 4,111 | 2,048 | 42 | 6,642 | 3,225 | 0 | 3,032 |
| reference-discrete-256x4-multinomial | evaluate | 280,696 | 112,695 | 2,048 | 24 | 3,888 | 1,880 | 0 | 3,020 |
| reference-discrete-256x4-multinomial | parse-evaluate | 283,771 | 112,695 | 2,048 | 42 | 6,658 | 3,233 | 0 | 2,956 |
| reference-discrete-64x4-gcd | evaluate | 29,585 | 1,039 | 512 | 48 | 7,776 | 1,880 | 0 | 3,044 |
| reference-discrete-64x4-gcd | parse-evaluate | 29,922 | 1,039 | 512 | 84 | 13,276 | 3,223 | 0 | 2,964 |
| reference-discrete-64x4-lcm | evaluate | 26,185 | 1,039 | 512 | 48 | 7,776 | 1,880 | 0 | 3,096 |
| reference-discrete-64x4-lcm | parse-evaluate | 26,197 | 1,039 | 512 | 84 | 13,276 | 3,223 | 0 | 2,948 |
| reference-discrete-64x4-multinomial | evaluate | 80,388 | 28,215 | 512 | 48 | 7,776 | 1,880 | 0 | 3,100 |
| reference-discrete-64x4-multinomial | parse-evaluate | 81,130 | 28,215 | 512 | 84 | 13,308 | 3,231 | 0 | 2,980 |
| scalar-discrete-combin | evaluate | 235,008 | 533 | 0 | 5,000 | 496,000 | 448 | 0 | 2,956 |
| scalar-discrete-combin | parse-evaluate | 235,139 | 533 | 0 | 9,000 | 961,000 | 881 | 0 | 2,976 |
| scalar-discrete-combina | evaluate | 1,606 | 20 | 0 | 5,000 | 496,000 | 448 | 0 | 2,992 |
| scalar-discrete-combina | parse-evaluate | 1,829 | 20 | 0 | 9,000 | 960,000 | 880 | 0 | 2,972 |
| scalar-discrete-delta | evaluate | 607 | 13 | 0 | 5,000 | 496,000 | 448 | 0 | 2,980 |
| scalar-discrete-delta | parse-evaluate | 841 | 13 | 0 | 9,000 | 955,000 | 875 | 0 | 2,976 |
| scalar-discrete-even | evaluate | 654 | 13 | 0 | 5,000 | 496,000 | 448 | 0 | 2,984 |
| scalar-discrete-even | parse-evaluate | 864 | 13 | 0 | 9,000 | 955,000 | 875 | 0 | 2,988 |
| scalar-discrete-fact | evaluate | 4,470 | 110 | 0 | 4,000 | 400,000 | 400 | 0 | 2,952 |
| scalar-discrete-fact | parse-evaluate | 4,654 | 110 | 0 | 8,000 | 858,000 | 826 | 0 | 2,964 |
| scalar-discrete-factdouble | evaluate | 2,202 | 67 | 0 | 4,000 | 400,000 | 400 | 0 | 2,964 |
| scalar-discrete-factdouble | parse-evaluate | 2,412 | 67 | 0 | 8,000 | 864,000 | 832 | 0 | 2,968 |
| scalar-discrete-gcd | evaluate | 1,050 | 51 | 0 | 6,000 | 848,000 | 624 | 0 | 2,964 |
| scalar-discrete-gcd | parse-evaluate | 1,268 | 51 | 0 | 10,000 | 1,338,000 | 1,082 | 0 | 2,976 |
| scalar-discrete-gestep | evaluate | 600 | 14 | 0 | 5,000 | 496,000 | 448 | 0 | 2,968 |
| scalar-discrete-gestep | parse-evaluate | 832 | 14 | 0 | 9,000 | 956,000 | 876 | 0 | 2,972 |
| scalar-discrete-lcm | evaluate | 53,130 | 6,168 | 0 | 6,000 | 848,000 | 624 | 0 | 2,964 |
| scalar-discrete-lcm | parse-evaluate | 53,287 | 6,168 | 0 | 10,000 | 1,311,000 | 1,055 | 0 | 2,980 |
| scalar-discrete-multinomial | evaluate | 5,640 | 200 | 0 | 6,000 | 848,000 | 624 | 0 | 2,964 |
| scalar-discrete-multinomial | parse-evaluate | 5,884 | 200 | 0 | 10,000 | 1,316,000 | 1,060 | 0 | 2,932 |
| scalar-discrete-odd | evaluate | 636 | 12 | 0 | 5,000 | 496,000 | 448 | 0 | 2,988 |
| scalar-discrete-odd | parse-evaluate | 834 | 12 | 0 | 9,000 | 954,000 | 874 | 0 | 2,968 |

## Nested projection scaling

The nested rows use `SUM(IF(range+1;GCD(range+0;other);0))`, with the corresponding LCM and MULTINOMIAL forms. Their work and resolver-read columns are retained to expose whether invariant reducer branches are projected once per outer evaluation. This is a bounded evaluator scaling observation, not a whole-workbook recalculation claim.

| case | elements | work/repeat | reference reads | time ns/repeat |
| --- | ---: | ---: | ---: | ---: |
| nested-discrete-64-gcd | 256 | 3,131 | 768 | 136,710 |
| nested-discrete-64-lcm | 256 | 3,131 | 768 | 132,898 |
| nested-discrete-64-multinomial | 256 | 30,307 | 768 | 188,465 |
| nested-discrete-256-gcd | 1024 | 12,347 | 3,072 | 553,512 |
| nested-discrete-256-lcm | 1024 | 12,347 | 3,072 | 536,972 |
| nested-discrete-256-multinomial | 1024 | 120,931 | 3,072 | 720,798 |
| nested-discrete-1024-gcd | 4096 | 49,211 | 12,288 | 2,230,261 |
| nested-discrete-1024-lcm | 4096 | 49,211 | 12,288 | 2,167,271 |
| nested-discrete-1024-multinomial | 4096 | 483,427 | 12,288 | 2,878,675 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed result and does not independently establish libm accuracy. The profile does not measure save, cache publication, native producer acceptance, cold filesystem state, or cross-platform bit identity.

Verification note: the frozen verifier's expected case tuple contains one `gstep` spelling while the harness emits `gestep`. `results/verify_retained.py` applies that single in-memory correction after SHA-checking the frozen verifier; the actual result is retained in `verification-receipt.json`. The DSUM RSS shift is recorded in `verification-notes.md`: +7.0% evaluate and +5.7% parse-evaluate, with unchanged allocation and live-budget metrics and unresolved cause.
