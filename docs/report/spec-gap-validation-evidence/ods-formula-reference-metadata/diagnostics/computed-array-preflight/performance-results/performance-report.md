# ODS reference-metadata evaluator performance profile

The baseline is committed `049c09cdde3978593149079c4257df047a3fa419`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 47,157 | 46,185 | -2.1% | 2,311 | 2,311 | 88 | 88 | 3,712 | 3,716 |
| array-control-16x16-arithmetic | parse-evaluate | 62,815 | 62,022 | -1.3% | 2,311 | 2,311 | 352 | 352 | 3,732 | 3,800 |
| array-control-16x16-sin | evaluate | 128,333 | 127,848 | -0.4% | 2,311 | 2,311 | 2,144 | 2,144 | 4,132 | 4,204 |
| array-control-16x16-sin | parse-evaluate | 144,795 | 143,010 | -1.2% | 2,311 | 2,311 | 2,412 | 2,412 | 4,160 | 4,220 |
| array-control-4x4-arithmetic | evaluate | 4,503 | 4,424 | -1.8% | 151 | 151 | 1,120 | 1,120 | 3,444 | 3,496 |
| array-control-4x4-arithmetic | parse-evaluate | 5,271 | 5,231 | -0.8% | 151 | 151 | 2,240 | 2,240 | 3,464 | 3,560 |
| array-control-4x4-sin | evaluate | 9,948 | 9,842 | -1.1% | 151 | 151 | 3,840 | 3,840 | 3,896 | 3,912 |
| array-control-4x4-sin | parse-evaluate | 10,815 | 10,773 | -0.4% | 151 | 151 | 5,040 | 5,040 | 4,064 | 3,936 |
| concat-borrowed-literals | evaluate | 722 | 714 | -1.1% | 21 | 21 | 6,000 | 6,000 | 3,420 | 3,480 |
| concat-borrowed-literals | parse-evaluate | 864 | 858 | -0.7% | 21 | 21 | 9,000 | 9,000 | 3,420 | 3,484 |
| concat-growth-chain | evaluate | 1,893 | 1,884 | -0.5% | 20 | 20 | 10,000 | 10,000 | 3,420 | 3,424 |
| concat-growth-chain | parse-evaluate | 2,257 | 2,262 | +0.2% | 20 | 20 | 17,000 | 17,000 | 3,364 | 3,440 |
| concat-owned-left | evaluate | 1,084 | 1,076 | -0.7% | 24 | 24 | 8,000 | 8,000 | 3,436 | 3,500 |
| concat-owned-left | parse-evaluate | 1,283 | 1,293 | +0.8% | 24 | 24 | 13,000 | 13,000 | 3,364 | 3,424 |
| concat-owned-right | evaluate | 1,087 | 1,082 | -0.5% | 24 | 24 | 8,000 | 8,000 | 3,408 | 3,476 |
| concat-owned-right | parse-evaluate | 1,361 | 1,354 | -0.5% | 24 | 24 | 13,000 | 13,000 | 3,352 | 3,424 |
| database-control-dstdev | evaluate | 4,040 | 3,990 | -1.2% | 30 | 30 | 20 | 20 | 3,648 | 3,624 |
| database-control-dstdev | parse-evaluate | 4,600 | 4,480 | -2.6% | 30 | 30 | 29 | 29 | 3,712 | 3,704 |
| database-control-dsum | evaluate | 4,110 | 4,060 | -1.2% | 28 | 28 | 20 | 20 | 3,716 | 3,600 |
| database-control-dsum | parse-evaluate | 4,670 | 4,570 | -2.1% | 28 | 28 | 29 | 29 | 3,732 | 3,628 |
| database-control-dvar | evaluate | 4,090 | 3,980 | -2.7% | 28 | 28 | 20 | 20 | 3,748 | 3,756 |
| database-control-dvar | parse-evaluate | 4,670 | 4,490 | -3.9% | 28 | 28 | 29 | 29 | 3,700 | 3,708 |
| literal-aggregate-4x1-sum | evaluate | 1,860 | 1,820 | -2.2% | 15 | 15 | 10 | 10 | 3,648 | 3,616 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,600 | 2,560 | -1.5% | 15 | 15 | 23 | 23 | 3,748 | 3,700 |
| reference-aggregate-64x4-sum | evaluate | 10,337 | 10,337 | +0.0% | 16 | 16 | 32 | 32 | 3,548 | 3,664 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,770 | 10,772 | +0.0% | 16 | 16 | 60 | 60 | 3,660 | 3,672 |
| reference-array-16x4-arithmetic | evaluate | 9,489 | 9,546 | +0.6% | 16 | 16 | 800 | 800 | 3,432 | 3,536 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,882 | 9,896 | +0.1% | 16 | 16 | 1,280 | 1,280 | 3,464 | 3,548 |
| reference-conditional-256x4-sumifs | evaluate | 83,295 | 88,180 | +5.9% | 52 | 52 | 42 | 42 | 3,712 | 3,740 |
| reference-conditional-256x4-sumifs | parse-evaluate | 84,835 | 83,835 | -1.2% | 52 | 52 | 68 | 68 | 3,688 | 3,712 |
| reference-control-average | evaluate | 10,385 | 10,282 | -1.0% | 20 | 20 | 32 | 32 | 3,648 | 3,616 |
| reference-control-average | parse-evaluate | 10,667 | 10,617 | -0.5% | 20 | 20 | 60 | 60 | 3,768 | 3,580 |
| reference-control-counta | evaluate | 8,702 | 8,662 | -0.5% | 19 | 19 | 32 | 32 | 3,652 | 3,624 |
| reference-control-counta | parse-evaluate | 9,177 | 9,170 | -0.1% | 19 | 19 | 60 | 60 | 3,680 | 3,564 |
| representative-median | evaluate | 1,589 | 1,603 | +0.9% | 16 | 16 | 9,000 | 9,000 | 3,524 | 3,544 |
| representative-median | parse-evaluate | 1,914 | 1,974 | +3.1% | 16 | 16 | 14,000 | 14,000 | 3,520 | 3,540 |
| representative-percentrank | evaluate | 979 | 965 | -1.4% | 19 | 19 | 7,000 | 7,000 | 3,512 | 3,520 |
| representative-percentrank | parse-evaluate | 1,227 | 1,224 | -0.2% | 19 | 19 | 11,000 | 11,000 | 3,520 | 3,540 |
| representative-rank | evaluate | 937 | 926 | -1.2% | 12 | 12 | 7,000 | 7,000 | 3,520 | 3,532 |
| representative-rank | parse-evaluate | 1,159 | 1,133 | -2.2% | 12 | 12 | 11,000 | 11,000 | 3,516 | 3,560 |
| scalar-aggregate-sum | evaluate | 575 | 576 | +0.2% | 10 | 10 | 4,000 | 4,000 | 3,500 | 3,512 |
| scalar-aggregate-sum | parse-evaluate | 760 | 767 | +0.9% | 10 | 10 | 8,000 | 8,000 | 3,496 | 3,544 |
| scalar-control-arithmetic | evaluate | 577 | 573 | -0.7% | 10 | 10 | 5,000 | 5,000 | 3,400 | 3,488 |
| scalar-control-arithmetic | parse-evaluate | 712 | 718 | +0.8% | 10 | 10 | 8,000 | 8,000 | 3,340 | 3,484 |
| scalar-control-average | evaluate | 1,026 | 1,017 | -0.9% | 23 | 23 | 6,000 | 6,000 | 3,484 | 3,548 |
| scalar-control-average | parse-evaluate | 1,260 | 1,265 | +0.4% | 23 | 23 | 10,000 | 10,000 | 3,504 | 3,552 |
| scalar-control-counta | evaluate | 838 | 840 | +0.2% | 24 | 24 | 6,000 | 6,000 | 3,460 | 3,524 |
| scalar-control-counta | parse-evaluate | 1,073 | 1,107 | +3.2% | 24 | 24 | 10,000 | 10,000 | 3,492 | 3,544 |
| scalar-control-imsum | evaluate | 1,426 | 1,406 | -1.4% | 41 | 41 | 8,000 | 8,000 | 3,444 | 3,476 |
| scalar-control-imsum | parse-evaluate | 2,029 | 1,998 | -1.5% | 41 | 41 | 17,000 | 17,000 | 3,460 | 3,504 |
| scalar-control-sin | evaluate | 471 | 467 | -0.8% | 10 | 10 | 4,000 | 4,000 | 3,820 | 3,852 |
| scalar-control-sin | parse-evaluate | 652 | 653 | +0.2% | 10 | 10 | 8,000 | 8,000 | 3,836 | 3,848 |
| scalar-control-stdev | evaluate | 865 | 852 | -1.5% | 21 | 21 | 6,000 | 6,000 | 3,500 | 3,528 |
| scalar-control-stdev | parse-evaluate | 1,087 | 1,062 | -2.3% | 21 | 21 | 10,000 | 10,000 | 3,492 | 3,540 |
| scalar-control-var | evaluate | 846 | 844 | -0.2% | 19 | 19 | 6,000 | 6,000 | 3,520 | 3,528 |
| scalar-control-var | parse-evaluate | 1,059 | 1,049 | -0.9% | 19 | 19 | 10,000 | 10,000 | 3,516 | 3,540 |

The only matched-control latency review flag above 5% is `reference-conditional-256x4-sumifs` in the `evaluate` phase: p50 `83,295` to `88,180` ns/repeat, or `+5.865%`. The same case in `parse-evaluate` is `-1.179%`. Across the fifteen child p50s, `evaluate` spans `77,935..91,085` ns/repeat for the baseline and `78,500..92,890` for the candidate; their medians are `83,295` and `88,180`. A simple independent 100,000-resample bootstrap of those fifteen child medians gives a descriptive 95% interval of `[+0.528%, +6.831%]` for the relative median difference. This interval is a variability estimate, not a causal regression claim: work (`1,839`), resolver reads (`1,792`), allocator calls (`42`), and bytes/repeat (`52`) are identical, and the host was not isolated. The pre-capture load average was `[0.91, 1.93, 2.25]` across 32 available CPUs; the retained process snapshot showed no matching cargo, rustc, boundary, or evidence workload.

## Reference-metadata workloads

The candidate matrix covers all eight reference-metadata functions over contiguous and 3-D reference descriptors, ordered reference lists, current coordinates, inline arrays, projected matrix coordinates, scalar refusal, large geometry, typed resource limits, and cancellation. Input/output bytes and the reviewed domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| reference-metadata-areas-3d | evaluate | 29 | 0 | 29 | 1,660 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,708 |
| reference-metadata-areas-3d | parse-evaluate | 29 | 0 | 29 | 2,250 | 16 | 0 | 19 | 5,378 | 5,282 | 0 | 3,616 |
| reference-metadata-areas-arity | evaluate | 8 | 0 | 8 | 460 | 6 | 0 | 4 | 632 | 632 | 0 | 3,476 |
| reference-metadata-areas-arity | parse-evaluate | 8 | 0 | 8 | 610 | 6 | 0 | 6 | 1,024 | 1,024 | 0 | 3,484 |
| reference-metadata-areas-array-refusal | evaluate | 17 | 0 | 17 | 1,870 | 25 | 0 | 11 | 5,016 | 4,392 | 0 | 3,568 |
| reference-metadata-areas-array-refusal | parse-evaluate | 17 | 0 | 17 | 3,100 | 25 | 0 | 20 | 6,409 | 5,241 | 0 | 3,608 |
| reference-metadata-areas-duplicate-list | evaluate | 21 | 0 | 21 | 2,230 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,684 |
| reference-metadata-areas-duplicate-list | parse-evaluate | 21 | 0 | 21 | 3,450 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,712 |
| reference-metadata-areas-list | evaluate | 29 | 0 | 29 | 2,360 | 21 | 0 | 14 | 4,648 | 4,152 | 0 | 3,692 |
| reference-metadata-areas-list | parse-evaluate | 29 | 0 | 29 | 3,110 | 21 | 0 | 24 | 6,793 | 5,881 | 0 | 3,728 |
| reference-metadata-areas-range | evaluate | 18 | 0 | 18 | 1,580 | 14 | 0 | 10 | 3,992 | 3,928 | 0 | 3,600 |
| reference-metadata-areas-range | parse-evaluate | 18 | 0 | 18 | 2,020 | 14 | 0 | 17 | 5,356 | 5,260 | 0 | 3,616 |
| reference-metadata-areas-refusal | evaluate | 10 | 0 | 10 | 1,090 | 14 | 0 | 7 | 3,592 | 3,528 | 0 | 3,620 |
| reference-metadata-areas-refusal | parse-evaluate | 10 | 0 | 10 | 1,280 | 14 | 0 | 11 | 4,050 | 3,954 | 0 | 3,612 |
| reference-metadata-areas-resource | evaluate | 20 | 0 | 20 | 715 | 9 | 0 | 24 | 11,488 | 2,808 | 0 | 3,720 |
| reference-metadata-areas-resource | parse-evaluate | 20 | 0 | 20 | 1,130 | 9 | 0 | 52 | 16,956 | 4,143 | 0 | 3,608 |
| reference-metadata-column-3d-cell | evaluate | 18 | 0 | 18 | 1,540 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,576 |
| reference-metadata-column-3d-cell | parse-evaluate | 18 | 0 | 18 | 1,960 | 13 | 0 | 17 | 5,359 | 5,263 | 0 | 3,664 |
| reference-metadata-column-array-refusal | evaluate | 16 | 0 | 16 | 1,740 | 23 | 0 | 11 | 4,928 | 4,304 | 0 | 3,660 |
| reference-metadata-column-array-refusal | parse-evaluate | 16 | 0 | 16 | 2,870 | 23 | 0 | 19 | 6,288 | 5,152 | 0 | 3,712 |
| reference-metadata-column-cell | evaluate | 14 | 0 | 14 | 1,540 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,664 |
| reference-metadata-column-cell | parse-evaluate | 14 | 0 | 14 | 1,900 | 13 | 0 | 16 | 5,351 | 5,255 | 0 | 3,668 |
| reference-metadata-column-current | evaluate | 9 | 0 | 9 | 550 | 9 | 0 | 4 | 632 | 632 | 0 | 3,708 |
| reference-metadata-column-current | parse-evaluate | 9 | 0 | 9 | 710 | 9 | 0 | 6 | 1,025 | 1,025 | 0 | 3,676 |
| reference-metadata-column-list-refusal | evaluate | 22 | 0 | 22 | 2,260 | 18 | 0 | 14 | 4,648 | 4,152 | 0 | 3,652 |
| reference-metadata-column-list-refusal | parse-evaluate | 22 | 0 | 22 | 3,500 | 18 | 0 | 22 | 6,784 | 5,872 | 0 | 3,664 |
| reference-metadata-column-matrix | evaluate | 29 | 0 | 29 | 1,860 | 20 | 0 | 11 | 4,256 | 3,928 | 264 | 3,708 |
| reference-metadata-column-matrix | parse-evaluate | 29 | 0 | 29 | 2,420 | 20 | 0 | 20 | 5,642 | 5,282 | 264 | 3,612 |
| reference-metadata-column-one-entry-list | evaluate | 30 | 0 | 30 | 3,350 | 34 | 0 | 20 | 5,448 | 4,568 | 0 | 3,688 |
| reference-metadata-column-one-entry-list | parse-evaluate | 30 | 0 | 30 | 4,160 | 34 | 0 | 30 | 7,657 | 6,329 | 0 | 3,596 |
| reference-metadata-column-projected | evaluate | 58 | 0 | 58 | 7,550 | 103 | 0 | 41 | 8,744 | 5,336 | 264 | 3,672 |
| reference-metadata-column-projected | parse-evaluate | 58 | 0 | 58 | 8,980 | 103 | 0 | 56 | 12,623 | 7,903 | 264 | 3,612 |
| reference-metadata-column-scalar-vector | evaluate | 18 | 0 | 18 | 1,570 | 15 | 0 | 10 | 3,992 | 3,928 | 0 | 3,656 |
| reference-metadata-column-scalar-vector | parse-evaluate | 18 | 0 | 18 | 2,000 | 15 | 0 | 17 | 5,356 | 5,260 | 0 | 3,700 |
| reference-metadata-columns-3d | evaluate | 31 | 0 | 31 | 1,660 | 18 | 0 | 10 | 3,992 | 3,928 | 0 | 3,712 |
| reference-metadata-columns-3d | parse-evaluate | 31 | 0 | 31 | 2,250 | 18 | 0 | 19 | 5,380 | 5,284 | 0 | 3,616 |
| reference-metadata-columns-arity | evaluate | 10 | 0 | 10 | 450 | 8 | 0 | 4 | 632 | 632 | 0 | 3,456 |
| reference-metadata-columns-arity | parse-evaluate | 10 | 0 | 10 | 640 | 8 | 0 | 6 | 1,026 | 1,026 | 0 | 3,492 |
| reference-metadata-columns-array | evaluate | 19 | 0 | 19 | 1,930 | 27 | 0 | 11 | 5,016 | 4,392 | 0 | 3,664 |
| reference-metadata-columns-array | parse-evaluate | 19 | 0 | 19 | 3,000 | 27 | 0 | 20 | 6,411 | 5,243 | 0 | 3,672 |
| reference-metadata-columns-formula-error | evaluate | 14 | 0 | 14 | 1,040 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,624 |
| reference-metadata-columns-formula-error | parse-evaluate | 14 | 0 | 14 | 1,300 | 18 | 0 | 11 | 4,054 | 3,958 | 0 | 3,588 |
| reference-metadata-columns-list-refusal | evaluate | 23 | 0 | 23 | 2,290 | 19 | 0 | 14 | 4,648 | 4,152 | 0 | 3,704 |
| reference-metadata-columns-list-refusal | parse-evaluate | 23 | 0 | 23 | 3,510 | 19 | 0 | 22 | 6,785 | 5,873 | 0 | 3,688 |
| reference-metadata-columns-one-entry-list | evaluate | 31 | 0 | 31 | 3,350 | 35 | 0 | 20 | 5,448 | 4,568 | 0 | 3,704 |
| reference-metadata-columns-one-entry-list | parse-evaluate | 31 | 0 | 31 | 4,160 | 35 | 0 | 30 | 7,658 | 6,330 | 0 | 3,672 |
| reference-metadata-columns-projected | evaluate | 41 | 0 | 41 | 4,420 | 50 | 0 | 27 | 5,192 | 4,648 | 176 | 3,640 |
| reference-metadata-columns-projected | parse-evaluate | 41 | 0 | 41 | 5,280 | 50 | 0 | 41 | 9,075 | 7,187 | 176 | 3,760 |
| reference-metadata-columns-range | evaluate | 20 | 0 | 20 | 1,560 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,708 |
| reference-metadata-columns-range | parse-evaluate | 20 | 0 | 20 | 2,040 | 16 | 0 | 17 | 5,358 | 5,262 | 0 | 3,736 |
| reference-metadata-columns-scalar-refusal | evaluate | 12 | 0 | 12 | 1,070 | 16 | 0 | 7 | 3,592 | 3,528 | 0 | 3,700 |
| reference-metadata-columns-scalar-refusal | parse-evaluate | 12 | 0 | 12 | 1,320 | 16 | 0 | 11 | 4,052 | 3,956 | 0 | 3,708 |
| reference-metadata-isref-arity | evaluate | 8 | 0 | 8 | 460 | 6 | 0 | 4 | 632 | 632 | 0 | 3,504 |
| reference-metadata-isref-arity | parse-evaluate | 8 | 0 | 8 | 620 | 6 | 0 | 6 | 1,024 | 1,024 | 0 | 3,484 |
| reference-metadata-isref-array | evaluate | 15 | 0 | 15 | 1,750 | 22 | 0 | 11 | 4,928 | 4,304 | 0 | 3,684 |
| reference-metadata-isref-array | parse-evaluate | 15 | 0 | 15 | 2,860 | 22 | 0 | 21 | 6,351 | 5,151 | 0 | 3,708 |
| reference-metadata-isref-error | evaluate | 12 | 0 | 12 | 1,070 | 16 | 0 | 7 | 3,592 | 3,528 | 0 | 3,712 |
| reference-metadata-isref-error | parse-evaluate | 12 | 0 | 12 | 1,280 | 16 | 0 | 11 | 4,052 | 3,956 | 0 | 3,736 |
| reference-metadata-isref-iferror-source | evaluate | 48 | 0 | 48 | 1,150 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,616 |
| reference-metadata-isref-iferror-source | parse-evaluate | 48 | 0 | 48 | 1,810 | 21 | 0 | 15 | 5,033 | 4,905 | 0 | 3,692 |
| reference-metadata-isref-iferror-source-arithmetic | evaluate | 50 | 0 | 50 | 880 | 18 | 0 | 6 | 3,080 | 2,888 | 0 | 3,772 |
| reference-metadata-isref-iferror-source-arithmetic | parse-evaluate | 50 | 0 | 50 | 1,700 | 18 | 0 | 16 | 5,355 | 4,683 | 0 | 3,604 |
| reference-metadata-isref-iferror-source-fallback | evaluate | 45 | 0 | 45 | 1,580 | 27 | 0 | 8 | 3,848 | 3,656 | 0 | 3,640 |
| reference-metadata-isref-iferror-source-fallback | parse-evaluate | 45 | 0 | 45 | 2,460 | 27 | 0 | 18 | 6,118 | 5,446 | 0 | 3,696 |
| reference-metadata-isref-ifna-source | evaluate | 45 | 0 | 45 | 1,160 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,608 |
| reference-metadata-isref-ifna-source | parse-evaluate | 45 | 0 | 45 | 1,770 | 18 | 0 | 15 | 5,030 | 4,902 | 0 | 3,712 |
| reference-metadata-isref-ifna-source-arithmetic | evaluate | 47 | 0 | 47 | 910 | 15 | 0 | 6 | 3,080 | 2,888 | 0 | 3,664 |
| reference-metadata-isref-ifna-source-arithmetic | parse-evaluate | 47 | 0 | 47 | 1,680 | 15 | 0 | 16 | 5,352 | 4,680 | 0 | 3,664 |
| reference-metadata-isref-ifna-source-fallback | evaluate | 43 | 0 | 43 | 1,390 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-isref-ifna-source-fallback | parse-evaluate | 43 | 0 | 43 | 2,030 | 23 | 0 | 15 | 5,028 | 4,900 | 0 | 3,612 |
| reference-metadata-isref-list | evaluate | 21 | 0 | 21 | 2,210 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,612 |
| reference-metadata-isref-list | parse-evaluate | 21 | 0 | 21 | 3,490 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,708 |
| reference-metadata-isref-number | evaluate | 10 | 0 | 10 | 1,070 | 14 | 0 | 7 | 3,592 | 3,528 | 0 | 3,588 |
| reference-metadata-isref-number | parse-evaluate | 10 | 0 | 10 | 1,320 | 14 | 0 | 11 | 4,050 | 3,954 | 0 | 3,688 |
| reference-metadata-isref-reference | evaluate | 18 | 0 | 18 | 1,590 | 14 | 0 | 10 | 3,992 | 3,928 | 0 | 3,620 |
| reference-metadata-isref-reference | parse-evaluate | 18 | 0 | 18 | 2,020 | 14 | 0 | 17 | 5,356 | 5,260 | 0 | 3,672 |
| reference-metadata-isref-selected-local | evaluate | 50 | 0 | 50 | 1,840 | 24 | 0 | 10 | 3,992 | 3,928 | 0 | 3,612 |
| reference-metadata-isref-selected-local | parse-evaluate | 50 | 0 | 50 | 2,720 | 24 | 0 | 20 | 6,204 | 5,692 | 0 | 3,756 |
| reference-metadata-isref-selected-source | evaluate | 49 | 0 | 49 | 1,360 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,612 |
| reference-metadata-isref-selected-source | parse-evaluate | 49 | 0 | 49 | 2,210 | 23 | 0 | 17 | 5,803 | 5,291 | 0 | 3,776 |
| reference-metadata-isref-source | evaluate | 32 | 0 | 32 | 1,080 | 12 | 0 | 7 | 3,592 | 3,528 | 0 | 3,712 |
| reference-metadata-isref-source | parse-evaluate | 32 | 0 | 32 | 1,641 | 12 | 0 | 14 | 4,985 | 4,889 | 0 | 3,612 |
| reference-metadata-isref-source-array-false-matrix | evaluate | 60 | 0 | 60 | 3,210 | 40 | 0 | 20 | 4,440 | 4,040 | 0 | 3,744 |
| reference-metadata-isref-source-array-false-matrix | parse-evaluate | 60 | 0 | 60 | 4,290 | 40 | 0 | 34 | 8,357 | 6,613 | 0 | 3,648 |
| reference-metadata-isref-source-array-true-matrix | evaluate | 60 | 0 | 60 | 4,330 | 53 | 0 | 25 | 5,608 | 4,760 | 0 | 3,704 |
| reference-metadata-isref-source-array-true-matrix | parse-evaluate | 60 | 0 | 60 | 5,430 | 53 | 0 | 39 | 9,525 | 7,333 | 0 | 3,740 |
| reference-metadata-large-range | evaluate | 19 | 0 | 19 | 1,600 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,676 |
| reference-metadata-large-range | parse-evaluate | 19 | 0 | 19 | 1,960 | 13 | 0 | 17 | 5,358 | 5,262 | 0 | 3,604 |
| reference-metadata-row-3d-cell | evaluate | 15 | 0 | 15 | 1,530 | 10 | 0 | 10 | 3,992 | 3,928 | 0 | 3,660 |
| reference-metadata-row-3d-cell | parse-evaluate | 15 | 0 | 15 | 1,950 | 10 | 0 | 17 | 5,356 | 5,260 | 0 | 3,744 |
| reference-metadata-row-array-refusal | evaluate | 13 | 0 | 13 | 1,740 | 20 | 0 | 11 | 4,928 | 4,304 | 0 | 3,744 |
| reference-metadata-row-array-refusal | parse-evaluate | 13 | 0 | 13 | 2,870 | 20 | 0 | 19 | 6,285 | 5,149 | 0 | 3,668 |
| reference-metadata-row-cell | evaluate | 11 | 0 | 11 | 1,540 | 10 | 0 | 10 | 3,992 | 3,928 | 0 | 3,752 |
| reference-metadata-row-cell | parse-evaluate | 11 | 0 | 11 | 1,900 | 10 | 0 | 16 | 5,348 | 5,252 | 0 | 3,656 |
| reference-metadata-row-current | evaluate | 6 | 0 | 6 | 530 | 6 | 0 | 4 | 632 | 632 | 0 | 3,744 |
| reference-metadata-row-current | parse-evaluate | 6 | 0 | 6 | 680 | 6 | 0 | 6 | 1,022 | 1,022 | 0 | 3,664 |
| reference-metadata-row-list-refusal | evaluate | 19 | 0 | 19 | 2,211 | 15 | 0 | 14 | 4,648 | 4,152 | 0 | 3,704 |
| reference-metadata-row-list-refusal | parse-evaluate | 19 | 0 | 19 | 3,480 | 15 | 0 | 22 | 6,781 | 5,869 | 0 | 3,604 |
| reference-metadata-row-matrix | evaluate | 26 | 0 | 26 | 1,840 | 17 | 0 | 11 | 4,256 | 3,928 | 264 | 3,660 |
| reference-metadata-row-matrix | parse-evaluate | 26 | 0 | 26 | 2,400 | 17 | 0 | 20 | 5,639 | 5,279 | 264 | 3,636 |
| reference-metadata-row-one-entry-list | evaluate | 27 | 0 | 27 | 3,370 | 31 | 0 | 20 | 5,448 | 4,568 | 0 | 3,604 |
| reference-metadata-row-one-entry-list | parse-evaluate | 27 | 0 | 27 | 4,150 | 31 | 0 | 30 | 7,654 | 6,326 | 0 | 3,600 |
| reference-metadata-row-projected | evaluate | 55 | 0 | 55 | 7,560 | 94 | 0 | 41 | 8,744 | 5,336 | 264 | 3,620 |
| reference-metadata-row-projected | parse-evaluate | 55 | 0 | 55 | 8,990 | 94 | 0 | 59 | 12,812 | 7,964 | 264 | 3,752 |
| reference-metadata-row-scalar-vector | evaluate | 15 | 0 | 15 | 1,520 | 12 | 0 | 10 | 3,992 | 3,928 | 0 | 3,700 |
| reference-metadata-row-scalar-vector | parse-evaluate | 15 | 0 | 15 | 2,000 | 12 | 0 | 17 | 5,353 | 5,257 | 0 | 3,584 |
| reference-metadata-rows-3d | evaluate | 28 | 0 | 28 | 1,680 | 15 | 0 | 10 | 3,992 | 3,928 | 0 | 3,664 |
| reference-metadata-rows-3d | parse-evaluate | 28 | 0 | 28 | 2,230 | 15 | 0 | 19 | 5,377 | 5,281 | 0 | 3,620 |
| reference-metadata-rows-array | evaluate | 16 | 0 | 16 | 1,810 | 24 | 0 | 11 | 5,016 | 4,392 | 0 | 3,664 |
| reference-metadata-rows-array | parse-evaluate | 16 | 0 | 16 | 3,110 | 24 | 0 | 20 | 6,408 | 5,240 | 0 | 3,640 |
| reference-metadata-rows-list-refusal | evaluate | 20 | 0 | 20 | 2,250 | 16 | 0 | 14 | 4,648 | 4,152 | 0 | 3,696 |
| reference-metadata-rows-list-refusal | parse-evaluate | 20 | 0 | 20 | 3,540 | 16 | 0 | 22 | 6,782 | 5,870 | 0 | 3,744 |
| reference-metadata-rows-one-entry-list | evaluate | 28 | 0 | 28 | 3,310 | 32 | 0 | 20 | 5,448 | 4,568 | 0 | 3,632 |
| reference-metadata-rows-one-entry-list | parse-evaluate | 28 | 0 | 28 | 4,080 | 32 | 0 | 30 | 7,655 | 6,327 | 0 | 3,732 |
| reference-metadata-rows-projected | evaluate | 38 | 0 | 38 | 4,430 | 47 | 0 | 27 | 5,192 | 4,648 | 176 | 3,668 |
| reference-metadata-rows-projected | parse-evaluate | 38 | 0 | 38 | 5,330 | 47 | 0 | 41 | 9,072 | 7,184 | 176 | 3,648 |
| reference-metadata-rows-range | evaluate | 17 | 0 | 17 | 1,540 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,584 |
| reference-metadata-rows-range | parse-evaluate | 17 | 0 | 17 | 2,010 | 13 | 0 | 17 | 5,355 | 5,259 | 0 | 3,604 |
| reference-metadata-rows-resource | evaluate | 19 | 0 | 19 | 710 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,708 |
| reference-metadata-rows-resource | parse-evaluate | 19 | 0 | 19 | 1,142 | 8 | 0 | 52 | 16,952 | 4,142 | 0 | 3,760 |
| reference-metadata-rows-scalar-refusal | evaluate | 9 | 0 | 9 | 1,070 | 13 | 0 | 7 | 3,592 | 3,528 | 0 | 3,736 |
| reference-metadata-rows-scalar-refusal | parse-evaluate | 9 | 0 | 9 | 1,250 | 13 | 0 | 11 | 4,049 | 3,953 | 0 | 3,696 |
| reference-metadata-sheet-3d | evaluate | 29 | 0 | 29 | 1,690 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,664 |
| reference-metadata-sheet-3d | parse-evaluate | 29 | 0 | 29 | 2,280 | 16 | 0 | 19 | 5,378 | 5,282 | 0 | 3,688 |
| reference-metadata-sheet-array-matrix | evaluate | 26 | 0 | 26 | 1,850 | 43 | 0 | 11 | 4,248 | 4,008 | 176 | 3,708 |
| reference-metadata-sheet-array-matrix | parse-evaluate | 26 | 0 | 26 | 2,370 | 43 | 0 | 18 | 5,554 | 4,834 | 176 | 3,572 |
| reference-metadata-sheet-array-refusal | evaluate | 15 | 0 | 15 | 2,000 | 25 | 0 | 12 | 4,929 | 4,304 | 0 | 3,612 |
| reference-metadata-sheet-array-refusal | parse-evaluate | 15 | 0 | 15 | 3,160 | 25 | 0 | 20 | 6,288 | 5,151 | 0 | 3,740 |
| reference-metadata-sheet-cancel | evaluate | 17 | 0 | 17 | 15 | 0 | 0 | 0 | 0 | 0 | 0 | 3,176 |
| reference-metadata-sheet-cancel | parse-evaluate | 17 | 0 | 17 | 390 | 0 | 0 | 28 | 5,464 | 1,366 | 0 | 3,192 |
| reference-metadata-sheet-computed-scalar | evaluate | 22 | 0 | 22 | 2,640 | 44 | 1 | 15 | 5,131 | 4,312 | 0 | 3,744 |
| reference-metadata-sheet-computed-scalar | parse-evaluate | 22 | 0 | 22 | 3,150 | 44 | 1 | 23 | 6,531 | 5,648 | 0 | 3,596 |
| reference-metadata-sheet-current | evaluate | 8 | 0 | 8 | 560 | 9 | 0 | 4 | 632 | 632 | 0 | 3,712 |
| reference-metadata-sheet-current | parse-evaluate | 8 | 0 | 8 | 740 | 9 | 0 | 6 | 1,024 | 1,024 | 0 | 3,612 |
| reference-metadata-sheet-hidden-text | evaluate | 16 | 0 | 16 | 1,070 | 25 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-sheet-hidden-text | parse-evaluate | 16 | 0 | 16 | 1,310 | 25 | 0 | 11 | 4,056 | 3,960 | 0 | 3,604 |
| reference-metadata-sheet-iferror-source | evaluate | 48 | 0 | 48 | 1,190 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,692 |
| reference-metadata-sheet-iferror-source | parse-evaluate | 48 | 0 | 48 | 1,810 | 21 | 0 | 15 | 5,033 | 4,905 | 0 | 3,688 |
| reference-metadata-sheet-iferror-source-fallback | evaluate | 45 | 0 | 45 | 1,570 | 27 | 0 | 8 | 3,848 | 3,656 | 0 | 3,752 |
| reference-metadata-sheet-iferror-source-fallback | parse-evaluate | 45 | 0 | 45 | 2,420 | 27 | 0 | 18 | 6,118 | 5,446 | 0 | 3,668 |
| reference-metadata-sheet-ifna-source | evaluate | 45 | 0 | 45 | 1,150 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,624 |
| reference-metadata-sheet-ifna-source | parse-evaluate | 45 | 0 | 45 | 1,760 | 18 | 0 | 15 | 5,030 | 4,902 | 0 | 3,592 |
| reference-metadata-sheet-ifna-source-fallback | evaluate | 43 | 0 | 43 | 1,390 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,708 |
| reference-metadata-sheet-ifna-source-fallback | parse-evaluate | 43 | 0 | 43 | 2,020 | 23 | 0 | 15 | 5,028 | 4,900 | 0 | 3,740 |
| reference-metadata-sheet-list-refusal | evaluate | 21 | 0 | 21 | 2,230 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,696 |
| reference-metadata-sheet-list-refusal | parse-evaluate | 21 | 0 | 21 | 3,570 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,700 |
| reference-metadata-sheet-logical-coercion | evaluate | 14 | 0 | 14 | 1,230 | 19 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-sheet-logical-coercion | parse-evaluate | 14 | 0 | 14 | 1,490 | 19 | 0 | 11 | 4,054 | 3,958 | 0 | 3,776 |
| reference-metadata-sheet-number-coercion | evaluate | 9 | 0 | 9 | 1,300 | 16 | 0 | 8 | 3,593 | 3,529 | 0 | 3,704 |
| reference-metadata-sheet-number-coercion | parse-evaluate | 9 | 0 | 9 | 1,530 | 16 | 0 | 12 | 4,050 | 3,954 | 0 | 3,620 |
| reference-metadata-sheet-reference | evaluate | 17 | 0 | 17 | 1,550 | 12 | 0 | 10 | 3,992 | 3,928 | 0 | 3,632 |
| reference-metadata-sheet-reference | parse-evaluate | 17 | 0 | 17 | 2,000 | 12 | 0 | 17 | 5,358 | 5,262 | 0 | 3,580 |
| reference-metadata-sheet-selected-local | evaluate | 54 | 0 | 54 | 1,850 | 24 | 0 | 10 | 3,992 | 3,928 | 0 | 3,616 |
| reference-metadata-sheet-selected-local | parse-evaluate | 54 | 0 | 54 | 2,830 | 24 | 0 | 21 | 6,212 | 5,700 | 0 | 3,576 |
| reference-metadata-sheet-selected-source | evaluate | 53 | 0 | 53 | 1,330 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,676 |
| reference-metadata-sheet-selected-source | parse-evaluate | 53 | 0 | 53 | 2,400 | 23 | 0 | 18 | 5,811 | 5,299 | 0 | 3,612 |
| reference-metadata-sheet-source | evaluate | 32 | 0 | 32 | 1,070 | 12 | 0 | 7 | 3,592 | 3,528 | 0 | 3,744 |
| reference-metadata-sheet-source | parse-evaluate | 32 | 0 | 32 | 1,620 | 12 | 0 | 14 | 4,985 | 4,889 | 0 | 3,676 |
| reference-metadata-sheet-source-array-false-matrix | evaluate | 60 | 0 | 60 | 3,240 | 40 | 0 | 20 | 4,440 | 4,040 | 0 | 3,764 |
| reference-metadata-sheet-source-array-false-matrix | parse-evaluate | 60 | 0 | 60 | 4,290 | 40 | 0 | 34 | 8,357 | 6,613 | 0 | 3,656 |
| reference-metadata-sheet-source-array-true-matrix | evaluate | 60 | 0 | 60 | 4,420 | 53 | 0 | 25 | 5,608 | 4,760 | 0 | 3,788 |
| reference-metadata-sheet-source-array-true-matrix | parse-evaluate | 60 | 0 | 60 | 5,440 | 53 | 0 | 39 | 9,525 | 7,333 | 0 | 3,756 |
| reference-metadata-sheet-text | evaluate | 14 | 0 | 14 | 1,110 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,664 |
| reference-metadata-sheet-text | parse-evaluate | 14 | 0 | 14 | 1,330 | 21 | 0 | 11 | 4,054 | 3,958 | 0 | 3,612 |
| reference-metadata-sheet-unknown-text | evaluate | 17 | 0 | 17 | 1,110 | 27 | 0 | 7 | 3,592 | 3,528 | 0 | 3,708 |
| reference-metadata-sheet-unknown-text | parse-evaluate | 17 | 0 | 17 | 1,330 | 27 | 0 | 11 | 4,057 | 3,961 | 0 | 3,660 |
| reference-metadata-sheets-3d | evaluate | 30 | 0 | 30 | 1,830 | 19 | 0 | 11 | 4,216 | 4,040 | 0 | 3,596 |
| reference-metadata-sheets-3d | parse-evaluate | 30 | 0 | 30 | 2,510 | 19 | 0 | 20 | 5,603 | 5,395 | 0 | 3,712 |
| reference-metadata-sheets-array-refusal | evaluate | 18 | 0 | 18 | 1,900 | 26 | 0 | 11 | 5,016 | 4,392 | 0 | 3,688 |
| reference-metadata-sheets-array-refusal | parse-evaluate | 18 | 0 | 18 | 3,140 | 26 | 0 | 20 | 6,410 | 5,242 | 0 | 3,668 |
| reference-metadata-sheets-current | evaluate | 9 | 0 | 9 | 570 | 10 | 0 | 4 | 632 | 632 | 0 | 3,668 |
| reference-metadata-sheets-current | parse-evaluate | 9 | 0 | 9 | 700 | 10 | 0 | 6 | 1,025 | 1,025 | 0 | 3,672 |
| reference-metadata-sheets-formula-error | evaluate | 14 | 0 | 14 | 1,060 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,704 |
| reference-metadata-sheets-formula-error | parse-evaluate | 14 | 0 | 14 | 1,300 | 18 | 0 | 11 | 4,054 | 3,958 | 0 | 3,624 |
| reference-metadata-sheets-iferror-source | evaluate | 49 | 0 | 49 | 1,150 | 22 | 0 | 7 | 3,592 | 3,528 | 0 | 3,620 |
| reference-metadata-sheets-iferror-source | parse-evaluate | 49 | 0 | 49 | 1,800 | 22 | 0 | 15 | 5,034 | 4,906 | 0 | 3,684 |
| reference-metadata-sheets-iferror-source-fallback | evaluate | 46 | 0 | 46 | 1,570 | 28 | 0 | 8 | 3,848 | 3,656 | 0 | 3,708 |
| reference-metadata-sheets-iferror-source-fallback | parse-evaluate | 46 | 0 | 46 | 2,400 | 28 | 0 | 18 | 6,119 | 5,447 | 0 | 3,708 |
| reference-metadata-sheets-ifna-source | evaluate | 46 | 0 | 46 | 1,170 | 19 | 0 | 7 | 3,592 | 3,528 | 0 | 3,648 |
| reference-metadata-sheets-ifna-source | parse-evaluate | 46 | 0 | 46 | 1,810 | 19 | 0 | 15 | 5,031 | 4,903 | 0 | 3,760 |
| reference-metadata-sheets-ifna-source-fallback | evaluate | 44 | 0 | 44 | 1,350 | 24 | 0 | 7 | 3,592 | 3,528 | 0 | 3,700 |
| reference-metadata-sheets-ifna-source-fallback | parse-evaluate | 44 | 0 | 44 | 2,010 | 24 | 0 | 15 | 5,029 | 4,901 | 0 | 3,684 |
| reference-metadata-sheets-list-refusal | evaluate | 22 | 0 | 22 | 2,210 | 18 | 0 | 14 | 4,648 | 4,152 | 0 | 3,756 |
| reference-metadata-sheets-list-refusal | parse-evaluate | 22 | 0 | 22 | 3,670 | 18 | 0 | 22 | 6,784 | 5,872 | 0 | 3,708 |
| reference-metadata-sheets-one-entry-list | evaluate | 30 | 0 | 30 | 3,370 | 34 | 0 | 20 | 5,448 | 4,568 | 0 | 3,660 |
| reference-metadata-sheets-one-entry-list | parse-evaluate | 30 | 0 | 30 | 4,150 | 34 | 0 | 30 | 7,657 | 6,329 | 0 | 3,660 |
| reference-metadata-sheets-reference | evaluate | 18 | 0 | 18 | 1,560 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,620 |
| reference-metadata-sheets-reference | parse-evaluate | 18 | 0 | 18 | 1,980 | 13 | 0 | 17 | 5,359 | 5,263 | 0 | 3,656 |
| reference-metadata-sheets-selected-local | evaluate | 55 | 0 | 55 | 1,840 | 25 | 0 | 10 | 3,992 | 3,928 | 0 | 3,672 |
| reference-metadata-sheets-selected-local | parse-evaluate | 55 | 0 | 55 | 2,850 | 25 | 0 | 21 | 6,213 | 5,701 | 0 | 3,708 |
| reference-metadata-sheets-selected-source | evaluate | 54 | 0 | 54 | 1,330 | 24 | 0 | 7 | 3,592 | 3,528 | 0 | 3,708 |
| reference-metadata-sheets-selected-source | parse-evaluate | 54 | 0 | 54 | 2,410 | 24 | 0 | 18 | 5,812 | 5,300 | 0 | 3,660 |
| reference-metadata-sheets-source | evaluate | 33 | 0 | 33 | 1,040 | 13 | 0 | 7 | 3,592 | 3,528 | 0 | 3,740 |
| reference-metadata-sheets-source | parse-evaluate | 33 | 0 | 33 | 1,660 | 13 | 0 | 14 | 4,986 | 4,890 | 0 | 3,744 |
| reference-metadata-sheets-source-array-false-matrix | evaluate | 61 | 0 | 61 | 3,200 | 41 | 0 | 20 | 4,440 | 4,040 | 0 | 3,668 |
| reference-metadata-sheets-source-array-false-matrix | parse-evaluate | 61 | 0 | 61 | 4,290 | 41 | 0 | 34 | 8,358 | 6,614 | 0 | 3,600 |
| reference-metadata-sheets-source-array-true-matrix | evaluate | 61 | 0 | 61 | 4,300 | 54 | 0 | 25 | 5,608 | 4,760 | 0 | 3,620 |
| reference-metadata-sheets-source-array-true-matrix | parse-evaluate | 61 | 0 | 61 | 5,440 | 54 | 0 | 39 | 9,526 | 7,334 | 0 | 3,688 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
