# ODS reference-metadata evaluator performance profile

The baseline is committed `049c09cdde3978593149079c4257df047a3fa419`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 46,730 | 46,675 | -0.1% | 2,311 | 2,311 | 88 | 88 | 3,740 | 3,600 |
| array-control-16x16-arithmetic | parse-evaluate | 61,705 | 61,317 | -0.6% | 2,311 | 2,311 | 352 | 352 | 3,748 | 3,760 |
| array-control-16x16-sin | evaluate | 128,495 | 128,293 | -0.2% | 2,311 | 2,311 | 2,144 | 2,144 | 4,136 | 4,224 |
| array-control-16x16-sin | parse-evaluate | 144,665 | 144,105 | -0.4% | 2,311 | 2,311 | 2,412 | 2,412 | 4,180 | 4,232 |
| array-control-4x4-arithmetic | evaluate | 4,450 | 4,445 | -0.1% | 151 | 151 | 1,120 | 1,120 | 3,532 | 3,340 |
| array-control-4x4-arithmetic | parse-evaluate | 5,213 | 5,241 | +0.5% | 151 | 151 | 2,240 | 2,240 | 3,496 | 3,512 |
| array-control-4x4-sin | evaluate | 9,825 | 9,853 | +0.3% | 151 | 151 | 3,840 | 3,840 | 3,936 | 3,960 |
| array-control-4x4-sin | parse-evaluate | 10,715 | 10,739 | +0.2% | 151 | 151 | 5,040 | 5,040 | 4,060 | 4,196 |
| concat-borrowed-literals | evaluate | 718 | 721 | +0.4% | 21 | 21 | 6,000 | 6,000 | 3,320 | 3,364 |
| concat-borrowed-literals | parse-evaluate | 853 | 862 | +1.1% | 21 | 21 | 9,000 | 9,000 | 3,328 | 3,376 |
| concat-growth-chain | evaluate | 1,895 | 1,896 | +0.1% | 20 | 20 | 10,000 | 10,000 | 3,316 | 3,364 |
| concat-growth-chain | parse-evaluate | 2,270 | 2,262 | -0.4% | 20 | 20 | 17,000 | 17,000 | 3,292 | 3,316 |
| concat-owned-left | evaluate | 1,068 | 1,071 | +0.3% | 24 | 24 | 8,000 | 8,000 | 3,344 | 3,392 |
| concat-owned-left | parse-evaluate | 1,305 | 1,276 | -2.2% | 24 | 24 | 13,000 | 13,000 | 3,328 | 3,444 |
| concat-owned-right | evaluate | 1,085 | 1,083 | -0.2% | 24 | 24 | 8,000 | 8,000 | 3,332 | 3,324 |
| concat-owned-right | parse-evaluate | 1,388 | 1,325 | -4.5% | 24 | 24 | 13,000 | 13,000 | 3,328 | 3,468 |
| database-control-dstdev | evaluate | 3,940 | 3,940 | +0.0% | 30 | 30 | 20 | 20 | 3,700 | 3,632 |
| database-control-dstdev | parse-evaluate | 4,460 | 4,490 | +0.7% | 30 | 30 | 29 | 29 | 3,604 | 3,772 |
| database-control-dsum | evaluate | 3,990 | 4,070 | +2.0% | 28 | 28 | 20 | 20 | 3,568 | 3,684 |
| database-control-dsum | parse-evaluate | 4,600 | 4,680 | +1.7% | 28 | 28 | 29 | 29 | 3,696 | 3,656 |
| database-control-dvar | evaluate | 3,940 | 3,920 | -0.5% | 28 | 28 | 20 | 20 | 3,728 | 3,656 |
| database-control-dvar | parse-evaluate | 4,540 | 4,610 | +1.5% | 28 | 28 | 29 | 29 | 3,612 | 3,652 |
| literal-aggregate-4x1-sum | evaluate | 1,830 | 1,840 | +0.5% | 15 | 15 | 10 | 10 | 3,664 | 3,656 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,560 | 2,540 | -0.8% | 15 | 15 | 23 | 23 | 3,640 | 3,684 |
| reference-aggregate-64x4-sum | evaluate | 10,290 | 10,492 | +2.0% | 16 | 16 | 32 | 32 | 3,552 | 3,716 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,742 | 10,845 | +1.0% | 16 | 16 | 60 | 60 | 3,624 | 3,604 |
| reference-array-16x4-arithmetic | evaluate | 9,481 | 9,532 | +0.5% | 16 | 16 | 800 | 800 | 3,500 | 3,368 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,819 | 9,869 | +0.5% | 16 | 16 | 1,280 | 1,280 | 3,472 | 3,428 |
| reference-conditional-256x4-sumifs | evaluate | 83,060 | 88,280 | +6.3% | 52 | 52 | 42 | 42 | 3,728 | 3,660 |
| reference-conditional-256x4-sumifs | parse-evaluate | 88,555 | 86,915 | -1.9% | 52 | 52 | 68 | 68 | 3,668 | 3,764 |
| reference-control-average | evaluate | 10,325 | 10,342 | +0.2% | 20 | 20 | 32 | 32 | 3,536 | 3,636 |
| reference-control-average | parse-evaluate | 10,757 | 10,630 | -1.2% | 20 | 20 | 60 | 60 | 3,540 | 3,652 |
| reference-control-counta | evaluate | 8,670 | 8,682 | +0.1% | 19 | 19 | 32 | 32 | 3,532 | 3,620 |
| reference-control-counta | parse-evaluate | 9,105 | 9,077 | -0.3% | 19 | 19 | 60 | 60 | 3,544 | 3,644 |
| representative-median | evaluate | 1,605 | 1,597 | -0.5% | 16 | 16 | 9,000 | 9,000 | 3,508 | 3,592 |
| representative-median | parse-evaluate | 1,886 | 1,876 | -0.5% | 16 | 16 | 14,000 | 14,000 | 3,456 | 3,564 |
| representative-percentrank | evaluate | 971 | 964 | -0.7% | 19 | 19 | 7,000 | 7,000 | 3,460 | 3,568 |
| representative-percentrank | parse-evaluate | 1,212 | 1,215 | +0.2% | 19 | 19 | 11,000 | 11,000 | 3,508 | 3,560 |
| representative-rank | evaluate | 919 | 918 | -0.1% | 12 | 12 | 7,000 | 7,000 | 3,476 | 3,564 |
| representative-rank | parse-evaluate | 1,127 | 1,131 | +0.4% | 12 | 12 | 11,000 | 11,000 | 3,464 | 3,528 |
| scalar-aggregate-sum | evaluate | 579 | 569 | -1.7% | 10 | 10 | 4,000 | 4,000 | 3,452 | 3,444 |
| scalar-aggregate-sum | parse-evaluate | 748 | 749 | +0.1% | 10 | 10 | 8,000 | 8,000 | 3,488 | 3,456 |
| scalar-control-arithmetic | evaluate | 570 | 564 | -1.1% | 10 | 10 | 5,000 | 5,000 | 3,408 | 3,484 |
| scalar-control-arithmetic | parse-evaluate | 703 | 699 | -0.6% | 10 | 10 | 8,000 | 8,000 | 3,304 | 3,460 |
| scalar-control-average | evaluate | 1,031 | 1,023 | -0.8% | 23 | 23 | 6,000 | 6,000 | 3,488 | 3,496 |
| scalar-control-average | parse-evaluate | 1,257 | 1,251 | -0.5% | 23 | 23 | 10,000 | 10,000 | 3,444 | 3,476 |
| scalar-control-counta | evaluate | 838 | 825 | -1.6% | 24 | 24 | 6,000 | 6,000 | 3,460 | 3,476 |
| scalar-control-counta | parse-evaluate | 1,083 | 1,073 | -0.9% | 24 | 24 | 10,000 | 10,000 | 3,420 | 3,420 |
| scalar-control-imsum | evaluate | 1,451 | 1,432 | -1.3% | 41 | 41 | 8,000 | 8,000 | 3,380 | 3,476 |
| scalar-control-imsum | parse-evaluate | 1,995 | 1,963 | -1.6% | 41 | 41 | 17,000 | 17,000 | 3,348 | 3,488 |
| scalar-control-sin | evaluate | 463 | 461 | -0.4% | 10 | 10 | 4,000 | 4,000 | 3,776 | 3,868 |
| scalar-control-sin | parse-evaluate | 650 | 649 | -0.2% | 10 | 10 | 8,000 | 8,000 | 3,780 | 3,840 |
| scalar-control-stdev | evaluate | 856 | 854 | -0.2% | 21 | 21 | 6,000 | 6,000 | 3,476 | 3,444 |
| scalar-control-stdev | parse-evaluate | 1,074 | 1,079 | +0.5% | 21 | 21 | 10,000 | 10,000 | 3,420 | 3,504 |
| scalar-control-var | evaluate | 846 | 847 | +0.1% | 19 | 19 | 6,000 | 6,000 | 3,400 | 3,464 |
| scalar-control-var | parse-evaluate | 1,054 | 1,055 | +0.1% | 19 | 19 | 10,000 | 10,000 | 3,368 | 3,472 |

## Reference-metadata workloads

The candidate matrix covers all eight reference-metadata functions over contiguous and 3-D reference descriptors, ordered reference lists, current coordinates, inline arrays, projected matrix coordinates, computed IF/IFERROR/IFNA arrays, scalar refusal, large geometry, typed resource limits, and cancellation. Input/output bytes and the reviewed domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| reference-metadata-areas-3d | evaluate | 29 | 0 | 29 | 1,650 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,712 |
| reference-metadata-areas-3d | parse-evaluate | 29 | 0 | 29 | 2,270 | 16 | 0 | 19 | 5,378 | 5,282 | 0 | 3,740 |
| reference-metadata-areas-arity | evaluate | 8 | 0 | 8 | 470 | 6 | 0 | 4 | 632 | 632 | 0 | 3,516 |
| reference-metadata-areas-arity | parse-evaluate | 8 | 0 | 8 | 630 | 6 | 0 | 6 | 1,024 | 1,024 | 0 | 3,500 |
| reference-metadata-areas-array-refusal | evaluate | 17 | 0 | 17 | 1,910 | 25 | 0 | 11 | 5,016 | 4,392 | 0 | 3,772 |
| reference-metadata-areas-array-refusal | parse-evaluate | 17 | 0 | 17 | 2,490 | 25 | 0 | 20 | 6,409 | 5,241 | 0 | 3,684 |
| reference-metadata-areas-duplicate-list | evaluate | 21 | 0 | 21 | 2,280 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,648 |
| reference-metadata-areas-duplicate-list | parse-evaluate | 21 | 0 | 21 | 2,860 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,628 |
| reference-metadata-areas-list | evaluate | 29 | 0 | 29 | 2,360 | 21 | 0 | 14 | 4,648 | 4,152 | 0 | 3,680 |
| reference-metadata-areas-list | parse-evaluate | 29 | 0 | 29 | 3,130 | 21 | 0 | 24 | 6,793 | 5,881 | 0 | 3,660 |
| reference-metadata-areas-range | evaluate | 18 | 0 | 18 | 1,610 | 14 | 0 | 10 | 3,992 | 3,928 | 0 | 3,676 |
| reference-metadata-areas-range | parse-evaluate | 18 | 0 | 18 | 2,020 | 14 | 0 | 17 | 5,356 | 5,260 | 0 | 3,664 |
| reference-metadata-areas-refusal | evaluate | 10 | 0 | 10 | 1,070 | 14 | 0 | 7 | 3,592 | 3,528 | 0 | 3,612 |
| reference-metadata-areas-refusal | parse-evaluate | 10 | 0 | 10 | 1,320 | 14 | 0 | 11 | 4,050 | 3,954 | 0 | 3,684 |
| reference-metadata-areas-resource | evaluate | 20 | 0 | 20 | 710 | 9 | 0 | 24 | 11,488 | 2,808 | 0 | 3,648 |
| reference-metadata-areas-resource | parse-evaluate | 20 | 0 | 20 | 1,122 | 9 | 0 | 52 | 16,956 | 4,143 | 0 | 3,736 |
| reference-metadata-column-3d-cell | evaluate | 18 | 0 | 18 | 1,550 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,680 |
| reference-metadata-column-3d-cell | parse-evaluate | 18 | 0 | 18 | 1,980 | 13 | 0 | 17 | 5,359 | 5,263 | 0 | 3,620 |
| reference-metadata-column-array-refusal | evaluate | 16 | 0 | 16 | 1,730 | 23 | 0 | 11 | 4,928 | 4,304 | 0 | 3,732 |
| reference-metadata-column-array-refusal | parse-evaluate | 16 | 0 | 16 | 2,370 | 23 | 0 | 19 | 6,288 | 5,152 | 0 | 3,700 |
| reference-metadata-column-cell | evaluate | 14 | 0 | 14 | 1,560 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,604 |
| reference-metadata-column-cell | parse-evaluate | 14 | 0 | 14 | 1,910 | 13 | 0 | 16 | 5,351 | 5,255 | 0 | 3,612 |
| reference-metadata-column-computed-if-matrix | evaluate | 56 | 0 | 56 | 6,850 | 72 | 2 | 41 | 6,872 | 4,744 | 176 | 3,820 |
| reference-metadata-column-computed-if-matrix | parse-evaluate | 56 | 0 | 56 | 8,250 | 72 | 2 | 59 | 10,994 | 7,362 | 176 | 3,736 |
| reference-metadata-column-computed-iferror-matrix | evaluate | 71 | 0 | 71 | 7,680 | 90 | 2 | 46 | 7,240 | 4,744 | 176 | 3,748 |
| reference-metadata-column-computed-iferror-matrix | parse-evaluate | 71 | 0 | 71 | 9,330 | 90 | 2 | 66 | 11,410 | 7,378 | 176 | 3,784 |
| reference-metadata-column-computed-ifna-matrix | evaluate | 68 | 0 | 68 | 7,770 | 87 | 2 | 46 | 7,240 | 4,744 | 176 | 3,732 |
| reference-metadata-column-computed-ifna-matrix | parse-evaluate | 68 | 0 | 68 | 9,350 | 87 | 2 | 66 | 11,407 | 7,375 | 176 | 3,744 |
| reference-metadata-column-current | evaluate | 9 | 0 | 9 | 570 | 9 | 0 | 4 | 632 | 632 | 0 | 3,656 |
| reference-metadata-column-current | parse-evaluate | 9 | 0 | 9 | 710 | 9 | 0 | 6 | 1,025 | 1,025 | 0 | 3,676 |
| reference-metadata-column-list-refusal | evaluate | 22 | 0 | 22 | 2,280 | 18 | 0 | 14 | 4,648 | 4,152 | 0 | 3,684 |
| reference-metadata-column-list-refusal | parse-evaluate | 22 | 0 | 22 | 2,990 | 18 | 0 | 22 | 6,784 | 5,872 | 0 | 3,676 |
| reference-metadata-column-matrix | evaluate | 29 | 0 | 29 | 1,900 | 20 | 0 | 11 | 4,256 | 3,928 | 264 | 3,628 |
| reference-metadata-column-matrix | parse-evaluate | 29 | 0 | 29 | 2,480 | 20 | 0 | 20 | 5,642 | 5,282 | 264 | 3,676 |
| reference-metadata-column-one-entry-list | evaluate | 30 | 0 | 30 | 3,420 | 34 | 0 | 20 | 5,448 | 4,568 | 0 | 3,656 |
| reference-metadata-column-one-entry-list | parse-evaluate | 30 | 0 | 30 | 4,250 | 34 | 0 | 30 | 7,657 | 6,329 | 0 | 3,732 |
| reference-metadata-column-projected | evaluate | 58 | 0 | 58 | 7,600 | 103 | 0 | 41 | 8,744 | 5,336 | 264 | 3,640 |
| reference-metadata-column-projected | parse-evaluate | 58 | 0 | 58 | 9,650 | 103 | 0 | 56 | 12,623 | 7,903 | 264 | 3,724 |
| reference-metadata-column-scalar-vector | evaluate | 18 | 0 | 18 | 1,570 | 15 | 0 | 10 | 3,992 | 3,928 | 0 | 3,620 |
| reference-metadata-column-scalar-vector | parse-evaluate | 18 | 0 | 18 | 2,030 | 15 | 0 | 17 | 5,356 | 5,260 | 0 | 3,632 |
| reference-metadata-columns-3d | evaluate | 31 | 0 | 31 | 1,660 | 18 | 0 | 10 | 3,992 | 3,928 | 0 | 3,632 |
| reference-metadata-columns-3d | parse-evaluate | 31 | 0 | 31 | 2,270 | 18 | 0 | 19 | 5,380 | 5,284 | 0 | 3,624 |
| reference-metadata-columns-arity | evaluate | 10 | 0 | 10 | 470 | 8 | 0 | 4 | 632 | 632 | 0 | 3,500 |
| reference-metadata-columns-arity | parse-evaluate | 10 | 0 | 10 | 640 | 8 | 0 | 6 | 1,026 | 1,026 | 0 | 3,500 |
| reference-metadata-columns-array | evaluate | 19 | 0 | 19 | 1,910 | 27 | 0 | 11 | 5,016 | 4,392 | 0 | 3,632 |
| reference-metadata-columns-array | parse-evaluate | 19 | 0 | 19 | 2,520 | 27 | 0 | 20 | 6,411 | 5,243 | 0 | 3,660 |
| reference-metadata-columns-formula-error | evaluate | 14 | 0 | 14 | 1,050 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,668 |
| reference-metadata-columns-formula-error | parse-evaluate | 14 | 0 | 14 | 1,320 | 18 | 0 | 11 | 4,054 | 3,958 | 0 | 3,664 |
| reference-metadata-columns-list-refusal | evaluate | 23 | 0 | 23 | 2,250 | 19 | 0 | 14 | 4,648 | 4,152 | 0 | 3,632 |
| reference-metadata-columns-list-refusal | parse-evaluate | 23 | 0 | 23 | 2,940 | 19 | 0 | 22 | 6,785 | 5,873 | 0 | 3,684 |
| reference-metadata-columns-one-entry-list | evaluate | 31 | 0 | 31 | 3,430 | 35 | 0 | 20 | 5,448 | 4,568 | 0 | 3,664 |
| reference-metadata-columns-one-entry-list | parse-evaluate | 31 | 0 | 31 | 4,280 | 35 | 0 | 30 | 7,658 | 6,330 | 0 | 3,632 |
| reference-metadata-columns-projected | evaluate | 41 | 0 | 41 | 4,450 | 50 | 0 | 27 | 5,192 | 4,648 | 176 | 3,740 |
| reference-metadata-columns-projected | parse-evaluate | 41 | 0 | 41 | 5,500 | 50 | 0 | 41 | 9,075 | 7,187 | 176 | 3,768 |
| reference-metadata-columns-range | evaluate | 20 | 0 | 20 | 1,610 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,620 |
| reference-metadata-columns-range | parse-evaluate | 20 | 0 | 20 | 2,020 | 16 | 0 | 17 | 5,358 | 5,262 | 0 | 3,708 |
| reference-metadata-columns-scalar-refusal | evaluate | 12 | 0 | 12 | 1,050 | 16 | 0 | 7 | 3,592 | 3,528 | 0 | 3,620 |
| reference-metadata-columns-scalar-refusal | parse-evaluate | 12 | 0 | 12 | 1,280 | 16 | 0 | 11 | 4,052 | 3,956 | 0 | 3,632 |
| reference-metadata-isref-arity | evaluate | 8 | 0 | 8 | 450 | 6 | 0 | 4 | 632 | 632 | 0 | 3,492 |
| reference-metadata-isref-arity | parse-evaluate | 8 | 0 | 8 | 640 | 6 | 0 | 6 | 1,024 | 1,024 | 0 | 3,480 |
| reference-metadata-isref-array | evaluate | 15 | 0 | 15 | 1,770 | 22 | 0 | 11 | 4,928 | 4,304 | 0 | 3,660 |
| reference-metadata-isref-array | parse-evaluate | 15 | 0 | 15 | 2,330 | 22 | 0 | 21 | 6,351 | 5,151 | 0 | 3,644 |
| reference-metadata-isref-error | evaluate | 12 | 0 | 12 | 1,070 | 16 | 0 | 7 | 3,592 | 3,528 | 0 | 3,772 |
| reference-metadata-isref-error | parse-evaluate | 12 | 0 | 12 | 1,290 | 16 | 0 | 11 | 4,052 | 3,956 | 0 | 3,664 |
| reference-metadata-isref-iferror-source | evaluate | 48 | 0 | 48 | 1,140 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,680 |
| reference-metadata-isref-iferror-source | parse-evaluate | 48 | 0 | 48 | 1,820 | 21 | 0 | 15 | 5,033 | 4,905 | 0 | 3,752 |
| reference-metadata-isref-iferror-source-arithmetic | evaluate | 50 | 0 | 50 | 930 | 18 | 0 | 6 | 3,080 | 2,888 | 0 | 3,664 |
| reference-metadata-isref-iferror-source-arithmetic | parse-evaluate | 50 | 0 | 50 | 1,690 | 18 | 0 | 16 | 5,355 | 4,683 | 0 | 3,680 |
| reference-metadata-isref-iferror-source-fallback | evaluate | 45 | 0 | 45 | 1,530 | 27 | 0 | 8 | 3,848 | 3,656 | 0 | 3,684 |
| reference-metadata-isref-iferror-source-fallback | parse-evaluate | 45 | 0 | 45 | 2,450 | 27 | 0 | 18 | 6,118 | 5,446 | 0 | 3,680 |
| reference-metadata-isref-ifna-source | evaluate | 45 | 0 | 45 | 1,150 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,668 |
| reference-metadata-isref-ifna-source | parse-evaluate | 45 | 0 | 45 | 1,810 | 18 | 0 | 15 | 5,030 | 4,902 | 0 | 3,620 |
| reference-metadata-isref-ifna-source-arithmetic | evaluate | 47 | 0 | 47 | 920 | 15 | 0 | 6 | 3,080 | 2,888 | 0 | 3,624 |
| reference-metadata-isref-ifna-source-arithmetic | parse-evaluate | 47 | 0 | 47 | 1,690 | 15 | 0 | 16 | 5,352 | 4,680 | 0 | 3,632 |
| reference-metadata-isref-ifna-source-fallback | evaluate | 43 | 0 | 43 | 1,330 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,672 |
| reference-metadata-isref-ifna-source-fallback | parse-evaluate | 43 | 0 | 43 | 2,050 | 23 | 0 | 15 | 5,028 | 4,900 | 0 | 3,644 |
| reference-metadata-isref-list | evaluate | 21 | 0 | 21 | 2,260 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,640 |
| reference-metadata-isref-list | parse-evaluate | 21 | 0 | 21 | 2,880 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,704 |
| reference-metadata-isref-number | evaluate | 10 | 0 | 10 | 1,050 | 14 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-isref-number | parse-evaluate | 10 | 0 | 10 | 1,270 | 14 | 0 | 11 | 4,050 | 3,954 | 0 | 3,660 |
| reference-metadata-isref-reference | evaluate | 18 | 0 | 18 | 1,580 | 14 | 0 | 10 | 3,992 | 3,928 | 0 | 3,640 |
| reference-metadata-isref-reference | parse-evaluate | 18 | 0 | 18 | 2,050 | 14 | 0 | 17 | 5,356 | 5,260 | 0 | 3,640 |
| reference-metadata-isref-selected-local | evaluate | 50 | 0 | 50 | 1,770 | 24 | 0 | 10 | 3,992 | 3,928 | 0 | 3,640 |
| reference-metadata-isref-selected-local | parse-evaluate | 50 | 0 | 50 | 2,680 | 24 | 0 | 20 | 6,204 | 5,692 | 0 | 3,708 |
| reference-metadata-isref-selected-source | evaluate | 49 | 0 | 49 | 1,330 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,680 |
| reference-metadata-isref-selected-source | parse-evaluate | 49 | 0 | 49 | 2,210 | 23 | 0 | 17 | 5,803 | 5,291 | 0 | 3,648 |
| reference-metadata-isref-source | evaluate | 32 | 0 | 32 | 1,130 | 12 | 0 | 7 | 3,592 | 3,528 | 0 | 3,624 |
| reference-metadata-isref-source | parse-evaluate | 32 | 0 | 32 | 1,650 | 12 | 0 | 14 | 4,985 | 4,889 | 0 | 3,632 |
| reference-metadata-isref-source-array-false-matrix | evaluate | 60 | 0 | 60 | 3,280 | 40 | 0 | 20 | 4,440 | 4,040 | 0 | 3,664 |
| reference-metadata-isref-source-array-false-matrix | parse-evaluate | 60 | 0 | 60 | 4,430 | 40 | 0 | 34 | 8,357 | 6,613 | 0 | 3,664 |
| reference-metadata-isref-source-array-true-matrix | evaluate | 60 | 0 | 60 | 4,400 | 53 | 0 | 25 | 5,608 | 4,760 | 0 | 3,732 |
| reference-metadata-isref-source-array-true-matrix | parse-evaluate | 60 | 0 | 60 | 5,430 | 53 | 0 | 39 | 9,525 | 7,333 | 0 | 3,692 |
| reference-metadata-large-range | evaluate | 19 | 0 | 19 | 1,580 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,680 |
| reference-metadata-large-range | parse-evaluate | 19 | 0 | 19 | 2,030 | 13 | 0 | 17 | 5,358 | 5,262 | 0 | 3,672 |
| reference-metadata-row-3d-cell | evaluate | 15 | 0 | 15 | 1,520 | 10 | 0 | 10 | 3,992 | 3,928 | 0 | 3,620 |
| reference-metadata-row-3d-cell | parse-evaluate | 15 | 0 | 15 | 1,950 | 10 | 0 | 17 | 5,356 | 5,260 | 0 | 3,680 |
| reference-metadata-row-array-refusal | evaluate | 13 | 0 | 13 | 1,790 | 20 | 0 | 11 | 4,928 | 4,304 | 0 | 3,816 |
| reference-metadata-row-array-refusal | parse-evaluate | 13 | 0 | 13 | 2,230 | 20 | 0 | 19 | 6,285 | 5,149 | 0 | 3,712 |
| reference-metadata-row-cell | evaluate | 11 | 0 | 11 | 1,540 | 10 | 0 | 10 | 3,992 | 3,928 | 0 | 3,616 |
| reference-metadata-row-cell | parse-evaluate | 11 | 0 | 11 | 1,910 | 10 | 0 | 16 | 5,348 | 5,252 | 0 | 3,664 |
| reference-metadata-row-computed-if-matrix | evaluate | 53 | 0 | 53 | 6,920 | 69 | 2 | 41 | 6,872 | 4,744 | 176 | 3,740 |
| reference-metadata-row-computed-if-matrix | parse-evaluate | 53 | 0 | 53 | 8,160 | 69 | 2 | 59 | 10,991 | 7,359 | 176 | 3,736 |
| reference-metadata-row-computed-iferror-matrix | evaluate | 68 | 0 | 68 | 7,620 | 87 | 2 | 46 | 7,240 | 4,744 | 176 | 3,684 |
| reference-metadata-row-computed-iferror-matrix | parse-evaluate | 68 | 0 | 68 | 9,190 | 87 | 2 | 66 | 11,407 | 7,375 | 176 | 3,632 |
| reference-metadata-row-computed-ifna-matrix | evaluate | 65 | 0 | 65 | 7,750 | 84 | 2 | 46 | 7,240 | 4,744 | 176 | 3,676 |
| reference-metadata-row-computed-ifna-matrix | parse-evaluate | 65 | 0 | 65 | 9,170 | 84 | 2 | 66 | 11,404 | 7,372 | 176 | 3,748 |
| reference-metadata-row-current | evaluate | 6 | 0 | 6 | 550 | 6 | 0 | 4 | 632 | 632 | 0 | 3,676 |
| reference-metadata-row-current | parse-evaluate | 6 | 0 | 6 | 690 | 6 | 0 | 6 | 1,022 | 1,022 | 0 | 3,656 |
| reference-metadata-row-list-refusal | evaluate | 19 | 0 | 19 | 2,250 | 15 | 0 | 14 | 4,648 | 4,152 | 0 | 3,656 |
| reference-metadata-row-list-refusal | parse-evaluate | 19 | 0 | 19 | 2,940 | 15 | 0 | 22 | 6,781 | 5,869 | 0 | 3,644 |
| reference-metadata-row-matrix | evaluate | 26 | 0 | 26 | 1,910 | 17 | 0 | 11 | 4,256 | 3,928 | 264 | 3,656 |
| reference-metadata-row-matrix | parse-evaluate | 26 | 0 | 26 | 2,410 | 17 | 0 | 20 | 5,639 | 5,279 | 264 | 3,624 |
| reference-metadata-row-one-entry-list | evaluate | 27 | 0 | 27 | 3,400 | 31 | 0 | 20 | 5,448 | 4,568 | 0 | 3,632 |
| reference-metadata-row-one-entry-list | parse-evaluate | 27 | 0 | 27 | 4,270 | 31 | 0 | 30 | 7,654 | 6,326 | 0 | 3,736 |
| reference-metadata-row-projected | evaluate | 55 | 0 | 55 | 7,720 | 94 | 0 | 41 | 8,744 | 5,336 | 264 | 3,672 |
| reference-metadata-row-projected | parse-evaluate | 55 | 0 | 55 | 9,030 | 94 | 0 | 59 | 12,812 | 7,964 | 264 | 3,676 |
| reference-metadata-row-scalar-vector | evaluate | 15 | 0 | 15 | 1,590 | 12 | 0 | 10 | 3,992 | 3,928 | 0 | 3,660 |
| reference-metadata-row-scalar-vector | parse-evaluate | 15 | 0 | 15 | 1,990 | 12 | 0 | 17 | 5,353 | 5,257 | 0 | 3,732 |
| reference-metadata-rows-3d | evaluate | 28 | 0 | 28 | 1,640 | 15 | 0 | 10 | 3,992 | 3,928 | 0 | 3,664 |
| reference-metadata-rows-3d | parse-evaluate | 28 | 0 | 28 | 2,220 | 15 | 0 | 19 | 5,377 | 5,281 | 0 | 3,636 |
| reference-metadata-rows-array | evaluate | 16 | 0 | 16 | 1,850 | 24 | 0 | 11 | 5,016 | 4,392 | 0 | 3,728 |
| reference-metadata-rows-array | parse-evaluate | 16 | 0 | 16 | 2,490 | 24 | 0 | 20 | 6,408 | 5,240 | 0 | 3,620 |
| reference-metadata-rows-list-refusal | evaluate | 20 | 0 | 20 | 2,290 | 16 | 0 | 14 | 4,648 | 4,152 | 0 | 3,644 |
| reference-metadata-rows-list-refusal | parse-evaluate | 20 | 0 | 20 | 2,900 | 16 | 0 | 22 | 6,782 | 5,870 | 0 | 3,628 |
| reference-metadata-rows-one-entry-list | evaluate | 28 | 0 | 28 | 3,430 | 32 | 0 | 20 | 5,448 | 4,568 | 0 | 3,740 |
| reference-metadata-rows-one-entry-list | parse-evaluate | 28 | 0 | 28 | 4,230 | 32 | 0 | 30 | 7,655 | 6,327 | 0 | 3,668 |
| reference-metadata-rows-projected | evaluate | 38 | 0 | 38 | 4,470 | 47 | 0 | 27 | 5,192 | 4,648 | 176 | 3,708 |
| reference-metadata-rows-projected | parse-evaluate | 38 | 0 | 38 | 5,430 | 47 | 0 | 41 | 9,072 | 7,184 | 176 | 3,680 |
| reference-metadata-rows-range | evaluate | 17 | 0 | 17 | 1,590 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,684 |
| reference-metadata-rows-range | parse-evaluate | 17 | 0 | 17 | 1,970 | 13 | 0 | 17 | 5,355 | 5,259 | 0 | 3,672 |
| reference-metadata-rows-resource | evaluate | 19 | 0 | 19 | 707 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,672 |
| reference-metadata-rows-resource | parse-evaluate | 19 | 0 | 19 | 1,127 | 8 | 0 | 52 | 16,952 | 4,142 | 0 | 3,684 |
| reference-metadata-rows-scalar-refusal | evaluate | 9 | 0 | 9 | 1,090 | 13 | 0 | 7 | 3,592 | 3,528 | 0 | 3,672 |
| reference-metadata-rows-scalar-refusal | parse-evaluate | 9 | 0 | 9 | 1,310 | 13 | 0 | 11 | 4,049 | 3,953 | 0 | 3,712 |
| reference-metadata-sheet-3d | evaluate | 29 | 0 | 29 | 1,680 | 16 | 0 | 10 | 3,992 | 3,928 | 0 | 3,660 |
| reference-metadata-sheet-3d | parse-evaluate | 29 | 0 | 29 | 2,290 | 16 | 0 | 19 | 5,378 | 5,282 | 0 | 3,636 |
| reference-metadata-sheet-array-matrix | evaluate | 26 | 0 | 26 | 1,900 | 43 | 0 | 11 | 4,248 | 4,008 | 176 | 3,668 |
| reference-metadata-sheet-array-matrix | parse-evaluate | 26 | 0 | 26 | 2,370 | 43 | 0 | 18 | 5,554 | 4,834 | 176 | 3,712 |
| reference-metadata-sheet-array-refusal | evaluate | 15 | 0 | 15 | 2,050 | 25 | 0 | 12 | 4,929 | 4,304 | 0 | 3,648 |
| reference-metadata-sheet-array-refusal | parse-evaluate | 15 | 0 | 15 | 2,490 | 25 | 0 | 20 | 6,288 | 5,151 | 0 | 3,800 |
| reference-metadata-sheet-cancel | evaluate | 17 | 0 | 17 | 12 | 0 | 0 | 0 | 0 | 0 | 0 | 2,996 |
| reference-metadata-sheet-cancel | parse-evaluate | 17 | 0 | 17 | 382 | 0 | 0 | 28 | 5,464 | 1,366 | 0 | 3,016 |
| reference-metadata-sheet-computed-if-matrix | evaluate | 60 | 0 | 60 | 13,500 | 224 | 0 | 69 | 9,608 | 5,048 | 176 | 3,724 |
| reference-metadata-sheet-computed-if-matrix | parse-evaluate | 60 | 0 | 60 | 15,630 | 224 | 0 | 84 | 12,836 | 6,772 | 176 | 3,772 |
| reference-metadata-sheet-computed-iferror-matrix | evaluate | 78 | 0 | 78 | 18,240 | 346 | 0 | 89 | 12,856 | 5,832 | 176 | 3,732 |
| reference-metadata-sheet-computed-iferror-matrix | parse-evaluate | 78 | 0 | 78 | 20,200 | 346 | 0 | 105 | 16,134 | 7,574 | 176 | 3,680 |
| reference-metadata-sheet-computed-ifna-matrix | evaluate | 75 | 0 | 75 | 18,080 | 340 | 0 | 89 | 12,856 | 5,832 | 176 | 3,632 |
| reference-metadata-sheet-computed-ifna-matrix | parse-evaluate | 75 | 0 | 75 | 20,000 | 340 | 0 | 105 | 16,131 | 7,571 | 176 | 3,752 |
| reference-metadata-sheet-computed-scalar | evaluate | 22 | 0 | 22 | 2,670 | 44 | 1 | 15 | 5,131 | 4,312 | 0 | 3,640 |
| reference-metadata-sheet-computed-scalar | parse-evaluate | 22 | 0 | 22 | 3,230 | 44 | 1 | 23 | 6,531 | 5,648 | 0 | 3,772 |
| reference-metadata-sheet-current | evaluate | 8 | 0 | 8 | 550 | 9 | 0 | 4 | 632 | 632 | 0 | 3,616 |
| reference-metadata-sheet-current | parse-evaluate | 8 | 0 | 8 | 730 | 9 | 0 | 6 | 1,024 | 1,024 | 0 | 3,636 |
| reference-metadata-sheet-hidden-text | evaluate | 16 | 0 | 16 | 1,080 | 25 | 0 | 7 | 3,592 | 3,528 | 0 | 3,620 |
| reference-metadata-sheet-hidden-text | parse-evaluate | 16 | 0 | 16 | 1,340 | 25 | 0 | 11 | 4,056 | 3,960 | 0 | 3,716 |
| reference-metadata-sheet-iferror-source | evaluate | 48 | 0 | 48 | 1,130 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,632 |
| reference-metadata-sheet-iferror-source | parse-evaluate | 48 | 0 | 48 | 1,810 | 21 | 0 | 15 | 5,033 | 4,905 | 0 | 3,660 |
| reference-metadata-sheet-iferror-source-fallback | evaluate | 45 | 0 | 45 | 1,520 | 27 | 0 | 8 | 3,848 | 3,656 | 0 | 3,652 |
| reference-metadata-sheet-iferror-source-fallback | parse-evaluate | 45 | 0 | 45 | 2,400 | 27 | 0 | 18 | 6,118 | 5,446 | 0 | 3,776 |
| reference-metadata-sheet-ifna-source | evaluate | 45 | 0 | 45 | 1,150 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,632 |
| reference-metadata-sheet-ifna-source | parse-evaluate | 45 | 0 | 45 | 1,820 | 18 | 0 | 15 | 5,030 | 4,902 | 0 | 3,616 |
| reference-metadata-sheet-ifna-source-fallback | evaluate | 43 | 0 | 43 | 1,330 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,668 |
| reference-metadata-sheet-ifna-source-fallback | parse-evaluate | 43 | 0 | 43 | 1,970 | 23 | 0 | 15 | 5,028 | 4,900 | 0 | 3,676 |
| reference-metadata-sheet-list-refusal | evaluate | 21 | 0 | 21 | 2,250 | 17 | 0 | 14 | 4,648 | 4,152 | 0 | 3,704 |
| reference-metadata-sheet-list-refusal | parse-evaluate | 21 | 0 | 21 | 2,980 | 17 | 0 | 22 | 6,783 | 5,871 | 0 | 3,728 |
| reference-metadata-sheet-logical-coercion | evaluate | 14 | 0 | 14 | 1,200 | 19 | 0 | 7 | 3,592 | 3,528 | 0 | 3,784 |
| reference-metadata-sheet-logical-coercion | parse-evaluate | 14 | 0 | 14 | 1,490 | 19 | 0 | 11 | 4,054 | 3,958 | 0 | 3,692 |
| reference-metadata-sheet-number-coercion | evaluate | 9 | 0 | 9 | 1,300 | 16 | 0 | 8 | 3,593 | 3,529 | 0 | 3,632 |
| reference-metadata-sheet-number-coercion | parse-evaluate | 9 | 0 | 9 | 1,540 | 16 | 0 | 12 | 4,050 | 3,954 | 0 | 3,628 |
| reference-metadata-sheet-reference | evaluate | 17 | 0 | 17 | 1,520 | 12 | 0 | 10 | 3,992 | 3,928 | 0 | 3,668 |
| reference-metadata-sheet-reference | parse-evaluate | 17 | 0 | 17 | 1,940 | 12 | 0 | 17 | 5,358 | 5,262 | 0 | 3,636 |
| reference-metadata-sheet-selected-local | evaluate | 54 | 0 | 54 | 1,800 | 24 | 0 | 10 | 3,992 | 3,928 | 0 | 3,764 |
| reference-metadata-sheet-selected-local | parse-evaluate | 54 | 0 | 54 | 2,790 | 24 | 0 | 21 | 6,212 | 5,700 | 0 | 3,660 |
| reference-metadata-sheet-selected-source | evaluate | 53 | 0 | 53 | 1,320 | 23 | 0 | 7 | 3,592 | 3,528 | 0 | 3,700 |
| reference-metadata-sheet-selected-source | parse-evaluate | 53 | 0 | 53 | 2,290 | 23 | 0 | 18 | 5,811 | 5,299 | 0 | 3,648 |
| reference-metadata-sheet-source | evaluate | 32 | 0 | 32 | 1,080 | 12 | 0 | 7 | 3,592 | 3,528 | 0 | 3,632 |
| reference-metadata-sheet-source | parse-evaluate | 32 | 0 | 32 | 1,640 | 12 | 0 | 14 | 4,985 | 4,889 | 0 | 3,652 |
| reference-metadata-sheet-source-array-false-matrix | evaluate | 60 | 0 | 60 | 3,280 | 40 | 0 | 20 | 4,440 | 4,040 | 0 | 3,620 |
| reference-metadata-sheet-source-array-false-matrix | parse-evaluate | 60 | 0 | 60 | 4,330 | 40 | 0 | 34 | 8,357 | 6,613 | 0 | 3,680 |
| reference-metadata-sheet-source-array-true-matrix | evaluate | 60 | 0 | 60 | 4,350 | 53 | 0 | 25 | 5,608 | 4,760 | 0 | 3,784 |
| reference-metadata-sheet-source-array-true-matrix | parse-evaluate | 60 | 0 | 60 | 5,390 | 53 | 0 | 39 | 9,525 | 7,333 | 0 | 3,700 |
| reference-metadata-sheet-text | evaluate | 14 | 0 | 14 | 1,110 | 21 | 0 | 7 | 3,592 | 3,528 | 0 | 3,660 |
| reference-metadata-sheet-text | parse-evaluate | 14 | 0 | 14 | 1,320 | 21 | 0 | 11 | 4,054 | 3,958 | 0 | 3,624 |
| reference-metadata-sheet-unknown-text | evaluate | 17 | 0 | 17 | 1,110 | 27 | 0 | 7 | 3,592 | 3,528 | 0 | 3,668 |
| reference-metadata-sheet-unknown-text | parse-evaluate | 17 | 0 | 17 | 1,350 | 27 | 0 | 11 | 4,057 | 3,961 | 0 | 3,624 |
| reference-metadata-sheets-3d | evaluate | 30 | 0 | 30 | 1,830 | 19 | 0 | 11 | 4,216 | 4,040 | 0 | 3,668 |
| reference-metadata-sheets-3d | parse-evaluate | 30 | 0 | 30 | 2,550 | 19 | 0 | 20 | 5,603 | 5,395 | 0 | 3,620 |
| reference-metadata-sheets-array-refusal | evaluate | 18 | 0 | 18 | 1,830 | 26 | 0 | 11 | 5,016 | 4,392 | 0 | 3,636 |
| reference-metadata-sheets-array-refusal | parse-evaluate | 18 | 0 | 18 | 2,520 | 26 | 0 | 20 | 6,410 | 5,242 | 0 | 3,624 |
| reference-metadata-sheets-current | evaluate | 9 | 0 | 9 | 560 | 10 | 0 | 4 | 632 | 632 | 0 | 3,736 |
| reference-metadata-sheets-current | parse-evaluate | 9 | 0 | 9 | 790 | 10 | 0 | 6 | 1,025 | 1,025 | 0 | 3,652 |
| reference-metadata-sheets-formula-error | evaluate | 14 | 0 | 14 | 1,070 | 18 | 0 | 7 | 3,592 | 3,528 | 0 | 3,660 |
| reference-metadata-sheets-formula-error | parse-evaluate | 14 | 0 | 14 | 1,300 | 18 | 0 | 11 | 4,054 | 3,958 | 0 | 3,604 |
| reference-metadata-sheets-iferror-source | evaluate | 49 | 0 | 49 | 1,130 | 22 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-sheets-iferror-source | parse-evaluate | 49 | 0 | 49 | 1,830 | 22 | 0 | 15 | 5,034 | 4,906 | 0 | 3,652 |
| reference-metadata-sheets-iferror-source-fallback | evaluate | 46 | 0 | 46 | 1,550 | 28 | 0 | 8 | 3,848 | 3,656 | 0 | 3,776 |
| reference-metadata-sheets-iferror-source-fallback | parse-evaluate | 46 | 0 | 46 | 2,410 | 28 | 0 | 18 | 6,119 | 5,447 | 0 | 3,616 |
| reference-metadata-sheets-ifna-source | evaluate | 46 | 0 | 46 | 1,130 | 19 | 0 | 7 | 3,592 | 3,528 | 0 | 3,680 |
| reference-metadata-sheets-ifna-source | parse-evaluate | 46 | 0 | 46 | 1,830 | 19 | 0 | 15 | 5,031 | 4,903 | 0 | 3,676 |
| reference-metadata-sheets-ifna-source-fallback | evaluate | 44 | 0 | 44 | 1,310 | 24 | 0 | 7 | 3,592 | 3,528 | 0 | 3,684 |
| reference-metadata-sheets-ifna-source-fallback | parse-evaluate | 44 | 0 | 44 | 2,050 | 24 | 0 | 15 | 5,029 | 4,901 | 0 | 3,804 |
| reference-metadata-sheets-list-refusal | evaluate | 22 | 0 | 22 | 2,300 | 18 | 0 | 14 | 4,648 | 4,152 | 0 | 3,652 |
| reference-metadata-sheets-list-refusal | parse-evaluate | 22 | 0 | 22 | 2,880 | 18 | 0 | 22 | 6,784 | 5,872 | 0 | 3,640 |
| reference-metadata-sheets-one-entry-list | evaluate | 30 | 0 | 30 | 3,390 | 34 | 0 | 20 | 5,448 | 4,568 | 0 | 3,680 |
| reference-metadata-sheets-one-entry-list | parse-evaluate | 30 | 0 | 30 | 4,340 | 34 | 0 | 30 | 7,657 | 6,329 | 0 | 3,728 |
| reference-metadata-sheets-reference | evaluate | 18 | 0 | 18 | 1,540 | 13 | 0 | 10 | 3,992 | 3,928 | 0 | 3,668 |
| reference-metadata-sheets-reference | parse-evaluate | 18 | 0 | 18 | 1,950 | 13 | 0 | 17 | 5,359 | 5,263 | 0 | 3,684 |
| reference-metadata-sheets-selected-local | evaluate | 55 | 0 | 55 | 1,820 | 25 | 0 | 10 | 3,992 | 3,928 | 0 | 3,680 |
| reference-metadata-sheets-selected-local | parse-evaluate | 55 | 0 | 55 | 2,810 | 25 | 0 | 21 | 6,213 | 5,701 | 0 | 3,640 |
| reference-metadata-sheets-selected-source | evaluate | 54 | 0 | 54 | 1,300 | 24 | 0 | 7 | 3,592 | 3,528 | 0 | 3,660 |
| reference-metadata-sheets-selected-source | parse-evaluate | 54 | 0 | 54 | 2,320 | 24 | 0 | 18 | 5,812 | 5,300 | 0 | 3,640 |
| reference-metadata-sheets-source | evaluate | 33 | 0 | 33 | 1,080 | 13 | 0 | 7 | 3,592 | 3,528 | 0 | 3,624 |
| reference-metadata-sheets-source | parse-evaluate | 33 | 0 | 33 | 1,690 | 13 | 0 | 14 | 4,986 | 4,890 | 0 | 3,684 |
| reference-metadata-sheets-source-array-false-matrix | evaluate | 61 | 0 | 61 | 3,210 | 41 | 0 | 20 | 4,440 | 4,040 | 0 | 3,620 |
| reference-metadata-sheets-source-array-false-matrix | parse-evaluate | 61 | 0 | 61 | 4,380 | 41 | 0 | 34 | 8,358 | 6,614 | 0 | 3,628 |
| reference-metadata-sheets-source-array-true-matrix | evaluate | 61 | 0 | 61 | 4,370 | 54 | 0 | 25 | 5,608 | 4,760 | 0 | 3,668 |
| reference-metadata-sheets-source-array-true-matrix | parse-evaluate | 61 | 0 | 61 | 5,460 | 54 | 0 | 39 | 9,526 | 7,334 | 0 | 3,744 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.

## Capture audit

The independent root audit passed all 4,920 retained rows after applying the reviewed exact read contract. The only matched-control latency flag above 5% is `reference-conditional-256x4-sumifs` in `evaluate`: `83,060.5` to `88,280.0` ns/repeat (`+6.284%`); its `parse-evaluate` delta is `-1.852%`. The canonical root audit used exact `elapsed_ns_p50/repeat` values and an independent unpaired median-ratio bootstrap (5,000 resamples, derived seed `5554865353495640452`), giving a descriptive 95% interval of `[-0.937%, +12.014%]`; a 100,000-resample replay with the same floats gives `[-0.728%, +12.032%]`. Work, resolver reads, allocator calls, and bytes/repeat are identical for this group, so the observation has no matching accounting change and no causal explanation is established. The revised capture and the superseded 99-case diagnostic together retain 9,570 timed rows; the full interpretation, host limitation, and setup/verifier diagnostics are in [`capture-analysis.md`](capture-analysis.md).
