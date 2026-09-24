# ODS value-inspection evaluator performance profile — latest complete capture

This report is derived from root session `87626`, baseline commit `d623f3c2ecc0c837017f700174656f0e443759a5`, three warmups, and fifteen fresh child samples in both `evaluate` and `parse-evaluate`. The preserved earlier complete report remains `performance-report.md`; this file describes `results/baseline-d623f3c2ecc0c837017f700174656f0e443759a5` and `results/candidate-final`.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 48,107 | 46,730 | -2.9% | 2,311 | 2,311 | 88 | 88 | 3,744 | 3,812 |
| array-control-16x16-arithmetic | parse-evaluate | 62,422 | 61,712 | -1.1% | 2,311 | 2,311 | 352 | 352 | 3,796 | 3,800 |
| array-control-16x16-sin | evaluate | 128,850 | 128,175 | -0.5% | 2,311 | 2,311 | 2,144 | 2,144 | 4,108 | 4,188 |
| array-control-16x16-sin | parse-evaluate | 144,428 | 143,705 | -0.5% | 2,311 | 2,311 | 2,412 | 2,412 | 4,120 | 4,176 |
| array-control-4x4-arithmetic | evaluate | 4,568 | 4,499 | -1.5% | 151 | 151 | 1,120 | 1,120 | 3,476 | 3,572 |
| array-control-4x4-arithmetic | parse-evaluate | 5,285 | 5,228 | -1.1% | 151 | 151 | 2,240 | 2,240 | 3,548 | 3,560 |
| array-control-4x4-sin | evaluate | 9,923 | 9,951 | +0.3% | 151 | 151 | 3,840 | 3,840 | 3,864 | 3,904 |
| array-control-4x4-sin | parse-evaluate | 10,731 | 10,725 | -0.1% | 151 | 151 | 5,040 | 5,040 | 3,824 | 3,964 |
| concat-borrowed-literals | evaluate | 724 | 717 | -1.0% | 21 | 21 | 6,000 | 6,000 | 3,176 | 3,276 |
| concat-borrowed-literals | parse-evaluate | 869 | 865 | -0.5% | 21 | 21 | 9,000 | 9,000 | 3,200 | 3,384 |
| concat-growth-chain | evaluate | 1,892 | 1,898 | +0.3% | 20 | 20 | 10,000 | 10,000 | 3,200 | 3,352 |
| concat-growth-chain | parse-evaluate | 2,299 | 2,302 | +0.1% | 20 | 20 | 17,000 | 17,000 | 3,212 | 3,276 |
| concat-owned-left | evaluate | 1,071 | 1,063 | -0.7% | 24 | 24 | 8,000 | 8,000 | 3,192 | 3,316 |
| concat-owned-left | parse-evaluate | 1,305 | 1,318 | +1.0% | 24 | 24 | 13,000 | 13,000 | 3,192 | 3,392 |
| concat-owned-right | evaluate | 1,083 | 1,082 | -0.1% | 24 | 24 | 8,000 | 8,000 | 3,220 | 3,288 |
| concat-owned-right | parse-evaluate | 1,308 | 1,322 | +1.1% | 24 | 24 | 13,000 | 13,000 | 3,184 | 3,264 |
| database-control-dstdev | evaluate | 3,940 | 3,850 | -2.3% | 30 | 30 | 20 | 20 | 3,564 | 3,624 |
| database-control-dstdev | parse-evaluate | 5,100 | 5,140 | +0.8% | 30 | 30 | 29 | 29 | 3,624 | 3,736 |
| database-control-dsum | evaluate | 4,040 | 3,950 | -2.2% | 28 | 28 | 20 | 20 | 3,632 | 3,716 |
| database-control-dsum | parse-evaluate | 5,210 | 5,130 | -1.5% | 28 | 28 | 29 | 29 | 3,620 | 3,624 |
| database-control-dvar | evaluate | 3,960 | 3,880 | -2.0% | 28 | 28 | 20 | 20 | 3,620 | 3,640 |
| database-control-dvar | parse-evaluate | 5,050 | 5,230 | +3.6% | 28 | 28 | 29 | 29 | 3,600 | 3,624 |
| literal-aggregate-4x1-sum | evaluate | 1,830 | 1,830 | +0.0% | 15 | 15 | 10 | 10 | 3,544 | 3,612 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,620 | 2,700 | +3.1% | 15 | 15 | 23 | 23 | 3,548 | 3,560 |
| reference-aggregate-64x4-sum | evaluate | 10,367 | 10,357 | -0.1% | 16 | 16 | 32 | 32 | 3,532 | 3,600 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,797 | 10,762 | -0.3% | 16 | 16 | 60 | 60 | 3,496 | 3,544 |
| reference-array-16x4-arithmetic | evaluate | 9,589 | 9,461 | -1.3% | 16 | 16 | 800 | 800 | 3,436 | 3,640 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,907 | 9,807 | -1.0% | 16 | 16 | 1,280 | 1,280 | 3,472 | 3,648 |
| reference-conditional-256x4-sumifs | evaluate | 85,355 | 86,365 | +1.2% | 52 | 52 | 42 | 42 | 3,600 | 3,688 |
| reference-conditional-256x4-sumifs | parse-evaluate | 86,100 | 87,110 | +1.2% | 52 | 52 | 68 | 68 | 3,600 | 3,648 |
| reference-control-average | evaluate | 10,390 | 10,262 | -1.2% | 20 | 20 | 32 | 32 | 3,588 | 3,664 |
| reference-control-average | parse-evaluate | 10,792 | 10,677 | -1.1% | 20 | 20 | 60 | 60 | 3,600 | 3,692 |
| reference-control-counta | evaluate | 8,737 | 8,675 | -0.7% | 19 | 19 | 32 | 32 | 3,564 | 3,548 |
| reference-control-counta | parse-evaluate | 9,167 | 9,052 | -1.3% | 19 | 19 | 60 | 60 | 3,596 | 3,540 |
| representative-median | evaluate | 1,558 | 1,553 | -0.3% | 16 | 16 | 9,000 | 9,000 | 3,364 | 3,444 |
| representative-median | parse-evaluate | 1,869 | 1,879 | +0.5% | 16 | 16 | 14,000 | 14,000 | 3,324 | 3,408 |
| representative-percentrank | evaluate | 987 | 968 | -1.9% | 19 | 19 | 7,000 | 7,000 | 3,316 | 3,484 |
| representative-percentrank | parse-evaluate | 1,208 | 1,215 | +0.6% | 19 | 19 | 11,000 | 11,000 | 3,340 | 3,488 |
| representative-rank | evaluate | 934 | 929 | -0.5% | 12 | 12 | 7,000 | 7,000 | 3,292 | 3,504 |
| representative-rank | parse-evaluate | 1,138 | 1,165 | +2.4% | 12 | 12 | 11,000 | 11,000 | 3,356 | 3,484 |
| scalar-aggregate-sum | evaluate | 581 | 579 | -0.3% | 10 | 10 | 4,000 | 4,000 | 3,348 | 3,316 |
| scalar-aggregate-sum | parse-evaluate | 752 | 763 | +1.5% | 10 | 10 | 8,000 | 8,000 | 3,312 | 3,308 |
| scalar-control-arithmetic | evaluate | 567 | 569 | +0.4% | 10 | 10 | 5,000 | 5,000 | 3,180 | 3,392 |
| scalar-control-arithmetic | parse-evaluate | 701 | 702 | +0.1% | 10 | 10 | 8,000 | 8,000 | 3,184 | 3,252 |
| scalar-control-average | evaluate | 1,028 | 1,028 | +0.0% | 23 | 23 | 6,000 | 6,000 | 3,284 | 3,324 |
| scalar-control-average | parse-evaluate | 1,290 | 1,251 | -3.0% | 23 | 23 | 10,000 | 10,000 | 3,356 | 3,332 |
| scalar-control-counta | evaluate | 820 | 837 | +2.1% | 24 | 24 | 6,000 | 6,000 | 3,284 | 3,300 |
| scalar-control-counta | parse-evaluate | 1,083 | 1,080 | -0.3% | 24 | 24 | 10,000 | 10,000 | 3,296 | 3,352 |
| scalar-control-imsum | evaluate | 1,396 | 1,373 | -1.6% | 41 | 41 | 8,000 | 8,000 | 3,292 | 3,304 |
| scalar-control-imsum | parse-evaluate | 1,948 | 1,912 | -1.8% | 41 | 41 | 17,000 | 17,000 | 3,228 | 3,244 |
| scalar-control-sin | evaluate | 462 | 466 | +0.9% | 10 | 10 | 4,000 | 4,000 | 3,684 | 3,796 |
| scalar-control-sin | parse-evaluate | 647 | 655 | +1.2% | 10 | 10 | 8,000 | 8,000 | 3,612 | 3,804 |
| scalar-control-stdev | evaluate | 853 | 857 | +0.5% | 21 | 21 | 6,000 | 6,000 | 3,312 | 3,280 |
| scalar-control-stdev | parse-evaluate | 1,074 | 1,086 | +1.1% | 21 | 21 | 10,000 | 10,000 | 3,304 | 3,420 |
| scalar-control-var | evaluate | 839 | 844 | +0.6% | 19 | 19 | 6,000 | 6,000 | 3,316 | 3,356 |
| scalar-control-var | parse-evaluate | 1,064 | 1,063 | -0.1% | 19 | 19 | 10,000 | 10,000 | 3,284 | 3,384 |

## Value-inspection and conversion workloads

The candidate matrix covers all sixteen value-inspection and conversion functions over scalar, borrowed reference, matrix, TYPE/N exception, typed refusal, lazy branch, date/fraction, separator, cancellation, and resource-refusal cases. The raw receipts retain exact checksums, output labels, resolver reads, work, allocation, and RSS values.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-inspection-isnumber | evaluate | 832 | 0 | 832 | 370 | 4 | 0 | 9 | 7,048 | 7,048 | 0 | 3,684 |
| cancellation-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 765 | 4 | 0 | 37 | 12,516 | 8,383 | 0 | 3,692 |
| cancellation-inspection-numbervalue | evaluate | 832 | 0 | 832 | 382 | 5 | 0 | 9 | 7,048 | 7,048 | 0 | 3,620 |
| cancellation-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 787 | 5 | 0 | 37 | 12,528 | 8,386 | 0 | 3,572 |
| cancellation-inspection-type | evaluate | 832 | 0 | 832 | 395 | 3 | 0 | 10 | 3,992 | 3,928 | 0 | 3,668 |
| cancellation-inspection-type | parse-evaluate | 832 | 0 | 832 | 797 | 3 | 0 | 38 | 9,444 | 5,259 | 0 | 3,608 |
| cancellation-inspection-value | evaluate | 832 | 0 | 832 | 357 | 3 | 0 | 9 | 7,048 | 7,048 | 0 | 3,668 |
| cancellation-inspection-value | parse-evaluate | 832 | 0 | 832 | 777 | 3 | 0 | 37 | 12,504 | 8,380 | 0 | 3,692 |
| lazy-if-cache-isblank | evaluate | 832 | 0 | 832 | 5,070 | 69 | 2 | 27 | 3,416 | 2,200 | 176 | 3,620 |
| lazy-if-cache-isblank | parse-evaluate | 832 | 0 | 832 | 6,130 | 69 | 2 | 41 | 7,305 | 4,745 | 176 | 3,684 |
| lazy-if-cache-isnumber | evaluate | 832 | 0 | 832 | 5,100 | 71 | 2 | 27 | 3,416 | 2,200 | 176 | 3,688 |
| lazy-if-cache-isnumber | parse-evaluate | 832 | 0 | 832 | 6,070 | 71 | 2 | 41 | 7,306 | 4,746 | 176 | 3,616 |
| lazy-if-cache-istext | evaluate | 832 | 0 | 832 | 5,120 | 67 | 2 | 27 | 3,416 | 2,200 | 176 | 3,672 |
| lazy-if-cache-istext | parse-evaluate | 832 | 0 | 832 | 6,100 | 67 | 2 | 41 | 7,304 | 4,744 | 176 | 3,616 |
| matrix-inspection-isblank | evaluate | 832 | 0 | 832 | 2,250 | 43 | 0 | 11 | 2,920 | 2,296 | 352 | 3,620 |
| matrix-inspection-isblank | parse-evaluate | 832 | 0 | 832 | 3,230 | 43 | 0 | 24 | 6,056 | 3,992 | 352 | 3,620 |
| matrix-inspection-iserror | evaluate | 832 | 0 | 832 | 2,260 | 43 | 0 | 11 | 2,920 | 2,296 | 352 | 3,692 |
| matrix-inspection-iserror | parse-evaluate | 832 | 0 | 832 | 3,190 | 43 | 0 | 24 | 6,056 | 3,992 | 352 | 3,688 |
| matrix-inspection-iseven | evaluate | 832 | 0 | 832 | 2,320 | 45 | 0 | 11 | 2,920 | 2,296 | 352 | 3,672 |
| matrix-inspection-iseven | parse-evaluate | 832 | 0 | 832 | 3,280 | 45 | 0 | 24 | 6,055 | 3,991 | 352 | 3,632 |
| matrix-inspection-isnumber | evaluate | 832 | 0 | 832 | 2,290 | 44 | 0 | 11 | 2,920 | 2,296 | 352 | 3,692 |
| matrix-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 3,210 | 44 | 0 | 24 | 6,057 | 3,993 | 352 | 3,656 |
| matrix-inspection-istext | evaluate | 832 | 0 | 832 | 2,250 | 42 | 0 | 11 | 2,920 | 2,296 | 352 | 3,624 |
| matrix-inspection-istext | parse-evaluate | 832 | 0 | 832 | 3,180 | 42 | 0 | 24 | 6,055 | 3,991 | 352 | 3,652 |
| matrix-inspection-n | evaluate | 832 | 0 | 832 | 2,090 | 32 | 0 | 11 | 5,016 | 4,392 | 0 | 3,616 |
| matrix-inspection-n | parse-evaluate | 832 | 0 | 832 | 3,350 | 32 | 0 | 24 | 8,146 | 6,082 | 0 | 3,692 |
| matrix-inspection-value | evaluate | 832 | 0 | 832 | 2,330 | 44 | 0 | 11 | 2,920 | 2,296 | 352 | 3,684 |
| matrix-inspection-value | parse-evaluate | 832 | 0 | 832 | 3,270 | 44 | 0 | 24 | 6,054 | 3,990 | 352 | 3,724 |
| n-reference-intersection | evaluate | 0 | 0 | 0 | 1,670 | 11 | 1 | 10 | 3,992 | 3,928 | 0 | 3,588 |
| n-reference-intersection | parse-evaluate | 0 | 0 | 0 | 2,070 | 11 | 1 | 17 | 5,351 | 5,255 | 0 | 3,564 |
| numbervalue-invalid-separator | evaluate | 0 | 0 | 0 | 1,750 | 22 | 0 | 11 | 2,680 | 2,232 | 176 | 3,688 |
| numbervalue-invalid-separator | parse-evaluate | 0 | 0 | 0 | 2,220 | 22 | 0 | 17 | 4,051 | 3,571 | 176 | 3,660 |
| numbervalue-transform-decimal-comma | evaluate | 12 | 0 | 12 | 1,376 | 40 | 0 | 8,000 | 2,113,000 | 1,656 | 0 | 3,620 |
| numbervalue-transform-decimal-comma | parse-evaluate | 12 | 0 | 12 | 1,647 | 40 | 0 | 12,000 | 2,593,000 | 2,104 | 0 | 3,608 |
| numbervalue-transform-grouped-space | evaluate | 12 | 0 | 12 | 1,376 | 42 | 0 | 8,000 | 2,114,000 | 1,656 | 0 | 3,692 |
| numbervalue-transform-grouped-space | parse-evaluate | 12 | 0 | 12 | 1,657 | 42 | 0 | 12,000 | 2,595,000 | 2,105 | 0 | 3,668 |
| numbervalue-transform-percent | evaluate | 12 | 0 | 12 | 1,369 | 42 | 0 | 8,000 | 2,114,000 | 1,656 | 0 | 3,688 |
| numbervalue-transform-percent | parse-evaluate | 12 | 0 | 12 | 1,656 | 42 | 0 | 12,000 | 2,595,000 | 2,105 | 0 | 3,628 |
| reference-inspection-error-type | evaluate | 832 | 0 | 832 | 7,485 | 208 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,564 |
| reference-inspection-error-type | parse-evaluate | 832 | 0 | 832 | 7,905 | 208 | 64 | 64 | 33,668 | 8,385 | 5,632 | 3,624 |
| reference-inspection-isblank | evaluate | 832 | 0 | 832 | 7,437 | 205 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,604 |
| reference-inspection-isblank | parse-evaluate | 832 | 0 | 832 | 7,847 | 205 | 64 | 64 | 33,656 | 8,382 | 5,632 | 3,680 |
| reference-inspection-iserr | evaluate | 832 | 0 | 832 | 7,547 | 203 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,632 |
| reference-inspection-iserr | parse-evaluate | 832 | 0 | 832 | 7,872 | 203 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,672 |
| reference-inspection-iserror | evaluate | 832 | 0 | 832 | 7,500 | 205 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,616 |
| reference-inspection-iserror | parse-evaluate | 832 | 0 | 832 | 7,865 | 205 | 64 | 64 | 33,656 | 8,382 | 5,632 | 3,660 |
| reference-inspection-iseven | evaluate | 832 | 0 | 832 | 8,122 | 292 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,692 |
| reference-inspection-iseven | parse-evaluate | 832 | 0 | 832 | 8,547 | 292 | 64 | 64 | 33,652 | 8,381 | 5,632 | 3,632 |
| reference-inspection-islogical | evaluate | 832 | 0 | 832 | 7,540 | 207 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,692 |
| reference-inspection-islogical | parse-evaluate | 832 | 0 | 832 | 7,872 | 207 | 64 | 64 | 33,664 | 8,384 | 5,632 | 3,624 |
| reference-inspection-isna | evaluate | 832 | 0 | 832 | 7,445 | 202 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,640 |
| reference-inspection-isna | parse-evaluate | 832 | 0 | 832 | 7,875 | 202 | 64 | 64 | 33,644 | 8,379 | 5,632 | 3,680 |
| reference-inspection-isnontext | evaluate | 832 | 0 | 832 | 7,472 | 207 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,588 |
| reference-inspection-isnontext | parse-evaluate | 832 | 0 | 832 | 7,825 | 207 | 64 | 64 | 33,664 | 8,384 | 5,632 | 3,604 |
| reference-inspection-isnumber | evaluate | 832 | 0 | 832 | 7,452 | 206 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,584 |
| reference-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 7,875 | 206 | 64 | 64 | 33,660 | 8,383 | 5,632 | 3,620 |
| reference-inspection-isodd | evaluate | 832 | 0 | 832 | 8,177 | 291 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,620 |
| reference-inspection-isodd | parse-evaluate | 832 | 0 | 832 | 8,635 | 291 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,628 |
| reference-inspection-istext | evaluate | 832 | 0 | 832 | 7,442 | 204 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,588 |
| reference-inspection-istext | parse-evaluate | 832 | 0 | 832 | 7,875 | 204 | 64 | 64 | 33,652 | 8,381 | 5,632 | 3,620 |
| reference-inspection-numbervalue | evaluate | 832 | 0 | 832 | 16,237 | 299 | 64 | 264 | 28,772 | 7,060 | 5,632 | 3,564 |
| reference-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 16,562 | 299 | 64 | 292 | 34,252 | 8,398 | 5,632 | 3,624 |
| reference-inspection-value | evaluate | 832 | 0 | 832 | 8,757 | 291 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,620 |
| reference-inspection-value | parse-evaluate | 832 | 0 | 832 | 9,095 | 291 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,660 |
| resource-inspection-isnumber | evaluate | 832 | 0 | 832 | 407 | 11 | 0 | 16 | 1,184 | 296 | 0 | 3,600 |
| resource-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 820 | 11 | 0 | 44 | 6,652 | 1,631 | 0 | 3,628 |
| resource-inspection-numbervalue | evaluate | 832 | 0 | 832 | 400 | 14 | 0 | 16 | 1,184 | 296 | 0 | 3,532 |
| resource-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 822 | 14 | 0 | 44 | 6,664 | 1,634 | 0 | 3,620 |
| resource-inspection-type | evaluate | 832 | 0 | 832 | 767 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,500 |
| resource-inspection-type | parse-evaluate | 832 | 0 | 832 | 1,087 | 8 | 0 | 52 | 16,940 | 4,139 | 0 | 3,476 |
| resource-inspection-value | evaluate | 832 | 0 | 832 | 410 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,624 |
| resource-inspection-value | parse-evaluate | 832 | 0 | 832 | 812 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,620 |
| scalar-inspection-error-type | evaluate | 0 | 0 | 0 | 697 | 19 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,608 |
| scalar-inspection-error-type | parse-evaluate | 0 | 0 | 0 | 933 | 19 | 0 | 9,000 | 1,481,000 | 1,449 | 0 | 3,648 |
| scalar-inspection-isblank | evaluate | 0 | 0 | 0 | 701 | 12 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| scalar-inspection-isblank | parse-evaluate | 0 | 0 | 0 | 911 | 12 | 0 | 9,000 | 1,476,000 | 1,444 | 0 | 3,584 |
| scalar-inspection-iserr | evaluate | 0 | 0 | 0 | 688 | 17 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,652 |
| scalar-inspection-iserr | parse-evaluate | 0 | 0 | 0 | 919 | 17 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,616 |
| scalar-inspection-iserror | evaluate | 0 | 0 | 0 | 697 | 16 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,692 |
| scalar-inspection-iserror | parse-evaluate | 0 | 0 | 0 | 915 | 16 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,640 |
| scalar-inspection-iseven | evaluate | 0 | 0 | 0 | 721 | 13 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,572 |
| scalar-inspection-iseven | parse-evaluate | 0 | 0 | 0 | 932 | 13 | 0 | 9,000 | 1,475,000 | 1,443 | 0 | 3,624 |
| scalar-inspection-islogical | evaluate | 0 | 0 | 0 | 824 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,652 |
| scalar-inspection-islogical | parse-evaluate | 0 | 0 | 0 | 1,083 | 20 | 0 | 9,000 | 1,482,000 | 1,450 | 0 | 3,588 |
| scalar-inspection-isna | evaluate | 0 | 0 | 0 | 695 | 13 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,580 |
| scalar-inspection-isna | parse-evaluate | 0 | 0 | 0 | 900 | 13 | 0 | 9,000 | 1,475,000 | 1,443 | 0 | 3,692 |
| scalar-inspection-isnontext | evaluate | 0 | 0 | 0 | 703 | 16 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| scalar-inspection-isnontext | parse-evaluate | 0 | 0 | 0 | 926 | 16 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,640 |
| scalar-inspection-isnumber | evaluate | 0 | 0 | 0 | 709 | 15 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| scalar-inspection-isnumber | parse-evaluate | 0 | 0 | 0 | 923 | 15 | 0 | 9,000 | 1,477,000 | 1,445 | 0 | 3,620 |
| scalar-inspection-isodd | evaluate | 0 | 0 | 0 | 723 | 12 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,692 |
| scalar-inspection-isodd | parse-evaluate | 0 | 0 | 0 | 935 | 12 | 0 | 9,000 | 1,474,000 | 1,442 | 0 | 3,672 |
| scalar-inspection-istext | evaluate | 0 | 0 | 0 | 709 | 14 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,620 |
| scalar-inspection-istext | parse-evaluate | 0 | 0 | 0 | 917 | 14 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,624 |
| scalar-inspection-n | evaluate | 0 | 0 | 0 | 1,120 | 14 | 0 | 7,000 | 3,592,000 | 3,528 | 0 | 3,688 |
| scalar-inspection-n | parse-evaluate | 0 | 0 | 0 | 1,364 | 14 | 0 | 11,000 | 4,050,000 | 3,954 | 0 | 3,628 |
| scalar-inspection-na | evaluate | 0 | 0 | 0 | 497 | 5 | 0 | 4,000 | 632,000 | 632 | 0 | 3,676 |
| scalar-inspection-na | parse-evaluate | 0 | 0 | 0 | 604 | 5 | 0 | 6,000 | 1,021,000 | 1,021 | 0 | 3,584 |
| scalar-inspection-numbervalue | evaluate | 0 | 0 | 0 | 1,374 | 40 | 0 | 8,000 | 2,113,000 | 1,656 | 0 | 3,660 |
| scalar-inspection-numbervalue | parse-evaluate | 0 | 0 | 0 | 1,661 | 40 | 0 | 12,000 | 2,593,000 | 2,104 | 0 | 3,624 |
| scalar-inspection-type | evaluate | 0 | 0 | 0 | 971 | 14 | 0 | 7,000 | 3,592,000 | 3,528 | 0 | 3,640 |
| scalar-inspection-type | parse-evaluate | 0 | 0 | 0 | 1,207 | 14 | 0 | 11,000 | 4,052,000 | 3,956 | 0 | 3,656 |
| scalar-inspection-value | evaluate | 0 | 0 | 0 | 773 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,700 |
| scalar-inspection-value | parse-evaluate | 0 | 0 | 0 | 996 | 20 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,692 |
| shape-refusal-numbervalue | evaluate | 0 | 0 | 0 | 2,030 | 25 | 0 | 12 | 1,944 | 1,576 | 0 | 3,620 |
| shape-refusal-numbervalue | parse-evaluate | 0 | 0 | 0 | 2,670 | 25 | 0 | 21 | 3,325 | 2,925 | 0 | 3,640 |
| shape-refusal-value | evaluate | 0 | 0 | 0 | 1,980 | 17 | 0 | 12 | 1,944 | 1,576 | 0 | 3,728 |
| shape-refusal-value | parse-evaluate | 0 | 0 | 0 | 2,600 | 17 | 0 | 21 | 3,319 | 2,919 | 0 | 3,640 |
| type-array-metadata | evaluate | 0 | 0 | 0 | 2,061 | 34 | 0 | 11 | 5,016 | 4,392 | 0 | 3,608 |
| type-array-metadata | parse-evaluate | 0 | 0 | 0 | 3,200 | 34 | 0 | 24 | 8,149 | 6,085 | 0 | 3,636 |
| type-reference-scan | evaluate | 832 | 0 | 832 | 3,177 | 76 | 64 | 40 | 15,968 | 3,928 | 0 | 3,620 |
| type-reference-scan | parse-evaluate | 832 | 0 | 832 | 3,595 | 76 | 64 | 68 | 21,420 | 5,259 | 0 | 3,628 |
| value-date-fraction-value-date | evaluate | 22 | 0 | 22 | 771 | 30 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| value-date-fraction-value-date | parse-evaluate | 22 | 0 | 22 | 986 | 30 | 0 | 9,000 | 1,484,000 | 1,452 | 0 | 3,628 |
| value-date-fraction-value-mixed-fraction | evaluate | 22 | 0 | 22 | 790 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| value-date-fraction-value-mixed-fraction | parse-evaluate | 22 | 0 | 22 | 995 | 20 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,608 |
| value-date-fraction-value-time | evaluate | 22 | 0 | 22 | 768 | 26 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,560 |
| value-date-fraction-value-time | parse-evaluate | 22 | 0 | 22 | 984 | 26 | 0 | 9,000 | 1,482,000 | 1,450 | 0 | 3,692 |

## Independent capture audit

`root-performance-audit.json` independently verified both complete frozen-source pairs (6900 samples total) with the same source/profile closure and candidate binary. It found unchanged accounting fields in all 56 matched groups per pair. The earlier pair retains the `reference-conditional-256x4-sumifs` parse latency review trigger (+7.7169%, bootstrap interval [−1.9621%, +10.9134%]) and one +172 KiB RSS trigger. The latest pair ranges from −3.0467% to +3.5644% latency and −1.5573% to +6.6667% RSS; it has no >5% latency trigger and eight RSS triggers ranging from +168 KiB to +212 KiB. These are retained review observations, not causal claims.

The resolver is an immutable borrowing fixture. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
