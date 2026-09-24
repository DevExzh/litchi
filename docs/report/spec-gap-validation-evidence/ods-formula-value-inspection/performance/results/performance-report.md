# ODS value-inspection evaluator performance profile

The baseline is committed `d623f3c2ecc0c837017f700174656f0e443759a5`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 47,332 | 46,380 | -2.0% | 2,311 | 2,311 | 88 | 88 | 3,764 | 3,824 |
| array-control-16x16-arithmetic | parse-evaluate | 61,985 | 62,460 | +0.8% | 2,311 | 2,311 | 352 | 352 | 3,752 | 3,776 |
| array-control-16x16-sin | evaluate | 128,643 | 127,948 | -0.5% | 2,311 | 2,311 | 2,144 | 2,144 | 4,100 | 4,208 |
| array-control-16x16-sin | parse-evaluate | 142,640 | 142,793 | +0.1% | 2,311 | 2,311 | 2,412 | 2,412 | 4,116 | 4,220 |
| array-control-4x4-arithmetic | evaluate | 4,509 | 4,375 | -3.0% | 151 | 151 | 1,120 | 1,120 | 3,516 | 3,584 |
| array-control-4x4-arithmetic | parse-evaluate | 5,185 | 5,171 | -0.3% | 151 | 151 | 2,240 | 2,240 | 3,508 | 3,584 |
| array-control-4x4-sin | evaluate | 9,941 | 9,910 | -0.3% | 151 | 151 | 3,840 | 3,840 | 3,864 | 3,928 |
| array-control-4x4-sin | parse-evaluate | 10,757 | 10,716 | -0.4% | 151 | 151 | 5,040 | 5,040 | 3,840 | 3,960 |
| concat-borrowed-literals | evaluate | 720 | 717 | -0.4% | 21 | 21 | 6,000 | 6,000 | 3,212 | 3,324 |
| concat-borrowed-literals | parse-evaluate | 895 | 864 | -3.5% | 21 | 21 | 9,000 | 9,000 | 3,188 | 3,344 |
| concat-growth-chain | evaluate | 1,916 | 1,911 | -0.3% | 20 | 20 | 10,000 | 10,000 | 3,204 | 3,344 |
| concat-growth-chain | parse-evaluate | 2,340 | 2,297 | -1.8% | 20 | 20 | 17,000 | 17,000 | 3,236 | 3,344 |
| concat-owned-left | evaluate | 1,081 | 1,065 | -1.5% | 24 | 24 | 8,000 | 8,000 | 3,224 | 3,200 |
| concat-owned-left | parse-evaluate | 1,331 | 1,301 | -2.3% | 24 | 24 | 13,000 | 13,000 | 3,236 | 3,212 |
| concat-owned-right | evaluate | 1,082 | 1,085 | +0.3% | 24 | 24 | 8,000 | 8,000 | 3,240 | 3,324 |
| concat-owned-right | parse-evaluate | 1,341 | 1,320 | -1.6% | 24 | 24 | 13,000 | 13,000 | 3,260 | 3,340 |
| database-control-dstdev | evaluate | 4,000 | 3,940 | -1.5% | 30 | 30 | 20 | 20 | 3,600 | 3,528 |
| database-control-dstdev | parse-evaluate | 5,180 | 5,140 | -0.8% | 30 | 30 | 29 | 29 | 3,680 | 3,508 |
| database-control-dsum | evaluate | 4,090 | 4,010 | -2.0% | 28 | 28 | 20 | 20 | 3,636 | 3,524 |
| database-control-dsum | parse-evaluate | 5,170 | 5,180 | +0.2% | 28 | 28 | 29 | 29 | 3,596 | 3,508 |
| database-control-dvar | evaluate | 3,920 | 3,910 | -0.3% | 28 | 28 | 20 | 20 | 3,600 | 3,640 |
| database-control-dvar | parse-evaluate | 5,070 | 5,040 | -0.6% | 28 | 28 | 29 | 29 | 3,588 | 3,628 |
| literal-aggregate-4x1-sum | evaluate | 1,820 | 1,860 | +2.2% | 15 | 15 | 10 | 10 | 3,624 | 3,548 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,650 | 2,660 | +0.4% | 15 | 15 | 23 | 23 | 3,636 | 3,580 |
| reference-aggregate-64x4-sum | evaluate | 10,335 | 10,357 | +0.2% | 16 | 16 | 32 | 32 | 3,612 | 3,612 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,692 | 10,752 | +0.6% | 16 | 16 | 60 | 60 | 3,556 | 3,580 |
| reference-array-16x4-arithmetic | evaluate | 9,495 | 9,555 | +0.6% | 16 | 16 | 800 | 800 | 3,504 | 3,540 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,834 | 9,828 | -0.1% | 16 | 16 | 1,280 | 1,280 | 3,528 | 3,528 |
| reference-conditional-256x4-sumifs | evaluate | 82,695 | 85,575 | +3.5% | 52 | 52 | 42 | 42 | 3,608 | 3,596 |
| reference-conditional-256x4-sumifs | parse-evaluate | 81,445 | 87,730 | +7.7% | 52 | 52 | 68 | 68 | 3,664 | 3,584 |
| reference-control-average | evaluate | 10,332 | 10,320 | -0.1% | 20 | 20 | 32 | 32 | 3,576 | 3,512 |
| reference-control-average | parse-evaluate | 10,755 | 10,725 | -0.3% | 20 | 20 | 60 | 60 | 3,632 | 3,612 |
| reference-control-counta | evaluate | 8,750 | 8,677 | -0.8% | 19 | 19 | 32 | 32 | 3,552 | 3,512 |
| reference-control-counta | parse-evaluate | 9,067 | 9,062 | -0.1% | 19 | 19 | 60 | 60 | 3,632 | 3,476 |
| representative-median | evaluate | 1,589 | 1,567 | -1.4% | 16 | 16 | 9,000 | 9,000 | 3,336 | 3,452 |
| representative-median | parse-evaluate | 1,923 | 1,876 | -2.4% | 16 | 16 | 14,000 | 14,000 | 3,368 | 3,416 |
| representative-percentrank | evaluate | 970 | 980 | +1.0% | 19 | 19 | 7,000 | 7,000 | 3,320 | 3,428 |
| representative-percentrank | parse-evaluate | 1,230 | 1,216 | -1.1% | 19 | 19 | 11,000 | 11,000 | 3,344 | 3,516 |
| representative-rank | evaluate | 936 | 929 | -0.7% | 12 | 12 | 7,000 | 7,000 | 3,336 | 3,416 |
| representative-rank | parse-evaluate | 1,172 | 1,141 | -2.6% | 12 | 12 | 11,000 | 11,000 | 3,328 | 3,460 |
| scalar-aggregate-sum | evaluate | 573 | 567 | -1.0% | 10 | 10 | 4,000 | 4,000 | 3,360 | 3,396 |
| scalar-aggregate-sum | parse-evaluate | 749 | 750 | +0.1% | 10 | 10 | 8,000 | 8,000 | 3,312 | 3,424 |
| scalar-control-arithmetic | evaluate | 571 | 570 | -0.2% | 10 | 10 | 5,000 | 5,000 | 3,220 | 3,296 |
| scalar-control-arithmetic | parse-evaluate | 703 | 704 | +0.1% | 10 | 10 | 8,000 | 8,000 | 3,204 | 3,276 |
| scalar-control-average | evaluate | 1,037 | 1,028 | -0.9% | 23 | 23 | 6,000 | 6,000 | 3,344 | 3,368 |
| scalar-control-average | parse-evaluate | 1,249 | 1,247 | -0.2% | 23 | 23 | 10,000 | 10,000 | 3,312 | 3,464 |
| scalar-control-counta | evaluate | 842 | 828 | -1.7% | 24 | 24 | 6,000 | 6,000 | 3,348 | 3,384 |
| scalar-control-counta | parse-evaluate | 1,065 | 1,081 | +1.5% | 24 | 24 | 10,000 | 10,000 | 3,356 | 3,376 |
| scalar-control-imsum | evaluate | 1,405 | 1,393 | -0.9% | 41 | 41 | 8,000 | 8,000 | 3,324 | 3,352 |
| scalar-control-imsum | parse-evaluate | 1,950 | 1,973 | +1.2% | 41 | 41 | 17,000 | 17,000 | 3,336 | 3,336 |
| scalar-control-sin | evaluate | 466 | 469 | +0.6% | 10 | 10 | 4,000 | 4,000 | 3,616 | 3,760 |
| scalar-control-sin | parse-evaluate | 645 | 653 | +1.2% | 10 | 10 | 8,000 | 8,000 | 3,624 | 3,716 |
| scalar-control-stdev | evaluate | 868 | 852 | -1.8% | 21 | 21 | 6,000 | 6,000 | 3,340 | 3,336 |
| scalar-control-stdev | parse-evaluate | 1,071 | 1,076 | +0.5% | 21 | 21 | 10,000 | 10,000 | 3,320 | 3,404 |
| scalar-control-var | evaluate | 856 | 845 | -1.3% | 19 | 19 | 6,000 | 6,000 | 3,388 | 3,432 |
| scalar-control-var | parse-evaluate | 1,054 | 1,057 | +0.3% | 19 | 19 | 10,000 | 10,000 | 3,328 | 3,420 |

## Value-inspection and conversion workloads

The candidate matrix covers all sixteen value-inspection and conversion functions over scalar identity, borrowed 64-cell references, matrix lifting, TYPE/N scalar-result exceptions, typed refusal, lazy projected branches, date/fraction parsing, separator transforms, sticky cancellation, and resource refusal. Input/output bytes and the reviewed domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-inspection-isnumber | evaluate | 832 | 0 | 832 | 367 | 4 | 0 | 9 | 7,048 | 7,048 | 0 | 3,528 |
| cancellation-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 780 | 4 | 0 | 37 | 12,516 | 8,383 | 0 | 3,512 |
| cancellation-inspection-numbervalue | evaluate | 832 | 0 | 832 | 372 | 5 | 0 | 9 | 7,048 | 7,048 | 0 | 3,516 |
| cancellation-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 780 | 5 | 0 | 37 | 12,528 | 8,386 | 0 | 3,512 |
| cancellation-inspection-type | evaluate | 832 | 0 | 832 | 392 | 3 | 0 | 10 | 3,992 | 3,928 | 0 | 3,660 |
| cancellation-inspection-type | parse-evaluate | 832 | 0 | 832 | 785 | 3 | 0 | 38 | 9,444 | 5,259 | 0 | 3,616 |
| cancellation-inspection-value | evaluate | 832 | 0 | 832 | 367 | 3 | 0 | 9 | 7,048 | 7,048 | 0 | 3,536 |
| cancellation-inspection-value | parse-evaluate | 832 | 0 | 832 | 760 | 3 | 0 | 37 | 12,504 | 8,380 | 0 | 3,516 |
| lazy-if-cache-isblank | evaluate | 832 | 0 | 832 | 5,120 | 69 | 2 | 27 | 3,416 | 2,200 | 176 | 3,576 |
| lazy-if-cache-isblank | parse-evaluate | 832 | 0 | 832 | 6,170 | 69 | 2 | 41 | 7,305 | 4,745 | 176 | 3,516 |
| lazy-if-cache-isnumber | evaluate | 832 | 0 | 832 | 5,210 | 71 | 2 | 27 | 3,416 | 2,200 | 176 | 3,528 |
| lazy-if-cache-isnumber | parse-evaluate | 832 | 0 | 832 | 6,070 | 71 | 2 | 41 | 7,306 | 4,746 | 176 | 3,596 |
| lazy-if-cache-istext | evaluate | 832 | 0 | 832 | 5,110 | 67 | 2 | 27 | 3,416 | 2,200 | 176 | 3,620 |
| lazy-if-cache-istext | parse-evaluate | 832 | 0 | 832 | 6,110 | 67 | 2 | 41 | 7,304 | 4,744 | 176 | 3,520 |
| matrix-inspection-isblank | evaluate | 832 | 0 | 832 | 2,250 | 43 | 0 | 11 | 2,920 | 2,296 | 352 | 3,528 |
| matrix-inspection-isblank | parse-evaluate | 832 | 0 | 832 | 3,170 | 43 | 0 | 24 | 6,056 | 3,992 | 352 | 3,588 |
| matrix-inspection-iserror | evaluate | 832 | 0 | 832 | 2,240 | 43 | 0 | 11 | 2,920 | 2,296 | 352 | 3,564 |
| matrix-inspection-iserror | parse-evaluate | 832 | 0 | 832 | 3,220 | 43 | 0 | 24 | 6,056 | 3,992 | 352 | 3,560 |
| matrix-inspection-iseven | evaluate | 832 | 0 | 832 | 2,290 | 45 | 0 | 11 | 2,920 | 2,296 | 352 | 3,532 |
| matrix-inspection-iseven | parse-evaluate | 832 | 0 | 832 | 3,240 | 45 | 0 | 24 | 6,055 | 3,991 | 352 | 3,548 |
| matrix-inspection-isnumber | evaluate | 832 | 0 | 832 | 2,220 | 44 | 0 | 11 | 2,920 | 2,296 | 352 | 3,576 |
| matrix-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 3,150 | 44 | 0 | 24 | 6,057 | 3,993 | 352 | 3,636 |
| matrix-inspection-istext | evaluate | 832 | 0 | 832 | 2,230 | 42 | 0 | 11 | 2,920 | 2,296 | 352 | 3,560 |
| matrix-inspection-istext | parse-evaluate | 832 | 0 | 832 | 3,200 | 42 | 0 | 24 | 6,055 | 3,991 | 352 | 3,568 |
| matrix-inspection-n | evaluate | 832 | 0 | 832 | 2,080 | 32 | 0 | 11 | 5,016 | 4,392 | 0 | 3,528 |
| matrix-inspection-n | parse-evaluate | 832 | 0 | 832 | 3,240 | 32 | 0 | 24 | 8,146 | 6,082 | 0 | 3,564 |
| matrix-inspection-value | evaluate | 832 | 0 | 832 | 2,340 | 44 | 0 | 11 | 2,920 | 2,296 | 352 | 3,544 |
| matrix-inspection-value | parse-evaluate | 832 | 0 | 832 | 3,280 | 44 | 0 | 24 | 6,054 | 3,990 | 352 | 3,576 |
| n-reference-intersection | evaluate | 0 | 0 | 0 | 1,650 | 11 | 1 | 10 | 3,992 | 3,928 | 0 | 3,576 |
| n-reference-intersection | parse-evaluate | 0 | 0 | 0 | 2,110 | 11 | 1 | 17 | 5,351 | 5,255 | 0 | 3,616 |
| numbervalue-invalid-separator | evaluate | 0 | 0 | 0 | 1,780 | 22 | 0 | 11 | 2,680 | 2,232 | 176 | 3,508 |
| numbervalue-invalid-separator | parse-evaluate | 0 | 0 | 0 | 2,250 | 22 | 0 | 17 | 4,051 | 3,571 | 176 | 3,512 |
| numbervalue-transform-decimal-comma | evaluate | 12 | 0 | 12 | 1,346 | 40 | 0 | 8,000 | 2,113,000 | 1,656 | 0 | 3,528 |
| numbervalue-transform-decimal-comma | parse-evaluate | 12 | 0 | 12 | 1,602 | 40 | 0 | 12,000 | 2,593,000 | 2,104 | 0 | 3,516 |
| numbervalue-transform-grouped-space | evaluate | 12 | 0 | 12 | 1,357 | 42 | 0 | 8,000 | 2,114,000 | 1,656 | 0 | 3,592 |
| numbervalue-transform-grouped-space | parse-evaluate | 12 | 0 | 12 | 1,622 | 42 | 0 | 12,000 | 2,595,000 | 2,105 | 0 | 3,584 |
| numbervalue-transform-percent | evaluate | 12 | 0 | 12 | 1,352 | 42 | 0 | 8,000 | 2,114,000 | 1,656 | 0 | 3,540 |
| numbervalue-transform-percent | parse-evaluate | 12 | 0 | 12 | 1,616 | 42 | 0 | 12,000 | 2,595,000 | 2,105 | 0 | 3,528 |
| reference-inspection-error-type | evaluate | 832 | 0 | 832 | 7,632 | 208 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,640 |
| reference-inspection-error-type | parse-evaluate | 832 | 0 | 832 | 8,067 | 208 | 64 | 64 | 33,668 | 8,385 | 5,632 | 3,508 |
| reference-inspection-isblank | evaluate | 832 | 0 | 832 | 7,512 | 205 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,516 |
| reference-inspection-isblank | parse-evaluate | 832 | 0 | 832 | 7,957 | 205 | 64 | 64 | 33,656 | 8,382 | 5,632 | 3,600 |
| reference-inspection-iserr | evaluate | 832 | 0 | 832 | 7,605 | 203 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,564 |
| reference-inspection-iserr | parse-evaluate | 832 | 0 | 832 | 7,945 | 203 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,528 |
| reference-inspection-iserror | evaluate | 832 | 0 | 832 | 7,585 | 205 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,528 |
| reference-inspection-iserror | parse-evaluate | 832 | 0 | 832 | 7,902 | 205 | 64 | 64 | 33,656 | 8,382 | 5,632 | 3,544 |
| reference-inspection-iseven | evaluate | 832 | 0 | 832 | 8,172 | 292 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,616 |
| reference-inspection-iseven | parse-evaluate | 832 | 0 | 832 | 8,650 | 292 | 64 | 64 | 33,652 | 8,381 | 5,632 | 3,532 |
| reference-inspection-islogical | evaluate | 832 | 0 | 832 | 7,542 | 207 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,508 |
| reference-inspection-islogical | parse-evaluate | 832 | 0 | 832 | 7,980 | 207 | 64 | 64 | 33,664 | 8,384 | 5,632 | 3,556 |
| reference-inspection-isna | evaluate | 832 | 0 | 832 | 7,557 | 202 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,528 |
| reference-inspection-isna | parse-evaluate | 832 | 0 | 832 | 7,927 | 202 | 64 | 64 | 33,644 | 8,379 | 5,632 | 3,528 |
| reference-inspection-isnontext | evaluate | 832 | 0 | 832 | 7,602 | 207 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,592 |
| reference-inspection-isnontext | parse-evaluate | 832 | 0 | 832 | 7,977 | 207 | 64 | 64 | 33,664 | 8,384 | 5,632 | 3,580 |
| reference-inspection-isnumber | evaluate | 832 | 0 | 832 | 7,565 | 206 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,532 |
| reference-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 7,937 | 206 | 64 | 64 | 33,660 | 8,383 | 5,632 | 3,564 |
| reference-inspection-isodd | evaluate | 832 | 0 | 832 | 8,225 | 291 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,528 |
| reference-inspection-isodd | parse-evaluate | 832 | 0 | 832 | 8,665 | 291 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,524 |
| reference-inspection-istext | evaluate | 832 | 0 | 832 | 7,542 | 204 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,528 |
| reference-inspection-istext | parse-evaluate | 832 | 0 | 832 | 7,972 | 204 | 64 | 64 | 33,652 | 8,381 | 5,632 | 3,580 |
| reference-inspection-numbervalue | evaluate | 832 | 0 | 832 | 16,270 | 299 | 64 | 264 | 28,772 | 7,060 | 5,632 | 3,516 |
| reference-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 16,695 | 299 | 64 | 292 | 34,252 | 8,398 | 5,632 | 3,500 |
| reference-inspection-value | evaluate | 832 | 0 | 832 | 8,742 | 291 | 64 | 36 | 28,192 | 7,048 | 5,632 | 3,528 |
| reference-inspection-value | parse-evaluate | 832 | 0 | 832 | 9,247 | 291 | 64 | 64 | 33,648 | 8,380 | 5,632 | 3,588 |
| resource-inspection-isnumber | evaluate | 832 | 0 | 832 | 402 | 11 | 0 | 16 | 1,184 | 296 | 0 | 3,580 |
| resource-inspection-isnumber | parse-evaluate | 832 | 0 | 832 | 815 | 11 | 0 | 44 | 6,652 | 1,631 | 0 | 3,512 |
| resource-inspection-numbervalue | evaluate | 832 | 0 | 832 | 397 | 14 | 0 | 16 | 1,184 | 296 | 0 | 3,508 |
| resource-inspection-numbervalue | parse-evaluate | 832 | 0 | 832 | 817 | 14 | 0 | 44 | 6,664 | 1,634 | 0 | 3,552 |
| resource-inspection-type | evaluate | 832 | 0 | 832 | 757 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,432 |
| resource-inspection-type | parse-evaluate | 832 | 0 | 832 | 1,082 | 8 | 0 | 52 | 16,940 | 4,139 | 0 | 3,452 |
| resource-inspection-value | evaluate | 832 | 0 | 832 | 405 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,520 |
| resource-inspection-value | parse-evaluate | 832 | 0 | 832 | 807 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,580 |
| scalar-inspection-error-type | evaluate | 0 | 0 | 0 | 692 | 19 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,572 |
| scalar-inspection-error-type | parse-evaluate | 0 | 0 | 0 | 915 | 19 | 0 | 9,000 | 1,481,000 | 1,449 | 0 | 3,564 |
| scalar-inspection-isblank | evaluate | 0 | 0 | 0 | 700 | 12 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,568 |
| scalar-inspection-isblank | parse-evaluate | 0 | 0 | 0 | 901 | 12 | 0 | 9,000 | 1,476,000 | 1,444 | 0 | 3,548 |
| scalar-inspection-iserr | evaluate | 0 | 0 | 0 | 692 | 17 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,600 |
| scalar-inspection-iserr | parse-evaluate | 0 | 0 | 0 | 911 | 17 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,600 |
| scalar-inspection-iserror | evaluate | 0 | 0 | 0 | 698 | 16 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,572 |
| scalar-inspection-iserror | parse-evaluate | 0 | 0 | 0 | 909 | 16 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,576 |
| scalar-inspection-iseven | evaluate | 0 | 0 | 0 | 712 | 13 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,636 |
| scalar-inspection-iseven | parse-evaluate | 0 | 0 | 0 | 921 | 13 | 0 | 9,000 | 1,475,000 | 1,443 | 0 | 3,584 |
| scalar-inspection-islogical | evaluate | 0 | 0 | 0 | 827 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,580 |
| scalar-inspection-islogical | parse-evaluate | 0 | 0 | 0 | 1,068 | 20 | 0 | 9,000 | 1,482,000 | 1,450 | 0 | 3,536 |
| scalar-inspection-isna | evaluate | 0 | 0 | 0 | 682 | 13 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,572 |
| scalar-inspection-isna | parse-evaluate | 0 | 0 | 0 | 889 | 13 | 0 | 9,000 | 1,475,000 | 1,443 | 0 | 3,580 |
| scalar-inspection-isnontext | evaluate | 0 | 0 | 0 | 706 | 16 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,660 |
| scalar-inspection-isnontext | parse-evaluate | 0 | 0 | 0 | 921 | 16 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,576 |
| scalar-inspection-isnumber | evaluate | 0 | 0 | 0 | 704 | 15 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,528 |
| scalar-inspection-isnumber | parse-evaluate | 0 | 0 | 0 | 923 | 15 | 0 | 9,000 | 1,477,000 | 1,445 | 0 | 3,616 |
| scalar-inspection-isodd | evaluate | 0 | 0 | 0 | 707 | 12 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,568 |
| scalar-inspection-isodd | parse-evaluate | 0 | 0 | 0 | 929 | 12 | 0 | 9,000 | 1,474,000 | 1,442 | 0 | 3,548 |
| scalar-inspection-istext | evaluate | 0 | 0 | 0 | 700 | 14 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,624 |
| scalar-inspection-istext | parse-evaluate | 0 | 0 | 0 | 923 | 14 | 0 | 9,000 | 1,478,000 | 1,446 | 0 | 3,616 |
| scalar-inspection-n | evaluate | 0 | 0 | 0 | 1,124 | 14 | 0 | 7,000 | 3,592,000 | 3,528 | 0 | 3,596 |
| scalar-inspection-n | parse-evaluate | 0 | 0 | 0 | 1,360 | 14 | 0 | 11,000 | 4,050,000 | 3,954 | 0 | 3,528 |
| scalar-inspection-na | evaluate | 0 | 0 | 0 | 504 | 5 | 0 | 4,000 | 632,000 | 632 | 0 | 3,600 |
| scalar-inspection-na | parse-evaluate | 0 | 0 | 0 | 603 | 5 | 0 | 6,000 | 1,021,000 | 1,021 | 0 | 3,516 |
| scalar-inspection-numbervalue | evaluate | 0 | 0 | 0 | 1,352 | 40 | 0 | 8,000 | 2,113,000 | 1,656 | 0 | 3,560 |
| scalar-inspection-numbervalue | parse-evaluate | 0 | 0 | 0 | 1,616 | 40 | 0 | 12,000 | 2,593,000 | 2,104 | 0 | 3,548 |
| scalar-inspection-type | evaluate | 0 | 0 | 0 | 965 | 14 | 0 | 7,000 | 3,592,000 | 3,528 | 0 | 3,628 |
| scalar-inspection-type | parse-evaluate | 0 | 0 | 0 | 1,190 | 14 | 0 | 11,000 | 4,052,000 | 3,956 | 0 | 3,592 |
| scalar-inspection-value | evaluate | 0 | 0 | 0 | 771 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,528 |
| scalar-inspection-value | parse-evaluate | 0 | 0 | 0 | 985 | 20 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,548 |
| shape-refusal-numbervalue | evaluate | 0 | 0 | 0 | 2,000 | 25 | 0 | 12 | 1,944 | 1,576 | 0 | 3,552 |
| shape-refusal-numbervalue | parse-evaluate | 0 | 0 | 0 | 2,670 | 25 | 0 | 21 | 3,325 | 2,925 | 0 | 3,608 |
| shape-refusal-value | evaluate | 0 | 0 | 0 | 2,000 | 17 | 0 | 12 | 1,944 | 1,576 | 0 | 3,624 |
| shape-refusal-value | parse-evaluate | 0 | 0 | 0 | 2,640 | 17 | 0 | 21 | 3,319 | 2,919 | 0 | 3,532 |
| type-array-metadata | evaluate | 0 | 0 | 0 | 2,010 | 34 | 0 | 11 | 5,016 | 4,392 | 0 | 3,548 |
| type-array-metadata | parse-evaluate | 0 | 0 | 0 | 3,250 | 34 | 0 | 24 | 8,149 | 6,085 | 0 | 3,608 |
| type-reference-scan | evaluate | 832 | 0 | 832 | 3,195 | 76 | 64 | 40 | 15,968 | 3,928 | 0 | 3,672 |
| type-reference-scan | parse-evaluate | 832 | 0 | 832 | 3,602 | 76 | 64 | 68 | 21,420 | 5,259 | 0 | 3,508 |
| value-date-fraction-value-date | evaluate | 22 | 0 | 22 | 761 | 30 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,512 |
| value-date-fraction-value-date | parse-evaluate | 22 | 0 | 22 | 978 | 30 | 0 | 9,000 | 1,484,000 | 1,452 | 0 | 3,520 |
| value-date-fraction-value-mixed-fraction | evaluate | 22 | 0 | 22 | 786 | 20 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,572 |
| value-date-fraction-value-mixed-fraction | parse-evaluate | 22 | 0 | 22 | 991 | 20 | 0 | 9,000 | 1,479,000 | 1,447 | 0 | 3,588 |
| value-date-fraction-value-time | evaluate | 22 | 0 | 22 | 757 | 26 | 0 | 5,000 | 1,016,000 | 1,016 | 0 | 3,528 |
| value-date-fraction-value-time | parse-evaluate | 22 | 0 | 22 | 982 | 26 | 0 | 9,000 | 1,482,000 | 1,450 | 0 | 3,588 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
