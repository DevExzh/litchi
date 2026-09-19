# ODS order-statistics evaluator performance profile

The baseline is committed `d6c0485365644a5525b7302ba0b3666d43df1b1f`. It contributes matched arithmetic, SIN, IMSUM, DSUM, SUM, SUMIFS, AVERAGE, COUNTA, DVAR, and DSTDEV controls. The eight order-statistics reducers are candidate-only evidence; each cell is the p50 across fifteen fresh child processes, with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,810 | 45,600 | -0.5% | 88 | 88 | 3,444 | 3,396 |
| array-control-16x16-arithmetic | parse-evaluate | 58,412 | 57,992 | -0.7% | 352 | 352 | 3,420 | 3,448 |
| array-control-16x16-sin | evaluate | 125,548 | 127,470 | +1.5% | 2,144 | 2,144 | 3,664 | 3,708 |
| array-control-16x16-sin | parse-evaluate | 139,103 | 138,938 | -0.1% | 2,412 | 2,412 | 3,624 | 3,612 |
| array-control-4x4-arithmetic | evaluate | 4,430 | 4,416 | -0.3% | 1,120 | 1,120 | 3,424 | 3,448 |
| array-control-4x4-arithmetic | parse-evaluate | 5,251 | 5,135 | -2.2% | 2,240 | 2,240 | 3,456 | 3,424 |
| array-control-4x4-sin | evaluate | 9,651 | 9,639 | -0.1% | 3,840 | 3,840 | 3,676 | 3,724 |
| array-control-4x4-sin | parse-evaluate | 10,540 | 10,537 | -0.0% | 5,040 | 5,040 | 3,668 | 3,724 |
| database-control-dstdev | evaluate | 3,800 | 3,770 | -0.8% | 20 | 20 | 3,420 | 3,484 |
| database-control-dstdev | parse-evaluate | 4,360 | 4,380 | +0.5% | 29 | 29 | 3,436 | 3,420 |
| database-control-dsum | evaluate | 3,840 | 3,850 | +0.3% | 20 | 20 | 3,424 | 3,420 |
| database-control-dsum | parse-evaluate | 4,420 | 4,410 | -0.2% | 29 | 29 | 3,404 | 3,448 |
| database-control-dvar | evaluate | 3,760 | 3,780 | +0.5% | 20 | 20 | 3,480 | 3,432 |
| database-control-dvar | parse-evaluate | 4,320 | 4,381 | +1.4% | 29 | 29 | 3,424 | 3,452 |
| literal-aggregate-4x1-sum | evaluate | 1,790 | 1,760 | -1.7% | 10 | 10 | 3,432 | 3,416 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,480 | 2,420 | -2.4% | 23 | 23 | 3,448 | 3,468 |
| reference-aggregate-64x4-sum | evaluate | 10,092 | 10,207 | +1.1% | 32 | 32 | 3,424 | 3,448 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,502 | 10,567 | +0.6% | 60 | 60 | 3,424 | 3,452 |
| reference-array-16x4-arithmetic | evaluate | 9,338 | 9,222 | -1.2% | 800 | 800 | 3,424 | 3,444 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,752 | 9,534 | -2.2% | 1,280 | 1,280 | 3,432 | 3,436 |
| reference-conditional-256x4-sumifs | evaluate | 78,765 | 79,185 | +0.5% | 42 | 42 | 3,428 | 3,456 |
| reference-conditional-256x4-sumifs | parse-evaluate | 74,820 | 79,600 | +6.4% | 68 | 68 | 3,444 | 3,444 |
| reference-control-average | evaluate | 10,155 | 10,217 | +0.6% | 32 | 32 | 3,424 | 3,456 |
| reference-control-average | parse-evaluate | 10,572 | 10,632 | +0.6% | 60 | 60 | 3,396 | 3,452 |
| reference-control-counta | evaluate | 8,547 | 8,582 | +0.4% | 32 | 32 | 3,436 | 3,456 |
| reference-control-counta | parse-evaluate | 9,000 | 8,980 | -0.2% | 60 | 60 | 3,448 | 3,440 |
| scalar-aggregate-sum | evaluate | 587 | 572 | -2.6% | 4,000 | 4,000 | 3,376 | 3,236 |
| scalar-aggregate-sum | parse-evaluate | 779 | 765 | -1.8% | 8,000 | 8,000 | 3,380 | 3,268 |
| scalar-control-arithmetic | evaluate | 560 | 560 | +0.0% | 5,000 | 5,000 | 3,232 | 3,172 |
| scalar-control-arithmetic | parse-evaluate | 691 | 690 | -0.1% | 8,000 | 8,000 | 3,260 | 3,168 |
| scalar-control-average | evaluate | 1,009 | 1,023 | +1.4% | 6,000 | 6,000 | 3,376 | 3,288 |
| scalar-control-average | parse-evaluate | 1,226 | 1,231 | +0.4% | 10,000 | 10,000 | 3,352 | 3,188 |
| scalar-control-counta | evaluate | 815 | 807 | -1.0% | 6,000 | 6,000 | 3,388 | 3,172 |
| scalar-control-counta | parse-evaluate | 1,059 | 1,057 | -0.2% | 10,000 | 10,000 | 3,372 | 3,172 |
| scalar-control-imsum | evaluate | 1,368 | 1,372 | +0.3% | 8,000 | 8,000 | 3,236 | 3,240 |
| scalar-control-imsum | parse-evaluate | 1,921 | 1,924 | +0.2% | 17,000 | 17,000 | 3,400 | 3,240 |
| scalar-control-sin | evaluate | 458 | 456 | -0.4% | 4,000 | 4,000 | 3,572 | 3,512 |
| scalar-control-sin | parse-evaluate | 643 | 642 | -0.2% | 8,000 | 8,000 | 3,556 | 3,540 |
| scalar-control-stdev | evaluate | 839 | 841 | +0.2% | 6,000 | 6,000 | 3,236 | 3,224 |
| scalar-control-stdev | parse-evaluate | 1,048 | 1,061 | +1.2% | 10,000 | 10,000 | 3,404 | 3,224 |
| scalar-control-var | evaluate | 824 | 833 | +1.1% | 6,000 | 6,000 | 3,356 | 3,176 |
| scalar-control-var | parse-evaluate | 1,044 | 1,038 | -0.6% | 10,000 | 10,000 | 3,372 | 3,320 |

## Candidate order-statistics reducer workloads

The bounded matrix covers scalar variadic calls, inline arrays (including array k/rank parameters), rectangular references, zero-read shape/domain refusals, formula-error precedence, typed reference-cell refusal, and projected reducers at 64 rows. MEDIAN, MODE, LARGE, and SMALL add sorted, reverse, duplicate, and deterministic pseudo-random reference lanes; MEDIAN, LARGE, SMALL, PERCENTILE, PERCENTRANK, QUARTILE, and RANK add 256-row and 1024-row projected scaling. A dedicated LARGE row lifts a projected `MUNIT([.G1:.G2])` parameter so complete reference caching and position-sensitive parameter evaluation remain observable. Resolver reads expose whether a reducer is projected once or rebuilt per output cell.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-order-large | evaluate | 2,570 | 48 | 0 | 13 | 5,064 | 4,424 | 0 | 3,432 |
| array-order-large | parse-evaluate | 2,940 | 48 | 0 | 21 | 6,427 | 5,275 | 0 | 3,472 |
| array-order-median | evaluate | 2,600 | 46 | 0 | 13 | 5,064 | 4,424 | 0 | 3,440 |
| array-order-median | parse-evaluate | 2,950 | 46 | 0 | 21 | 6,426 | 5,274 | 0 | 3,452 |
| array-order-mode | evaluate | 3,010 | 71 | 0 | 16 | 7,336 | 5,640 | 0 | 3,472 |
| array-order-mode | parse-evaluate | 3,820 | 71 | 0 | 27 | 10,396 | 7,324 | 0 | 3,464 |
| array-order-percentile | evaluate | 2,750 | 56 | 0 | 13 | 5,064 | 4,424 | 0 | 3,436 |
| array-order-percentile | parse-evaluate | 3,090 | 56 | 0 | 21 | 6,435 | 5,283 | 0 | 3,480 |
| array-order-percentrank | evaluate | 2,330 | 46 | 0 | 12 | 5,720 | 4,840 | 0 | 3,432 |
| array-order-percentrank | parse-evaluate | 2,960 | 46 | 0 | 21 | 8,627 | 6,467 | 0 | 3,444 |
| array-order-quartile | evaluate | 2,690 | 51 | 0 | 13 | 5,064 | 4,424 | 0 | 3,452 |
| array-order-quartile | parse-evaluate | 3,120 | 51 | 0 | 21 | 6,430 | 5,278 | 0 | 3,456 |
| array-order-rank | evaluate | 2,230 | 36 | 0 | 12 | 6,552 | 5,160 | 0 | 3,480 |
| array-order-rank | parse-evaluate | 2,670 | 36 | 0 | 20 | 7,914 | 6,010 | 0 | 3,456 |
| array-order-small | evaluate | 2,540 | 48 | 0 | 13 | 5,064 | 4,424 | 0 | 3,456 |
| array-order-small | parse-evaluate | 2,990 | 48 | 0 | 21 | 6,427 | 5,275 | 0 | 3,460 |
| array-parameter-order-large | evaluate | 3,580 | 68 | 0 | 18 | 6,360 | 4,856 | 0 | 3,460 |
| array-parameter-order-large | parse-evaluate | 4,520 | 68 | 0 | 31 | 9,492 | 6,548 | 0 | 3,464 |
| array-parameter-order-small | evaluate | 3,560 | 68 | 0 | 18 | 6,360 | 4,856 | 0 | 3,460 |
| array-parameter-order-small | parse-evaluate | 4,560 | 68 | 0 | 31 | 9,492 | 6,548 | 0 | 3,480 |
| error-order-large | evaluate | 4,251 | 82 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,420 |
| error-order-large | parse-evaluate | 4,689 | 82 | 64 | 1,920 | 513,120 | 5,774 | 0 | 3,444 |
| error-order-median | evaluate | 4,064 | 80 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,464 |
| error-order-median | parse-evaluate | 4,552 | 80 | 64 | 1,920 | 513,040 | 5,773 | 0 | 3,432 |
| error-order-mode | evaluate | 4,048 | 78 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,456 |
| error-order-mode | parse-evaluate | 4,466 | 78 | 64 | 1,920 | 512,880 | 5,771 | 0 | 3,440 |
| error-order-percentile | evaluate | 4,258 | 89 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,456 |
| error-order-percentile | parse-evaluate | 4,699 | 89 | 64 | 1,920 | 513,680 | 5,781 | 0 | 3,460 |
| error-order-percentrank | evaluate | 4,513 | 156 | 64 | 1,040 | 420,480 | 4,680 | 0 | 3,460 |
| error-order-percentrank | parse-evaluate | 4,956 | 156 | 64 | 1,600 | 530,560 | 6,024 | 0 | 3,456 |
| error-order-quartile | evaluate | 4,243 | 85 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,452 |
| error-order-quartile | parse-evaluate | 4,654 | 85 | 64 | 1,920 | 513,360 | 5,777 | 0 | 3,440 |
| error-order-rank | evaluate | 4,361 | 149 | 64 | 960 | 400,000 | 4,552 | 0 | 3,444 |
| error-order-rank | parse-evaluate | 4,786 | 149 | 64 | 1,520 | 509,520 | 5,889 | 0 | 3,460 |
| error-order-small | evaluate | 4,275 | 82 | 64 | 1,360 | 403,840 | 4,440 | 0 | 3,468 |
| error-order-small | parse-evaluate | 4,675 | 82 | 64 | 1,920 | 513,120 | 5,774 | 0 | 3,432 |
| nested-projected-order-1024-large | evaluate | 2,275,001 | 161,817 | 8,192 | 4,160 | 1,119,864 | 758,088 | 0 | 4,072 |
| nested-projected-order-1024-large | parse-evaluate | 2,276,072 | 161,817 | 8,192 | 4,173 | 1,122,154 | 759,866 | 0 | 4,040 |
| nested-projected-order-1024-median | evaluate | 2,287,431 | 161,810 | 8,192 | 4,160 | 1,119,864 | 758,088 | 0 | 3,988 |
| nested-projected-order-1024-median | parse-evaluate | 2,284,561 | 161,810 | 8,192 | 4,173 | 1,122,153 | 759,865 | 0 | 4,068 |
| nested-projected-order-1024-percentile | evaluate | 2,287,621 | 161,824 | 8,192 | 4,160 | 1,119,864 | 758,088 | 0 | 3,972 |
| nested-projected-order-1024-percentile | parse-evaluate | 2,292,951 | 161,824 | 8,192 | 4,173 | 1,122,161 | 759,873 | 0 | 3,988 |
| nested-projected-order-1024-percentrank | evaluate | 1,101,535 | 28,737 | 8,192 | 4,153 | 1,055,848 | 726,120 | 0 | 4,088 |
| nested-projected-order-1024-percentrank | parse-evaluate | 1,109,996 | 28,737 | 8,192 | 4,166 | 1,058,148 | 727,908 | 0 | 3,996 |
| nested-projected-order-1024-quartile | evaluate | 2,286,671 | 161,820 | 8,192 | 4,160 | 1,119,864 | 758,088 | 0 | 3,976 |
| nested-projected-order-1024-quartile | parse-evaluate | 2,282,261 | 161,820 | 8,192 | 4,173 | 1,122,157 | 759,869 | 0 | 4,044 |
| nested-projected-order-1024-rank | evaluate | 1,090,905 | 28,730 | 8,192 | 4,153 | 1,055,848 | 726,120 | 0 | 4,032 |
| nested-projected-order-1024-rank | parse-evaluate | 1,097,185 | 28,730 | 8,192 | 4,166 | 1,058,141 | 727,901 | 0 | 4,056 |
| nested-projected-order-1024-small | evaluate | 2,283,631 | 161,817 | 8,192 | 4,160 | 1,119,864 | 758,088 | 0 | 4,104 |
| nested-projected-order-1024-small | parse-evaluate | 2,280,971 | 161,817 | 8,192 | 4,173 | 1,122,154 | 759,866 | 0 | 4,084 |
| nested-projected-order-256-large | evaluate | 510,037 | 34,194 | 2,048 | 2,164 | 568,560 | 192,840 | 0 | 3,480 |
| nested-projected-order-256-large | parse-evaluate | 515,822 | 34,194 | 2,048 | 2,190 | 573,136 | 194,616 | 0 | 3,444 |
| nested-projected-order-256-median | evaluate | 509,722 | 34,187 | 2,048 | 2,164 | 568,560 | 192,840 | 0 | 3,480 |
| nested-projected-order-256-median | parse-evaluate | 513,777 | 34,187 | 2,048 | 2,190 | 573,134 | 194,615 | 0 | 3,404 |
| nested-projected-order-256-percentile | evaluate | 512,327 | 34,201 | 2,048 | 2,164 | 568,560 | 192,840 | 0 | 3,460 |
| nested-projected-order-256-percentile | parse-evaluate | 512,887 | 34,201 | 2,048 | 2,190 | 573,150 | 194,623 | 0 | 3,468 |
| nested-projected-order-256-percentrank | evaluate | 264,516 | 7,233 | 2,048 | 2,154 | 538,832 | 185,448 | 0 | 3,468 |
| nested-projected-order-256-percentrank | parse-evaluate | 265,391 | 7,233 | 2,048 | 2,180 | 543,428 | 187,234 | 0 | 3,452 |
| nested-projected-order-256-quartile | evaluate | 510,982 | 34,197 | 2,048 | 2,164 | 568,560 | 192,840 | 0 | 3,456 |
| nested-projected-order-256-quartile | parse-evaluate | 511,107 | 34,197 | 2,048 | 2,190 | 573,142 | 194,619 | 0 | 3,456 |
| nested-projected-order-256-rank | evaluate | 262,476 | 7,226 | 2,048 | 2,154 | 538,832 | 185,448 | 0 | 3,468 |
| nested-projected-order-256-rank | parse-evaluate | 262,791 | 7,226 | 2,048 | 2,180 | 543,414 | 187,227 | 0 | 3,496 |
| nested-projected-order-256-small | evaluate | 509,817 | 34,194 | 2,048 | 2,164 | 568,560 | 192,840 | 0 | 3,452 |
| nested-projected-order-256-small | parse-evaluate | 511,327 | 34,194 | 2,048 | 2,190 | 573,136 | 194,616 | 0 | 3,484 |
| nested-projected-order-64-large | evaluate | 116,345 | 6,988 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,460 |
| nested-projected-order-64-large | parse-evaluate | 118,278 | 6,988 | 512 | 1,284 | 310,680 | 53,302 | 0 | 3,440 |
| nested-projected-order-64-median | evaluate | 117,273 | 6,981 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,444 |
| nested-projected-order-64-median | parse-evaluate | 118,005 | 6,981 | 512 | 1,284 | 310,676 | 53,301 | 0 | 3,452 |
| nested-projected-order-64-mode | evaluate | 113,195 | 6,589 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,452 |
| nested-projected-order-64-mode | parse-evaluate | 114,525 | 6,589 | 512 | 1,284 | 310,668 | 53,299 | 0 | 3,460 |
| nested-projected-order-64-percentile | evaluate | 116,700 | 6,995 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,444 |
| nested-projected-order-64-percentile | parse-evaluate | 118,303 | 6,995 | 512 | 1,284 | 310,708 | 53,309 | 0 | 3,448 |
| nested-projected-order-64-percentrank | evaluate | 71,680 | 1,857 | 512 | 1,220 | 291,232 | 50,280 | 0 | 3,476 |
| nested-projected-order-64-percentrank | parse-evaluate | 71,203 | 1,857 | 512 | 1,272 | 300,416 | 52,064 | 0 | 3,448 |
| nested-projected-order-64-quartile | evaluate | 117,238 | 6,991 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,452 |
| nested-projected-order-64-quartile | parse-evaluate | 118,253 | 6,991 | 512 | 1,284 | 310,692 | 53,305 | 0 | 3,456 |
| nested-projected-order-64-rank | evaluate | 71,002 | 1,850 | 512 | 1,220 | 291,232 | 50,280 | 0 | 3,448 |
| nested-projected-order-64-rank | parse-evaluate | 72,513 | 1,850 | 512 | 1,272 | 300,388 | 52,057 | 0 | 3,452 |
| nested-projected-order-64-small | evaluate | 117,080 | 6,988 | 512 | 1,232 | 301,536 | 51,528 | 0 | 3,452 |
| nested-projected-order-64-small | parse-evaluate | 118,925 | 6,988 | 512 | 1,284 | 310,680 | 53,302 | 0 | 3,452 |
| nested-projected-order-parameter-large | evaluate | 14,710 | 281 | 26 | 69 | 12,856 | 6,216 | 0 | 3,416 |
| nested-projected-order-parameter-large | parse-evaluate | 16,250 | 281 | 26 | 92 | 17,045 | 8,869 | 0 | 3,468 |
| reference-order-duplicates-256-large | evaluate | 250,566 | 25,319 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,416 |
| reference-order-duplicates-256-large | parse-evaluate | 252,946 | 25,319 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,424 |
| reference-order-duplicates-256-median | evaluate | 252,261 | 25,317 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,472 |
| reference-order-duplicates-256-median | parse-evaluate | 251,276 | 25,317 | 1,024 | 56 | 43,548 | 13,454 | 0 | 3,436 |
| reference-order-duplicates-256-mode | evaluate | 261,441 | 26,337 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,452 |
| reference-order-duplicates-256-mode | parse-evaluate | 259,641 | 26,337 | 1,024 | 56 | 43,544 | 13,452 | 0 | 3,444 |
| reference-order-duplicates-256-small | evaluate | 253,131 | 25,319 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,456 |
| reference-order-duplicates-256-small | parse-evaluate | 253,651 | 25,319 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,436 |
| reference-order-duplicates-64-large | evaluate | 52,972 | 5,034 | 256 | 76 | 32,480 | 5,976 | 0 | 3,456 |
| reference-order-duplicates-64-large | parse-evaluate | 53,292 | 5,034 | 256 | 104 | 37,944 | 7,310 | 0 | 3,432 |
| reference-order-duplicates-64-median | evaluate | 52,947 | 5,032 | 256 | 76 | 32,480 | 5,976 | 0 | 3,468 |
| reference-order-duplicates-64-median | parse-evaluate | 53,357 | 5,032 | 256 | 104 | 37,940 | 7,309 | 0 | 3,456 |
| reference-order-duplicates-64-mode | evaluate | 54,882 | 5,284 | 256 | 76 | 32,480 | 5,976 | 0 | 3,440 |
| reference-order-duplicates-64-mode | parse-evaluate | 55,377 | 5,284 | 256 | 104 | 37,932 | 7,307 | 0 | 3,448 |
| reference-order-duplicates-64-small | evaluate | 52,920 | 5,034 | 256 | 76 | 32,480 | 5,976 | 0 | 3,456 |
| reference-order-duplicates-64-small | parse-evaluate | 53,312 | 5,034 | 256 | 104 | 37,944 | 7,310 | 0 | 3,436 |
| reference-order-random-256-large | evaluate | 274,716 | 27,735 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,448 |
| reference-order-random-256-large | parse-evaluate | 275,376 | 27,735 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,452 |
| reference-order-random-256-median | evaluate | 274,431 | 27,733 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,420 |
| reference-order-random-256-median | parse-evaluate | 275,386 | 27,733 | 1,024 | 56 | 43,548 | 13,454 | 0 | 3,464 |
| reference-order-random-256-mode | evaluate | 284,121 | 28,753 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,460 |
| reference-order-random-256-mode | parse-evaluate | 282,851 | 28,753 | 1,024 | 56 | 43,544 | 13,452 | 0 | 3,456 |
| reference-order-random-256-small | evaluate | 274,561 | 27,735 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,452 |
| reference-order-random-256-small | parse-evaluate | 274,926 | 27,735 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,480 |
| reference-order-random-64-large | evaluate | 56,635 | 5,433 | 256 | 76 | 32,480 | 5,976 | 0 | 3,456 |
| reference-order-random-64-large | parse-evaluate | 56,922 | 5,433 | 256 | 104 | 37,944 | 7,310 | 0 | 3,460 |
| reference-order-random-64-median | evaluate | 56,462 | 5,431 | 256 | 76 | 32,480 | 5,976 | 0 | 3,452 |
| reference-order-random-64-median | parse-evaluate | 56,925 | 5,431 | 256 | 104 | 37,940 | 7,309 | 0 | 3,444 |
| reference-order-random-64-mode | evaluate | 58,607 | 5,683 | 256 | 76 | 32,480 | 5,976 | 0 | 3,456 |
| reference-order-random-64-mode | parse-evaluate | 58,925 | 5,683 | 256 | 104 | 37,932 | 7,307 | 0 | 3,468 |
| reference-order-random-64-small | evaluate | 56,547 | 5,433 | 256 | 76 | 32,480 | 5,976 | 0 | 3,460 |
| reference-order-random-64-small | parse-evaluate | 56,957 | 5,433 | 256 | 104 | 37,944 | 7,310 | 0 | 3,468 |
| reference-order-reverse-256-large | evaluate | 258,166 | 25,992 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,444 |
| reference-order-reverse-256-large | parse-evaluate | 259,751 | 25,992 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,448 |
| reference-order-reverse-256-median | evaluate | 258,876 | 25,990 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,456 |
| reference-order-reverse-256-median | parse-evaluate | 257,456 | 25,990 | 1,024 | 56 | 43,548 | 13,454 | 0 | 3,476 |
| reference-order-reverse-256-mode | evaluate | 266,261 | 27,010 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,444 |
| reference-order-reverse-256-mode | parse-evaluate | 268,046 | 27,010 | 1,024 | 56 | 43,544 | 13,452 | 0 | 3,456 |
| reference-order-reverse-256-small | evaluate | 256,901 | 25,992 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,472 |
| reference-order-reverse-256-small | parse-evaluate | 259,386 | 25,992 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,460 |
| reference-order-reverse-64-large | evaluate | 52,882 | 5,023 | 256 | 76 | 32,480 | 5,976 | 0 | 3,464 |
| reference-order-reverse-64-large | parse-evaluate | 53,285 | 5,023 | 256 | 104 | 37,944 | 7,310 | 0 | 3,444 |
| reference-order-reverse-64-median | evaluate | 52,887 | 5,021 | 256 | 76 | 32,480 | 5,976 | 0 | 3,460 |
| reference-order-reverse-64-median | parse-evaluate | 53,310 | 5,021 | 256 | 104 | 37,940 | 7,309 | 0 | 3,456 |
| reference-order-reverse-64-mode | evaluate | 54,812 | 5,273 | 256 | 76 | 32,480 | 5,976 | 0 | 3,480 |
| reference-order-reverse-64-mode | parse-evaluate | 55,225 | 5,273 | 256 | 104 | 37,932 | 7,307 | 0 | 3,468 |
| reference-order-reverse-64-small | evaluate | 52,932 | 5,023 | 256 | 76 | 32,480 | 5,976 | 0 | 3,440 |
| reference-order-reverse-64-small | parse-evaluate | 53,325 | 5,023 | 256 | 104 | 37,944 | 7,310 | 0 | 3,448 |
| reference-order-sorted-64-large | evaluate | 58,622 | 5,678 | 256 | 76 | 32,480 | 5,976 | 0 | 3,452 |
| reference-order-sorted-64-large | parse-evaluate | 59,047 | 5,678 | 256 | 104 | 37,944 | 7,310 | 0 | 3,480 |
| reference-order-sorted-64-median | evaluate | 59,835 | 5,676 | 256 | 76 | 32,480 | 5,976 | 0 | 3,460 |
| reference-order-sorted-64-median | parse-evaluate | 59,100 | 5,676 | 256 | 104 | 37,940 | 7,309 | 0 | 3,460 |
| reference-order-sorted-64-mode | evaluate | 60,580 | 5,928 | 256 | 76 | 32,480 | 5,976 | 0 | 3,456 |
| reference-order-sorted-64-mode | parse-evaluate | 60,892 | 5,928 | 256 | 104 | 37,932 | 7,307 | 0 | 3,448 |
| reference-order-sorted-64-small | evaluate | 58,690 | 5,678 | 256 | 76 | 32,480 | 5,976 | 0 | 3,464 |
| reference-order-sorted-64-small | parse-evaluate | 59,110 | 5,678 | 256 | 104 | 37,944 | 7,310 | 0 | 3,456 |
| reference-order-sorted-large | evaluate | 283,846 | 29,044 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,420 |
| reference-order-sorted-large | parse-evaluate | 284,411 | 29,044 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,468 |
| reference-order-sorted-median | evaluate | 285,116 | 29,042 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,456 |
| reference-order-sorted-median | parse-evaluate | 285,221 | 29,042 | 1,024 | 56 | 43,548 | 13,454 | 0 | 3,452 |
| reference-order-sorted-mode | evaluate | 293,096 | 30,062 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,444 |
| reference-order-sorted-mode | parse-evaluate | 294,486 | 30,062 | 1,024 | 56 | 43,544 | 13,452 | 0 | 3,468 |
| reference-order-sorted-percentile | evaluate | 285,996 | 29,051 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,420 |
| reference-order-sorted-percentile | parse-evaluate | 286,326 | 29,051 | 1,024 | 56 | 43,564 | 13,462 | 0 | 3,464 |
| reference-order-sorted-percentrank | evaluate | 43,390 | 2,078 | 1,024 | 26 | 10,512 | 4,680 | 0 | 3,452 |
| reference-order-sorted-percentrank | parse-evaluate | 43,890 | 2,078 | 1,024 | 40 | 13,266 | 6,025 | 0 | 3,452 |
| reference-order-sorted-quartile | evaluate | 285,061 | 29,047 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,464 |
| reference-order-sorted-quartile | parse-evaluate | 285,886 | 29,047 | 1,024 | 56 | 43,556 | 13,458 | 0 | 3,456 |
| reference-order-sorted-rank | evaluate | 42,135 | 2,071 | 1,024 | 24 | 10,000 | 4,552 | 0 | 3,440 |
| reference-order-sorted-rank | parse-evaluate | 42,530 | 2,071 | 1,024 | 38 | 12,740 | 5,890 | 0 | 3,464 |
| reference-order-sorted-small | evaluate | 285,951 | 29,044 | 1,024 | 42 | 40,816 | 12,120 | 0 | 3,452 |
| reference-order-sorted-small | parse-evaluate | 285,951 | 29,044 | 1,024 | 56 | 43,550 | 13,455 | 0 | 3,464 |
| resource-order-large | evaluate | 697 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,480 |
| resource-order-large | parse-evaluate | 1,117 | 10 | 0 | 52 | 16,952 | 4,142 | 0 | 3,460 |
| resource-order-median | evaluate | 727 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,428 |
| resource-order-median | parse-evaluate | 1,115 | 10 | 0 | 52 | 16,948 | 4,141 | 0 | 3,460 |
| resource-order-mode | evaluate | 707 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,460 |
| resource-order-mode | parse-evaluate | 1,072 | 8 | 0 | 52 | 16,940 | 4,139 | 0 | 3,456 |
| resource-order-percentile | evaluate | 695 | 15 | 0 | 24 | 11,488 | 2,808 | 0 | 3,460 |
| resource-order-percentile | parse-evaluate | 1,125 | 15 | 0 | 52 | 16,980 | 4,149 | 0 | 3,460 |
| resource-order-percentrank | evaluate | 845 | 17 | 0 | 28 | 12,512 | 2,936 | 0 | 3,460 |
| resource-order-percentrank | parse-evaluate | 1,295 | 17 | 0 | 56 | 18,016 | 4,280 | 0 | 3,448 |
| resource-order-quartile | evaluate | 697 | 13 | 0 | 24 | 11,488 | 2,808 | 0 | 3,504 |
| resource-order-quartile | parse-evaluate | 1,117 | 13 | 0 | 52 | 16,964 | 4,145 | 0 | 3,464 |
| resource-order-rank | evaluate | 865 | 14 | 0 | 28 | 13,024 | 3,192 | 0 | 3,464 |
| resource-order-rank | parse-evaluate | 1,290 | 14 | 0 | 56 | 18,500 | 4,529 | 0 | 3,456 |
| resource-order-small | evaluate | 695 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,448 |
| resource-order-small | parse-evaluate | 1,125 | 10 | 0 | 52 | 16,952 | 4,142 | 0 | 3,464 |
| scalar-order-large | evaluate | 752 | 16 | 0 | 6,000 | 512,000 | 464 | 0 | 3,208 |
| scalar-order-large | parse-evaluate | 962 | 16 | 0 | 10,000 | 971,000 | 891 | 0 | 3,160 |
| scalar-order-median | evaluate | 1,572 | 43 | 0 | 9,000 | 1,088,000 | 752 | 0 | 3,308 |
| scalar-order-median | parse-evaluate | 1,850 | 43 | 0 | 14,000 | 2,320,000 | 1,568 | 0 | 3,272 |
| scalar-order-mode | evaluate | 1,812 | 71 | 0 | 11,000 | 1,856,000 | 1,136 | 0 | 3,168 |
| scalar-order-mode | parse-evaluate | 2,138 | 71 | 0 | 17,000 | 3,170,000 | 1,970 | 0 | 3,172 |
| scalar-order-percentile | evaluate | 744 | 23 | 0 | 6,000 | 512,000 | 464 | 0 | 3,188 |
| scalar-order-percentile | parse-evaluate | 971 | 23 | 0 | 10,000 | 978,000 | 898 | 0 | 3,176 |
| scalar-order-percentrank | evaluate | 938 | 26 | 0 | 7,000 | 864,000 | 640 | 0 | 3,228 |
| scalar-order-percentrank | parse-evaluate | 1,180 | 26 | 0 | 11,000 | 1,331,000 | 1,075 | 0 | 3,172 |
| scalar-order-quartile | evaluate | 743 | 19 | 0 | 6,000 | 512,000 | 464 | 0 | 3,204 |
| scalar-order-quartile | parse-evaluate | 956 | 19 | 0 | 10,000 | 974,000 | 894 | 0 | 3,192 |
| scalar-order-rank | evaluate | 903 | 20 | 0 | 7,000 | 864,000 | 640 | 0 | 3,172 |
| scalar-order-rank | parse-evaluate | 1,111 | 20 | 0 | 11,000 | 1,324,000 | 1,068 | 0 | 3,200 |
| scalar-order-small | evaluate | 757 | 16 | 0 | 6,000 | 512,000 | 464 | 0 | 3,176 |
| scalar-order-small | parse-evaluate | 963 | 16 | 0 | 10,000 | 971,000 | 891 | 0 | 3,220 |
| type-refusal-order-large | evaluate | 750 | 16 | 0 | 6,000 | 512,000 | 464 | 0 | 3,172 |
| type-refusal-order-large | parse-evaluate | 964 | 16 | 0 | 10,000 | 971,000 | 891 | 0 | 3,168 |
| type-refusal-order-median | evaluate | 399 | 8 | 0 | 4,000 | 400,000 | 400 | 0 | 3,164 |
| type-refusal-order-median | parse-evaluate | 518 | 8 | 0 | 6,000 | 793,000 | 793 | 0 | 3,224 |
| type-refusal-order-mode | evaluate | 2,310 | 18 | 0 | 14 | 4,648 | 4,152 | 0 | 3,460 |
| type-refusal-order-mode | parse-evaluate | 2,970 | 18 | 0 | 23 | 6,022 | 5,494 | 0 | 3,464 |
| type-refusal-order-percentile | evaluate | 745 | 23 | 0 | 6,000 | 512,000 | 464 | 0 | 3,180 |
| type-refusal-order-percentile | parse-evaluate | 967 | 23 | 0 | 10,000 | 978,000 | 898 | 0 | 3,164 |
| type-refusal-order-percentrank | evaluate | 907 | 26 | 0 | 7,000 | 864,000 | 640 | 0 | 3,256 |
| type-refusal-order-percentrank | parse-evaluate | 1,128 | 26 | 0 | 11,000 | 1,331,000 | 1,075 | 0 | 3,172 |
| type-refusal-order-quartile | evaluate | 2,350 | 25 | 0 | 14 | 4,648 | 4,152 | 0 | 3,460 |
| type-refusal-order-quartile | parse-evaluate | 3,130 | 25 | 0 | 24 | 6,796 | 5,884 | 0 | 3,440 |
| type-refusal-order-rank | evaluate | 739 | 16 | 0 | 6,000 | 512,000 | 464 | 0 | 3,156 |
| type-refusal-order-rank | parse-evaluate | 945 | 16 | 0 | 10,000 | 970,000 | 890 | 0 | 3,220 |
| type-refusal-order-small | evaluate | 752 | 16 | 0 | 6,000 | 512,000 | 464 | 0 | 3,216 |
| type-refusal-order-small | parse-evaluate | 964 | 16 | 0 | 10,000 | 971,000 | 891 | 0 | 3,240 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed finite result or formula error before timing. The profile does not measure save, recalculation, cache publication, native producer acceptance, cold filesystem state, wildcard/regular-expression host profiles, or cross-platform bit identity.
