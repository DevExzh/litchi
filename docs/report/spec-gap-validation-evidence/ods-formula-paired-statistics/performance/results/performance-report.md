# ODS paired-statistics evaluator performance profile

The baseline is committed `aa48eee68cb6ee0904523394d4ff8015dce6e595`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,867 | 46,242 | +0.8% | 88 | 88 | 3,432 | 3,492 |
| array-control-16x16-arithmetic | parse-evaluate | 58,670 | 58,115 | -0.9% | 352 | 352 | 3,448 | 3,520 |
| array-control-16x16-sin | evaluate | 126,038 | 126,958 | +0.7% | 2,144 | 2,144 | 3,768 | 3,812 |
| array-control-16x16-sin | parse-evaluate | 138,863 | 140,200 | +1.0% | 2,412 | 2,412 | 3,748 | 3,796 |
| array-control-4x4-arithmetic | evaluate | 4,316 | 4,352 | +0.8% | 1,120 | 1,120 | 3,424 | 3,488 |
| array-control-4x4-arithmetic | parse-evaluate | 5,273 | 5,339 | +1.3% | 2,240 | 2,240 | 3,436 | 3,492 |
| array-control-4x4-sin | evaluate | 9,676 | 9,680 | +0.0% | 3,840 | 3,840 | 3,792 | 3,800 |
| array-control-4x4-sin | parse-evaluate | 10,601 | 10,792 | +1.8% | 5,040 | 5,040 | 3,740 | 3,796 |
| database-control-dstdev | evaluate | 3,840 | 3,870 | +0.8% | 20 | 20 | 3,452 | 3,528 |
| database-control-dstdev | parse-evaluate | 4,390 | 4,490 | +2.3% | 29 | 29 | 3,496 | 3,500 |
| database-control-dsum | evaluate | 3,890 | 3,930 | +1.0% | 20 | 20 | 3,464 | 3,480 |
| database-control-dsum | parse-evaluate | 4,480 | 4,600 | +2.7% | 29 | 29 | 3,436 | 3,512 |
| database-control-dvar | evaluate | 3,770 | 3,850 | +2.1% | 20 | 20 | 3,512 | 3,492 |
| database-control-dvar | parse-evaluate | 4,370 | 4,490 | +2.7% | 29 | 29 | 3,476 | 3,524 |
| literal-aggregate-4x1-sum | evaluate | 1,780 | 1,810 | +1.7% | 10 | 10 | 3,432 | 3,496 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,440 | 2,480 | +1.6% | 23 | 23 | 3,456 | 3,516 |
| matched-descriptive-extreme-inline-avedev | evaluate | 3,690 | 3,390 | -8.1% | 11 | 11 | 3,436 | 3,492 |
| matched-descriptive-extreme-inline-avedev | parse-evaluate | 4,240 | 3,990 | -5.9% | 20 | 20 | 3,460 | 3,512 |
| matched-descriptive-extreme-inline-devsq | evaluate | 4,140 | 3,390 | -18.1% | 11 | 11 | 3,472 | 3,524 |
| matched-descriptive-extreme-inline-devsq | parse-evaluate | 4,660 | 4,040 | -13.3% | 20 | 20 | 3,432 | 3,520 |
| matched-descriptive-extreme-inline-kurt | evaluate | 10,880 | 8,610 | -20.9% | 11 | 11 | 3,448 | 3,492 |
| matched-descriptive-extreme-inline-kurt | parse-evaluate | 11,450 | 9,240 | -19.3% | 20 | 20 | 3,460 | 3,480 |
| matched-descriptive-extreme-inline-skew | evaluate | 5,470 | 3,210 | -41.3% | 11 | 11 | 3,448 | 3,492 |
| matched-descriptive-extreme-inline-skew | parse-evaluate | 6,020 | 3,820 | -36.5% | 20 | 20 | 3,464 | 3,492 |
| matched-descriptive-extreme-inline-skewp | evaluate | 5,420 | 3,160 | -41.7% | 11 | 11 | 3,460 | 3,512 |
| matched-descriptive-extreme-inline-skewp | parse-evaluate | 6,040 | 3,770 | -37.6% | 20 | 20 | 3,468 | 3,516 |
| matched-descriptive-reference-avedev | evaluate | 48,350 | 49,417 | +2.2% | 32 | 32 | 3,444 | 3,488 |
| matched-descriptive-reference-avedev | parse-evaluate | 48,457 | 49,430 | +2.0% | 60 | 60 | 3,436 | 3,500 |
| matched-descriptive-reference-devsq | evaluate | 17,702 | 18,107 | +2.3% | 32 | 32 | 3,440 | 3,516 |
| matched-descriptive-reference-devsq | parse-evaluate | 18,157 | 18,582 | +2.3% | 60 | 60 | 3,456 | 3,524 |
| matched-descriptive-reference-kurt | evaluate | 32,117 | 32,775 | +2.0% | 32 | 32 | 3,436 | 3,492 |
| matched-descriptive-reference-kurt | parse-evaluate | 32,627 | 33,215 | +1.8% | 60 | 60 | 3,424 | 3,504 |
| matched-descriptive-reference-skew | evaluate | 20,997 | 21,500 | +2.4% | 32 | 32 | 3,452 | 3,484 |
| matched-descriptive-reference-skew | parse-evaluate | 21,577 | 21,900 | +1.5% | 60 | 60 | 3,428 | 3,492 |
| matched-descriptive-reference-skewp | evaluate | 21,202 | 21,300 | +0.5% | 32 | 32 | 3,436 | 3,492 |
| matched-descriptive-reference-skewp | parse-evaluate | 21,630 | 21,905 | +1.3% | 60 | 60 | 3,436 | 3,524 |
| matched-descriptive-sensitive-inline-avedev | evaluate | 3,400 | 3,320 | -2.4% | 11 | 11 | 3,464 | 3,492 |
| matched-descriptive-sensitive-inline-avedev | parse-evaluate | 3,920 | 3,810 | -2.8% | 19 | 19 | 3,456 | 3,516 |
| matched-descriptive-sensitive-inline-devsq | evaluate | 3,210 | 3,480 | +8.4% | 11 | 11 | 3,464 | 3,508 |
| matched-descriptive-sensitive-inline-devsq | parse-evaluate | 3,740 | 4,050 | +8.3% | 19 | 19 | 3,432 | 3,488 |
| matched-descriptive-sensitive-inline-kurt | evaluate | 12,880 | 12,980 | +0.8% | 11 | 11 | 3,468 | 3,528 |
| matched-descriptive-sensitive-inline-kurt | parse-evaluate | 13,410 | 13,500 | +0.7% | 19 | 19 | 3,456 | 3,492 |
| matched-descriptive-sensitive-inline-skew | evaluate | 4,300 | 4,430 | +3.0% | 11 | 11 | 3,456 | 3,496 |
| matched-descriptive-sensitive-inline-skew | parse-evaluate | 4,850 | 5,000 | +3.1% | 19 | 19 | 3,428 | 3,492 |
| matched-descriptive-sensitive-inline-skewp | evaluate | 4,320 | 4,420 | +2.3% | 11 | 11 | 3,480 | 3,520 |
| matched-descriptive-sensitive-inline-skewp | parse-evaluate | 4,830 | 5,080 | +5.2% | 19 | 19 | 3,452 | 3,508 |
| reference-aggregate-64x4-sum | evaluate | 10,222 | 10,262 | +0.4% | 32 | 32 | 3,488 | 3,488 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,675 | 10,702 | +0.3% | 60 | 60 | 3,448 | 3,492 |
| reference-array-16x4-arithmetic | evaluate | 9,295 | 9,326 | +0.3% | 800 | 800 | 3,456 | 3,492 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,544 | 9,597 | +0.6% | 1,280 | 1,280 | 3,416 | 3,496 |
| reference-conditional-256x4-sumifs | evaluate | 79,335 | 80,780 | +1.8% | 42 | 42 | 3,436 | 3,512 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,015 | 81,480 | +1.8% | 68 | 68 | 3,464 | 3,528 |
| reference-control-average | evaluate | 10,245 | 10,267 | +0.2% | 32 | 32 | 3,480 | 3,516 |
| reference-control-average | parse-evaluate | 10,602 | 10,615 | +0.1% | 60 | 60 | 3,452 | 3,492 |
| reference-control-counta | evaluate | 8,632 | 8,702 | +0.8% | 32 | 32 | 3,468 | 3,504 |
| reference-control-counta | parse-evaluate | 9,047 | 9,107 | +0.7% | 60 | 60 | 3,428 | 3,500 |
| representative-median | evaluate | 1,623 | 1,622 | -0.1% | 9,000 | 9,000 | 3,268 | 3,376 |
| representative-median | parse-evaluate | 1,912 | 1,890 | -1.2% | 14,000 | 14,000 | 3,276 | 3,288 |
| representative-percentrank | evaluate | 997 | 981 | -1.6% | 7,000 | 7,000 | 3,396 | 3,268 |
| representative-percentrank | parse-evaluate | 1,237 | 1,225 | -1.0% | 11,000 | 11,000 | 3,352 | 3,300 |
| representative-rank | evaluate | 935 | 912 | -2.5% | 7,000 | 7,000 | 3,380 | 3,220 |
| representative-rank | parse-evaluate | 1,167 | 1,131 | -3.1% | 11,000 | 11,000 | 3,352 | 3,332 |
| scalar-aggregate-sum | evaluate | 585 | 569 | -2.7% | 4,000 | 4,000 | 3,400 | 3,248 |
| scalar-aggregate-sum | parse-evaluate | 764 | 780 | +2.1% | 8,000 | 8,000 | 3,396 | 3,408 |
| scalar-control-arithmetic | evaluate | 560 | 563 | +0.5% | 5,000 | 5,000 | 3,136 | 3,172 |
| scalar-control-arithmetic | parse-evaluate | 691 | 708 | +2.5% | 8,000 | 8,000 | 3,136 | 3,184 |
| scalar-control-average | evaluate | 1,039 | 1,006 | -3.2% | 6,000 | 6,000 | 3,424 | 3,264 |
| scalar-control-average | parse-evaluate | 1,267 | 1,247 | -1.6% | 10,000 | 10,000 | 3,392 | 3,284 |
| scalar-control-counta | evaluate | 859 | 802 | -6.6% | 6,000 | 6,000 | 3,392 | 3,268 |
| scalar-control-counta | parse-evaluate | 1,089 | 1,068 | -1.9% | 10,000 | 10,000 | 3,228 | 3,252 |
| scalar-control-imsum | evaluate | 1,497 | 1,379 | -7.9% | 8,000 | 8,000 | 3,388 | 3,200 |
| scalar-control-imsum | parse-evaluate | 2,051 | 2,011 | -2.0% | 17,000 | 17,000 | 3,376 | 3,272 |
| scalar-control-sin | evaluate | 458 | 467 | +2.0% | 4,000 | 4,000 | 3,572 | 3,376 |
| scalar-control-sin | parse-evaluate | 646 | 660 | +2.2% | 8,000 | 8,000 | 3,620 | 3,552 |
| scalar-control-stdev | evaluate | 866 | 836 | -3.5% | 6,000 | 6,000 | 3,260 | 3,288 |
| scalar-control-stdev | parse-evaluate | 1,074 | 1,059 | -1.4% | 10,000 | 10,000 | 3,396 | 3,412 |
| scalar-control-var | evaluate | 859 | 821 | -4.4% | 6,000 | 6,000 | 3,244 | 3,396 |
| scalar-control-var | parse-evaluate | 1,065 | 1,030 | -3.3% | 10,000 | 10,000 | 3,396 | 3,380 |

## Paired-statistics workloads

The candidate matrix covers paired small/large references, finite extreme values, offset origins, aligned pairwise skips, shape/resource refusal, sticky cancellation, and FORECAST query-array/cache behavior. Numerical labels and pair-count behavior are taken from the frozen contract and are retained with the raw case receipts.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-paired-correl | evaluate | 555 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,524 |
| cancellation-paired-correl | parse-evaluate | 1,120 | 5 | 0 | 49 | 9,904 | 5,674 | 0 | 3,520 |
| cancellation-paired-covar | evaluate | 555 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,532 |
| cancellation-paired-covar | parse-evaluate | 1,127 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,488 |
| cancellation-paired-forecast | evaluate | 640 | 6 | 0 | 14 | 5,352 | 4,904 | 0 | 3,528 |
| cancellation-paired-forecast | parse-evaluate | 1,207 | 6 | 0 | 50 | 10,880 | 6,254 | 0 | 3,520 |
| cancellation-paired-intercept | evaluate | 557 | 6 | 0 | 13 | 4,392 | 4,328 | 0 | 3,492 |
| cancellation-paired-intercept | parse-evaluate | 1,127 | 6 | 0 | 49 | 9,916 | 5,677 | 0 | 3,492 |
| cancellation-paired-pearson | evaluate | 562 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,492 |
| cancellation-paired-pearson | parse-evaluate | 1,125 | 5 | 0 | 49 | 9,908 | 5,675 | 0 | 3,492 |
| cancellation-paired-rsq | evaluate | 560 | 4 | 0 | 13 | 4,392 | 4,328 | 0 | 3,492 |
| cancellation-paired-rsq | parse-evaluate | 1,115 | 4 | 0 | 49 | 9,892 | 5,671 | 0 | 3,492 |
| cancellation-paired-slope | evaluate | 562 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,536 |
| cancellation-paired-slope | parse-evaluate | 1,127 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,488 |
| cancellation-paired-steyx | evaluate | 562 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,512 |
| cancellation-paired-steyx | parse-evaluate | 1,127 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,492 |
| extreme-paired-correl | evaluate | 5,800 | 15,997 | 0 | 14 | 6,952 | 5,512 | 0 | 3,496 |
| extreme-paired-correl | parse-evaluate | 6,680 | 15,997 | 0 | 26 | 10,053 | 7,205 | 0 | 3,476 |
| extreme-paired-covar | evaluate | 3,880 | 5,254 | 0 | 14 | 6,952 | 5,512 | 0 | 3,516 |
| extreme-paired-covar | parse-evaluate | 4,750 | 5,254 | 0 | 26 | 10,052 | 7,204 | 0 | 3,460 |
| extreme-paired-forecast | evaluate | 6,010 | 13,371 | 0 | 14 | 7,144 | 5,704 | 0 | 3,492 |
| extreme-paired-forecast | parse-evaluate | 6,920 | 13,371 | 0 | 26 | 10,249 | 7,401 | 0 | 3,504 |
| extreme-paired-intercept | evaluate | 5,310 | 10,924 | 0 | 14 | 6,952 | 5,512 | 0 | 3,500 |
| extreme-paired-intercept | parse-evaluate | 6,220 | 10,924 | 0 | 26 | 10,056 | 7,208 | 0 | 3,484 |
| extreme-paired-pearson | evaluate | 5,790 | 15,998 | 0 | 14 | 6,952 | 5,512 | 0 | 3,484 |
| extreme-paired-pearson | parse-evaluate | 6,640 | 15,998 | 0 | 26 | 10,054 | 7,206 | 0 | 3,492 |
| extreme-paired-rsq | evaluate | 6,240 | 15,994 | 0 | 14 | 6,952 | 5,512 | 0 | 3,496 |
| extreme-paired-rsq | parse-evaluate | 7,070 | 15,994 | 0 | 26 | 10,050 | 7,202 | 0 | 3,492 |
| extreme-paired-slope | evaluate | 5,260 | 10,920 | 0 | 14 | 6,952 | 5,512 | 0 | 3,516 |
| extreme-paired-slope | parse-evaluate | 6,130 | 10,920 | 0 | 26 | 10,052 | 7,204 | 0 | 3,488 |
| extreme-paired-steyx | evaluate | 5,740 | 15,996 | 0 | 14 | 6,952 | 5,512 | 0 | 3,472 |
| extreme-paired-steyx | parse-evaluate | 6,550 | 15,996 | 0 | 26 | 10,052 | 7,204 | 0 | 3,528 |
| forecast-query-array | evaluate | 14,290 | 28,521 | 128 | 18 | 6,008 | 5,384 | 176 | 3,496 |
| forecast-query-array | parse-evaluate | 15,260 | 28,521 | 128 | 32 | 8,322 | 7,154 | 176 | 3,472 |
| forecast-query-array-cache | evaluate | 18,660 | 28,585 | 128 | 42 | 15,464 | 13,720 | 176 | 3,488 |
| forecast-query-array-cache | parse-evaluate | 19,910 | 28,585 | 128 | 62 | 19,592 | 16,344 | 176 | 3,532 |
| large-paired-correl | evaluate | 125,480 | 300,554 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,488 |
| large-paired-correl | parse-evaluate | 124,325 | 300,554 | 2,048 | 44 | 11,544 | 5,676 | 0 | 3,480 |
| large-paired-covar | evaluate | 111,076 | 153,131 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,492 |
| large-paired-covar | parse-evaluate | 113,991 | 153,131 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,520 |
| large-paired-forecast | evaluate | 118,495 | 229,588 | 2,048 | 28 | 10,704 | 4,904 | 0 | 3,520 |
| large-paired-forecast | parse-evaluate | 118,596 | 229,588 | 2,048 | 46 | 13,472 | 6,256 | 0 | 3,512 |
| large-paired-intercept | evaluate | 116,505 | 227,141 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,476 |
| large-paired-intercept | parse-evaluate | 117,981 | 227,141 | 2,048 | 44 | 11,550 | 5,679 | 0 | 3,520 |
| large-paired-pearson | evaluate | 127,126 | 300,555 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,492 |
| large-paired-pearson | parse-evaluate | 123,380 | 300,555 | 2,048 | 44 | 11,546 | 5,677 | 0 | 3,500 |
| large-paired-rsq | evaluate | 124,085 | 300,551 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,508 |
| large-paired-rsq | parse-evaluate | 124,396 | 300,551 | 2,048 | 44 | 11,538 | 5,673 | 0 | 3,480 |
| large-paired-slope | evaluate | 117,275 | 227,137 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,484 |
| large-paired-slope | parse-evaluate | 118,000 | 227,137 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,516 |
| large-paired-steyx | evaluate | 126,001 | 300,553 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,508 |
| large-paired-steyx | parse-evaluate | 123,905 | 300,553 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,488 |
| offset-paired-correl | evaluate | 12,412 | 32,714 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| offset-paired-correl | parse-evaluate | 12,931 | 32,714 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,492 |
| offset-paired-covar | evaluate | 9,991 | 13,931 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,524 |
| offset-paired-covar | parse-evaluate | 10,487 | 13,931 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,524 |
| offset-paired-forecast | evaluate | 12,414 | 26,068 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,496 |
| offset-paired-forecast | parse-evaluate | 12,953 | 26,068 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,492 |
| offset-paired-intercept | evaluate | 11,670 | 23,621 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| offset-paired-intercept | parse-evaluate | 12,184 | 23,621 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,496 |
| offset-paired-pearson | evaluate | 12,367 | 32,715 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,488 |
| offset-paired-pearson | parse-evaluate | 12,915 | 32,715 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,488 |
| offset-paired-rsq | evaluate | 12,904 | 32,711 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,552 |
| offset-paired-rsq | parse-evaluate | 13,397 | 32,711 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,492 |
| offset-paired-slope | evaluate | 11,502 | 23,617 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| offset-paired-slope | parse-evaluate | 12,088 | 23,617 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,488 |
| offset-paired-steyx | evaluate | 12,348 | 32,713 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| offset-paired-steyx | parse-evaluate | 12,882 | 32,713 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,492 |
| pairwise-skip-paired-correl | evaluate | 11,398 | 23,914 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| pairwise-skip-paired-correl | parse-evaluate | 11,987 | 23,914 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,496 |
| pairwise-skip-paired-covar | evaluate | 9,220 | 9,419 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,500 |
| pairwise-skip-paired-covar | parse-evaluate | 9,755 | 9,419 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,504 |
| pairwise-skip-paired-forecast | evaluate | 11,561 | 19,412 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,520 |
| pairwise-skip-paired-forecast | parse-evaluate | 12,083 | 19,412 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,492 |
| pairwise-skip-paired-intercept | evaluate | 10,673 | 16,965 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,488 |
| pairwise-skip-paired-intercept | parse-evaluate | 11,296 | 16,965 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,516 |
| pairwise-skip-paired-pearson | evaluate | 11,248 | 23,915 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,500 |
| pairwise-skip-paired-pearson | parse-evaluate | 11,856 | 23,915 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,484 |
| pairwise-skip-paired-rsq | evaluate | 11,850 | 23,911 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,468 |
| pairwise-skip-paired-rsq | parse-evaluate | 12,350 | 23,911 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,488 |
| pairwise-skip-paired-slope | evaluate | 10,626 | 16,961 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,520 |
| pairwise-skip-paired-slope | parse-evaluate | 11,220 | 16,961 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,488 |
| pairwise-skip-paired-steyx | evaluate | 11,316 | 23,913 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| pairwise-skip-paired-steyx | parse-evaluate | 11,893 | 23,913 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,528 |
| resource-paired-correl | evaluate | 712 | 11 | 0 | 24 | 11,488 | 2,808 | 0 | 3,476 |
| resource-paired-correl | parse-evaluate | 1,275 | 11 | 0 | 60 | 17,000 | 4,154 | 0 | 3,524 |
| resource-paired-covar | evaluate | 720 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,492 |
| resource-paired-covar | parse-evaluate | 1,265 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,488 |
| resource-paired-forecast | evaluate | 887 | 16 | 0 | 28 | 13,024 | 3,192 | 0 | 3,492 |
| resource-paired-forecast | parse-evaluate | 1,455 | 16 | 0 | 64 | 18,552 | 4,542 | 0 | 3,488 |
| resource-paired-intercept | evaluate | 720 | 14 | 0 | 24 | 11,488 | 2,808 | 0 | 3,468 |
| resource-paired-intercept | parse-evaluate | 1,280 | 14 | 0 | 60 | 17,012 | 4,157 | 0 | 3,484 |
| resource-paired-pearson | evaluate | 707 | 12 | 0 | 24 | 11,488 | 2,808 | 0 | 3,488 |
| resource-paired-pearson | parse-evaluate | 1,275 | 12 | 0 | 60 | 17,004 | 4,155 | 0 | 3,508 |
| resource-paired-rsq | evaluate | 700 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,500 |
| resource-paired-rsq | parse-evaluate | 1,257 | 8 | 0 | 60 | 16,988 | 4,151 | 0 | 3,524 |
| resource-paired-slope | evaluate | 705 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,528 |
| resource-paired-slope | parse-evaluate | 1,262 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,440 |
| resource-paired-steyx | evaluate | 692 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,468 |
| resource-paired-steyx | parse-evaluate | 1,272 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,496 |
| shape-reject-paired-correl | evaluate | 2,697 | 26 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,496 |
| shape-reject-paired-correl | parse-evaluate | 3,528 | 26 | 0 | 2,320 | 576,560 | 6,295 | 0 | 3,476 |
| shape-reject-paired-covar | evaluate | 2,724 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,524 |
| shape-reject-paired-covar | parse-evaluate | 3,521 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,496 |
| shape-reject-paired-forecast | evaluate | 2,626 | 26 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,528 |
| shape-reject-paired-forecast | parse-evaluate | 3,430 | 26 | 0 | 2,320 | 576,720 | 6,297 | 0 | 3,516 |
| shape-reject-paired-intercept | evaluate | 2,697 | 29 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,528 |
| shape-reject-paired-intercept | parse-evaluate | 3,527 | 29 | 0 | 2,320 | 576,800 | 6,298 | 0 | 3,492 |
| shape-reject-paired-pearson | evaluate | 2,699 | 27 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,536 |
| shape-reject-paired-pearson | parse-evaluate | 3,520 | 27 | 0 | 2,320 | 576,640 | 6,296 | 0 | 3,492 |
| shape-reject-paired-rsq | evaluate | 1,960 | 17 | 0 | 1,040 | 351,360 | 4,328 | 0 | 3,520 |
| shape-reject-paired-rsq | parse-evaluate | 2,637 | 17 | 0 | 1,760 | 461,360 | 5,671 | 0 | 3,488 |
| shape-reject-paired-slope | evaluate | 2,694 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,508 |
| shape-reject-paired-slope | parse-evaluate | 3,486 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,488 |
| shape-reject-paired-steyx | evaluate | 2,681 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,488 |
| shape-reject-paired-steyx | parse-evaluate | 3,518 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,528 |
| small-paired-correl | evaluate | 12,393 | 32,714 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,520 |
| small-paired-correl | parse-evaluate | 12,970 | 32,714 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,508 |
| small-paired-covar | evaluate | 9,978 | 13,931 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,524 |
| small-paired-covar | parse-evaluate | 10,530 | 13,931 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,520 |
| small-paired-forecast | evaluate | 12,479 | 26,068 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,496 |
| small-paired-forecast | parse-evaluate | 13,005 | 26,068 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,528 |
| small-paired-intercept | evaluate | 11,637 | 23,621 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,492 |
| small-paired-intercept | parse-evaluate | 12,154 | 23,621 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,512 |
| small-paired-pearson | evaluate | 12,414 | 32,715 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,516 |
| small-paired-pearson | parse-evaluate | 13,005 | 32,715 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,492 |
| small-paired-rsq | evaluate | 12,864 | 32,711 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,528 |
| small-paired-rsq | parse-evaluate | 13,476 | 32,711 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,492 |
| small-paired-slope | evaluate | 11,538 | 23,617 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,524 |
| small-paired-slope | parse-evaluate | 12,091 | 23,617 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,504 |
| small-paired-steyx | evaluate | 12,378 | 32,713 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,504 |
| small-paired-steyx | parse-evaluate | 12,928 | 32,713 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,508 |

The resolver is an immutable borrowing fixture. Each child validates one finite numerical result or typed failure before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
