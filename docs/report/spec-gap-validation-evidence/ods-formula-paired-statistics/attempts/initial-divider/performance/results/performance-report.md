# ODS paired-statistics evaluator performance profile

The baseline is committed `aa48eee68cb6ee0904523394d4ff8015dce6e595`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,442 | 45,642 | +0.4% | 88 | 88 | 3,488 | 3,400 |
| array-control-16x16-arithmetic | parse-evaluate | 58,885 | 58,010 | -1.5% | 352 | 352 | 3,464 | 3,460 |
| array-control-16x16-sin | evaluate | 126,833 | 126,775 | -0.0% | 2,144 | 2,144 | 3,808 | 3,716 |
| array-control-16x16-sin | parse-evaluate | 141,310 | 139,573 | -1.2% | 2,412 | 2,412 | 3,804 | 3,740 |
| array-control-4x4-arithmetic | evaluate | 4,310 | 4,331 | +0.5% | 1,120 | 1,120 | 3,460 | 3,400 |
| array-control-4x4-arithmetic | parse-evaluate | 5,277 | 5,284 | +0.1% | 2,240 | 2,240 | 3,456 | 3,380 |
| array-control-4x4-sin | evaluate | 9,732 | 9,663 | -0.7% | 3,840 | 3,840 | 3,716 | 3,744 |
| array-control-4x4-sin | parse-evaluate | 10,610 | 10,727 | +1.1% | 5,040 | 5,040 | 3,792 | 3,748 |
| database-control-dstdev | evaluate | 3,850 | 3,830 | -0.5% | 20 | 20 | 3,472 | 3,408 |
| database-control-dstdev | parse-evaluate | 4,410 | 4,410 | +0.0% | 29 | 29 | 3,508 | 3,476 |
| database-control-dsum | evaluate | 3,871 | 3,880 | +0.2% | 20 | 20 | 3,476 | 3,532 |
| database-control-dsum | parse-evaluate | 4,530 | 4,550 | +0.4% | 29 | 29 | 3,464 | 3,452 |
| database-control-dvar | evaluate | 3,840 | 3,850 | +0.3% | 20 | 20 | 3,500 | 3,548 |
| database-control-dvar | parse-evaluate | 4,430 | 4,470 | +0.9% | 29 | 29 | 3,428 | 3,524 |
| literal-aggregate-4x1-sum | evaluate | 1,840 | 1,790 | -2.7% | 10 | 10 | 3,484 | 3,464 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,480 | 2,500 | +0.8% | 23 | 23 | 3,492 | 3,464 |
| matched-descriptive-extreme-inline-avedev | evaluate | 3,780 | 5,010 | +32.5% | 11 | 11 | 3,464 | 3,460 |
| matched-descriptive-extreme-inline-avedev | parse-evaluate | 4,400 | 5,650 | +28.4% | 20 | 20 | 3,492 | 3,476 |
| matched-descriptive-extreme-inline-devsq | evaluate | 4,110 | 6,120 | +48.9% | 11 | 11 | 3,480 | 3,396 |
| matched-descriptive-extreme-inline-devsq | parse-evaluate | 4,810 | 6,710 | +39.5% | 20 | 20 | 3,488 | 3,524 |
| matched-descriptive-extreme-inline-kurt | evaluate | 10,790 | 11,100 | +2.9% | 11 | 11 | 3,504 | 3,572 |
| matched-descriptive-extreme-inline-kurt | parse-evaluate | 11,380 | 11,680 | +2.6% | 20 | 20 | 3,480 | 3,572 |
| matched-descriptive-extreme-inline-skew | evaluate | 5,560 | 5,570 | +0.2% | 11 | 11 | 3,496 | 3,480 |
| matched-descriptive-extreme-inline-skew | parse-evaluate | 6,160 | 6,170 | +0.2% | 20 | 20 | 3,480 | 3,468 |
| matched-descriptive-extreme-inline-skewp | evaluate | 5,530 | 5,590 | +1.1% | 11 | 11 | 3,464 | 3,476 |
| matched-descriptive-extreme-inline-skewp | parse-evaluate | 6,180 | 6,180 | +0.0% | 20 | 20 | 3,464 | 3,484 |
| matched-descriptive-reference-avedev | evaluate | 49,267 | 50,555 | +2.6% | 32 | 32 | 3,488 | 3,512 |
| matched-descriptive-reference-avedev | parse-evaluate | 49,735 | 50,717 | +2.0% | 60 | 60 | 3,464 | 3,476 |
| matched-descriptive-reference-devsq | evaluate | 18,242 | 20,050 | +9.9% | 32 | 32 | 3,464 | 3,456 |
| matched-descriptive-reference-devsq | parse-evaluate | 18,632 | 20,375 | +9.4% | 60 | 60 | 3,476 | 3,456 |
| matched-descriptive-reference-kurt | evaluate | 32,565 | 32,487 | -0.2% | 32 | 32 | 3,472 | 3,528 |
| matched-descriptive-reference-kurt | parse-evaluate | 32,867 | 32,932 | +0.2% | 60 | 60 | 3,476 | 3,400 |
| matched-descriptive-reference-skew | evaluate | 21,690 | 21,487 | -0.9% | 32 | 32 | 3,448 | 3,548 |
| matched-descriptive-reference-skew | parse-evaluate | 22,050 | 21,752 | -1.4% | 60 | 60 | 3,460 | 3,524 |
| matched-descriptive-reference-skewp | evaluate | 21,640 | 21,250 | -1.8% | 32 | 32 | 3,476 | 3,484 |
| matched-descriptive-reference-skewp | parse-evaluate | 22,092 | 21,867 | -1.0% | 60 | 60 | 3,528 | 3,452 |
| matched-descriptive-sensitive-inline-avedev | evaluate | 3,510 | 4,850 | +38.2% | 11 | 11 | 3,488 | 3,428 |
| matched-descriptive-sensitive-inline-avedev | parse-evaluate | 4,040 | 5,450 | +34.9% | 19 | 19 | 3,484 | 3,476 |
| matched-descriptive-sensitive-inline-devsq | evaluate | 3,220 | 5,460 | +69.6% | 11 | 11 | 3,512 | 3,500 |
| matched-descriptive-sensitive-inline-devsq | parse-evaluate | 3,800 | 6,020 | +58.4% | 19 | 19 | 3,512 | 3,488 |
| matched-descriptive-sensitive-inline-kurt | evaluate | 12,900 | 13,000 | +0.8% | 11 | 11 | 3,448 | 3,476 |
| matched-descriptive-sensitive-inline-kurt | parse-evaluate | 13,350 | 13,560 | +1.6% | 19 | 19 | 3,496 | 3,472 |
| matched-descriptive-sensitive-inline-skew | evaluate | 4,390 | 4,440 | +1.1% | 11 | 11 | 3,424 | 3,532 |
| matched-descriptive-sensitive-inline-skew | parse-evaluate | 4,980 | 4,940 | -0.8% | 19 | 19 | 3,484 | 3,544 |
| matched-descriptive-sensitive-inline-skewp | evaluate | 4,390 | 4,420 | +0.7% | 11 | 11 | 3,496 | 3,464 |
| matched-descriptive-sensitive-inline-skewp | parse-evaluate | 4,910 | 4,990 | +1.6% | 19 | 19 | 3,500 | 3,516 |
| reference-aggregate-64x4-sum | evaluate | 10,510 | 10,275 | -2.2% | 32 | 32 | 3,512 | 3,400 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,897 | 10,702 | -1.8% | 60 | 60 | 3,496 | 3,472 |
| reference-array-16x4-arithmetic | evaluate | 9,249 | 9,275 | +0.3% | 800 | 800 | 3,516 | 3,460 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,642 | 9,512 | -1.3% | 1,280 | 1,280 | 3,468 | 3,472 |
| reference-conditional-256x4-sumifs | evaluate | 80,995 | 78,635 | -2.9% | 42 | 42 | 3,504 | 3,456 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,910 | 79,510 | -1.7% | 68 | 68 | 3,512 | 3,560 |
| reference-control-average | evaluate | 10,487 | 10,220 | -2.5% | 32 | 32 | 3,512 | 3,488 |
| reference-control-average | parse-evaluate | 10,892 | 10,572 | -2.9% | 60 | 60 | 3,496 | 3,464 |
| reference-control-counta | evaluate | 8,647 | 8,590 | -0.7% | 32 | 32 | 3,468 | 3,436 |
| reference-control-counta | parse-evaluate | 9,017 | 9,080 | +0.7% | 60 | 60 | 3,512 | 3,496 |
| representative-median | evaluate | 1,666 | 1,630 | -2.2% | 9,000 | 9,000 | 3,424 | 3,100 |
| representative-median | parse-evaluate | 1,935 | 1,911 | -1.2% | 14,000 | 14,000 | 3,456 | 3,332 |
| representative-percentrank | evaluate | 968 | 1,000 | +3.3% | 7,000 | 7,000 | 3,428 | 3,296 |
| representative-percentrank | parse-evaluate | 1,202 | 1,205 | +0.2% | 11,000 | 11,000 | 3,460 | 3,092 |
| representative-rank | evaluate | 924 | 951 | +2.9% | 7,000 | 7,000 | 3,504 | 3,208 |
| representative-rank | parse-evaluate | 1,148 | 1,166 | +1.6% | 11,000 | 11,000 | 3,460 | 3,188 |
| scalar-aggregate-sum | evaluate | 563 | 560 | -0.5% | 4,000 | 4,000 | 3,432 | 3,268 |
| scalar-aggregate-sum | parse-evaluate | 759 | 756 | -0.4% | 8,000 | 8,000 | 3,408 | 3,120 |
| scalar-control-arithmetic | evaluate | 560 | 560 | +0.0% | 5,000 | 5,000 | 3,096 | 3,084 |
| scalar-control-arithmetic | parse-evaluate | 697 | 695 | -0.3% | 8,000 | 8,000 | 3,132 | 3,016 |
| scalar-control-average | evaluate | 1,021 | 1,005 | -1.6% | 6,000 | 6,000 | 3,448 | 3,212 |
| scalar-control-average | parse-evaluate | 1,249 | 1,229 | -1.6% | 10,000 | 10,000 | 3,392 | 3,328 |
| scalar-control-counta | evaluate | 826 | 827 | +0.1% | 6,000 | 6,000 | 3,412 | 3,312 |
| scalar-control-counta | parse-evaluate | 1,076 | 1,070 | -0.6% | 10,000 | 10,000 | 3,320 | 3,220 |
| scalar-control-imsum | evaluate | 1,450 | 1,378 | -5.0% | 8,000 | 8,000 | 3,256 | 3,204 |
| scalar-control-imsum | parse-evaluate | 2,031 | 1,934 | -4.8% | 17,000 | 17,000 | 3,324 | 3,240 |
| scalar-control-sin | evaluate | 462 | 465 | +0.6% | 4,000 | 4,000 | 3,740 | 3,564 |
| scalar-control-sin | parse-evaluate | 648 | 651 | +0.5% | 8,000 | 8,000 | 3,700 | 3,552 |
| scalar-control-stdev | evaluate | 855 | 847 | -0.9% | 6,000 | 6,000 | 3,416 | 3,340 |
| scalar-control-stdev | parse-evaluate | 1,086 | 1,058 | -2.6% | 10,000 | 10,000 | 3,356 | 3,176 |
| scalar-control-var | evaluate | 839 | 849 | +1.2% | 6,000 | 6,000 | 3,392 | 3,308 |
| scalar-control-var | parse-evaluate | 1,049 | 1,049 | +0.0% | 10,000 | 10,000 | 3,424 | 3,256 |

## Paired-statistics workloads

The candidate matrix covers paired small/large references, finite extreme values, offset origins, aligned pairwise skips, shape/resource refusal, sticky cancellation, and FORECAST query-array/cache behavior. Numerical labels and pair-count behavior are taken from the frozen contract and are retained with the raw case receipts.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-paired-correl | evaluate | 560 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,540 |
| cancellation-paired-correl | parse-evaluate | 1,110 | 5 | 0 | 49 | 9,904 | 5,674 | 0 | 3,524 |
| cancellation-paired-covar | evaluate | 557 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,448 |
| cancellation-paired-covar | parse-evaluate | 1,107 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,520 |
| cancellation-paired-forecast | evaluate | 620 | 6 | 0 | 14 | 5,352 | 4,904 | 0 | 3,472 |
| cancellation-paired-forecast | parse-evaluate | 1,182 | 6 | 0 | 50 | 10,880 | 6,254 | 0 | 3,536 |
| cancellation-paired-intercept | evaluate | 547 | 6 | 0 | 13 | 4,392 | 4,328 | 0 | 3,544 |
| cancellation-paired-intercept | parse-evaluate | 1,120 | 6 | 0 | 49 | 9,916 | 5,677 | 0 | 3,428 |
| cancellation-paired-pearson | evaluate | 547 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,552 |
| cancellation-paired-pearson | parse-evaluate | 1,115 | 5 | 0 | 49 | 9,908 | 5,675 | 0 | 3,464 |
| cancellation-paired-rsq | evaluate | 555 | 4 | 0 | 13 | 4,392 | 4,328 | 0 | 3,476 |
| cancellation-paired-rsq | parse-evaluate | 1,095 | 4 | 0 | 49 | 9,892 | 5,671 | 0 | 3,488 |
| cancellation-paired-slope | evaluate | 547 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,476 |
| cancellation-paired-slope | parse-evaluate | 1,115 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,536 |
| cancellation-paired-steyx | evaluate | 542 | 5 | 0 | 13 | 4,392 | 4,328 | 0 | 3,456 |
| cancellation-paired-steyx | parse-evaluate | 1,107 | 5 | 0 | 49 | 9,900 | 5,673 | 0 | 3,452 |
| extreme-paired-correl | evaluate | 5,680 | 15,997 | 0 | 14 | 6,952 | 5,512 | 0 | 3,592 |
| extreme-paired-correl | parse-evaluate | 6,630 | 15,997 | 0 | 26 | 10,053 | 7,205 | 0 | 3,468 |
| extreme-paired-covar | evaluate | 8,130 | 5,254 | 0 | 14 | 6,952 | 5,512 | 0 | 3,524 |
| extreme-paired-covar | parse-evaluate | 8,190 | 5,254 | 0 | 26 | 10,052 | 7,204 | 0 | 3,552 |
| extreme-paired-forecast | evaluate | 12,731 | 13,371 | 0 | 14 | 7,144 | 5,704 | 0 | 3,556 |
| extreme-paired-forecast | parse-evaluate | 13,620 | 13,371 | 0 | 26 | 10,249 | 7,401 | 0 | 3,540 |
| extreme-paired-intercept | evaluate | 11,100 | 10,924 | 0 | 14 | 6,952 | 5,512 | 0 | 3,496 |
| extreme-paired-intercept | parse-evaluate | 12,400 | 10,924 | 0 | 26 | 10,056 | 7,208 | 0 | 3,612 |
| extreme-paired-pearson | evaluate | 5,710 | 15,998 | 0 | 14 | 6,952 | 5,512 | 0 | 3,512 |
| extreme-paired-pearson | parse-evaluate | 6,570 | 15,998 | 0 | 26 | 10,054 | 7,206 | 0 | 3,540 |
| extreme-paired-rsq | evaluate | 15,730 | 15,994 | 0 | 14 | 6,952 | 5,512 | 0 | 3,536 |
| extreme-paired-rsq | parse-evaluate | 16,540 | 15,994 | 0 | 26 | 10,050 | 7,202 | 0 | 3,456 |
| extreme-paired-slope | evaluate | 9,470 | 10,920 | 0 | 14 | 6,952 | 5,512 | 0 | 3,556 |
| extreme-paired-slope | parse-evaluate | 10,470 | 10,920 | 0 | 26 | 10,052 | 7,204 | 0 | 3,540 |
| extreme-paired-steyx | evaluate | 5,780 | 15,996 | 0 | 14 | 6,952 | 5,512 | 0 | 3,524 |
| extreme-paired-steyx | parse-evaluate | 6,580 | 15,996 | 0 | 26 | 10,052 | 7,204 | 0 | 3,544 |
| forecast-query-array | evaluate | 20,370 | 28,521 | 128 | 18 | 6,008 | 5,384 | 176 | 3,520 |
| forecast-query-array | parse-evaluate | 21,380 | 28,521 | 128 | 32 | 8,322 | 7,154 | 176 | 3,520 |
| forecast-query-array-cache | evaluate | 24,561 | 28,585 | 128 | 42 | 15,464 | 13,720 | 176 | 3,496 |
| forecast-query-array-cache | parse-evaluate | 25,970 | 28,585 | 128 | 62 | 19,592 | 16,344 | 176 | 3,552 |
| large-paired-correl | evaluate | 122,080 | 300,554 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,544 |
| large-paired-correl | parse-evaluate | 123,146 | 300,554 | 2,048 | 44 | 11,544 | 5,676 | 0 | 3,456 |
| large-paired-covar | evaluate | 116,985 | 153,131 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,504 |
| large-paired-covar | parse-evaluate | 113,006 | 153,131 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,468 |
| large-paired-forecast | evaluate | 122,095 | 229,588 | 2,048 | 28 | 10,704 | 4,904 | 0 | 3,540 |
| large-paired-forecast | parse-evaluate | 122,900 | 229,588 | 2,048 | 46 | 13,472 | 6,256 | 0 | 3,544 |
| large-paired-intercept | evaluate | 122,665 | 227,141 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,464 |
| large-paired-intercept | parse-evaluate | 122,195 | 227,141 | 2,048 | 44 | 11,550 | 5,679 | 0 | 3,544 |
| large-paired-pearson | evaluate | 125,611 | 300,555 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,568 |
| large-paired-pearson | parse-evaluate | 126,315 | 300,555 | 2,048 | 44 | 11,546 | 5,677 | 0 | 3,476 |
| large-paired-rsq | evaluate | 135,145 | 300,551 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,484 |
| large-paired-rsq | parse-evaluate | 133,500 | 300,551 | 2,048 | 44 | 11,538 | 5,673 | 0 | 3,524 |
| large-paired-slope | evaluate | 124,821 | 227,137 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,548 |
| large-paired-slope | parse-evaluate | 122,256 | 227,137 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,588 |
| large-paired-steyx | evaluate | 122,571 | 300,553 | 2,048 | 26 | 8,784 | 4,328 | 0 | 3,520 |
| large-paired-steyx | parse-evaluate | 127,726 | 300,553 | 2,048 | 44 | 11,542 | 5,675 | 0 | 3,520 |
| offset-paired-correl | evaluate | 12,262 | 32,714 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,588 |
| offset-paired-correl | parse-evaluate | 12,826 | 32,714 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,576 |
| offset-paired-covar | evaluate | 13,299 | 13,931 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,468 |
| offset-paired-covar | parse-evaluate | 13,655 | 13,931 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,480 |
| offset-paired-forecast | evaluate | 15,575 | 26,068 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,580 |
| offset-paired-forecast | parse-evaluate | 16,116 | 26,068 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,520 |
| offset-paired-intercept | evaluate | 17,383 | 23,621 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,552 |
| offset-paired-intercept | parse-evaluate | 17,888 | 23,621 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,540 |
| offset-paired-pearson | evaluate | 12,233 | 32,715 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,552 |
| offset-paired-pearson | parse-evaluate | 12,839 | 32,715 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,540 |
| offset-paired-rsq | evaluate | 22,451 | 32,711 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,476 |
| offset-paired-rsq | parse-evaluate | 23,041 | 32,711 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,468 |
| offset-paired-slope | evaluate | 15,757 | 23,617 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,528 |
| offset-paired-slope | parse-evaluate | 16,198 | 23,617 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,560 |
| offset-paired-steyx | evaluate | 12,414 | 32,713 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,576 |
| offset-paired-steyx | parse-evaluate | 12,901 | 32,713 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,564 |
| pairwise-skip-paired-correl | evaluate | 11,234 | 23,914 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,580 |
| pairwise-skip-paired-correl | parse-evaluate | 11,875 | 23,914 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,540 |
| pairwise-skip-paired-covar | evaluate | 11,162 | 9,419 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,528 |
| pairwise-skip-paired-covar | parse-evaluate | 11,726 | 9,419 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,456 |
| pairwise-skip-paired-forecast | evaluate | 14,640 | 19,412 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,556 |
| pairwise-skip-paired-forecast | parse-evaluate | 15,302 | 19,412 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,524 |
| pairwise-skip-paired-intercept | evaluate | 16,720 | 16,965 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,512 |
| pairwise-skip-paired-intercept | parse-evaluate | 17,378 | 16,965 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,560 |
| pairwise-skip-paired-pearson | evaluate | 11,267 | 23,915 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,528 |
| pairwise-skip-paired-pearson | parse-evaluate | 11,787 | 23,915 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,540 |
| pairwise-skip-paired-rsq | evaluate | 21,446 | 23,911 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,532 |
| pairwise-skip-paired-rsq | parse-evaluate | 21,974 | 23,911 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,584 |
| pairwise-skip-paired-slope | evaluate | 14,840 | 16,961 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,592 |
| pairwise-skip-paired-slope | parse-evaluate | 15,384 | 16,961 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,528 |
| pairwise-skip-paired-steyx | evaluate | 11,323 | 23,913 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,476 |
| pairwise-skip-paired-steyx | parse-evaluate | 11,891 | 23,913 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,528 |
| resource-paired-correl | evaluate | 672 | 11 | 0 | 24 | 11,488 | 2,808 | 0 | 3,356 |
| resource-paired-correl | parse-evaluate | 1,232 | 11 | 0 | 60 | 17,000 | 4,154 | 0 | 3,520 |
| resource-paired-covar | evaluate | 672 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,392 |
| resource-paired-covar | parse-evaluate | 1,230 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,536 |
| resource-paired-forecast | evaluate | 852 | 16 | 0 | 28 | 13,024 | 3,192 | 0 | 3,400 |
| resource-paired-forecast | parse-evaluate | 1,422 | 16 | 0 | 64 | 18,552 | 4,542 | 0 | 3,400 |
| resource-paired-intercept | evaluate | 682 | 14 | 0 | 24 | 11,488 | 2,808 | 0 | 3,476 |
| resource-paired-intercept | parse-evaluate | 1,232 | 14 | 0 | 60 | 17,012 | 4,157 | 0 | 3,452 |
| resource-paired-pearson | evaluate | 682 | 12 | 0 | 24 | 11,488 | 2,808 | 0 | 3,396 |
| resource-paired-pearson | parse-evaluate | 1,242 | 12 | 0 | 60 | 17,004 | 4,155 | 0 | 3,476 |
| resource-paired-rsq | evaluate | 675 | 8 | 0 | 24 | 11,488 | 2,808 | 0 | 3,372 |
| resource-paired-rsq | parse-evaluate | 1,205 | 8 | 0 | 60 | 16,988 | 4,151 | 0 | 3,476 |
| resource-paired-slope | evaluate | 672 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,460 |
| resource-paired-slope | parse-evaluate | 1,235 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,472 |
| resource-paired-steyx | evaluate | 680 | 10 | 0 | 24 | 11,488 | 2,808 | 0 | 3,516 |
| resource-paired-steyx | parse-evaluate | 1,227 | 10 | 0 | 60 | 16,996 | 4,153 | 0 | 3,428 |
| shape-reject-paired-correl | evaluate | 2,617 | 26 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,416 |
| shape-reject-paired-correl | parse-evaluate | 3,440 | 26 | 0 | 2,320 | 576,560 | 6,295 | 0 | 3,460 |
| shape-reject-paired-covar | evaluate | 2,606 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,452 |
| shape-reject-paired-covar | parse-evaluate | 3,431 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,472 |
| shape-reject-paired-forecast | evaluate | 2,586 | 26 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,440 |
| shape-reject-paired-forecast | parse-evaluate | 3,401 | 26 | 0 | 2,320 | 576,720 | 6,297 | 0 | 3,516 |
| shape-reject-paired-intercept | evaluate | 2,598 | 29 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,564 |
| shape-reject-paired-intercept | parse-evaluate | 3,453 | 29 | 0 | 2,320 | 576,800 | 6,298 | 0 | 3,464 |
| shape-reject-paired-pearson | evaluate | 2,642 | 27 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,412 |
| shape-reject-paired-pearson | parse-evaluate | 3,496 | 27 | 0 | 2,320 | 576,640 | 6,296 | 0 | 3,564 |
| shape-reject-paired-rsq | evaluate | 1,897 | 17 | 0 | 1,040 | 351,360 | 4,328 | 0 | 3,428 |
| shape-reject-paired-rsq | parse-evaluate | 2,437 | 17 | 0 | 1,760 | 461,360 | 5,671 | 0 | 3,460 |
| shape-reject-paired-slope | evaluate | 2,628 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,532 |
| shape-reject-paired-slope | parse-evaluate | 3,486 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,548 |
| shape-reject-paired-steyx | evaluate | 2,642 | 25 | 0 | 1,360 | 403,840 | 4,552 | 0 | 3,544 |
| shape-reject-paired-steyx | parse-evaluate | 3,507 | 25 | 0 | 2,320 | 576,480 | 6,294 | 0 | 3,528 |
| small-paired-correl | evaluate | 12,196 | 32,714 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,524 |
| small-paired-correl | parse-evaluate | 12,822 | 32,714 | 128 | 1,760 | 461,600 | 5,674 | 0 | 3,480 |
| small-paired-covar | evaluate | 12,527 | 13,931 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,448 |
| small-paired-covar | parse-evaluate | 12,453 | 13,931 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,564 |
| small-paired-forecast | evaluate | 15,445 | 26,068 | 128 | 1,120 | 428,160 | 4,904 | 0 | 3,516 |
| small-paired-forecast | parse-evaluate | 15,992 | 26,068 | 128 | 1,840 | 538,720 | 6,254 | 0 | 3,560 |
| small-paired-intercept | evaluate | 17,461 | 23,621 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,548 |
| small-paired-intercept | parse-evaluate | 17,965 | 23,621 | 128 | 1,760 | 461,840 | 5,677 | 0 | 3,588 |
| small-paired-pearson | evaluate | 12,316 | 32,715 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,544 |
| small-paired-pearson | parse-evaluate | 12,891 | 32,715 | 128 | 1,760 | 461,680 | 5,675 | 0 | 3,568 |
| small-paired-rsq | evaluate | 22,489 | 32,711 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,436 |
| small-paired-rsq | parse-evaluate | 23,062 | 32,711 | 128 | 1,760 | 461,360 | 5,671 | 0 | 3,456 |
| small-paired-slope | evaluate | 15,887 | 23,617 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,464 |
| small-paired-slope | parse-evaluate | 16,288 | 23,617 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,520 |
| small-paired-steyx | evaluate | 12,301 | 32,713 | 128 | 1,040 | 351,360 | 4,328 | 0 | 3,548 |
| small-paired-steyx | parse-evaluate | 12,922 | 32,713 | 128 | 1,760 | 461,520 | 5,673 | 0 | 3,528 |

The resolver is an immutable borrowing fixture. Each child validates one finite numerical result or typed failure before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
