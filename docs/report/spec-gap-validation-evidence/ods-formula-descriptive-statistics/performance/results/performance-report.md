# ODS descriptive-statistics evaluator performance profile

The baseline is committed `b8e5d5fe257fd95747c69a3c44a53cedd96f77ed`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,712 | 46,107 | +0.9% | 88 | 88 | 3,712 | 3,788 |
| array-control-16x16-arithmetic | parse-evaluate | 58,545 | 59,092 | +0.9% | 352 | 352 | 3,740 | 3,840 |
| array-control-16x16-sin | evaluate | 126,205 | 127,405 | +1.0% | 2,144 | 2,144 | 3,816 | 3,840 |
| array-control-16x16-sin | parse-evaluate | 139,605 | 140,290 | +0.5% | 2,412 | 2,412 | 3,812 | 3,840 |
| array-control-4x4-arithmetic | evaluate | 4,375 | 4,324 | -1.2% | 1,120 | 1,120 | 3,492 | 3,528 |
| array-control-4x4-arithmetic | parse-evaluate | 5,173 | 5,189 | +0.3% | 2,240 | 2,240 | 3,508 | 3,492 |
| array-control-4x4-sin | evaluate | 9,739 | 9,650 | -0.9% | 3,840 | 3,840 | 3,544 | 3,548 |
| array-control-4x4-sin | parse-evaluate | 10,583 | 10,636 | +0.5% | 5,040 | 5,040 | 3,560 | 3,564 |
| database-control-dstdev | evaluate | 3,770 | 3,930 | +4.2% | 20 | 20 | 3,576 | 3,512 |
| database-control-dstdev | parse-evaluate | 4,290 | 4,430 | +3.3% | 29 | 29 | 3,560 | 3,496 |
| database-control-dsum | evaluate | 3,870 | 3,980 | +2.8% | 20 | 20 | 3,512 | 3,520 |
| database-control-dsum | parse-evaluate | 4,440 | 4,450 | +0.2% | 29 | 29 | 3,564 | 3,532 |
| database-control-dvar | evaluate | 3,810 | 3,890 | +2.1% | 20 | 20 | 3,560 | 3,536 |
| database-control-dvar | parse-evaluate | 4,330 | 4,390 | +1.4% | 29 | 29 | 3,548 | 3,492 |
| literal-aggregate-4x1-sum | evaluate | 1,740 | 1,760 | +1.1% | 10 | 10 | 3,548 | 3,532 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,420 | 2,460 | +1.7% | 23 | 23 | 3,504 | 3,548 |
| reference-aggregate-64x4-sum | evaluate | 10,197 | 10,217 | +0.2% | 32 | 32 | 3,396 | 3,556 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,540 | 10,690 | +1.4% | 60 | 60 | 3,500 | 3,548 |
| reference-array-16x4-arithmetic | evaluate | 9,270 | 9,307 | +0.4% | 800 | 800 | 3,504 | 3,568 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,584 | 9,788 | +2.1% | 1,280 | 1,280 | 3,448 | 3,520 |
| reference-conditional-256x4-sumifs | evaluate | 80,145 | 78,800 | -1.7% | 42 | 42 | 3,580 | 3,548 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,375 | 79,905 | -0.6% | 68 | 68 | 3,576 | 3,556 |
| reference-control-average | evaluate | 10,252 | 10,215 | -0.4% | 32 | 32 | 3,448 | 3,484 |
| reference-control-average | parse-evaluate | 10,567 | 10,700 | +1.3% | 60 | 60 | 3,540 | 3,548 |
| reference-control-counta | evaluate | 8,607 | 8,672 | +0.8% | 32 | 32 | 3,500 | 3,536 |
| reference-control-counta | parse-evaluate | 8,970 | 9,165 | +2.2% | 60 | 60 | 3,548 | 3,548 |
| representative-median | evaluate | 1,533 | 1,587 | +3.5% | 9,000 | 9,000 | 3,312 | 3,316 |
| representative-median | parse-evaluate | 1,795 | 1,845 | +2.8% | 14,000 | 14,000 | 3,336 | 3,320 |
| representative-percentrank | evaluate | 942 | 979 | +3.9% | 7,000 | 7,000 | 3,272 | 3,448 |
| representative-percentrank | parse-evaluate | 1,181 | 1,224 | +3.6% | 11,000 | 11,000 | 3,384 | 3,300 |
| representative-rank | evaluate | 905 | 912 | +0.8% | 7,000 | 7,000 | 3,312 | 3,324 |
| representative-rank | parse-evaluate | 1,118 | 1,144 | +2.3% | 11,000 | 11,000 | 3,332 | 3,324 |
| scalar-aggregate-sum | evaluate | 559 | 574 | +2.7% | 4,000 | 4,000 | 3,332 | 3,388 |
| scalar-aggregate-sum | parse-evaluate | 748 | 754 | +0.8% | 8,000 | 8,000 | 3,376 | 3,324 |
| scalar-control-arithmetic | evaluate | 560 | 561 | +0.2% | 5,000 | 5,000 | 3,388 | 3,316 |
| scalar-control-arithmetic | parse-evaluate | 692 | 702 | +1.4% | 8,000 | 8,000 | 3,332 | 3,304 |
| scalar-control-average | evaluate | 1,010 | 1,012 | +0.2% | 6,000 | 6,000 | 3,324 | 3,316 |
| scalar-control-average | parse-evaluate | 1,231 | 1,241 | +0.8% | 10,000 | 10,000 | 3,332 | 3,340 |
| scalar-control-counta | evaluate | 824 | 806 | -2.2% | 6,000 | 6,000 | 3,320 | 3,340 |
| scalar-control-counta | parse-evaluate | 1,062 | 1,070 | +0.8% | 10,000 | 10,000 | 3,320 | 3,300 |
| scalar-control-imsum | evaluate | 1,397 | 1,402 | +0.4% | 8,000 | 8,000 | 3,324 | 3,292 |
| scalar-control-imsum | parse-evaluate | 1,959 | 1,977 | +0.9% | 17,000 | 17,000 | 3,324 | 3,452 |
| scalar-control-sin | evaluate | 460 | 460 | +0.0% | 4,000 | 4,000 | 3,348 | 3,456 |
| scalar-control-sin | parse-evaluate | 644 | 642 | -0.3% | 8,000 | 8,000 | 3,368 | 3,444 |
| scalar-control-stdev | evaluate | 837 | 847 | +1.2% | 6,000 | 6,000 | 3,320 | 3,340 |
| scalar-control-stdev | parse-evaluate | 1,054 | 1,060 | +0.6% | 10,000 | 10,000 | 3,320 | 3,336 |
| scalar-control-var | evaluate | 826 | 830 | +0.5% | 6,000 | 6,000 | 3,364 | 3,324 |
| scalar-control-var | parse-evaluate | 1,034 | 1,057 | +2.2% | 10,000 | 10,000 | 3,328 | 3,320 |

## Descriptive-statistics workloads

The candidate matrix covers scalar and inline inputs, 64/256/1024-row borrowed references, projected cache scaling, domain/list-shape admission or refusal, formula-error continuation, cancellation, and resource limits. Numerical labels and pass-count behavior are taken from the frozen contract and are retained with the raw case receipts.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-descriptive-avedev | evaluate | 322 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,492 |
| cancellation-descriptive-avedev | parse-evaluate | 732 | 3 | 0 | 36 | 6,876 | 2,749 | 0 | 3,544 |
| cancellation-descriptive-devsq | evaluate | 322 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,516 |
| cancellation-descriptive-devsq | parse-evaluate | 725 | 3 | 0 | 36 | 6,872 | 2,748 | 0 | 3,504 |
| cancellation-descriptive-geomean | evaluate | 320 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,484 |
| cancellation-descriptive-geomean | parse-evaluate | 732 | 3 | 0 | 36 | 6,880 | 2,750 | 0 | 3,552 |
| cancellation-descriptive-harmean | evaluate | 315 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,548 |
| cancellation-descriptive-harmean | parse-evaluate | 742 | 3 | 0 | 36 | 6,880 | 2,750 | 0 | 3,496 |
| cancellation-descriptive-kurt | evaluate | 322 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,564 |
| cancellation-descriptive-kurt | parse-evaluate | 712 | 3 | 0 | 36 | 6,868 | 2,747 | 0 | 3,540 |
| cancellation-descriptive-skew | evaluate | 317 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,548 |
| cancellation-descriptive-skew | parse-evaluate | 720 | 3 | 0 | 36 | 6,868 | 2,747 | 0 | 3,548 |
| cancellation-descriptive-skewp | evaluate | 317 | 3 | 0 | 8 | 1,416 | 1,416 | 0 | 3,568 |
| cancellation-descriptive-skewp | parse-evaluate | 727 | 3 | 0 | 36 | 6,872 | 2,748 | 0 | 3,544 |
| domain-descriptive-avedev | evaluate | 2,134 | 29 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,460 |
| domain-descriptive-avedev | parse-evaluate | 2,412 | 29 | 0 | 12,000 | 2,271,000 | 1,535 | 0 | 3,480 |
| domain-descriptive-devsq | evaluate | 3,081 | 28 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,636 |
| domain-descriptive-devsq | parse-evaluate | 3,364 | 28 | 0 | 12,000 | 2,270,000 | 1,534 | 0 | 3,572 |
| domain-descriptive-geomean | evaluate | 743 | 24 | 0 | 5,000 | 496,000 | 448 | 0 | 3,388 |
| domain-descriptive-geomean | parse-evaluate | 963 | 24 | 0 | 9,000 | 958,000 | 878 | 0 | 3,324 |
| domain-descriptive-harmean | evaluate | 1,358 | 1,854 | 0 | 10,000 | 888,000 | 656 | 0 | 3,308 |
| domain-descriptive-harmean | parse-evaluate | 1,671 | 1,854 | 0 | 15,000 | 2,120,000 | 1,472 | 0 | 3,332 |
| domain-descriptive-kurt | evaluate | 1,007 | 25 | 0 | 6,000 | 848,000 | 624 | 0 | 3,580 |
| domain-descriptive-kurt | parse-evaluate | 1,222 | 25 | 0 | 10,000 | 1,308,000 | 1,052 | 0 | 3,580 |
| domain-descriptive-skew | evaluate | 758 | 19 | 0 | 5,000 | 496,000 | 448 | 0 | 3,580 |
| domain-descriptive-skew | parse-evaluate | 968 | 19 | 0 | 9,000 | 954,000 | 874 | 0 | 3,580 |
| domain-descriptive-skewp | evaluate | 770 | 20 | 0 | 5,000 | 496,000 | 448 | 0 | 3,600 |
| domain-descriptive-skewp | parse-evaluate | 988 | 20 | 0 | 9,000 | 955,000 | 875 | 0 | 3,640 |
| error-descriptive-avedev | evaluate | 4,336 | 141 | 64 | 640 | 113,280 | 1,416 | 0 | 3,492 |
| error-descriptive-avedev | parse-evaluate | 4,760 | 141 | 64 | 1,200 | 222,480 | 2,749 | 0 | 3,536 |
| error-descriptive-devsq | evaluate | 5,041 | 140 | 64 | 640 | 113,280 | 1,416 | 0 | 3,528 |
| error-descriptive-devsq | parse-evaluate | 5,450 | 140 | 64 | 1,200 | 222,400 | 2,748 | 0 | 3,492 |
| error-descriptive-geomean | evaluate | 4,281 | 142 | 64 | 640 | 113,280 | 1,416 | 0 | 3,532 |
| error-descriptive-geomean | parse-evaluate | 4,663 | 142 | 64 | 1,200 | 222,560 | 2,750 | 0 | 3,544 |
| error-descriptive-harmean | evaluate | 4,439 | 142 | 64 | 640 | 113,280 | 1,416 | 0 | 3,524 |
| error-descriptive-harmean | parse-evaluate | 4,848 | 142 | 64 | 1,200 | 222,560 | 2,750 | 0 | 3,544 |
| error-descriptive-kurt | evaluate | 6,440 | 139 | 64 | 640 | 113,280 | 1,416 | 0 | 3,552 |
| error-descriptive-kurt | parse-evaluate | 6,754 | 139 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,492 |
| error-descriptive-skew | evaluate | 5,699 | 139 | 64 | 640 | 113,280 | 1,416 | 0 | 3,540 |
| error-descriptive-skew | parse-evaluate | 5,995 | 139 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,492 |
| error-descriptive-skewp | evaluate | 5,758 | 140 | 64 | 640 | 113,280 | 1,416 | 0 | 3,504 |
| error-descriptive-skewp | parse-evaluate | 6,095 | 140 | 64 | 1,200 | 222,400 | 2,748 | 0 | 3,552 |
| inline-descriptive-avedev | evaluate | 2,900 | 44 | 0 | 11 | 5,016 | 4,392 | 0 | 3,792 |
| inline-descriptive-avedev | parse-evaluate | 3,290 | 44 | 0 | 19 | 6,378 | 5,242 | 0 | 3,804 |
| inline-descriptive-devsq | evaluate | 3,090 | 34 | 0 | 11 | 5,016 | 4,392 | 0 | 3,756 |
| inline-descriptive-devsq | parse-evaluate | 3,580 | 34 | 0 | 19 | 6,377 | 5,241 | 0 | 3,732 |
| inline-descriptive-geomean | evaluate | 2,040 | 36 | 0 | 11 | 5,016 | 4,392 | 0 | 3,780 |
| inline-descriptive-geomean | parse-evaluate | 2,490 | 36 | 0 | 19 | 6,379 | 5,243 | 0 | 3,756 |
| inline-descriptive-harmean | evaluate | 3,000 | 36 | 0 | 11 | 5,016 | 4,392 | 0 | 3,544 |
| inline-descriptive-harmean | parse-evaluate | 3,450 | 36 | 0 | 19 | 6,379 | 5,243 | 0 | 3,548 |
| inline-descriptive-kurt | evaluate | 12,820 | 33 | 0 | 11 | 5,016 | 4,392 | 0 | 3,756 |
| inline-descriptive-kurt | parse-evaluate | 13,201 | 33 | 0 | 19 | 6,376 | 5,240 | 0 | 3,784 |
| inline-descriptive-skew | evaluate | 4,540 | 33 | 0 | 11 | 5,016 | 4,392 | 0 | 3,808 |
| inline-descriptive-skew | parse-evaluate | 4,990 | 33 | 0 | 19 | 6,376 | 5,240 | 0 | 3,840 |
| inline-descriptive-skewp | evaluate | 4,500 | 34 | 0 | 11 | 5,016 | 4,392 | 0 | 3,772 |
| inline-descriptive-skewp | parse-evaluate | 4,950 | 34 | 0 | 19 | 6,377 | 5,241 | 0 | 3,780 |
| list-admit-descriptive-avedev | evaluate | 5,823 | 85 | 32 | 960 | 155,520 | 1,576 | 0 | 3,740 |
| list-admit-descriptive-avedev | parse-evaluate | 6,345 | 85 | 32 | 1,680 | 265,600 | 2,920 | 0 | 3,840 |
| list-admit-descriptive-geomean | evaluate | 2,711 | 53 | 16 | 960 | 155,520 | 1,576 | 0 | 3,804 |
| list-admit-descriptive-geomean | parse-evaluate | 3,250 | 53 | 16 | 1,680 | 265,680 | 2,921 | 0 | 3,772 |
| list-admit-descriptive-harmean | evaluate | 3,656 | 53 | 16 | 960 | 155,520 | 1,576 | 0 | 3,480 |
| list-admit-descriptive-harmean | parse-evaluate | 4,207 | 53 | 16 | 1,680 | 265,680 | 2,921 | 0 | 3,484 |
| list-admit-descriptive-kurt | evaluate | 13,968 | 50 | 16 | 960 | 155,520 | 1,576 | 0 | 3,780 |
| list-admit-descriptive-kurt | parse-evaluate | 14,427 | 50 | 16 | 1,680 | 265,440 | 2,918 | 0 | 3,804 |
| list-admit-descriptive-skew | evaluate | 5,314 | 50 | 16 | 960 | 155,520 | 1,576 | 0 | 3,764 |
| list-admit-descriptive-skew | parse-evaluate | 5,908 | 50 | 16 | 1,680 | 265,440 | 2,918 | 0 | 3,744 |
| list-refusal-descriptive-devsq | evaluate | 1,792 | 17 | 0 | 960 | 155,520 | 1,576 | 0 | 3,508 |
| list-refusal-descriptive-devsq | parse-evaluate | 2,315 | 17 | 0 | 1,680 | 265,520 | 2,919 | 0 | 3,548 |
| list-refusal-descriptive-skewp | evaluate | 1,802 | 17 | 0 | 960 | 155,520 | 1,576 | 0 | 3,516 |
| list-refusal-descriptive-skewp | parse-evaluate | 2,351 | 17 | 0 | 1,680 | 265,520 | 2,919 | 0 | 3,556 |
| projected-descriptive-1024-avedev | evaluate | 777,934 | 16,445 | 8,192 | 28 | 3,384 | 2,360 | 0 | 3,748 |
| projected-descriptive-1024-avedev | parse-evaluate | 783,914 | 16,445 | 8,192 | 44 | 7,434 | 4,970 | 0 | 3,788 |
| projected-descriptive-1024-devsq | evaluate | 256,361 | 8,251 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,732 |
| projected-descriptive-1024-devsq | parse-evaluate | 257,101 | 8,251 | 4,096 | 44 | 7,433 | 4,969 | 0 | 3,804 |
| projected-descriptive-1024-geomean | evaluate | 200,781 | 8,253 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,796 |
| projected-descriptive-1024-geomean | parse-evaluate | 201,291 | 8,253 | 4,096 | 44 | 7,435 | 4,971 | 0 | 3,800 |
| projected-descriptive-1024-harmean | evaluate | 221,561 | 8,253 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,564 |
| projected-descriptive-1024-harmean | parse-evaluate | 224,311 | 8,253 | 4,096 | 44 | 7,435 | 4,971 | 0 | 3,532 |
| projected-descriptive-1024-kurt | evaluate | 349,731 | 8,250 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,780 |
| projected-descriptive-1024-kurt | parse-evaluate | 350,872 | 8,250 | 4,096 | 44 | 7,432 | 4,968 | 0 | 3,800 |
| projected-descriptive-1024-skew | evaluate | 292,602 | 8,250 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,804 |
| projected-descriptive-1024-skew | parse-evaluate | 295,821 | 8,250 | 4,096 | 44 | 7,432 | 4,968 | 0 | 3,796 |
| projected-descriptive-1024-skewp | evaluate | 295,221 | 8,251 | 4,096 | 28 | 3,384 | 2,360 | 0 | 3,776 |
| projected-descriptive-1024-skewp | parse-evaluate | 294,881 | 8,251 | 4,096 | 44 | 7,433 | 4,969 | 0 | 3,804 |
| projected-descriptive-256-avedev | evaluate | 196,391 | 4,157 | 2,048 | 28 | 3,384 | 2,360 | 0 | 3,744 |
| projected-descriptive-256-avedev | parse-evaluate | 198,831 | 4,157 | 2,048 | 44 | 7,433 | 4,969 | 0 | 3,756 |
| projected-descriptive-256-devsq | evaluate | 67,691 | 2,107 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,792 |
| projected-descriptive-256-devsq | parse-evaluate | 69,400 | 2,107 | 1,024 | 44 | 7,432 | 4,968 | 0 | 3,800 |
| projected-descriptive-256-geomean | evaluate | 53,741 | 2,109 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,796 |
| projected-descriptive-256-geomean | parse-evaluate | 54,731 | 2,109 | 1,024 | 44 | 7,434 | 4,970 | 0 | 3,760 |
| projected-descriptive-256-harmean | evaluate | 59,030 | 2,109 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,540 |
| projected-descriptive-256-harmean | parse-evaluate | 59,850 | 2,109 | 1,024 | 44 | 7,434 | 4,970 | 0 | 3,552 |
| projected-descriptive-256-kurt | evaluate | 98,101 | 2,106 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,808 |
| projected-descriptive-256-kurt | parse-evaluate | 98,970 | 2,106 | 1,024 | 44 | 7,431 | 4,967 | 0 | 3,748 |
| projected-descriptive-256-skew | evaluate | 79,270 | 2,106 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,812 |
| projected-descriptive-256-skew | parse-evaluate | 79,201 | 2,106 | 1,024 | 44 | 7,431 | 4,967 | 0 | 3,796 |
| projected-descriptive-256-skewp | evaluate | 79,190 | 2,107 | 1,024 | 28 | 3,384 | 2,360 | 0 | 3,804 |
| projected-descriptive-256-skewp | parse-evaluate | 79,020 | 2,107 | 1,024 | 44 | 7,432 | 4,968 | 0 | 3,840 |
| projected-descriptive-64-avedev | evaluate | 53,140 | 1,085 | 512 | 28 | 3,384 | 2,360 | 0 | 3,812 |
| projected-descriptive-64-avedev | parse-evaluate | 53,430 | 1,085 | 512 | 44 | 7,432 | 4,968 | 0 | 3,800 |
| projected-descriptive-64-devsq | evaluate | 21,690 | 571 | 256 | 28 | 3,384 | 2,360 | 0 | 3,804 |
| projected-descriptive-64-devsq | parse-evaluate | 22,690 | 571 | 256 | 44 | 7,431 | 4,967 | 0 | 3,808 |
| projected-descriptive-64-geomean | evaluate | 17,200 | 573 | 256 | 28 | 3,384 | 2,360 | 0 | 3,832 |
| projected-descriptive-64-geomean | parse-evaluate | 18,080 | 573 | 256 | 44 | 7,433 | 4,969 | 0 | 3,808 |
| projected-descriptive-64-harmean | evaluate | 19,460 | 573 | 256 | 28 | 3,384 | 2,360 | 0 | 3,492 |
| projected-descriptive-64-harmean | parse-evaluate | 20,091 | 573 | 256 | 44 | 7,433 | 4,969 | 0 | 3,556 |
| projected-descriptive-64-kurt | evaluate | 36,060 | 570 | 256 | 28 | 3,384 | 2,360 | 0 | 3,748 |
| projected-descriptive-64-kurt | parse-evaluate | 36,880 | 570 | 256 | 44 | 7,430 | 4,966 | 0 | 3,804 |
| projected-descriptive-64-skew | evaluate | 25,170 | 570 | 256 | 28 | 3,384 | 2,360 | 0 | 3,748 |
| projected-descriptive-64-skew | parse-evaluate | 26,100 | 570 | 256 | 44 | 7,430 | 4,966 | 0 | 3,804 |
| projected-descriptive-64-skewp | evaluate | 25,120 | 571 | 256 | 28 | 3,384 | 2,360 | 0 | 3,804 |
| projected-descriptive-64-skewp | parse-evaluate | 26,160 | 571 | 256 | 44 | 7,431 | 4,967 | 0 | 3,804 |
| reference-descriptive-1024-avedev | evaluate | 774,904 | 16,399 | 8,192 | 8 | 1,416 | 1,416 | 0 | 3,752 |
| reference-descriptive-1024-avedev | parse-evaluate | 788,144 | 16,399 | 8,192 | 15 | 2,783 | 2,751 | 0 | 3,792 |
| reference-descriptive-1024-devsq | evaluate | 249,452 | 8,205 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,788 |
| reference-descriptive-1024-devsq | parse-evaluate | 247,482 | 8,205 | 4,096 | 15 | 2,782 | 2,750 | 0 | 3,796 |
| reference-descriptive-1024-geomean | evaluate | 197,551 | 8,207 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,768 |
| reference-descriptive-1024-geomean | parse-evaluate | 197,511 | 8,207 | 4,096 | 15 | 2,784 | 2,752 | 0 | 3,844 |
| reference-descriptive-1024-harmean | evaluate | 217,611 | 8,207 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,548 |
| reference-descriptive-1024-harmean | parse-evaluate | 215,931 | 8,207 | 4,096 | 15 | 2,784 | 2,752 | 0 | 3,552 |
| reference-descriptive-1024-kurt | evaluate | 340,092 | 8,204 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,748 |
| reference-descriptive-1024-kurt | parse-evaluate | 344,752 | 8,204 | 4,096 | 15 | 2,781 | 2,749 | 0 | 3,856 |
| reference-descriptive-1024-skew | evaluate | 291,652 | 8,204 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,804 |
| reference-descriptive-1024-skew | parse-evaluate | 291,922 | 8,204 | 4,096 | 15 | 2,781 | 2,749 | 0 | 3,804 |
| reference-descriptive-1024-skewp | evaluate | 291,731 | 8,205 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,784 |
| reference-descriptive-1024-skewp | parse-evaluate | 291,902 | 8,205 | 4,096 | 15 | 2,782 | 2,750 | 0 | 3,836 |
| reference-descriptive-256-avedev | evaluate | 193,481 | 4,111 | 2,048 | 16 | 2,832 | 1,416 | 0 | 3,836 |
| reference-descriptive-256-avedev | parse-evaluate | 193,461 | 4,111 | 2,048 | 30 | 5,564 | 2,750 | 0 | 3,736 |
| reference-descriptive-256-devsq | evaluate | 64,740 | 2,061 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,808 |
| reference-descriptive-256-devsq | parse-evaluate | 64,900 | 2,061 | 1,024 | 30 | 5,562 | 2,749 | 0 | 3,796 |
| reference-descriptive-256-geomean | evaluate | 50,235 | 2,063 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,776 |
| reference-descriptive-256-geomean | parse-evaluate | 50,910 | 2,063 | 1,024 | 30 | 5,566 | 2,751 | 0 | 3,784 |
| reference-descriptive-256-harmean | evaluate | 56,220 | 2,063 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,508 |
| reference-descriptive-256-harmean | parse-evaluate | 56,100 | 2,063 | 1,024 | 30 | 5,566 | 2,751 | 0 | 3,552 |
| reference-descriptive-256-kurt | evaluate | 95,245 | 2,060 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,804 |
| reference-descriptive-256-kurt | parse-evaluate | 95,250 | 2,060 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,804 |
| reference-descriptive-256-skew | evaluate | 75,225 | 2,060 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,804 |
| reference-descriptive-256-skew | parse-evaluate | 75,905 | 2,060 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,828 |
| reference-descriptive-256-skewp | evaluate | 75,270 | 2,061 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,800 |
| reference-descriptive-256-skewp | parse-evaluate | 75,755 | 2,061 | 1,024 | 30 | 5,562 | 2,749 | 0 | 3,796 |
| reference-descriptive-64-avedev | evaluate | 49,482 | 1,039 | 512 | 32 | 5,664 | 1,416 | 0 | 3,740 |
| reference-descriptive-64-avedev | parse-evaluate | 49,492 | 1,039 | 512 | 60 | 11,124 | 2,749 | 0 | 3,804 |
| reference-descriptive-64-devsq | evaluate | 18,090 | 525 | 256 | 32 | 5,664 | 1,416 | 0 | 3,808 |
| reference-descriptive-64-devsq | parse-evaluate | 18,480 | 525 | 256 | 60 | 11,120 | 2,748 | 0 | 3,756 |
| reference-descriptive-64-geomean | evaluate | 13,427 | 527 | 256 | 32 | 5,664 | 1,416 | 0 | 3,804 |
| reference-descriptive-64-geomean | parse-evaluate | 13,885 | 527 | 256 | 60 | 11,128 | 2,750 | 0 | 3,768 |
| reference-descriptive-64-harmean | evaluate | 15,592 | 527 | 256 | 32 | 5,664 | 1,416 | 0 | 3,544 |
| reference-descriptive-64-harmean | parse-evaluate | 16,105 | 527 | 256 | 60 | 11,128 | 2,750 | 0 | 3,548 |
| reference-descriptive-64-kurt | evaluate | 34,877 | 524 | 256 | 32 | 5,664 | 1,416 | 0 | 3,788 |
| reference-descriptive-64-kurt | parse-evaluate | 32,995 | 524 | 256 | 60 | 11,116 | 2,747 | 0 | 3,792 |
| reference-descriptive-64-skew | evaluate | 21,530 | 524 | 256 | 32 | 5,664 | 1,416 | 0 | 3,788 |
| reference-descriptive-64-skew | parse-evaluate | 21,947 | 524 | 256 | 60 | 11,116 | 2,747 | 0 | 3,804 |
| reference-descriptive-64-skewp | evaluate | 21,540 | 525 | 256 | 32 | 5,664 | 1,416 | 0 | 3,812 |
| reference-descriptive-64-skewp | parse-evaluate | 21,925 | 525 | 256 | 60 | 11,120 | 2,748 | 0 | 3,736 |
| resource-descriptive-avedev | evaluate | 390 | 9 | 0 | 16 | 1,184 | 296 | 0 | 3,596 |
| resource-descriptive-avedev | parse-evaluate | 807 | 9 | 0 | 44 | 6,644 | 1,629 | 0 | 3,536 |
| resource-descriptive-devsq | evaluate | 390 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,544 |
| resource-descriptive-devsq | parse-evaluate | 822 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,548 |
| resource-descriptive-geomean | evaluate | 392 | 10 | 0 | 16 | 1,184 | 296 | 0 | 3,540 |
| resource-descriptive-geomean | parse-evaluate | 820 | 10 | 0 | 44 | 6,648 | 1,630 | 0 | 3,552 |
| resource-descriptive-harmean | evaluate | 397 | 10 | 0 | 16 | 1,184 | 296 | 0 | 3,480 |
| resource-descriptive-harmean | parse-evaluate | 812 | 10 | 0 | 44 | 6,648 | 1,630 | 0 | 3,540 |
| resource-descriptive-kurt | evaluate | 392 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,528 |
| resource-descriptive-kurt | parse-evaluate | 800 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,552 |
| resource-descriptive-skew | evaluate | 395 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,552 |
| resource-descriptive-skew | parse-evaluate | 817 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,532 |
| resource-descriptive-skewp | evaluate | 392 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,548 |
| resource-descriptive-skewp | parse-evaluate | 820 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,532 |
| scalar-descriptive-avedev | evaluate | 1,655 | 33 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,472 |
| scalar-descriptive-avedev | parse-evaluate | 1,952 | 33 | 0 | 12,000 | 2,272,000 | 1,536 | 0 | 3,376 |
| scalar-descriptive-devsq | evaluate | 2,075 | 32 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,648 |
| scalar-descriptive-devsq | parse-evaluate | 2,365 | 32 | 0 | 12,000 | 2,271,000 | 1,535 | 0 | 3,580 |
| scalar-descriptive-geomean | evaluate | 1,050 | 34 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,340 |
| scalar-descriptive-geomean | parse-evaluate | 1,318 | 34 | 0 | 12,000 | 2,273,000 | 1,537 | 0 | 3,324 |
| scalar-descriptive-harmean | evaluate | 1,516 | 34 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,328 |
| scalar-descriptive-harmean | parse-evaluate | 1,810 | 34 | 0 | 12,000 | 2,273,000 | 1,537 | 0 | 3,304 |
| scalar-descriptive-kurt | evaluate | 11,893 | 31 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,692 |
| scalar-descriptive-kurt | parse-evaluate | 12,173 | 31 | 0 | 12,000 | 2,270,000 | 1,534 | 0 | 3,556 |
| scalar-descriptive-skew | evaluate | 3,454 | 31 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,560 |
| scalar-descriptive-skew | parse-evaluate | 3,749 | 31 | 0 | 12,000 | 2,270,000 | 1,534 | 0 | 3,704 |
| scalar-descriptive-skewp | evaluate | 3,505 | 32 | 0 | 7,000 | 1,040,000 | 720 | 0 | 3,580 |
| scalar-descriptive-skewp | parse-evaluate | 3,769 | 32 | 0 | 12,000 | 2,271,000 | 1,535 | 0 | 3,704 |

The resolver is an immutable borrowing fixture. Each child validates one finite numerical result or typed failure before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.

## Review disposition

Across the 48 matched-control phase groups, elapsed p50 deltas range from
`-2.18%` to `+4.24%`; allocator calls, evaluator work, and normalized
resolver reads are unchanged. One RSS threshold flag is retained:
`representative-percentrank` evaluate is `3,272 -> 3,448 KiB` (`+5.38%`).
The paired parse-evaluate row is `3,384 -> 3,300 KiB` (`-2.48%`). This is an
isolated RSS observation with no corresponding allocation, work, or read
change; it remains in the raw receipts and was not cleared by rerunning.

The frozen verifier's cancellation normalization defect is documented in
`verification-notes.md`; an exact-total-read diagnostic check passes all 4,170
rows while retaining the authoritative one-read-per-child receipt.
