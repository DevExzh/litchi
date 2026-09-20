# ODS byte-position text evaluator performance profile

The baseline is committed `3844f235bac545ff0ae1580b97612883c1fd9f89`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 47,260 | 46,047 | -2.6% | 2,311 | 2,311 | 88 | 88 | 4,000 | 3,924 |
| array-control-16x16-arithmetic | parse-evaluate | 56,520 | 56,567 | +0.1% | 2,311 | 2,311 | 352 | 352 | 3,952 | 3,888 |
| array-control-16x16-sin | evaluate | 128,008 | 129,650 | +1.3% | 2,311 | 2,311 | 2,144 | 2,144 | 4,276 | 4,344 |
| array-control-16x16-sin | parse-evaluate | 138,975 | 139,108 | +0.1% | 2,311 | 2,311 | 2,412 | 2,412 | 4,264 | 4,312 |
| array-control-4x4-arithmetic | evaluate | 4,477 | 4,463 | -0.3% | 151 | 151 | 1,120 | 1,120 | 3,708 | 3,708 |
| array-control-4x4-arithmetic | parse-evaluate | 5,288 | 5,233 | -1.0% | 151 | 151 | 2,240 | 2,240 | 3,744 | 3,724 |
| array-control-4x4-sin | evaluate | 9,811 | 9,956 | +1.5% | 151 | 151 | 3,840 | 3,840 | 3,988 | 4,092 |
| array-control-4x4-sin | parse-evaluate | 10,758 | 10,858 | +0.9% | 151 | 151 | 5,040 | 5,040 | 4,024 | 4,040 |
| concat-borrowed-literals | evaluate | 730 | 730 | +0.0% | 21 | 21 | 6,000 | 6,000 | 3,516 | 3,408 |
| concat-borrowed-literals | parse-evaluate | 869 | 865 | -0.5% | 21 | 21 | 9,000 | 9,000 | 3,576 | 3,416 |
| concat-growth-chain | evaluate | 1,887 | 1,928 | +2.2% | 20 | 20 | 10,000 | 10,000 | 3,512 | 3,404 |
| concat-growth-chain | parse-evaluate | 2,272 | 2,276 | +0.2% | 20 | 20 | 17,000 | 17,000 | 3,520 | 3,404 |
| concat-owned-left | evaluate | 1,092 | 1,091 | -0.1% | 24 | 24 | 8,000 | 8,000 | 3,520 | 3,416 |
| concat-owned-left | parse-evaluate | 1,344 | 1,310 | -2.5% | 24 | 24 | 13,000 | 13,000 | 3,588 | 3,420 |
| concat-owned-right | evaluate | 1,080 | 1,106 | +2.4% | 24 | 24 | 8,000 | 8,000 | 3,512 | 3,436 |
| concat-owned-right | parse-evaluate | 1,392 | 1,387 | -0.4% | 24 | 24 | 13,000 | 13,000 | 3,608 | 3,392 |
| database-control-dstdev | evaluate | 4,130 | 3,910 | -5.3% | 30 | 30 | 20 | 20 | 3,732 | 3,764 |
| database-control-dstdev | parse-evaluate | 4,650 | 4,550 | -2.2% | 30 | 30 | 29 | 29 | 3,712 | 3,768 |
| database-control-dsum | evaluate | 4,160 | 3,990 | -4.1% | 28 | 28 | 20 | 20 | 3,684 | 3,724 |
| database-control-dsum | parse-evaluate | 4,700 | 4,570 | -2.8% | 28 | 28 | 29 | 29 | 3,736 | 3,792 |
| database-control-dvar | evaluate | 4,070 | 3,920 | -3.7% | 28 | 28 | 20 | 20 | 3,752 | 3,768 |
| database-control-dvar | parse-evaluate | 4,660 | 4,540 | -2.6% | 28 | 28 | 29 | 29 | 3,732 | 3,752 |
| literal-aggregate-4x1-sum | evaluate | 1,850 | 1,850 | +0.0% | 15 | 15 | 10 | 10 | 3,764 | 3,764 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,600 | 2,520 | -3.1% | 15 | 15 | 23 | 23 | 3,768 | 3,840 |
| reference-aggregate-64x4-sum | evaluate | 10,295 | 10,202 | -0.9% | 16 | 16 | 32 | 32 | 3,728 | 3,752 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,757 | 10,650 | -1.0% | 16 | 16 | 60 | 60 | 3,728 | 3,852 |
| reference-array-16x4-arithmetic | evaluate | 9,516 | 9,429 | -0.9% | 16 | 16 | 800 | 800 | 3,704 | 3,652 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,859 | 9,846 | -0.1% | 16 | 16 | 1,280 | 1,280 | 3,696 | 3,684 |
| reference-conditional-256x4-sumifs | evaluate | 79,640 | 78,855 | -1.0% | 52 | 52 | 42 | 42 | 3,704 | 3,708 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,585 | 79,670 | -1.1% | 52 | 52 | 68 | 68 | 3,744 | 3,728 |
| reference-control-average | evaluate | 10,355 | 10,285 | -0.7% | 20 | 20 | 32 | 32 | 3,744 | 3,768 |
| reference-control-average | parse-evaluate | 10,717 | 10,657 | -0.6% | 20 | 20 | 60 | 60 | 3,752 | 3,748 |
| reference-control-counta | evaluate | 8,687 | 8,620 | -0.8% | 19 | 19 | 32 | 32 | 3,744 | 3,696 |
| reference-control-counta | parse-evaluate | 9,157 | 9,057 | -1.1% | 19 | 19 | 60 | 60 | 3,748 | 3,724 |
| representative-median | evaluate | 1,624 | 1,605 | -1.2% | 16 | 16 | 9,000 | 9,000 | 3,632 | 3,564 |
| representative-median | parse-evaluate | 1,946 | 1,970 | +1.2% | 16 | 16 | 14,000 | 14,000 | 3,668 | 3,580 |
| representative-percentrank | evaluate | 1,012 | 999 | -1.3% | 19 | 19 | 7,000 | 7,000 | 3,684 | 3,504 |
| representative-percentrank | parse-evaluate | 1,251 | 1,258 | +0.6% | 19 | 19 | 11,000 | 11,000 | 3,656 | 3,604 |
| representative-rank | evaluate | 942 | 933 | -1.0% | 12 | 12 | 7,000 | 7,000 | 3,584 | 3,504 |
| representative-rank | parse-evaluate | 1,172 | 1,161 | -0.9% | 12 | 12 | 11,000 | 11,000 | 3,648 | 3,512 |
| scalar-aggregate-sum | evaluate | 550 | 566 | +2.9% | 10 | 10 | 4,000 | 4,000 | 3,576 | 3,532 |
| scalar-aggregate-sum | parse-evaluate | 727 | 751 | +3.3% | 10 | 10 | 8,000 | 8,000 | 3,648 | 3,600 |
| scalar-control-arithmetic | evaluate | 583 | 587 | +0.7% | 10 | 10 | 5,000 | 5,000 | 3,588 | 3,404 |
| scalar-control-arithmetic | parse-evaluate | 717 | 717 | +0.0% | 10 | 10 | 8,000 | 8,000 | 3,492 | 3,388 |
| scalar-control-average | evaluate | 1,035 | 1,024 | -1.1% | 23 | 23 | 6,000 | 6,000 | 3,652 | 3,632 |
| scalar-control-average | parse-evaluate | 1,252 | 1,252 | +0.0% | 23 | 23 | 10,000 | 10,000 | 3,648 | 3,540 |
| scalar-control-counta | evaluate | 828 | 819 | -1.1% | 24 | 24 | 6,000 | 6,000 | 3,656 | 3,500 |
| scalar-control-counta | parse-evaluate | 1,058 | 1,065 | +0.7% | 24 | 24 | 10,000 | 10,000 | 3,576 | 3,556 |
| scalar-control-imsum | evaluate | 1,406 | 1,406 | +0.0% | 41 | 41 | 8,000 | 8,000 | 3,640 | 3,512 |
| scalar-control-imsum | parse-evaluate | 1,976 | 1,944 | -1.6% | 41 | 41 | 17,000 | 17,000 | 3,704 | 3,416 |
| scalar-control-sin | evaluate | 466 | 468 | +0.4% | 10 | 10 | 4,000 | 4,000 | 3,864 | 3,820 |
| scalar-control-sin | parse-evaluate | 643 | 649 | +0.9% | 10 | 10 | 8,000 | 8,000 | 3,844 | 3,820 |
| scalar-control-stdev | evaluate | 861 | 853 | -0.9% | 21 | 21 | 6,000 | 6,000 | 3,584 | 3,488 |
| scalar-control-stdev | parse-evaluate | 1,062 | 1,066 | +0.4% | 21 | 21 | 10,000 | 10,000 | 3,684 | 3,568 |
| scalar-control-var | evaluate | 851 | 845 | -0.7% | 19 | 19 | 6,000 | 6,000 | 3,640 | 3,428 |
| scalar-control-var | parse-evaluate | 1,044 | 1,051 | +0.7% | 19 | 19 | 10,000 | 10,000 | 3,516 | 3,568 |

## Byte-position and projected-reducer workloads

The candidate matrix covers all seven byte functions over scalar and large Unicode/ASCII inputs, borrowed 64-cell references, matrix broadcast, typed refusal, sticky cancellation, resource refusal, worst-case searches, REPLACEB growth, and projected AVERAGE/SUM reducers. Input/output bytes and the reviewed UTF-8/domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| cancellation-byte-lenb | evaluate | 0 | 0 | 0 | 387 | 3 | 0 | 10 | 7,224 | 7,224 | 0 | 3,896 |
| cancellation-byte-lenb | parse-evaluate | 0 | 0 | 0 | 775 | 3 | 0 | 38 | 12,676 | 8,555 | 0 | 3,712 |
| cancellation-byte-midb | evaluate | 0 | 0 | 0 | 480 | 4 | 0 | 12 | 8,400 | 7,952 | 0 | 3,708 |
| cancellation-byte-midb | parse-evaluate | 0 | 0 | 0 | 897 | 4 | 0 | 40 | 13,868 | 9,287 | 0 | 3,716 |
| cancellation-byte-replaceb | evaluate | 0 | 0 | 0 | 522 | 6 | 0 | 13 | 8,936 | 8,360 | 0 | 3,708 |
| cancellation-byte-replaceb | parse-evaluate | 0 | 0 | 0 | 1,012 | 6 | 0 | 45 | 17,516 | 10,089 | 0 | 3,772 |
| cancellation-byte-searchb | evaluate | 0 | 0 | 0 | 492 | 5 | 0 | 12 | 8,400 | 7,952 | 0 | 3,696 |
| cancellation-byte-searchb | parse-evaluate | 0 | 0 | 0 | 937 | 5 | 0 | 40 | 13,888 | 9,292 | 0 | 3,760 |
| large-ascii-byte-findb | evaluate | 11,520 | 0 | 11,520 | 4,039 | 23,063 | 0 | 6,000 | 848,000 | 624 | 0 | 3,668 |
| large-ascii-byte-findb | parse-evaluate | 11,520 | 0 | 11,520 | 7,581 | 23,063 | 0 | 10,000 | 12,834,000 | 12,578 | 0 | 3,628 |
| large-ascii-byte-leftb | evaluate | 11,520 | 7,000 | 11,527 | 3,827 | 23,054 | 0 | 5,000 | 496,000 | 448 | 0 | 3,660 |
| large-ascii-byte-leftb | parse-evaluate | 11,520 | 7,000 | 11,527 | 7,291 | 23,054 | 0 | 9,000 | 12,476,000 | 12,396 | 0 | 3,620 |
| large-ascii-byte-lenb | evaluate | 11,520 | 0 | 11,520 | 3,632 | 23,049 | 0 | 4,000 | 400,000 | 400 | 0 | 3,624 |
| large-ascii-byte-lenb | parse-evaluate | 11,520 | 0 | 11,520 | 7,063 | 23,049 | 0 | 8,000 | 12,377,000 | 12,345 | 0 | 3,636 |
| large-ascii-byte-midb | evaluate | 11,520 | 5,000 | 11,525 | 3,982 | 23,057 | 0 | 6,000 | 848,000 | 624 | 0 | 3,664 |
| large-ascii-byte-midb | parse-evaluate | 11,520 | 5,000 | 11,525 | 7,473 | 23,057 | 0 | 10,000 | 12,829,000 | 12,573 | 0 | 3,636 |
| large-ascii-byte-replaceb | evaluate | 11,520 | 11,521,000 | 23,041 | 14,670 | 34,591 | 0 | 8,000 | 12,561,000 | 12,241 | 11,521 | 3,660 |
| large-ascii-byte-replaceb | parse-evaluate | 11,520 | 11,521,000 | 23,041 | 18,190 | 34,591 | 0 | 13,000 | 25,320,000 | 24,584 | 11,521 | 3,576 |
| large-ascii-byte-rightb | evaluate | 11,520 | 7,000 | 11,527 | 3,820 | 23,055 | 0 | 5,000 | 496,000 | 448 | 0 | 3,608 |
| large-ascii-byte-rightb | parse-evaluate | 11,520 | 7,000 | 11,527 | 7,294 | 23,055 | 0 | 9,000 | 12,477,000 | 12,397 | 0 | 3,580 |
| large-ascii-byte-searchb | evaluate | 11,520 | 0 | 11,520 | 4,652 | 23,092 | 0 | 9,000 | 932,000 | 708 | 0 | 3,668 |
| large-ascii-byte-searchb | parse-evaluate | 11,520 | 0 | 11,520 | 8,143 | 23,092 | 0 | 13,000 | 12,920,000 | 12,664 | 0 | 3,704 |
| large-unicode-byte-findb | evaluate | 2,880 | 0 | 2,880 | 1,711 | 5,781 | 0 | 6,000 | 848,000 | 624 | 0 | 3,660 |
| large-unicode-byte-findb | parse-evaluate | 2,880 | 0 | 2,880 | 2,842 | 5,781 | 0 | 10,000 | 4,193,000 | 3,937 | 0 | 3,628 |
| large-unicode-byte-leftb | evaluate | 2,880 | 6,000 | 2,886 | 1,488 | 5,774 | 0 | 5,000 | 496,000 | 448 | 0 | 3,640 |
| large-unicode-byte-leftb | parse-evaluate | 2,880 | 6,000 | 2,886 | 2,547 | 5,774 | 0 | 9,000 | 3,836,000 | 3,756 | 0 | 3,560 |
| large-unicode-byte-lenb | evaluate | 2,880 | 0 | 2,880 | 1,292 | 5,769 | 0 | 4,000 | 400,000 | 400 | 0 | 3,644 |
| large-unicode-byte-lenb | parse-evaluate | 2,880 | 0 | 2,880 | 2,320 | 5,769 | 0 | 8,000 | 3,737,000 | 3,705 | 0 | 3,640 |
| large-unicode-byte-midb | evaluate | 2,880 | 4,000 | 2,884 | 1,636 | 5,777 | 0 | 6,000 | 848,000 | 624 | 0 | 3,584 |
| large-unicode-byte-midb | parse-evaluate | 2,880 | 4,000 | 2,884 | 2,730 | 5,777 | 0 | 10,000 | 4,189,000 | 3,933 | 0 | 3,660 |
| large-unicode-byte-replaceb | evaluate | 2,880 | 2,881,000 | 5,761 | 4,608 | 8,671 | 0 | 8,000 | 3,921,000 | 3,601 | 2,881 | 3,660 |
| large-unicode-byte-replaceb | parse-evaluate | 2,880 | 2,881,000 | 5,761 | 5,735 | 8,671 | 0 | 13,000 | 8,040,000 | 7,304 | 2,881 | 3,576 |
| large-unicode-byte-rightb | evaluate | 2,880 | 7,000 | 2,887 | 1,505 | 5,775 | 0 | 5,000 | 496,000 | 448 | 0 | 3,684 |
| large-unicode-byte-rightb | parse-evaluate | 2,880 | 7,000 | 2,887 | 2,560 | 5,775 | 0 | 9,000 | 3,837,000 | 3,757 | 0 | 3,508 |
| large-unicode-byte-searchb | evaluate | 2,880 | 0 | 2,880 | 2,045 | 5,792 | 0 | 9,000 | 876,000 | 652 | 0 | 3,640 |
| large-unicode-byte-searchb | parse-evaluate | 2,880 | 0 | 2,880 | 3,163 | 5,792 | 0 | 13,000 | 4,223,000 | 3,967 | 0 | 3,640 |
| matrix-broadcast-byte-findb | evaluate | 833 | 0 | 833 | 25,365 | 1,489 | 64 | 1,120 | 714,240 | 8,304 | 5,632 | 3,772 |
| matrix-broadcast-byte-findb | parse-evaluate | 833 | 0 | 833 | 25,686 | 1,489 | 64 | 1,680 | 823,840 | 9,642 | 5,632 | 3,860 |
| matrix-broadcast-byte-leftb | evaluate | 832 | 9,120 | 946 | 21,811 | 1,359 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,904 |
| matrix-broadcast-byte-leftb | parse-evaluate | 832 | 9,120 | 946 | 22,290 | 1,359 | 128 | 1,920 | 744,400 | 9,209 | 5,632 | 3,796 |
| matrix-broadcast-byte-lenb | evaluate | 832 | 0 | 832 | 14,877 | 1,162 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,872 |
| matrix-broadcast-byte-lenb | parse-evaluate | 832 | 0 | 832 | 15,336 | 1,162 | 64 | 1,440 | 701,040 | 8,731 | 5,632 | 3,904 |
| matrix-broadcast-byte-midb | evaluate | 832 | 8,960 | 944 | 24,884 | 1,489 | 128 | 1,360 | 746,240 | 8,704 | 5,632 | 3,804 |
| matrix-broadcast-byte-midb | parse-evaluate | 832 | 8,960 | 944 | 25,525 | 1,489 | 128 | 2,080 | 856,480 | 10,050 | 5,632 | 3,908 |
| matrix-broadcast-byte-replaceb | evaluate | 832 | 62,720 | 1,616 | 39,633 | 2,472 | 128 | 6,560 | 851,840 | 9,896 | 6,416 | 3,856 |
| matrix-broadcast-byte-replaceb | parse-evaluate | 832 | 62,720 | 1,616 | 40,733 | 2,472 | 128 | 7,360 | 1,024,160 | 11,634 | 6,416 | 3,832 |
| matrix-broadcast-byte-rightb | evaluate | 832 | 8,400 | 937 | 21,698 | 1,360 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,856 |
| matrix-broadcast-byte-rightb | parse-evaluate | 832 | 8,400 | 937 | 22,306 | 1,360 | 128 | 1,920 | 744,480 | 9,210 | 5,632 | 3,784 |
| matrix-broadcast-byte-searchb | evaluate | 833 | 0 | 833 | 45,719 | 2,043 | 64 | 16,480 | 857,600 | 8,332 | 5,632 | 3,860 |
| matrix-broadcast-byte-searchb | parse-evaluate | 833 | 0 | 833 | 45,608 | 2,043 | 64 | 17,040 | 967,360 | 9,672 | 5,632 | 3,932 |
| projected-statistical-average-len | evaluate | 46 | 0 | 46 | 7,460 | 114 | 2 | 38 | 5,592 | 2,760 | 176 | 3,792 |
| projected-statistical-average-len | parse-evaluate | 46 | 0 | 46 | 8,490 | 114 | 2 | 54 | 9,640 | 5,368 | 176 | 3,820 |
| projected-statistical-average-lenb | evaluate | 47 | 0 | 47 | 7,440 | 116 | 2 | 38 | 5,592 | 2,760 | 176 | 3,760 |
| projected-statistical-average-lenb | parse-evaluate | 47 | 0 | 47 | 8,500 | 116 | 2 | 54 | 9,641 | 5,369 | 176 | 3,840 |
| projected-statistical-sum-lenb | evaluate | 43 | 0 | 43 | 8,180 | 152 | 4 | 40 | 7,880 | 5,304 | 176 | 3,920 |
| projected-statistical-sum-lenb | parse-evaluate | 43 | 0 | 43 | 9,610 | 152 | 4 | 56 | 11,925 | 7,909 | 176 | 3,776 |
| reference-64-byte-findb | evaluate | 833 | 0 | 833 | 25,170 | 1,489 | 64 | 1,120 | 714,240 | 8,304 | 5,632 | 3,768 |
| reference-64-byte-findb | parse-evaluate | 833 | 0 | 833 | 25,727 | 1,489 | 64 | 1,680 | 823,840 | 9,642 | 5,632 | 3,856 |
| reference-64-byte-leftb | evaluate | 832 | 33,920 | 1,256 | 19,928 | 1,294 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,864 |
| reference-64-byte-leftb | parse-evaluate | 832 | 33,920 | 1,256 | 20,383 | 1,294 | 64 | 1,520 | 711,520 | 8,798 | 5,632 | 3,892 |
| reference-64-byte-lenb | evaluate | 832 | 0 | 832 | 14,783 | 1,162 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,848 |
| reference-64-byte-lenb | parse-evaluate | 832 | 0 | 832 | 15,283 | 1,162 | 64 | 1,440 | 701,040 | 8,731 | 5,632 | 3,792 |
| reference-64-byte-midb | evaluate | 832 | 22,400 | 1,112 | 22,654 | 1,424 | 64 | 1,120 | 714,240 | 8,304 | 5,632 | 3,816 |
| reference-64-byte-midb | parse-evaluate | 832 | 22,400 | 1,112 | 23,139 | 1,424 | 64 | 1,680 | 823,600 | 9,639 | 5,632 | 3,776 |
| reference-64-byte-replaceb | evaluate | 832 | 72,960 | 1,744 | 37,669 | 2,665 | 64 | 6,320 | 830,080 | 9,624 | 6,544 | 3,928 |
| reference-64-byte-replaceb | parse-evaluate | 832 | 72,960 | 1,744 | 38,307 | 2,665 | 64 | 6,960 | 1,001,680 | 11,353 | 6,544 | 3,812 |
| reference-64-byte-rightb | evaluate | 832 | 35,840 | 1,280 | 19,415 | 1,295 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,764 |
| reference-64-byte-rightb | parse-evaluate | 832 | 35,840 | 1,280 | 19,792 | 1,295 | 64 | 1,520 | 711,600 | 8,799 | 5,632 | 3,856 |
| reference-64-byte-searchb | evaluate | 833 | 0 | 833 | 45,162 | 2,043 | 64 | 16,480 | 857,600 | 8,332 | 5,632 | 3,896 |
| reference-64-byte-searchb | parse-evaluate | 833 | 0 | 833 | 45,409 | 2,043 | 64 | 17,040 | 967,360 | 9,672 | 5,632 | 3,764 |
| refusal-byte-findb | evaluate | 0 | 0 | 0 | 2,309 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,820 |
| refusal-byte-findb | parse-evaluate | 0 | 0 | 0 | 2,872 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,716 |
| refusal-byte-leftb | evaluate | 0 | 0 | 0 | 2,279 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,768 |
| refusal-byte-leftb | parse-evaluate | 0 | 0 | 0 | 2,854 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,876 |
| refusal-byte-lenb | evaluate | 0 | 0 | 0 | 2,240 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,908 |
| refusal-byte-lenb | parse-evaluate | 0 | 0 | 0 | 2,831 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,768 |
| refusal-byte-midb | evaluate | 0 | 0 | 0 | 2,310 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,744 |
| refusal-byte-midb | parse-evaluate | 0 | 0 | 0 | 2,836 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,764 |
| refusal-byte-replaceb | evaluate | 0 | 0 | 0 | 2,275 | 22 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,696 |
| refusal-byte-replaceb | parse-evaluate | 0 | 0 | 0 | 2,829 | 22 | 0 | 1,920 | 324,800 | 2,924 | 0 | 3,772 |
| refusal-byte-rightb | evaluate | 0 | 0 | 0 | 2,258 | 20 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,824 |
| refusal-byte-rightb | parse-evaluate | 0 | 0 | 0 | 2,828 | 20 | 0 | 1,920 | 324,640 | 2,922 | 0 | 3,924 |
| refusal-byte-searchb | evaluate | 0 | 0 | 0 | 2,286 | 21 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,764 |
| refusal-byte-searchb | parse-evaluate | 0 | 0 | 0 | 2,900 | 21 | 0 | 1,920 | 324,720 | 2,923 | 0 | 3,708 |
| replaceb-growth | evaluate | 15,616 | 15,616,000 | 31,232 | 19,446 | 46,872 | 0 | 8,000 | 16,656,000 | 16,336 | 15,616 | 3,784 |
| replaceb-growth | parse-evaluate | 15,616 | 15,616,000 | 31,232 | 24,117 | 46,872 | 0 | 13,000 | 33,508,000 | 32,772 | 15,616 | 3,956 |
| resource-byte-lenb | evaluate | 0 | 0 | 0 | 381 | 7 | 0 | 320 | 23,680 | 296 | 0 | 3,716 |
| resource-byte-lenb | parse-evaluate | 0 | 0 | 0 | 754 | 7 | 0 | 880 | 132,720 | 1,627 | 0 | 3,748 |
| resource-byte-midb | evaluate | 0 | 0 | 0 | 517 | 9 | 0 | 400 | 33,920 | 360 | 0 | 3,708 |
| resource-byte-midb | parse-evaluate | 0 | 0 | 0 | 922 | 9 | 0 | 960 | 143,280 | 1,695 | 0 | 3,704 |
| resource-byte-replaceb | evaluate | 0 | 0 | 0 | 637 | 14 | 0 | 480 | 54,400 | 488 | 0 | 3,752 |
| resource-byte-replaceb | parse-evaluate | 0 | 0 | 0 | 1,135 | 14 | 0 | 1,120 | 226,000 | 2,217 | 0 | 3,716 |
| resource-byte-searchb | evaluate | 0 | 0 | 0 | 668 | 14 | 0 | 480 | 64,640 | 744 | 0 | 3,752 |
| resource-byte-searchb | parse-evaluate | 0 | 0 | 0 | 1,098 | 14 | 0 | 1,040 | 174,400 | 2,084 | 0 | 3,720 |
| search-worstcase-byte-findb | evaluate | 4,229 | 0 | 4,229 | 6,799 | 8,487 | 0 | 6,000 | 848,000 | 624 | 0 | 3,700 |
| search-worstcase-byte-findb | parse-evaluate | 4,229 | 0 | 4,229 | 8,334 | 8,487 | 0 | 10,000 | 5,546,000 | 5,290 | 0 | 3,632 |
| search-worstcase-byte-searchb | evaluate | 4,229 | 0 | 4,229 | 97,385 | 17,084 | 0 | 9,000 | 4,572,000 | 4,348 | 0 | 3,688 |
| search-worstcase-byte-searchb | parse-evaluate | 4,229 | 0 | 4,229 | 98,480 | 17,084 | 0 | 13,000 | 9,272,000 | 9,016 | 0 | 3,668 |
| tiny-byte-findb | evaluate | 10 | 0 | 10 | 951 | 43 | 0 | 6,000 | 848,000 | 624 | 0 | 3,580 |
| tiny-byte-findb | parse-evaluate | 10 | 0 | 10 | 1,188 | 43 | 0 | 10,000 | 1,324,000 | 1,068 | 0 | 3,640 |
| tiny-byte-leftb | evaluate | 10 | 6,000 | 16 | 714 | 34 | 0 | 5,000 | 496,000 | 448 | 0 | 3,572 |
| tiny-byte-leftb | parse-evaluate | 10 | 6,000 | 16 | 932 | 34 | 0 | 9,000 | 966,000 | 886 | 0 | 3,608 |
| tiny-byte-lenb | evaluate | 10 | 0 | 10 | 511 | 29 | 0 | 4,000 | 400,000 | 400 | 0 | 3,624 |
| tiny-byte-lenb | parse-evaluate | 10 | 0 | 10 | 711 | 29 | 0 | 8,000 | 867,000 | 835 | 0 | 3,672 |
| tiny-byte-midb | evaluate | 10 | 5,000 | 15 | 899 | 37 | 0 | 6,000 | 848,000 | 624 | 0 | 3,660 |
| tiny-byte-midb | parse-evaluate | 10 | 5,000 | 15 | 1,117 | 37 | 0 | 10,000 | 1,319,000 | 1,063 | 0 | 3,572 |
| tiny-byte-replaceb | evaluate | 10 | 11,000 | 21 | 1,254 | 61 | 0 | 8,000 | 1,051,000 | 731 | 11 | 3,672 |
| tiny-byte-replaceb | parse-evaluate | 10 | 11,000 | 21 | 1,491 | 61 | 0 | 13,000 | 2,300,000 | 1,564 | 11 | 3,660 |
| tiny-byte-rightb | evaluate | 10 | 7,000 | 17 | 712 | 35 | 0 | 5,000 | 496,000 | 448 | 0 | 3,580 |
| tiny-byte-rightb | parse-evaluate | 10 | 7,000 | 17 | 924 | 35 | 0 | 9,000 | 967,000 | 887 | 0 | 3,628 |
| tiny-byte-searchb | evaluate | 10 | 0 | 10 | 1,247 | 50 | 0 | 9,000 | 876,000 | 652 | 0 | 3,656 |
| tiny-byte-searchb | parse-evaluate | 10 | 0 | 10 | 1,507 | 50 | 0 | 13,000 | 1,354,000 | 1,098 | 0 | 3,680 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
