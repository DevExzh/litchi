# ODS text-functions evaluator performance profile

The baseline is committed `8f09231e36982248eface4d599432143a67f6e49`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 47,627 | 47,297 | -0.7% | 2,311 | 2,311 | 88 | 88 | 3,736 | 4,044 |
| array-control-16x16-arithmetic | parse-evaluate | 57,475 | 57,812 | +0.6% | 2,311 | 2,311 | 352 | 352 | 3,688 | 3,964 |
| array-control-16x16-sin | evaluate | 126,828 | 127,763 | +0.7% | 2,311 | 2,311 | 2,144 | 2,144 | 4,044 | 4,264 |
| array-control-16x16-sin | parse-evaluate | 138,118 | 138,633 | +0.4% | 2,311 | 2,311 | 2,412 | 2,412 | 4,032 | 4,292 |
| array-control-4x4-arithmetic | evaluate | 4,505 | 4,546 | +0.9% | 151 | 151 | 1,120 | 1,120 | 3,468 | 3,752 |
| array-control-4x4-arithmetic | parse-evaluate | 5,318 | 5,383 | +1.2% | 151 | 151 | 2,240 | 2,240 | 3,472 | 3,704 |
| array-control-4x4-sin | evaluate | 9,823 | 9,827 | +0.0% | 151 | 151 | 3,840 | 3,840 | 3,788 | 4,068 |
| array-control-4x4-sin | parse-evaluate | 10,630 | 10,798 | +1.6% | 151 | 151 | 5,040 | 5,040 | 3,784 | 4,028 |
| concat-borrowed-literals | evaluate | 733 | 732 | -0.1% | 21 | 21 | 6,000 | 6,000 | 3,432 | 3,488 |
| concat-borrowed-literals | parse-evaluate | 875 | 870 | -0.6% | 21 | 21 | 9,000 | 9,000 | 3,424 | 3,664 |
| concat-growth-chain | evaluate | 1,923 | 1,898 | -1.3% | 20 | 20 | 10,000 | 10,000 | 3,436 | 3,668 |
| concat-growth-chain | parse-evaluate | 2,333 | 2,303 | -1.3% | 20 | 20 | 17,000 | 17,000 | 3,448 | 3,628 |
| concat-owned-left | evaluate | 1,087 | 1,067 | -1.8% | 24 | 24 | 8,000 | 8,000 | 3,428 | 3,512 |
| concat-owned-left | parse-evaluate | 1,338 | 1,317 | -1.6% | 24 | 24 | 13,000 | 13,000 | 3,432 | 3,680 |
| concat-owned-right | evaluate | 1,086 | 1,083 | -0.3% | 24 | 24 | 8,000 | 8,000 | 3,428 | 3,576 |
| concat-owned-right | parse-evaluate | 1,350 | 1,331 | -1.4% | 24 | 24 | 13,000 | 13,000 | 3,376 | 3,576 |
| database-control-dstdev | evaluate | 3,910 | 3,930 | +0.5% | 30 | 30 | 20 | 20 | 3,484 | 3,752 |
| database-control-dstdev | parse-evaluate | 4,570 | 4,670 | +2.2% | 30 | 30 | 29 | 29 | 3,492 | 3,712 |
| database-control-dsum | evaluate | 3,950 | 3,940 | -0.3% | 28 | 28 | 20 | 20 | 3,480 | 3,828 |
| database-control-dsum | parse-evaluate | 4,630 | 4,690 | +1.3% | 28 | 28 | 29 | 29 | 3,496 | 3,716 |
| database-control-dvar | evaluate | 3,940 | 3,880 | -1.5% | 28 | 28 | 20 | 20 | 3,496 | 3,752 |
| database-control-dvar | parse-evaluate | 4,501 | 4,550 | +1.1% | 28 | 28 | 29 | 29 | 3,516 | 3,736 |
| literal-aggregate-4x1-sum | evaluate | 1,810 | 1,840 | +1.7% | 15 | 15 | 10 | 10 | 3,480 | 3,720 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,590 | 2,530 | -2.3% | 15 | 15 | 23 | 23 | 3,496 | 3,800 |
| reference-aggregate-64x4-sum | evaluate | 10,285 | 10,225 | -0.6% | 16 | 16 | 32 | 32 | 3,476 | 3,700 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,710 | 10,627 | -0.8% | 16 | 16 | 60 | 60 | 3,492 | 3,764 |
| reference-array-16x4-arithmetic | evaluate | 9,545 | 9,592 | +0.5% | 16 | 16 | 800 | 800 | 3,468 | 3,728 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,796 | 9,885 | +0.9% | 16 | 16 | 1,280 | 1,280 | 3,464 | 3,716 |
| reference-conditional-256x4-sumifs | evaluate | 80,785 | 79,110 | -2.1% | 52 | 52 | 42 | 42 | 3,480 | 3,712 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,985 | 80,520 | -0.6% | 52 | 52 | 68 | 68 | 3,484 | 3,728 |
| reference-control-average | evaluate | 10,322 | 10,247 | -0.7% | 20 | 20 | 32 | 32 | 3,480 | 3,724 |
| reference-control-average | parse-evaluate | 10,675 | 10,732 | +0.5% | 20 | 20 | 60 | 60 | 3,500 | 3,812 |
| reference-control-counta | evaluate | 8,677 | 8,657 | -0.2% | 19 | 19 | 32 | 32 | 3,484 | 3,832 |
| reference-control-counta | parse-evaluate | 9,125 | 9,130 | +0.1% | 19 | 19 | 60 | 60 | 3,472 | 3,752 |
| representative-median | evaluate | 1,574 | 1,585 | +0.7% | 16 | 16 | 9,000 | 9,000 | 3,460 | 3,632 |
| representative-median | parse-evaluate | 1,862 | 1,917 | +3.0% | 16 | 16 | 14,000 | 14,000 | 3,420 | 3,592 |
| representative-percentrank | evaluate | 970 | 983 | +1.3% | 19 | 19 | 7,000 | 7,000 | 3,496 | 3,536 |
| representative-percentrank | parse-evaluate | 1,201 | 1,246 | +3.7% | 19 | 19 | 11,000 | 11,000 | 3,484 | 3,644 |
| representative-rank | evaluate | 927 | 926 | -0.1% | 12 | 12 | 7,000 | 7,000 | 3,448 | 3,604 |
| representative-rank | parse-evaluate | 1,149 | 1,157 | +0.7% | 12 | 12 | 11,000 | 11,000 | 3,472 | 3,592 |
| scalar-aggregate-sum | evaluate | 561 | 546 | -2.7% | 10 | 10 | 4,000 | 4,000 | 3,468 | 3,684 |
| scalar-aggregate-sum | parse-evaluate | 759 | 729 | -4.0% | 10 | 10 | 8,000 | 8,000 | 3,436 | 3,684 |
| scalar-control-arithmetic | evaluate | 570 | 573 | +0.5% | 10 | 10 | 5,000 | 5,000 | 3,436 | 3,692 |
| scalar-control-arithmetic | parse-evaluate | 719 | 714 | -0.7% | 10 | 10 | 8,000 | 8,000 | 3,456 | 3,648 |
| scalar-control-average | evaluate | 1,029 | 1,031 | +0.2% | 23 | 23 | 6,000 | 6,000 | 3,468 | 3,604 |
| scalar-control-average | parse-evaluate | 1,257 | 1,254 | -0.2% | 23 | 23 | 10,000 | 10,000 | 3,440 | 3,532 |
| scalar-control-counta | evaluate | 833 | 818 | -1.8% | 24 | 24 | 6,000 | 6,000 | 3,468 | 3,664 |
| scalar-control-counta | parse-evaluate | 1,067 | 1,056 | -1.0% | 24 | 24 | 10,000 | 10,000 | 3,488 | 3,632 |
| scalar-control-imsum | evaluate | 1,404 | 1,395 | -0.6% | 41 | 41 | 8,000 | 8,000 | 3,480 | 3,648 |
| scalar-control-imsum | parse-evaluate | 1,978 | 1,956 | -1.1% | 41 | 41 | 17,000 | 17,000 | 3,444 | 3,664 |
| scalar-control-sin | evaluate | 468 | 460 | -1.7% | 10 | 10 | 4,000 | 4,000 | 3,732 | 3,944 |
| scalar-control-sin | parse-evaluate | 652 | 641 | -1.7% | 10 | 10 | 8,000 | 8,000 | 3,744 | 3,940 |
| scalar-control-stdev | evaluate | 867 | 855 | -1.4% | 21 | 21 | 6,000 | 6,000 | 3,456 | 3,592 |
| scalar-control-stdev | parse-evaluate | 1,075 | 1,070 | -0.5% | 21 | 21 | 10,000 | 10,000 | 3,448 | 3,680 |
| scalar-control-var | evaluate | 848 | 844 | -0.5% | 19 | 19 | 6,000 | 6,000 | 3,448 | 3,668 |
| scalar-control-var | parse-evaluate | 1,058 | 1,049 | -0.9% | 19 | 19 | 10,000 | 10,000 | 3,468 | 3,532 |

## Text-function workloads

The candidate matrix covers scalar and large Unicode inputs, borrowed 64-cell references, mapped matrix outputs, typed refusal, sticky cancellation, resource refusal, worst-case searches, REPT growth, and ASC/JIS width conversion. Text input/output bytes and the reviewed Unicode/domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| asc-jis-expansion-asc | evaluate | 16 | 9,000 | 25 | 710 | 45 | 0 | 5,000 | 409,000 | 409 | 9 | 3,524 |
| asc-jis-expansion-asc | parse-evaluate | 16 | 9,000 | 25 | 908 | 45 | 0 | 9,000 | 877,000 | 845 | 9 | 3,628 |
| asc-jis-expansion-jis | evaluate | 16 | 12,000 | 28 | 704 | 42 | 0 | 5,000 | 412,000 | 412 | 12 | 3,644 |
| asc-jis-expansion-jis | parse-evaluate | 16 | 12,000 | 28 | 903 | 42 | 0 | 9,000 | 877,000 | 845 | 12 | 3,680 |
| cancellation-text-concatenate | evaluate | 0 | 0 | 0 | 430 | 6 | 0 | 11 | 7,352 | 7,288 | 0 | 3,732 |
| cancellation-text-concatenate | parse-evaluate | 0 | 0 | 0 | 870 | 6 | 0 | 39 | 12,864 | 8,634 | 0 | 3,720 |
| cancellation-text-len | evaluate | 0 | 0 | 0 | 380 | 2 | 0 | 10 | 7,224 | 7,224 | 0 | 3,724 |
| cancellation-text-len | parse-evaluate | 0 | 0 | 0 | 782 | 2 | 0 | 38 | 12,672 | 8,554 | 0 | 3,716 |
| cancellation-text-search | evaluate | 0 | 0 | 0 | 445 | 4 | 0 | 11 | 7,352 | 7,288 | 0 | 3,700 |
| cancellation-text-search | parse-evaluate | 0 | 0 | 0 | 867 | 4 | 0 | 39 | 12,828 | 8,625 | 0 | 3,696 |
| cancellation-text-substitute | evaluate | 0 | 0 | 0 | 475 | 6 | 0 | 12 | 8,400 | 7,952 | 0 | 3,744 |
| cancellation-text-substitute | parse-evaluate | 0 | 0 | 0 | 942 | 6 | 0 | 40 | 13,908 | 9,297 | 0 | 3,764 |
| large-unicode-text-asc | evaluate | 2,880 | 2,880,000 | 5,760 | 11,255 | 5,768 | 0 | 4,000 | 400,000 | 400 | 0 | 3,640 |
| large-unicode-text-asc | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 12,351 | 5,768 | 0 | 8,000 | 3,736,000 | 3,704 | 0 | 3,632 |
| large-unicode-text-char | evaluate | 2,880 | 1,000 | 2,881 | 626 | 12 | 0 | 5,000 | 401,000 | 401 | 1 | 3,664 |
| large-unicode-text-char | parse-evaluate | 2,880 | 1,000 | 2,881 | 821 | 12 | 0 | 9,000 | 858,000 | 826 | 1 | 3,688 |
| large-unicode-text-clean | evaluate | 2,880 | 2,880,000 | 5,760 | 12,150 | 5,770 | 0 | 4,000 | 400,000 | 400 | 0 | 3,668 |
| large-unicode-text-clean | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 13,194 | 5,770 | 0 | 8,000 | 3,738,000 | 3,706 | 0 | 3,680 |
| large-unicode-text-code | evaluate | 2,880 | 0 | 2,880 | 1,288 | 5,769 | 0 | 4,000 | 400,000 | 400 | 0 | 3,648 |
| large-unicode-text-code | parse-evaluate | 2,880 | 0 | 2,880 | 2,316 | 5,769 | 0 | 8,000 | 3,737,000 | 3,705 | 0 | 3,648 |
| large-unicode-text-concatenate | evaluate | 2,880 | 2,885,000 | 5,765 | 4,260 | 8,674 | 0 | 6,000 | 3,381,000 | 3,333 | 2,885 | 3,600 |
| large-unicode-text-concatenate | parse-evaluate | 2,880 | 2,885,000 | 5,765 | 5,360 | 8,674 | 0 | 10,000 | 6,733,000 | 6,653 | 2,885 | 3,664 |
| large-unicode-text-dollar | evaluate | 2,880 | 6,000 | 2,886 | 1,077 | 22 | 0 | 6,000 | 502,000 | 454 | 6 | 3,540 |
| large-unicode-text-dollar | parse-evaluate | 2,880 | 6,000 | 2,886 | 1,298 | 22 | 0 | 10,000 | 963,000 | 883 | 6 | 3,672 |
| large-unicode-text-exact | evaluate | 2,880 | 0 | 2,880 | 2,326 | 11,533 | 0 | 5,000 | 496,000 | 448 | 0 | 3,660 |
| large-unicode-text-exact | parse-evaluate | 2,880 | 0 | 2,880 | 4,199 | 11,533 | 0 | 9,000 | 6,717,000 | 6,637 | 0 | 3,648 |
| large-unicode-text-find | evaluate | 2,880 | 0 | 2,880 | 2,522 | 5,774 | 0 | 5,000 | 496,000 | 448 | 0 | 3,664 |
| large-unicode-text-find | parse-evaluate | 2,880 | 0 | 2,880 | 3,651 | 5,774 | 0 | 9,000 | 3,837,000 | 3,757 | 0 | 3,532 |
| large-unicode-text-fixed | evaluate | 2,880 | 5,000 | 2,885 | 1,247 | 26 | 0 | 7,000 | 853,000 | 629 | 5 | 3,684 |
| large-unicode-text-fixed | parse-evaluate | 2,880 | 5,000 | 2,885 | 1,511 | 26 | 0 | 11,000 | 1,320,000 | 1,064 | 5 | 3,664 |
| large-unicode-text-jis | evaluate | 2,880 | 4,800,000 | 7,680 | 32,681 | 12,488 | 0 | 5,000 | 5,200,000 | 5,200 | 4,800 | 3,668 |
| large-unicode-text-jis | parse-evaluate | 2,880 | 4,800,000 | 7,680 | 33,809 | 12,488 | 0 | 9,000 | 8,536,000 | 8,504 | 4,800 | 3,628 |
| large-unicode-text-left | evaluate | 2,880 | 2,000 | 2,882 | 2,443 | 5,773 | 0 | 5,000 | 496,000 | 448 | 0 | 3,648 |
| large-unicode-text-left | parse-evaluate | 2,880 | 2,000 | 2,882 | 3,525 | 5,773 | 0 | 9,000 | 3,835,000 | 3,755 | 0 | 3,648 |
| large-unicode-text-len | evaluate | 2,880 | 0 | 2,880 | 2,258 | 5,768 | 0 | 4,000 | 400,000 | 400 | 0 | 3,668 |
| large-unicode-text-len | parse-evaluate | 2,880 | 0 | 2,880 | 3,279 | 5,768 | 0 | 8,000 | 3,736,000 | 3,704 | 0 | 3,676 |
| large-unicode-text-lower | evaluate | 2,880 | 2,880,000 | 5,760 | 47,333 | 8,650 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,668 |
| large-unicode-text-lower | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 48,231 | 8,650 | 0 | 9,000 | 6,618,000 | 6,586 | 2,880 | 3,528 |
| large-unicode-text-mid | evaluate | 2,880 | 2,000 | 2,882 | 2,609 | 5,776 | 0 | 6,000 | 848,000 | 624 | 0 | 3,656 |
| large-unicode-text-mid | parse-evaluate | 2,880 | 2,000 | 2,882 | 3,705 | 5,776 | 0 | 10,000 | 4,188,000 | 3,932 | 0 | 3,668 |
| large-unicode-text-proper | evaluate | 2,880 | 2,880,000 | 5,760 | 53,404 | 8,651 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,488 |
| large-unicode-text-proper | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 54,413 | 8,651 | 0 | 9,000 | 6,619,000 | 6,587 | 2,880 | 3,604 |
| large-unicode-text-replace | evaluate | 2,880 | 2,880,000 | 5,760 | 5,570 | 8,665 | 0 | 8,000 | 3,920,000 | 3,600 | 2,880 | 3,700 |
| large-unicode-text-replace | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 6,757 | 8,665 | 0 | 13,000 | 8,036,000 | 7,300 | 2,880 | 3,592 |
| large-unicode-text-rept | evaluate | 2,880 | 5,760,000 | 8,640 | 6,761 | 11,533 | 0 | 6,000 | 6,256,000 | 6,208 | 5,760 | 3,648 |
| large-unicode-text-rept | parse-evaluate | 2,880 | 5,760,000 | 8,640 | 7,842 | 11,533 | 0 | 10,000 | 9,595,000 | 9,515 | 5,760 | 3,592 |
| large-unicode-text-right | evaluate | 2,880 | 2,000 | 2,882 | 3,502 | 5,774 | 0 | 5,000 | 496,000 | 448 | 0 | 3,660 |
| large-unicode-text-right | parse-evaluate | 2,880 | 2,000 | 2,882 | 4,590 | 5,774 | 0 | 9,000 | 3,836,000 | 3,756 | 0 | 3,644 |
| large-unicode-text-search | evaluate | 2,880 | 0 | 2,880 | 2,812 | 5,784 | 0 | 8,000 | 524,000 | 476 | 0 | 3,596 |
| large-unicode-text-search | parse-evaluate | 2,880 | 0 | 2,880 | 3,929 | 5,784 | 0 | 12,000 | 3,867,000 | 3,787 | 0 | 3,644 |
| large-unicode-text-substitute | evaluate | 2,880 | 2,880,000 | 5,760 | 13,561 | 8,665 | 0 | 7,000 | 3,728,000 | 3,504 | 2,880 | 3,628 |
| large-unicode-text-substitute | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 14,648 | 8,665 | 0 | 11,000 | 7,079,000 | 6,823 | 2,880 | 3,604 |
| large-unicode-text-t | evaluate | 2,880 | 2,880,000 | 5,760 | 3,826 | 5,766 | 0 | 4,000 | 400,000 | 400 | 0 | 3,668 |
| large-unicode-text-t | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 4,828 | 5,766 | 0 | 8,000 | 3,734,000 | 3,702 | 0 | 3,636 |
| large-unicode-text-text | evaluate | 2,880 | 5,000 | 2,885 | 1,471 | 25 | 0 | 6,000 | 501,000 | 453 | 5 | 3,700 |
| large-unicode-text-text | parse-evaluate | 2,880 | 5,000 | 2,885 | 1,691 | 25 | 0 | 10,000 | 965,000 | 885 | 5 | 3,696 |
| large-unicode-text-trim | evaluate | 2,880 | 2,879,000 | 5,759 | 13,156 | 8,648 | 0 | 5,000 | 3,279,000 | 3,279 | 2,879 | 3,600 |
| large-unicode-text-trim | parse-evaluate | 2,880 | 2,879,000 | 5,759 | 14,195 | 8,648 | 0 | 9,000 | 6,616,000 | 6,584 | 2,879 | 3,532 |
| large-unicode-text-unichar | evaluate | 2,880 | 1,000 | 2,881 | 632 | 15 | 0 | 5,000 | 401,000 | 401 | 1 | 3,648 |
| large-unicode-text-unichar | parse-evaluate | 2,880 | 1,000 | 2,881 | 841 | 15 | 0 | 9,000 | 861,000 | 829 | 1 | 3,656 |
| large-unicode-text-unicode | evaluate | 2,880 | 0 | 2,880 | 1,290 | 5,772 | 0 | 4,000 | 400,000 | 400 | 0 | 3,596 |
| large-unicode-text-unicode | parse-evaluate | 2,880 | 0 | 2,880 | 2,329 | 5,772 | 0 | 8,000 | 3,740,000 | 3,708 | 0 | 3,536 |
| large-unicode-text-upper | evaluate | 2,880 | 2,880,000 | 5,760 | 45,977 | 8,650 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,604 |
| large-unicode-text-upper | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 47,117 | 8,650 | 0 | 9,000 | 6,618,000 | 6,586 | 2,880 | 3,668 |
| matrix-broadcast-text-concatenate | evaluate | 784 | 67,840 | 1,632 | 47,204 | 2,293 | 128 | 11,440 | 707,200 | 8,713 | 6,480 | 3,764 |
| matrix-broadcast-text-concatenate | parse-evaluate | 784 | 67,840 | 1,632 | 47,804 | 2,293 | 128 | 12,160 | 817,840 | 10,064 | 6,480 | 3,736 |
| matrix-broadcast-text-exact | evaluate | 784 | 0 | 784 | 35,816 | 1,503 | 128 | 6,320 | 639,360 | 7,865 | 5,632 | 3,740 |
| matrix-broadcast-text-exact | parse-evaluate | 784 | 0 | 784 | 36,316 | 1,503 | 128 | 7,040 | 749,520 | 9,210 | 5,632 | 3,744 |
| matrix-broadcast-text-find | evaluate | 784 | 0 | 784 | 22,074 | 1,309 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,732 |
| matrix-broadcast-text-find | parse-evaluate | 784 | 0 | 784 | 22,499 | 1,309 | 64 | 1,520 | 711,600 | 8,799 | 5,632 | 3,720 |
| matrix-broadcast-text-left | evaluate | 784 | 12,640 | 942 | 22,273 | 1,310 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,712 |
| matrix-broadcast-text-left | parse-evaluate | 784 | 12,640 | 942 | 22,818 | 1,310 | 128 | 1,920 | 744,320 | 9,208 | 5,632 | 3,708 |
| matrix-broadcast-text-len | evaluate | 784 | 0 | 784 | 15,577 | 1,113 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,736 |
| matrix-broadcast-text-len | parse-evaluate | 784 | 0 | 784 | 15,889 | 1,113 | 64 | 1,440 | 700,960 | 8,730 | 5,632 | 3,764 |
| matrix-broadcast-text-lower | evaluate | 784 | 62,720 | 1,568 | 32,391 | 1,483 | 64 | 3,440 | 621,440 | 7,768 | 6,000 | 3,752 |
| matrix-broadcast-text-lower | parse-evaluate | 784 | 62,720 | 1,568 | 32,865 | 1,483 | 64 | 4,000 | 730,560 | 9,100 | 6,000 | 3,776 |
| matrix-broadcast-text-mid | evaluate | 784 | 12,240 | 937 | 27,550 | 1,440 | 128 | 1,360 | 746,240 | 8,704 | 5,632 | 3,732 |
| matrix-broadcast-text-mid | parse-evaluate | 784 | 12,240 | 937 | 27,985 | 1,440 | 128 | 2,080 | 856,400 | 10,049 | 5,632 | 3,704 |
| matrix-broadcast-text-proper | evaluate | 784 | 62,720 | 1,568 | 39,475 | 1,676 | 64 | 4,720 | 636,800 | 7,960 | 6,192 | 3,716 |
| matrix-broadcast-text-proper | parse-evaluate | 784 | 62,720 | 1,568 | 40,101 | 1,676 | 64 | 5,280 | 746,000 | 9,293 | 6,192 | 3,764 |
| matrix-broadcast-text-replace | evaluate | 784 | 61,680 | 1,555 | 43,491 | 2,410 | 128 | 6,560 | 850,800 | 9,883 | 6,403 | 3,700 |
| matrix-broadcast-text-replace | parse-evaluate | 784 | 61,680 | 1,555 | 44,063 | 2,410 | 128 | 7,360 | 1,023,040 | 11,620 | 6,403 | 3,700 |
| matrix-broadcast-text-right | evaluate | 784 | 15,680 | 980 | 23,111 | 1,311 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,696 |
| matrix-broadcast-text-right | parse-evaluate | 784 | 15,680 | 980 | 23,709 | 1,311 | 128 | 1,920 | 744,400 | 9,209 | 5,632 | 3,720 |
| matrix-broadcast-text-search | evaluate | 784 | 0 | 784 | 40,868 | 1,799 | 64 | 16,320 | 745,600 | 7,492 | 5,632 | 3,736 |
| matrix-broadcast-text-search | parse-evaluate | 784 | 0 | 784 | 41,171 | 1,799 | 64 | 16,880 | 855,120 | 8,829 | 5,632 | 3,748 |
| matrix-broadcast-text-substitute | evaluate | 784 | 62,720 | 1,568 | 34,980 | 1,934 | 64 | 4,320 | 748,160 | 8,728 | 6,056 | 3,736 |
| matrix-broadcast-text-substitute | parse-evaluate | 784 | 62,720 | 1,568 | 35,818 | 1,934 | 64 | 4,880 | 858,320 | 10,073 | 6,056 | 3,748 |
| matrix-broadcast-text-trim | evaluate | 784 | 61,440 | 1,552 | 18,121 | 1,194 | 64 | 1,520 | 598,400 | 7,480 | 5,712 | 3,724 |
| matrix-broadcast-text-trim | parse-evaluate | 784 | 61,440 | 1,552 | 18,419 | 1,194 | 64 | 2,080 | 707,440 | 8,811 | 5,712 | 3,764 |
| reference-64-text-asc | evaluate | 784 | 58,880 | 1,520 | 19,975 | 1,241 | 64 | 1,520 | 597,120 | 7,464 | 5,696 | 3,764 |
| reference-64-text-asc | parse-evaluate | 784 | 58,880 | 1,520 | 20,373 | 1,241 | 64 | 2,080 | 706,080 | 8,794 | 5,696 | 3,700 |
| reference-64-text-char | evaluate | 784 | 5,120 | 848 | 24,633 | 394 | 64 | 6,000 | 597,120 | 7,464 | 5,696 | 3,728 |
| reference-64-text-char | parse-evaluate | 784 | 5,120 | 848 | 24,849 | 394 | 64 | 6,560 | 706,160 | 8,795 | 5,696 | 3,724 |
| reference-64-text-clean | evaluate | 784 | 62,720 | 1,568 | 18,772 | 1,115 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,700 |
| reference-64-text-clean | parse-evaluate | 784 | 62,720 | 1,568 | 19,152 | 1,115 | 64 | 1,440 | 701,120 | 8,732 | 5,632 | 3,728 |
| reference-64-text-code | evaluate | 784 | 0 | 784 | 15,147 | 1,114 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,776 |
| reference-64-text-code | parse-evaluate | 784 | 0 | 784 | 15,522 | 1,114 | 64 | 1,440 | 701,040 | 8,731 | 5,632 | 3,828 |
| reference-64-text-concatenate | evaluate | 784 | 88,320 | 1,888 | 32,260 | 2,680 | 64 | 6,080 | 690,560 | 8,568 | 6,736 | 3,720 |
| reference-64-text-concatenate | parse-evaluate | 784 | 88,320 | 1,888 | 32,601 | 2,680 | 64 | 6,640 | 800,800 | 9,914 | 6,736 | 3,728 |
| reference-64-text-dollar | evaluate | 784 | 25,600 | 1,104 | 42,683 | 719 | 64 | 6,080 | 627,840 | 7,784 | 5,952 | 3,708 |
| reference-64-text-dollar | parse-evaluate | 784 | 25,600 | 1,104 | 43,465 | 719 | 64 | 6,640 | 737,200 | 9,119 | 5,952 | 3,764 |
| reference-64-text-exact | evaluate | 1,568 | 0 | 1,568 | 21,821 | 2,095 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,764 |
| reference-64-text-exact | parse-evaluate | 1,568 | 0 | 1,568 | 22,382 | 2,095 | 128 | 1,920 | 744,400 | 9,209 | 5,632 | 3,756 |
| reference-64-text-find | evaluate | 784 | 0 | 784 | 22,015 | 1,309 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,780 |
| reference-64-text-find | parse-evaluate | 784 | 0 | 784 | 22,522 | 1,309 | 64 | 1,520 | 711,600 | 8,799 | 5,632 | 3,752 |
| reference-64-text-fixed | evaluate | 784 | 20,480 | 1,040 | 44,830 | 726 | 64 | 6,240 | 734,720 | 8,560 | 5,888 | 3,716 |
| reference-64-text-fixed | parse-evaluate | 784 | 20,480 | 1,040 | 45,431 | 726 | 64 | 6,800 | 844,560 | 9,901 | 5,888 | 3,760 |
| reference-64-text-jis | evaluate | 784 | 133,120 | 2,448 | 33,629 | 3,393 | 64 | 6,000 | 725,120 | 9,064 | 7,296 | 3,780 |
| reference-64-text-jis | parse-evaluate | 784 | 133,120 | 2,448 | 34,107 | 3,393 | 64 | 6,560 | 834,080 | 10,394 | 7,296 | 3,724 |
| reference-64-text-left | evaluate | 784 | 13,440 | 952 | 19,604 | 1,245 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,768 |
| reference-64-text-left | parse-evaluate | 784 | 13,440 | 952 | 20,225 | 1,245 | 64 | 1,520 | 711,440 | 8,797 | 5,632 | 3,720 |
| reference-64-text-len | evaluate | 784 | 0 | 784 | 15,495 | 1,113 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,756 |
| reference-64-text-len | parse-evaluate | 784 | 0 | 784 | 15,905 | 1,113 | 64 | 1,440 | 700,960 | 8,730 | 5,632 | 3,732 |
| reference-64-text-lower | evaluate | 784 | 62,720 | 1,568 | 32,353 | 1,483 | 64 | 3,440 | 621,440 | 7,768 | 6,000 | 3,812 |
| reference-64-text-lower | parse-evaluate | 784 | 62,720 | 1,568 | 32,729 | 1,483 | 64 | 4,000 | 730,560 | 9,100 | 6,000 | 3,760 |
| reference-64-text-mid | evaluate | 784 | 11,520 | 928 | 23,791 | 1,375 | 64 | 1,120 | 714,240 | 8,304 | 5,632 | 3,712 |
| reference-64-text-mid | parse-evaluate | 784 | 11,520 | 928 | 24,253 | 1,375 | 64 | 1,680 | 823,520 | 9,638 | 5,632 | 3,760 |
| reference-64-text-proper | evaluate | 784 | 62,720 | 1,568 | 39,586 | 1,676 | 64 | 4,720 | 636,800 | 7,960 | 6,192 | 3,764 |
| reference-64-text-proper | parse-evaluate | 784 | 62,720 | 1,568 | 39,919 | 1,676 | 64 | 5,280 | 746,000 | 9,293 | 6,192 | 3,724 |
| reference-64-text-replace | evaluate | 784 | 61,440 | 1,552 | 39,006 | 2,342 | 64 | 6,320 | 818,560 | 9,480 | 6,400 | 3,760 |
| reference-64-text-replace | parse-evaluate | 784 | 61,440 | 1,552 | 39,370 | 2,342 | 64 | 6,960 | 989,920 | 11,206 | 6,400 | 3,724 |
| reference-64-text-rept | evaluate | 784 | 125,440 | 2,352 | 28,411 | 2,813 | 64 | 6,080 | 727,680 | 9,032 | 7,200 | 3,760 |
| reference-64-text-rept | parse-evaluate | 784 | 125,440 | 2,352 | 28,803 | 2,813 | 64 | 6,640 | 836,880 | 10,365 | 7,200 | 3,736 |
| reference-64-text-right | evaluate | 784 | 16,000 | 984 | 20,360 | 1,246 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,764 |
| reference-64-text-right | parse-evaluate | 784 | 16,000 | 984 | 20,826 | 1,246 | 64 | 1,520 | 711,520 | 8,798 | 5,632 | 3,764 |
| reference-64-text-search | evaluate | 784 | 0 | 784 | 40,653 | 1,799 | 64 | 16,320 | 745,600 | 7,492 | 5,632 | 3,716 |
| reference-64-text-search | parse-evaluate | 784 | 0 | 784 | 40,965 | 1,799 | 64 | 16,880 | 855,120 | 8,829 | 5,632 | 3,724 |
| reference-64-text-substitute | evaluate | 784 | 62,720 | 1,568 | 35,079 | 1,934 | 64 | 4,320 | 748,160 | 8,728 | 6,056 | 3,764 |
| reference-64-text-substitute | parse-evaluate | 784 | 62,720 | 1,568 | 35,742 | 1,934 | 64 | 4,880 | 858,320 | 10,073 | 6,056 | 3,728 |
| reference-64-text-t | evaluate | 784 | 62,720 | 1,568 | 14,256 | 1,111 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,764 |
| reference-64-text-t | parse-evaluate | 784 | 62,720 | 1,568 | 14,654 | 1,111 | 64 | 1,440 | 700,800 | 8,728 | 5,632 | 3,700 |
| reference-64-text-text | evaluate | 784 | 20,480 | 1,040 | 58,963 | 848 | 64 | 6,080 | 622,720 | 7,720 | 5,888 | 3,764 |
| reference-64-text-text | parse-evaluate | 784 | 20,480 | 1,040 | 60,038 | 848 | 64 | 6,640 | 732,320 | 9,058 | 5,888 | 3,684 |
| reference-64-text-trim | evaluate | 784 | 61,440 | 1,552 | 18,044 | 1,194 | 64 | 1,520 | 598,400 | 7,480 | 5,712 | 3,724 |
| reference-64-text-trim | parse-evaluate | 784 | 61,440 | 1,552 | 18,443 | 1,194 | 64 | 2,080 | 707,440 | 8,811 | 5,712 | 3,744 |
| reference-64-text-unichar | evaluate | 784 | 5,120 | 848 | 24,932 | 397 | 64 | 6,000 | 597,120 | 7,464 | 5,696 | 3,768 |
| reference-64-text-unichar | parse-evaluate | 784 | 5,120 | 848 | 25,336 | 397 | 64 | 6,560 | 706,400 | 8,798 | 5,696 | 3,716 |
| reference-64-text-unicode | evaluate | 784 | 0 | 784 | 15,509 | 1,117 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,764 |
| reference-64-text-unicode | parse-evaluate | 784 | 0 | 784 | 15,905 | 1,117 | 64 | 1,440 | 701,280 | 8,734 | 5,632 | 3,720 |
| reference-64-text-upper | evaluate | 784 | 62,720 | 1,568 | 37,788 | 1,747 | 64 | 5,360 | 642,560 | 8,032 | 6,264 | 3,764 |
| reference-64-text-upper | parse-evaluate | 784 | 62,720 | 1,568 | 38,226 | 1,747 | 64 | 5,920 | 751,680 | 9,364 | 6,264 | 3,720 |
| refusal-text-asc | evaluate | 0 | 0 | 0 | 2,329 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,752 |
| refusal-text-asc | parse-evaluate | 0 | 0 | 0 | 2,895 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,760 |
| refusal-text-char | evaluate | 0 | 0 | 0 | 1,353 | 14 | 0 | 720 | 150,400 | 1,464 | 0 | 3,728 |
| refusal-text-char | parse-evaluate | 0 | 0 | 0 | 1,546 | 14 | 0 | 1,040 | 186,960 | 1,889 | 0 | 3,716 |
| refusal-text-clean | evaluate | 0 | 0 | 0 | 2,526 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,700 |
| refusal-text-clean | parse-evaluate | 0 | 0 | 0 | 2,910 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,724 |
| refusal-text-code | evaluate | 0 | 0 | 0 | 2,342 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,792 |
| refusal-text-code | parse-evaluate | 0 | 0 | 0 | 2,893 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,728 |
| refusal-text-concatenate | evaluate | 0 | 0 | 0 | 2,445 | 25 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,716 |
| refusal-text-concatenate | parse-evaluate | 0 | 0 | 0 | 2,910 | 25 | 0 | 1,920 | 325,040 | 2,927 | 0 | 3,752 |
| refusal-text-dollar | evaluate | 0 | 0 | 0 | 1,391 | 22 | 0 | 720 | 150,400 | 1,464 | 0 | 3,716 |
| refusal-text-dollar | parse-evaluate | 0 | 0 | 0 | 1,603 | 22 | 0 | 1,040 | 187,680 | 1,898 | 0 | 3,740 |
| refusal-text-exact | evaluate | 0 | 0 | 0 | 2,359 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,708 |
| refusal-text-exact | parse-evaluate | 0 | 0 | 0 | 2,908 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,716 |
| refusal-text-find | evaluate | 0 | 0 | 0 | 2,325 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,696 |
| refusal-text-find | parse-evaluate | 0 | 0 | 0 | 2,902 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,760 |
| refusal-text-fixed | evaluate | 0 | 0 | 0 | 1,404 | 21 | 0 | 720 | 150,400 | 1,464 | 0 | 3,720 |
| refusal-text-fixed | parse-evaluate | 0 | 0 | 0 | 1,615 | 21 | 0 | 1,040 | 187,600 | 1,897 | 0 | 3,708 |
| refusal-text-jis | evaluate | 0 | 0 | 0 | 2,332 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,728 |
| refusal-text-jis | parse-evaluate | 0 | 0 | 0 | 2,871 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,716 |
| refusal-text-left | evaluate | 0 | 0 | 0 | 2,343 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,708 |
| refusal-text-left | parse-evaluate | 0 | 0 | 0 | 2,886 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,724 |
| refusal-text-len | evaluate | 0 | 0 | 0 | 2,354 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,756 |
| refusal-text-len | parse-evaluate | 0 | 0 | 0 | 2,894 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,700 |
| refusal-text-lower | evaluate | 0 | 0 | 0 | 2,405 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,692 |
| refusal-text-lower | parse-evaluate | 0 | 0 | 0 | 2,905 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,700 |
| refusal-text-mid | evaluate | 0 | 0 | 0 | 2,315 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,716 |
| refusal-text-mid | parse-evaluate | 0 | 0 | 0 | 2,904 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,716 |
| refusal-text-proper | evaluate | 0 | 0 | 0 | 2,333 | 20 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,724 |
| refusal-text-proper | parse-evaluate | 0 | 0 | 0 | 2,931 | 20 | 0 | 1,920 | 324,640 | 2,922 | 0 | 3,768 |
| refusal-text-replace | evaluate | 0 | 0 | 0 | 2,325 | 21 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,764 |
| refusal-text-replace | parse-evaluate | 0 | 0 | 0 | 2,923 | 21 | 0 | 1,920 | 324,720 | 2,923 | 0 | 3,708 |
| refusal-text-rept | evaluate | 0 | 0 | 0 | 2,437 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,764 |
| refusal-text-rept | parse-evaluate | 0 | 0 | 0 | 2,897 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,748 |
| refusal-text-right | evaluate | 0 | 0 | 0 | 2,387 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,724 |
| refusal-text-right | parse-evaluate | 0 | 0 | 0 | 2,953 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,760 |
| refusal-text-search | evaluate | 0 | 0 | 0 | 2,340 | 20 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,764 |
| refusal-text-search | parse-evaluate | 0 | 0 | 0 | 2,996 | 20 | 0 | 1,920 | 324,640 | 2,922 | 0 | 3,716 |
| refusal-text-substitute | evaluate | 0 | 0 | 0 | 2,340 | 24 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,760 |
| refusal-text-substitute | parse-evaluate | 0 | 0 | 0 | 2,912 | 24 | 0 | 1,920 | 324,960 | 2,926 | 0 | 3,764 |
| refusal-text-t | evaluate | 0 | 0 | 0 | 2,321 | 15 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,708 |
| refusal-text-t | parse-evaluate | 0 | 0 | 0 | 2,856 | 15 | 0 | 1,920 | 324,240 | 2,917 | 0 | 3,760 |
| refusal-text-text | evaluate | 0 | 0 | 0 | 1,452 | 21 | 0 | 720 | 150,400 | 1,464 | 0 | 3,728 |
| refusal-text-text | parse-evaluate | 0 | 0 | 0 | 1,653 | 21 | 0 | 1,040 | 187,520 | 1,896 | 0 | 3,700 |
| refusal-text-trim | evaluate | 0 | 0 | 0 | 2,455 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,796 |
| refusal-text-trim | parse-evaluate | 0 | 0 | 0 | 2,889 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,744 |
| refusal-text-unichar | evaluate | 0 | 0 | 0 | 1,352 | 17 | 0 | 720 | 150,400 | 1,464 | 0 | 3,716 |
| refusal-text-unichar | parse-evaluate | 0 | 0 | 0 | 1,572 | 17 | 0 | 1,040 | 187,200 | 1,892 | 0 | 3,720 |
| refusal-text-unicode | evaluate | 0 | 0 | 0 | 2,349 | 21 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,716 |
| refusal-text-unicode | parse-evaluate | 0 | 0 | 0 | 2,922 | 21 | 0 | 1,920 | 324,720 | 2,923 | 0 | 3,712 |
| refusal-text-upper | evaluate | 0 | 0 | 0 | 2,345 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,760 |
| refusal-text-upper | parse-evaluate | 0 | 0 | 0 | 2,925 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,732 |
| rept-growth | evaluate | 8,192 | 8,192,000 | 16,384 | 26,410 | 8,210 | 0 | 6,000 | 8,688,000 | 8,640 | 8,192 | 3,540 |
| rept-growth | parse-evaluate | 8,192 | 8,192,000 | 16,384 | 26,683 | 8,210 | 0 | 10,000 | 9,151,000 | 9,071 | 8,192 | 3,644 |
| resource-text-concatenate | evaluate | 0 | 0 | 0 | 506 | 15 | 0 | 400 | 33,920 | 360 | 0 | 3,728 |
| resource-text-concatenate | parse-evaluate | 0 | 0 | 0 | 957 | 15 | 0 | 960 | 144,160 | 1,706 | 0 | 3,728 |
| resource-text-len | evaluate | 0 | 0 | 0 | 378 | 6 | 0 | 320 | 23,680 | 296 | 0 | 3,736 |
| resource-text-len | parse-evaluate | 0 | 0 | 0 | 745 | 6 | 0 | 880 | 132,640 | 1,626 | 0 | 3,712 |
| resource-text-search | evaluate | 0 | 0 | 0 | 647 | 12 | 0 | 480 | 64,640 | 744 | 0 | 3,724 |
| resource-text-search | parse-evaluate | 0 | 0 | 0 | 1,076 | 12 | 0 | 1,040 | 174,160 | 2,081 | 0 | 3,736 |
| resource-text-substitute | evaluate | 0 | 0 | 0 | 522 | 15 | 0 | 400 | 33,920 | 360 | 0 | 3,780 |
| resource-text-substitute | parse-evaluate | 0 | 0 | 0 | 981 | 15 | 0 | 960 | 144,080 | 1,705 | 0 | 3,728 |
| search-worstcase-exact | evaluate | 4,229 | 0 | 4,229 | 2,979 | 16,421 | 0 | 5,000 | 496,000 | 448 | 0 | 3,636 |
| search-worstcase-exact | parse-evaluate | 4,229 | 0 | 4,229 | 5,519 | 16,421 | 0 | 9,000 | 9,161,000 | 9,081 | 0 | 3,576 |
| search-worstcase-find | evaluate | 4,229 | 0 | 4,229 | 9,489 | 8,482 | 0 | 5,000 | 496,000 | 448 | 0 | 3,620 |
| search-worstcase-find | parse-evaluate | 4,229 | 0 | 4,229 | 10,906 | 8,482 | 0 | 9,000 | 5,191,000 | 5,111 | 0 | 3,592 |
| search-worstcase-search | evaluate | 4,229 | 0 | 4,229 | 96,827 | 17,079 | 0 | 8,000 | 4,220,000 | 4,172 | 0 | 3,692 |
| search-worstcase-search | parse-evaluate | 4,229 | 0 | 4,229 | 98,349 | 17,079 | 0 | 12,000 | 8,917,000 | 8,837 | 0 | 3,688 |
| search-worstcase-substitute | evaluate | 4,229 | 3,970,000 | 8,199 | 14,298 | 12,463 | 0 | 7,000 | 4,818,000 | 4,594 | 3,970 | 3,656 |
| search-worstcase-substitute | parse-evaluate | 4,229 | 3,970,000 | 8,199 | 15,775 | 12,463 | 0 | 11,000 | 9,523,000 | 9,267 | 3,970 | 3,536 |
| text-format-fraction-six | evaluate | 31 | 3 | 34 | 1,580 | 60 | 0 | 6 | 499 | 451 | 3 | 3,692 |
| text-format-fraction-six | parse-evaluate | 31 | 3 | 34 | 1,940 | 60 | 0 | 10 | 988 | 908 | 3 | 3,660 |
| tiny-text-asc | evaluate | 12 | 11,000 | 23 | 546 | 30 | 0 | 4,000 | 400,000 | 400 | 0 | 3,644 |
| tiny-text-asc | parse-evaluate | 12 | 11,000 | 23 | 740 | 30 | 0 | 8,000 | 867,000 | 835 | 0 | 3,640 |
| tiny-text-char | evaluate | 12 | 1,000 | 13 | 631 | 12 | 0 | 5,000 | 401,000 | 401 | 1 | 3,572 |
| tiny-text-char | parse-evaluate | 12 | 1,000 | 13 | 820 | 12 | 0 | 9,000 | 858,000 | 826 | 1 | 3,644 |
| tiny-text-clean | evaluate | 12 | 11,000 | 23 | 562 | 32 | 0 | 4,000 | 400,000 | 400 | 0 | 3,672 |
| tiny-text-clean | parse-evaluate | 12 | 11,000 | 23 | 771 | 32 | 0 | 8,000 | 869,000 | 837 | 0 | 3,656 |
| tiny-text-code | evaluate | 12 | 0 | 12 | 510 | 31 | 0 | 4,000 | 400,000 | 400 | 0 | 3,516 |
| tiny-text-code | parse-evaluate | 12 | 0 | 12 | 706 | 31 | 0 | 8,000 | 868,000 | 836 | 0 | 3,668 |
| tiny-text-concatenate | evaluate | 12 | 16,000 | 28 | 915 | 67 | 0 | 6,000 | 512,000 | 464 | 16 | 3,644 |
| tiny-text-concatenate | parse-evaluate | 12 | 16,000 | 28 | 1,145 | 67 | 0 | 10,000 | 995,000 | 915 | 16 | 3,476 |
| tiny-text-dollar | evaluate | 12 | 6,000 | 18 | 1,074 | 22 | 0 | 6,000 | 502,000 | 454 | 6 | 3,668 |
| tiny-text-dollar | parse-evaluate | 12 | 6,000 | 18 | 1,306 | 22 | 0 | 10,000 | 963,000 | 883 | 6 | 3,724 |
| tiny-text-exact | evaluate | 12 | 0 | 12 | 713 | 57 | 0 | 5,000 | 496,000 | 448 | 0 | 3,664 |
| tiny-text-exact | parse-evaluate | 12 | 0 | 12 | 945 | 57 | 0 | 9,000 | 979,000 | 899 | 0 | 3,560 |
| tiny-text-find | evaluate | 12 | 0 | 12 | 750 | 36 | 0 | 5,000 | 496,000 | 448 | 0 | 3,636 |
| tiny-text-find | parse-evaluate | 12 | 0 | 12 | 998 | 36 | 0 | 9,000 | 968,000 | 888 | 0 | 3,652 |
| tiny-text-fixed | evaluate | 12 | 5,000 | 17 | 1,249 | 26 | 0 | 7,000 | 853,000 | 629 | 5 | 3,684 |
| tiny-text-fixed | parse-evaluate | 12 | 5,000 | 17 | 1,515 | 26 | 0 | 11,000 | 1,320,000 | 1,064 | 5 | 3,676 |
| tiny-text-jis | evaluate | 12 | 27,000 | 39 | 793 | 67 | 0 | 5,000 | 427,000 | 427 | 27 | 3,540 |
| tiny-text-jis | parse-evaluate | 12 | 27,000 | 39 | 998 | 67 | 0 | 9,000 | 894,000 | 862 | 27 | 3,516 |
| tiny-text-left | evaluate | 12 | 2,000 | 14 | 691 | 35 | 0 | 5,000 | 496,000 | 448 | 0 | 3,672 |
| tiny-text-left | parse-evaluate | 12 | 2,000 | 14 | 956 | 35 | 0 | 9,000 | 966,000 | 886 | 0 | 3,632 |
| tiny-text-len | evaluate | 12 | 0 | 12 | 509 | 30 | 0 | 4,000 | 400,000 | 400 | 0 | 3,664 |
| tiny-text-len | parse-evaluate | 12 | 0 | 12 | 704 | 30 | 0 | 8,000 | 867,000 | 835 | 0 | 3,660 |
| tiny-text-lower | evaluate | 12 | 11,000 | 23 | 886 | 43 | 0 | 5,000 | 411,000 | 411 | 11 | 3,664 |
| tiny-text-lower | parse-evaluate | 12 | 11,000 | 23 | 1,098 | 43 | 0 | 9,000 | 880,000 | 848 | 11 | 3,608 |
| tiny-text-mid | evaluate | 12 | 2,000 | 14 | 880 | 38 | 0 | 6,000 | 848,000 | 624 | 0 | 3,604 |
| tiny-text-mid | parse-evaluate | 12 | 2,000 | 14 | 1,120 | 38 | 0 | 10,000 | 1,319,000 | 1,063 | 0 | 3,664 |
| tiny-text-proper | evaluate | 12 | 11,000 | 23 | 933 | 44 | 0 | 5,000 | 411,000 | 411 | 11 | 3,548 |
| tiny-text-proper | parse-evaluate | 12 | 11,000 | 23 | 1,156 | 44 | 0 | 9,000 | 881,000 | 849 | 11 | 3,604 |
| tiny-text-replace | evaluate | 12 | 11,000 | 23 | 1,210 | 58 | 0 | 8,000 | 1,051,000 | 731 | 11 | 3,660 |
| tiny-text-replace | parse-evaluate | 12 | 11,000 | 23 | 1,585 | 58 | 0 | 13,000 | 2,298,000 | 1,562 | 11 | 3,636 |
| tiny-text-rept | evaluate | 12 | 22,000 | 34 | 877 | 57 | 0 | 6,000 | 518,000 | 470 | 22 | 3,536 |
| tiny-text-rept | parse-evaluate | 12 | 22,000 | 34 | 1,091 | 57 | 0 | 10,000 | 988,000 | 908 | 22 | 3,604 |
| tiny-text-right | evaluate | 12 | 2,000 | 14 | 710 | 36 | 0 | 5,000 | 496,000 | 448 | 0 | 3,604 |
| tiny-text-right | parse-evaluate | 12 | 2,000 | 14 | 958 | 36 | 0 | 9,000 | 967,000 | 887 | 0 | 3,604 |
| tiny-text-search | evaluate | 12 | 0 | 12 | 1,065 | 46 | 0 | 8,000 | 524,000 | 476 | 0 | 3,632 |
| tiny-text-search | parse-evaluate | 12 | 0 | 12 | 1,294 | 46 | 0 | 12,000 | 998,000 | 918 | 0 | 3,668 |
| tiny-text-substitute | evaluate | 12 | 11,000 | 23 | 1,168 | 58 | 0 | 7,000 | 859,000 | 635 | 11 | 3,616 |
| tiny-text-substitute | parse-evaluate | 12 | 11,000 | 23 | 1,441 | 58 | 0 | 11,000 | 1,341,000 | 1,085 | 11 | 3,668 |
| tiny-text-t | evaluate | 12 | 11,000 | 23 | 492 | 28 | 0 | 4,000 | 400,000 | 400 | 0 | 3,536 |
| tiny-text-t | parse-evaluate | 12 | 11,000 | 23 | 682 | 28 | 0 | 8,000 | 865,000 | 833 | 0 | 3,636 |
| tiny-text-text | evaluate | 12 | 5,000 | 17 | 1,465 | 25 | 0 | 6,000 | 501,000 | 453 | 5 | 3,688 |
| tiny-text-text | parse-evaluate | 12 | 5,000 | 17 | 1,698 | 25 | 0 | 10,000 | 965,000 | 885 | 5 | 3,656 |
| tiny-text-trim | evaluate | 12 | 11,000 | 23 | 520 | 31 | 0 | 4,000 | 400,000 | 400 | 0 | 3,668 |
| tiny-text-trim | parse-evaluate | 12 | 11,000 | 23 | 722 | 31 | 0 | 8,000 | 868,000 | 836 | 0 | 3,520 |
| tiny-text-unichar | evaluate | 12 | 1,000 | 13 | 629 | 15 | 0 | 5,000 | 401,000 | 401 | 1 | 3,648 |
| tiny-text-unichar | parse-evaluate | 12 | 1,000 | 13 | 838 | 15 | 0 | 9,000 | 861,000 | 829 | 1 | 3,544 |
| tiny-text-unicode | evaluate | 12 | 0 | 12 | 514 | 34 | 0 | 4,000 | 400,000 | 400 | 0 | 3,572 |
| tiny-text-unicode | parse-evaluate | 12 | 0 | 12 | 720 | 34 | 0 | 8,000 | 871,000 | 839 | 0 | 3,476 |
| tiny-text-upper | evaluate | 12 | 11,000 | 23 | 892 | 43 | 0 | 5,000 | 411,000 | 411 | 11 | 3,528 |
| tiny-text-upper | parse-evaluate | 12 | 11,000 | 23 | 1,102 | 43 | 0 | 9,000 | 880,000 | 848 | 11 | 3,628 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
