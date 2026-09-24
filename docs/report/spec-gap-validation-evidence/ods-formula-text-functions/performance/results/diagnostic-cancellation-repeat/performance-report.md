# ODS text-functions evaluator performance profile

The baseline is committed `8f09231e36982248eface4d599432143a67f6e49`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases; every row below is the p50 across those fresh children with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 47,682 | 46,372 | -2.7% | 2,311 | 2,311 | 88 | 88 | 3,688 | 4,004 |
| array-control-16x16-arithmetic | parse-evaluate | 57,710 | 58,630 | +1.6% | 2,311 | 2,311 | 352 | 352 | 3,684 | 4,020 |
| array-control-16x16-sin | evaluate | 128,680 | 129,610 | +0.7% | 2,311 | 2,311 | 2,144 | 2,144 | 4,048 | 4,300 |
| array-control-16x16-sin | parse-evaluate | 139,130 | 141,143 | +1.4% | 2,311 | 2,311 | 2,412 | 2,412 | 4,048 | 4,232 |
| array-control-4x4-arithmetic | evaluate | 4,497 | 4,475 | -0.5% | 151 | 151 | 1,120 | 1,120 | 3,408 | 3,668 |
| array-control-4x4-arithmetic | parse-evaluate | 5,366 | 5,342 | -0.4% | 151 | 151 | 2,240 | 2,240 | 3,416 | 3,716 |
| array-control-4x4-sin | evaluate | 9,855 | 9,920 | +0.7% | 151 | 151 | 3,840 | 3,840 | 3,780 | 4,048 |
| array-control-4x4-sin | parse-evaluate | 10,730 | 10,851 | +1.1% | 151 | 151 | 5,040 | 5,040 | 3,792 | 3,992 |
| concat-borrowed-literals | evaluate | 733 | 732 | -0.1% | 21 | 21 | 6,000 | 6,000 | 3,412 | 3,536 |
| concat-borrowed-literals | parse-evaluate | 868 | 869 | +0.1% | 21 | 21 | 9,000 | 9,000 | 3,420 | 3,584 |
| concat-growth-chain | evaluate | 1,892 | 1,928 | +1.9% | 20 | 20 | 10,000 | 10,000 | 3,388 | 3,476 |
| concat-growth-chain | parse-evaluate | 2,328 | 2,307 | -0.9% | 20 | 20 | 17,000 | 17,000 | 3,424 | 3,492 |
| concat-owned-left | evaluate | 1,078 | 1,074 | -0.4% | 24 | 24 | 8,000 | 8,000 | 3,424 | 3,504 |
| concat-owned-left | parse-evaluate | 1,316 | 1,323 | +0.5% | 24 | 24 | 13,000 | 13,000 | 3,408 | 3,484 |
| concat-owned-right | evaluate | 1,086 | 1,084 | -0.2% | 24 | 24 | 8,000 | 8,000 | 3,408 | 3,620 |
| concat-owned-right | parse-evaluate | 1,335 | 1,344 | +0.7% | 24 | 24 | 13,000 | 13,000 | 3,392 | 3,532 |
| database-control-dstdev | evaluate | 3,810 | 3,910 | +2.6% | 30 | 30 | 20 | 20 | 3,504 | 3,764 |
| database-control-dstdev | parse-evaluate | 4,500 | 4,660 | +3.6% | 30 | 30 | 29 | 29 | 3,480 | 3,780 |
| database-control-dsum | evaluate | 3,930 | 4,060 | +3.3% | 28 | 28 | 20 | 20 | 3,440 | 3,740 |
| database-control-dsum | parse-evaluate | 4,590 | 4,700 | +2.4% | 28 | 28 | 29 | 29 | 3,484 | 3,748 |
| database-control-dvar | evaluate | 3,870 | 3,980 | +2.8% | 28 | 28 | 20 | 20 | 3,452 | 3,736 |
| database-control-dvar | parse-evaluate | 4,500 | 4,720 | +4.9% | 28 | 28 | 29 | 29 | 3,468 | 3,768 |
| literal-aggregate-4x1-sum | evaluate | 1,820 | 1,810 | -0.5% | 15 | 15 | 10 | 10 | 3,448 | 3,780 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,610 | 2,541 | -2.6% | 15 | 15 | 23 | 23 | 3,456 | 3,744 |
| reference-aggregate-64x4-sum | evaluate | 10,240 | 10,295 | +0.5% | 16 | 16 | 32 | 32 | 3,440 | 3,760 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,665 | 10,667 | +0.0% | 16 | 16 | 60 | 60 | 3,444 | 3,692 |
| reference-array-16x4-arithmetic | evaluate | 9,487 | 9,623 | +1.4% | 16 | 16 | 800 | 800 | 3,420 | 3,732 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,875 | 9,908 | +0.3% | 16 | 16 | 1,280 | 1,280 | 3,420 | 3,716 |
| reference-conditional-256x4-sumifs | evaluate | 80,205 | 79,355 | -1.1% | 52 | 52 | 42 | 42 | 3,480 | 3,748 |
| reference-conditional-256x4-sumifs | parse-evaluate | 80,110 | 80,340 | +0.3% | 52 | 52 | 68 | 68 | 3,460 | 3,780 |
| reference-control-average | evaluate | 10,337 | 10,332 | -0.0% | 20 | 20 | 32 | 32 | 3,484 | 3,784 |
| reference-control-average | parse-evaluate | 10,745 | 10,642 | -1.0% | 20 | 20 | 60 | 60 | 3,464 | 3,760 |
| reference-control-counta | evaluate | 8,645 | 8,687 | +0.5% | 19 | 19 | 32 | 32 | 3,548 | 3,748 |
| reference-control-counta | parse-evaluate | 9,100 | 9,085 | -0.2% | 19 | 19 | 60 | 60 | 3,460 | 3,740 |
| representative-median | evaluate | 1,566 | 1,562 | -0.3% | 16 | 16 | 9,000 | 9,000 | 3,392 | 3,636 |
| representative-median | parse-evaluate | 1,852 | 1,849 | -0.2% | 16 | 16 | 14,000 | 14,000 | 3,420 | 3,544 |
| representative-percentrank | evaluate | 973 | 986 | +1.3% | 19 | 19 | 7,000 | 7,000 | 3,428 | 3,580 |
| representative-percentrank | parse-evaluate | 1,213 | 1,221 | +0.7% | 19 | 19 | 11,000 | 11,000 | 3,412 | 3,588 |
| representative-rank | evaluate | 932 | 937 | +0.5% | 12 | 12 | 7,000 | 7,000 | 3,420 | 3,588 |
| representative-rank | parse-evaluate | 1,169 | 1,138 | -2.7% | 12 | 12 | 11,000 | 11,000 | 3,420 | 3,488 |
| scalar-aggregate-sum | evaluate | 569 | 546 | -4.0% | 10 | 10 | 4,000 | 4,000 | 3,460 | 3,600 |
| scalar-aggregate-sum | parse-evaluate | 758 | 730 | -3.7% | 10 | 10 | 8,000 | 8,000 | 3,412 | 3,536 |
| scalar-control-arithmetic | evaluate | 574 | 574 | +0.0% | 10 | 10 | 5,000 | 5,000 | 3,368 | 3,568 |
| scalar-control-arithmetic | parse-evaluate | 715 | 733 | +2.5% | 10 | 10 | 8,000 | 8,000 | 3,412 | 3,604 |
| scalar-control-average | evaluate | 1,032 | 1,045 | +1.3% | 23 | 23 | 6,000 | 6,000 | 3,384 | 3,616 |
| scalar-control-average | parse-evaluate | 1,276 | 1,268 | -0.6% | 23 | 23 | 10,000 | 10,000 | 3,388 | 3,496 |
| scalar-control-counta | evaluate | 832 | 830 | -0.2% | 24 | 24 | 6,000 | 6,000 | 3,384 | 3,504 |
| scalar-control-counta | parse-evaluate | 1,079 | 1,077 | -0.2% | 24 | 24 | 10,000 | 10,000 | 3,420 | 3,556 |
| scalar-control-imsum | evaluate | 1,428 | 1,400 | -2.0% | 41 | 41 | 8,000 | 8,000 | 3,404 | 3,540 |
| scalar-control-imsum | parse-evaluate | 1,995 | 1,988 | -0.4% | 41 | 41 | 17,000 | 17,000 | 3,416 | 3,520 |
| scalar-control-sin | evaluate | 460 | 462 | +0.4% | 10 | 10 | 4,000 | 4,000 | 3,708 | 3,916 |
| scalar-control-sin | parse-evaluate | 646 | 640 | -0.9% | 10 | 10 | 8,000 | 8,000 | 3,712 | 3,824 |
| scalar-control-stdev | evaluate | 870 | 867 | -0.3% | 21 | 21 | 6,000 | 6,000 | 3,408 | 3,516 |
| scalar-control-stdev | parse-evaluate | 1,085 | 1,085 | +0.0% | 21 | 21 | 10,000 | 10,000 | 3,392 | 3,504 |
| scalar-control-var | evaluate | 853 | 851 | -0.2% | 19 | 19 | 6,000 | 6,000 | 3,388 | 3,496 |
| scalar-control-var | parse-evaluate | 1,068 | 1,057 | -1.0% | 19 | 19 | 10,000 | 10,000 | 3,372 | 3,592 |

## Text-function workloads

The candidate matrix covers scalar and large Unicode inputs, borrowed 64-cell references, mapped matrix outputs, typed refusal, sticky cancellation, resource refusal, worst-case searches, REPT growth, and ASC/JIS width conversion. Text input/output bytes and the reviewed Unicode/domain labels remain with the raw case receipts.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| asc-jis-expansion-asc | evaluate | 16 | 9,000 | 25 | 697 | 45 | 0 | 5,000 | 409,000 | 409 | 9 | 3,588 |
| asc-jis-expansion-asc | parse-evaluate | 16 | 9,000 | 25 | 891 | 45 | 0 | 9,000 | 877,000 | 845 | 9 | 3,548 |
| asc-jis-expansion-jis | evaluate | 16 | 12,000 | 28 | 698 | 42 | 0 | 5,000 | 412,000 | 412 | 12 | 3,524 |
| asc-jis-expansion-jis | parse-evaluate | 16 | 12,000 | 28 | 888 | 42 | 0 | 9,000 | 877,000 | 845 | 12 | 3,600 |
| cancellation-text-concatenate | evaluate | 0 | 0 | 0 | 29 | 0 | 0 | 11 | 7,352 | 7,288 | 0 | 3,692 |
| cancellation-text-concatenate | parse-evaluate | 0 | 0 | 0 | 413 | 0 | 0 | 571 | 117,592 | 8,634 | 0 | 3,744 |
| cancellation-text-len | evaluate | 0 | 0 | 0 | 27 | 0 | 0 | 10 | 7,224 | 7,224 | 0 | 3,744 |
| cancellation-text-len | parse-evaluate | 0 | 0 | 0 | 374 | 0 | 0 | 570 | 116,184 | 8,554 | 0 | 3,760 |
| cancellation-text-search | evaluate | 0 | 0 | 0 | 31 | 0 | 0 | 11 | 7,352 | 7,288 | 0 | 3,724 |
| cancellation-text-search | parse-evaluate | 0 | 0 | 0 | 400 | 0 | 0 | 571 | 116,872 | 8,625 | 0 | 3,800 |
| cancellation-text-substitute | evaluate | 0 | 0 | 0 | 32 | 0 | 0 | 12 | 8,400 | 7,952 | 0 | 3,772 |
| cancellation-text-substitute | parse-evaluate | 0 | 0 | 0 | 423 | 0 | 0 | 572 | 118,560 | 9,297 | 0 | 3,752 |
| large-unicode-text-asc | evaluate | 2,880 | 2,880,000 | 5,760 | 11,883 | 5,768 | 0 | 4,000 | 400,000 | 400 | 0 | 3,536 |
| large-unicode-text-asc | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 12,951 | 5,768 | 0 | 8,000 | 3,736,000 | 3,704 | 0 | 3,640 |
| large-unicode-text-char | evaluate | 2,880 | 1,000 | 2,881 | 628 | 12 | 0 | 5,000 | 401,000 | 401 | 1 | 3,620 |
| large-unicode-text-char | parse-evaluate | 2,880 | 1,000 | 2,881 | 813 | 12 | 0 | 9,000 | 858,000 | 826 | 1 | 3,536 |
| large-unicode-text-clean | evaluate | 2,880 | 2,880,000 | 5,760 | 13,414 | 5,770 | 0 | 4,000 | 400,000 | 400 | 0 | 3,596 |
| large-unicode-text-clean | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 14,424 | 5,770 | 0 | 8,000 | 3,738,000 | 3,706 | 0 | 3,664 |
| large-unicode-text-code | evaluate | 2,880 | 0 | 2,880 | 1,284 | 5,769 | 0 | 4,000 | 400,000 | 400 | 0 | 3,580 |
| large-unicode-text-code | parse-evaluate | 2,880 | 0 | 2,880 | 2,308 | 5,769 | 0 | 8,000 | 3,737,000 | 3,705 | 0 | 3,532 |
| large-unicode-text-concatenate | evaluate | 2,880 | 2,885,000 | 5,765 | 4,263 | 8,674 | 0 | 6,000 | 3,381,000 | 3,333 | 2,885 | 3,528 |
| large-unicode-text-concatenate | parse-evaluate | 2,880 | 2,885,000 | 5,765 | 5,353 | 8,674 | 0 | 10,000 | 6,733,000 | 6,653 | 2,885 | 3,576 |
| large-unicode-text-dollar | evaluate | 2,880 | 6,000 | 2,886 | 1,096 | 22 | 0 | 6,000 | 502,000 | 454 | 6 | 3,668 |
| large-unicode-text-dollar | parse-evaluate | 2,880 | 6,000 | 2,886 | 1,314 | 22 | 0 | 10,000 | 963,000 | 883 | 6 | 3,600 |
| large-unicode-text-exact | evaluate | 2,880 | 0 | 2,880 | 2,315 | 11,533 | 0 | 5,000 | 496,000 | 448 | 0 | 3,588 |
| large-unicode-text-exact | parse-evaluate | 2,880 | 0 | 2,880 | 4,196 | 11,533 | 0 | 9,000 | 6,717,000 | 6,637 | 0 | 3,664 |
| large-unicode-text-find | evaluate | 2,880 | 0 | 2,880 | 2,556 | 5,774 | 0 | 5,000 | 496,000 | 448 | 0 | 3,572 |
| large-unicode-text-find | parse-evaluate | 2,880 | 0 | 2,880 | 3,626 | 5,774 | 0 | 9,000 | 3,837,000 | 3,757 | 0 | 3,552 |
| large-unicode-text-fixed | evaluate | 2,880 | 5,000 | 2,885 | 1,227 | 26 | 0 | 7,000 | 853,000 | 629 | 5 | 3,644 |
| large-unicode-text-fixed | parse-evaluate | 2,880 | 5,000 | 2,885 | 1,489 | 26 | 0 | 11,000 | 1,320,000 | 1,064 | 5 | 3,544 |
| large-unicode-text-jis | evaluate | 2,880 | 4,800,000 | 7,680 | 33,828 | 12,488 | 0 | 5,000 | 5,200,000 | 5,200 | 4,800 | 3,496 |
| large-unicode-text-jis | parse-evaluate | 2,880 | 4,800,000 | 7,680 | 34,950 | 12,488 | 0 | 9,000 | 8,536,000 | 8,504 | 4,800 | 3,640 |
| large-unicode-text-left | evaluate | 2,880 | 2,000 | 2,882 | 2,450 | 5,773 | 0 | 5,000 | 496,000 | 448 | 0 | 3,524 |
| large-unicode-text-left | parse-evaluate | 2,880 | 2,000 | 2,882 | 3,528 | 5,773 | 0 | 9,000 | 3,835,000 | 3,755 | 0 | 3,528 |
| large-unicode-text-len | evaluate | 2,880 | 0 | 2,880 | 2,258 | 5,768 | 0 | 4,000 | 400,000 | 400 | 0 | 3,540 |
| large-unicode-text-len | parse-evaluate | 2,880 | 0 | 2,880 | 3,276 | 5,768 | 0 | 8,000 | 3,736,000 | 3,704 | 0 | 3,536 |
| large-unicode-text-lower | evaluate | 2,880 | 2,880,000 | 5,760 | 49,025 | 8,650 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,648 |
| large-unicode-text-lower | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 49,993 | 8,650 | 0 | 9,000 | 6,618,000 | 6,586 | 2,880 | 3,544 |
| large-unicode-text-mid | evaluate | 2,880 | 2,000 | 2,882 | 2,623 | 5,776 | 0 | 6,000 | 848,000 | 624 | 0 | 3,612 |
| large-unicode-text-mid | parse-evaluate | 2,880 | 2,000 | 2,882 | 3,690 | 5,776 | 0 | 10,000 | 4,188,000 | 3,932 | 0 | 3,524 |
| large-unicode-text-proper | evaluate | 2,880 | 2,880,000 | 5,760 | 53,571 | 8,651 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,564 |
| large-unicode-text-proper | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 54,944 | 8,651 | 0 | 9,000 | 6,619,000 | 6,587 | 2,880 | 3,668 |
| large-unicode-text-replace | evaluate | 2,880 | 2,880,000 | 5,760 | 5,591 | 8,665 | 0 | 8,000 | 3,920,000 | 3,600 | 2,880 | 3,596 |
| large-unicode-text-replace | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 6,761 | 8,665 | 0 | 13,000 | 8,036,000 | 7,300 | 2,880 | 3,668 |
| large-unicode-text-rept | evaluate | 2,880 | 5,760,000 | 8,640 | 6,792 | 11,533 | 0 | 6,000 | 6,256,000 | 6,208 | 5,760 | 3,496 |
| large-unicode-text-rept | parse-evaluate | 2,880 | 5,760,000 | 8,640 | 7,839 | 11,533 | 0 | 10,000 | 9,595,000 | 9,515 | 5,760 | 3,508 |
| large-unicode-text-right | evaluate | 2,880 | 2,000 | 2,882 | 3,534 | 5,774 | 0 | 5,000 | 496,000 | 448 | 0 | 3,592 |
| large-unicode-text-right | parse-evaluate | 2,880 | 2,000 | 2,882 | 4,615 | 5,774 | 0 | 9,000 | 3,836,000 | 3,756 | 0 | 3,540 |
| large-unicode-text-search | evaluate | 2,880 | 0 | 2,880 | 2,799 | 5,784 | 0 | 8,000 | 524,000 | 476 | 0 | 3,672 |
| large-unicode-text-search | parse-evaluate | 2,880 | 0 | 2,880 | 3,901 | 5,784 | 0 | 12,000 | 3,867,000 | 3,787 | 0 | 3,592 |
| large-unicode-text-substitute | evaluate | 2,880 | 2,880,000 | 5,760 | 14,379 | 8,665 | 0 | 7,000 | 3,728,000 | 3,504 | 2,880 | 3,500 |
| large-unicode-text-substitute | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 15,337 | 8,665 | 0 | 11,000 | 7,079,000 | 6,823 | 2,880 | 3,524 |
| large-unicode-text-t | evaluate | 2,880 | 2,880,000 | 5,760 | 3,838 | 5,766 | 0 | 4,000 | 400,000 | 400 | 0 | 3,560 |
| large-unicode-text-t | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 4,852 | 5,766 | 0 | 8,000 | 3,734,000 | 3,702 | 0 | 3,568 |
| large-unicode-text-text | evaluate | 2,880 | 5,000 | 2,885 | 1,422 | 25 | 0 | 6,000 | 501,000 | 453 | 5 | 3,640 |
| large-unicode-text-text | parse-evaluate | 2,880 | 5,000 | 2,885 | 1,646 | 25 | 0 | 10,000 | 965,000 | 885 | 5 | 3,556 |
| large-unicode-text-trim | evaluate | 2,880 | 2,879,000 | 5,759 | 13,242 | 8,648 | 0 | 5,000 | 3,279,000 | 3,279 | 2,879 | 3,616 |
| large-unicode-text-trim | parse-evaluate | 2,880 | 2,879,000 | 5,759 | 14,343 | 8,648 | 0 | 9,000 | 6,616,000 | 6,584 | 2,879 | 3,516 |
| large-unicode-text-unichar | evaluate | 2,880 | 1,000 | 2,881 | 624 | 15 | 0 | 5,000 | 401,000 | 401 | 1 | 3,588 |
| large-unicode-text-unichar | parse-evaluate | 2,880 | 1,000 | 2,881 | 826 | 15 | 0 | 9,000 | 861,000 | 829 | 1 | 3,540 |
| large-unicode-text-unicode | evaluate | 2,880 | 0 | 2,880 | 1,290 | 5,772 | 0 | 4,000 | 400,000 | 400 | 0 | 3,540 |
| large-unicode-text-unicode | parse-evaluate | 2,880 | 0 | 2,880 | 2,312 | 5,772 | 0 | 8,000 | 3,740,000 | 3,708 | 0 | 3,516 |
| large-unicode-text-upper | evaluate | 2,880 | 2,880,000 | 5,760 | 46,476 | 8,650 | 0 | 5,000 | 3,280,000 | 3,280 | 2,880 | 3,648 |
| large-unicode-text-upper | parse-evaluate | 2,880 | 2,880,000 | 5,760 | 47,515 | 8,650 | 0 | 9,000 | 6,618,000 | 6,586 | 2,880 | 3,540 |
| matrix-broadcast-text-concatenate | evaluate | 784 | 67,840 | 1,632 | 44,859 | 2,293 | 128 | 11,440 | 707,200 | 8,713 | 6,480 | 3,748 |
| matrix-broadcast-text-concatenate | parse-evaluate | 784 | 67,840 | 1,632 | 45,553 | 2,293 | 128 | 12,160 | 817,840 | 10,064 | 6,480 | 3,688 |
| matrix-broadcast-text-exact | evaluate | 784 | 0 | 784 | 34,172 | 1,503 | 128 | 6,320 | 639,360 | 7,865 | 5,632 | 3,740 |
| matrix-broadcast-text-exact | parse-evaluate | 784 | 0 | 784 | 35,050 | 1,503 | 128 | 7,040 | 749,520 | 9,210 | 5,632 | 3,784 |
| matrix-broadcast-text-find | evaluate | 784 | 0 | 784 | 22,305 | 1,309 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,732 |
| matrix-broadcast-text-find | parse-evaluate | 784 | 0 | 784 | 22,644 | 1,309 | 64 | 1,520 | 711,600 | 8,799 | 5,632 | 3,704 |
| matrix-broadcast-text-left | evaluate | 784 | 12,640 | 942 | 22,831 | 1,310 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,692 |
| matrix-broadcast-text-left | parse-evaluate | 784 | 12,640 | 942 | 23,559 | 1,310 | 128 | 1,920 | 744,320 | 9,208 | 5,632 | 3,740 |
| matrix-broadcast-text-len | evaluate | 784 | 0 | 784 | 15,503 | 1,113 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,688 |
| matrix-broadcast-text-len | parse-evaluate | 784 | 0 | 784 | 15,976 | 1,113 | 64 | 1,440 | 700,960 | 8,730 | 5,632 | 3,744 |
| matrix-broadcast-text-lower | evaluate | 784 | 62,720 | 1,568 | 33,099 | 1,483 | 64 | 3,440 | 621,440 | 7,768 | 6,000 | 3,744 |
| matrix-broadcast-text-lower | parse-evaluate | 784 | 62,720 | 1,568 | 33,665 | 1,483 | 64 | 4,000 | 730,560 | 9,100 | 6,000 | 3,776 |
| matrix-broadcast-text-mid | evaluate | 784 | 12,240 | 937 | 26,923 | 1,440 | 128 | 1,360 | 746,240 | 8,704 | 5,632 | 3,720 |
| matrix-broadcast-text-mid | parse-evaluate | 784 | 12,240 | 937 | 27,683 | 1,440 | 128 | 2,080 | 856,400 | 10,049 | 5,632 | 3,748 |
| matrix-broadcast-text-proper | evaluate | 784 | 62,720 | 1,568 | 40,843 | 1,676 | 64 | 4,720 | 636,800 | 7,960 | 6,192 | 3,776 |
| matrix-broadcast-text-proper | parse-evaluate | 784 | 62,720 | 1,568 | 41,425 | 1,676 | 64 | 5,280 | 746,000 | 9,293 | 6,192 | 3,820 |
| matrix-broadcast-text-replace | evaluate | 784 | 61,680 | 1,555 | 43,226 | 2,410 | 128 | 6,560 | 850,800 | 9,883 | 6,403 | 3,744 |
| matrix-broadcast-text-replace | parse-evaluate | 784 | 61,680 | 1,555 | 43,912 | 2,410 | 128 | 7,360 | 1,023,040 | 11,620 | 6,403 | 3,756 |
| matrix-broadcast-text-right | evaluate | 784 | 15,680 | 980 | 23,234 | 1,311 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,736 |
| matrix-broadcast-text-right | parse-evaluate | 784 | 15,680 | 980 | 23,779 | 1,311 | 128 | 1,920 | 744,400 | 9,209 | 5,632 | 3,772 |
| matrix-broadcast-text-search | evaluate | 784 | 0 | 784 | 40,473 | 1,799 | 64 | 16,320 | 745,600 | 7,492 | 5,632 | 3,716 |
| matrix-broadcast-text-search | parse-evaluate | 784 | 0 | 784 | 41,221 | 1,799 | 64 | 16,880 | 855,120 | 8,829 | 5,632 | 3,768 |
| matrix-broadcast-text-substitute | evaluate | 784 | 62,720 | 1,568 | 35,511 | 1,934 | 64 | 4,320 | 748,160 | 8,728 | 6,056 | 3,736 |
| matrix-broadcast-text-substitute | parse-evaluate | 784 | 62,720 | 1,568 | 36,039 | 1,934 | 64 | 4,880 | 858,320 | 10,073 | 6,056 | 3,776 |
| matrix-broadcast-text-trim | evaluate | 784 | 61,440 | 1,552 | 18,205 | 1,194 | 64 | 1,520 | 598,400 | 7,480 | 5,712 | 3,776 |
| matrix-broadcast-text-trim | parse-evaluate | 784 | 61,440 | 1,552 | 18,544 | 1,194 | 64 | 2,080 | 707,440 | 8,811 | 5,712 | 3,776 |
| reference-64-text-asc | evaluate | 784 | 58,880 | 1,520 | 19,772 | 1,241 | 64 | 1,520 | 597,120 | 7,464 | 5,696 | 3,740 |
| reference-64-text-asc | parse-evaluate | 784 | 58,880 | 1,520 | 20,114 | 1,241 | 64 | 2,080 | 706,080 | 8,794 | 5,696 | 3,684 |
| reference-64-text-char | evaluate | 784 | 5,120 | 848 | 24,591 | 394 | 64 | 6,000 | 597,120 | 7,464 | 5,696 | 3,684 |
| reference-64-text-char | parse-evaluate | 784 | 5,120 | 848 | 24,940 | 394 | 64 | 6,560 | 706,160 | 8,795 | 5,696 | 3,764 |
| reference-64-text-clean | evaluate | 784 | 62,720 | 1,568 | 19,869 | 1,115 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,800 |
| reference-64-text-clean | parse-evaluate | 784 | 62,720 | 1,568 | 20,055 | 1,115 | 64 | 1,440 | 701,120 | 8,732 | 5,632 | 3,800 |
| reference-64-text-code | evaluate | 784 | 0 | 784 | 15,393 | 1,114 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,688 |
| reference-64-text-code | parse-evaluate | 784 | 0 | 784 | 15,730 | 1,114 | 64 | 1,440 | 701,040 | 8,731 | 5,632 | 3,744 |
| reference-64-text-concatenate | evaluate | 784 | 88,320 | 1,888 | 32,048 | 2,680 | 64 | 6,080 | 690,560 | 8,568 | 6,736 | 3,692 |
| reference-64-text-concatenate | parse-evaluate | 784 | 88,320 | 1,888 | 32,539 | 2,680 | 64 | 6,640 | 800,800 | 9,914 | 6,736 | 3,744 |
| reference-64-text-dollar | evaluate | 784 | 25,600 | 1,104 | 44,412 | 719 | 64 | 6,080 | 627,840 | 7,784 | 5,952 | 3,776 |
| reference-64-text-dollar | parse-evaluate | 784 | 25,600 | 1,104 | 44,767 | 719 | 64 | 6,640 | 737,200 | 9,119 | 5,952 | 3,760 |
| reference-64-text-exact | evaluate | 1,568 | 0 | 1,568 | 21,947 | 2,095 | 128 | 1,200 | 634,240 | 7,864 | 5,632 | 3,780 |
| reference-64-text-exact | parse-evaluate | 1,568 | 0 | 1,568 | 22,393 | 2,095 | 128 | 1,920 | 744,400 | 9,209 | 5,632 | 3,736 |
| reference-64-text-find | evaluate | 784 | 0 | 784 | 22,227 | 1,309 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,780 |
| reference-64-text-find | parse-evaluate | 784 | 0 | 784 | 22,700 | 1,309 | 64 | 1,520 | 711,600 | 8,799 | 5,632 | 3,740 |
| reference-64-text-fixed | evaluate | 784 | 20,480 | 1,040 | 46,258 | 726 | 64 | 6,240 | 734,720 | 8,560 | 5,888 | 3,744 |
| reference-64-text-fixed | parse-evaluate | 784 | 20,480 | 1,040 | 46,815 | 726 | 64 | 6,800 | 844,560 | 9,901 | 5,888 | 3,760 |
| reference-64-text-jis | evaluate | 784 | 133,120 | 2,448 | 34,022 | 3,393 | 64 | 6,000 | 725,120 | 9,064 | 7,296 | 3,740 |
| reference-64-text-jis | parse-evaluate | 784 | 133,120 | 2,448 | 34,358 | 3,393 | 64 | 6,560 | 834,080 | 10,394 | 7,296 | 3,736 |
| reference-64-text-left | evaluate | 784 | 13,440 | 952 | 20,256 | 1,245 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,752 |
| reference-64-text-left | parse-evaluate | 784 | 13,440 | 952 | 20,493 | 1,245 | 64 | 1,520 | 711,440 | 8,797 | 5,632 | 3,708 |
| reference-64-text-len | evaluate | 784 | 0 | 784 | 15,540 | 1,113 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,784 |
| reference-64-text-len | parse-evaluate | 784 | 0 | 784 | 15,950 | 1,113 | 64 | 1,440 | 700,960 | 8,730 | 5,632 | 3,780 |
| reference-64-text-lower | evaluate | 784 | 62,720 | 1,568 | 33,154 | 1,483 | 64 | 3,440 | 621,440 | 7,768 | 6,000 | 3,780 |
| reference-64-text-lower | parse-evaluate | 784 | 62,720 | 1,568 | 33,638 | 1,483 | 64 | 4,000 | 730,560 | 9,100 | 6,000 | 3,800 |
| reference-64-text-mid | evaluate | 784 | 11,520 | 928 | 23,852 | 1,375 | 64 | 1,120 | 714,240 | 8,304 | 5,632 | 3,736 |
| reference-64-text-mid | parse-evaluate | 784 | 11,520 | 928 | 24,230 | 1,375 | 64 | 1,680 | 823,520 | 9,638 | 5,632 | 3,792 |
| reference-64-text-proper | evaluate | 784 | 62,720 | 1,568 | 40,976 | 1,676 | 64 | 4,720 | 636,800 | 7,960 | 6,192 | 3,756 |
| reference-64-text-proper | parse-evaluate | 784 | 62,720 | 1,568 | 41,584 | 1,676 | 64 | 5,280 | 746,000 | 9,293 | 6,192 | 3,752 |
| reference-64-text-replace | evaluate | 784 | 61,440 | 1,552 | 39,640 | 2,342 | 64 | 6,320 | 818,560 | 9,480 | 6,400 | 3,740 |
| reference-64-text-replace | parse-evaluate | 784 | 61,440 | 1,552 | 40,112 | 2,342 | 64 | 6,960 | 989,920 | 11,206 | 6,400 | 3,740 |
| reference-64-text-rept | evaluate | 784 | 125,440 | 2,352 | 28,567 | 2,813 | 64 | 6,080 | 727,680 | 9,032 | 7,200 | 3,684 |
| reference-64-text-rept | parse-evaluate | 784 | 125,440 | 2,352 | 29,080 | 2,813 | 64 | 6,640 | 836,880 | 10,365 | 7,200 | 3,792 |
| reference-64-text-right | evaluate | 784 | 16,000 | 984 | 20,684 | 1,246 | 64 | 960 | 602,240 | 7,464 | 5,632 | 3,716 |
| reference-64-text-right | parse-evaluate | 784 | 16,000 | 984 | 21,052 | 1,246 | 64 | 1,520 | 711,520 | 8,798 | 5,632 | 3,728 |
| reference-64-text-search | evaluate | 784 | 0 | 784 | 40,424 | 1,799 | 64 | 16,320 | 745,600 | 7,492 | 5,632 | 3,744 |
| reference-64-text-search | parse-evaluate | 784 | 0 | 784 | 41,254 | 1,799 | 64 | 16,880 | 855,120 | 8,829 | 5,632 | 3,768 |
| reference-64-text-substitute | evaluate | 784 | 62,720 | 1,568 | 35,387 | 1,934 | 64 | 4,320 | 748,160 | 8,728 | 6,056 | 3,736 |
| reference-64-text-substitute | parse-evaluate | 784 | 62,720 | 1,568 | 36,102 | 1,934 | 64 | 4,880 | 858,320 | 10,073 | 6,056 | 3,776 |
| reference-64-text-t | evaluate | 784 | 62,720 | 1,568 | 14,249 | 1,111 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,700 |
| reference-64-text-t | parse-evaluate | 784 | 62,720 | 1,568 | 14,597 | 1,111 | 64 | 1,440 | 700,800 | 8,728 | 5,632 | 3,732 |
| reference-64-text-text | evaluate | 784 | 20,480 | 1,040 | 61,105 | 848 | 64 | 6,080 | 622,720 | 7,720 | 5,888 | 3,772 |
| reference-64-text-text | parse-evaluate | 784 | 20,480 | 1,040 | 62,417 | 848 | 64 | 6,640 | 732,320 | 9,058 | 5,888 | 3,664 |
| reference-64-text-trim | evaluate | 784 | 61,440 | 1,552 | 18,217 | 1,194 | 64 | 1,520 | 598,400 | 7,480 | 5,712 | 3,716 |
| reference-64-text-trim | parse-evaluate | 784 | 61,440 | 1,552 | 18,541 | 1,194 | 64 | 2,080 | 707,440 | 8,811 | 5,712 | 3,772 |
| reference-64-text-unichar | evaluate | 784 | 5,120 | 848 | 24,543 | 397 | 64 | 6,000 | 597,120 | 7,464 | 5,696 | 3,740 |
| reference-64-text-unichar | parse-evaluate | 784 | 5,120 | 848 | 25,092 | 397 | 64 | 6,560 | 706,400 | 8,798 | 5,696 | 3,716 |
| reference-64-text-unicode | evaluate | 784 | 0 | 784 | 15,818 | 1,117 | 64 | 880 | 592,000 | 7,400 | 5,632 | 3,744 |
| reference-64-text-unicode | parse-evaluate | 784 | 0 | 784 | 16,214 | 1,117 | 64 | 1,440 | 701,280 | 8,734 | 5,632 | 3,788 |
| reference-64-text-upper | evaluate | 784 | 62,720 | 1,568 | 38,011 | 1,747 | 64 | 5,360 | 642,560 | 8,032 | 6,264 | 3,792 |
| reference-64-text-upper | parse-evaluate | 784 | 62,720 | 1,568 | 38,358 | 1,747 | 64 | 5,920 | 751,680 | 9,364 | 6,264 | 3,760 |
| refusal-text-asc | evaluate | 0 | 0 | 0 | 2,411 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,740 |
| refusal-text-asc | parse-evaluate | 0 | 0 | 0 | 2,953 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,720 |
| refusal-text-char | evaluate | 0 | 0 | 0 | 1,362 | 14 | 0 | 720 | 150,400 | 1,464 | 0 | 3,740 |
| refusal-text-char | parse-evaluate | 0 | 0 | 0 | 1,587 | 14 | 0 | 1,040 | 186,960 | 1,889 | 0 | 3,732 |
| refusal-text-clean | evaluate | 0 | 0 | 0 | 2,432 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,776 |
| refusal-text-clean | parse-evaluate | 0 | 0 | 0 | 3,000 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,744 |
| refusal-text-code | evaluate | 0 | 0 | 0 | 2,407 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,728 |
| refusal-text-code | parse-evaluate | 0 | 0 | 0 | 2,978 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,740 |
| refusal-text-concatenate | evaluate | 0 | 0 | 0 | 2,397 | 25 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,720 |
| refusal-text-concatenate | parse-evaluate | 0 | 0 | 0 | 2,995 | 25 | 0 | 1,920 | 325,040 | 2,927 | 0 | 3,708 |
| refusal-text-dollar | evaluate | 0 | 0 | 0 | 1,493 | 22 | 0 | 720 | 150,400 | 1,464 | 0 | 3,772 |
| refusal-text-dollar | parse-evaluate | 0 | 0 | 0 | 1,712 | 22 | 0 | 1,040 | 187,680 | 1,898 | 0 | 3,772 |
| refusal-text-exact | evaluate | 0 | 0 | 0 | 2,460 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,776 |
| refusal-text-exact | parse-evaluate | 0 | 0 | 0 | 3,023 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,760 |
| refusal-text-find | evaluate | 0 | 0 | 0 | 2,423 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,740 |
| refusal-text-find | parse-evaluate | 0 | 0 | 0 | 3,018 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,772 |
| refusal-text-fixed | evaluate | 0 | 0 | 0 | 1,438 | 21 | 0 | 720 | 150,400 | 1,464 | 0 | 3,736 |
| refusal-text-fixed | parse-evaluate | 0 | 0 | 0 | 1,648 | 21 | 0 | 1,040 | 187,600 | 1,897 | 0 | 3,740 |
| refusal-text-jis | evaluate | 0 | 0 | 0 | 2,500 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,748 |
| refusal-text-jis | parse-evaluate | 0 | 0 | 0 | 2,927 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,744 |
| refusal-text-left | evaluate | 0 | 0 | 0 | 2,499 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,720 |
| refusal-text-left | parse-evaluate | 0 | 0 | 0 | 2,976 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,796 |
| refusal-text-len | evaluate | 0 | 0 | 0 | 2,407 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,736 |
| refusal-text-len | parse-evaluate | 0 | 0 | 0 | 2,928 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,744 |
| refusal-text-lower | evaluate | 0 | 0 | 0 | 2,432 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,800 |
| refusal-text-lower | parse-evaluate | 0 | 0 | 0 | 3,027 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,744 |
| refusal-text-mid | evaluate | 0 | 0 | 0 | 2,445 | 17 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,720 |
| refusal-text-mid | parse-evaluate | 0 | 0 | 0 | 3,026 | 17 | 0 | 1,920 | 324,400 | 2,919 | 0 | 3,696 |
| refusal-text-proper | evaluate | 0 | 0 | 0 | 2,417 | 20 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,708 |
| refusal-text-proper | parse-evaluate | 0 | 0 | 0 | 2,956 | 20 | 0 | 1,920 | 324,640 | 2,922 | 0 | 3,732 |
| refusal-text-replace | evaluate | 0 | 0 | 0 | 2,435 | 21 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,744 |
| refusal-text-replace | parse-evaluate | 0 | 0 | 0 | 3,020 | 21 | 0 | 1,920 | 324,720 | 2,923 | 0 | 3,796 |
| refusal-text-rept | evaluate | 0 | 0 | 0 | 2,421 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,744 |
| refusal-text-rept | parse-evaluate | 0 | 0 | 0 | 3,002 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,728 |
| refusal-text-right | evaluate | 0 | 0 | 0 | 2,434 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,744 |
| refusal-text-right | parse-evaluate | 0 | 0 | 0 | 3,004 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,772 |
| refusal-text-search | evaluate | 0 | 0 | 0 | 2,444 | 20 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,772 |
| refusal-text-search | parse-evaluate | 0 | 0 | 0 | 2,996 | 20 | 0 | 1,920 | 324,640 | 2,922 | 0 | 3,668 |
| refusal-text-substitute | evaluate | 0 | 0 | 0 | 2,465 | 24 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,740 |
| refusal-text-substitute | parse-evaluate | 0 | 0 | 0 | 2,998 | 24 | 0 | 1,920 | 324,960 | 2,926 | 0 | 3,740 |
| refusal-text-t | evaluate | 0 | 0 | 0 | 2,379 | 15 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,744 |
| refusal-text-t | parse-evaluate | 0 | 0 | 0 | 2,910 | 15 | 0 | 1,920 | 324,240 | 2,917 | 0 | 3,724 |
| refusal-text-text | evaluate | 0 | 0 | 0 | 1,490 | 21 | 0 | 720 | 150,400 | 1,464 | 0 | 3,740 |
| refusal-text-text | parse-evaluate | 0 | 0 | 0 | 1,666 | 21 | 0 | 1,040 | 187,520 | 1,896 | 0 | 3,728 |
| refusal-text-trim | evaluate | 0 | 0 | 0 | 2,403 | 18 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,732 |
| refusal-text-trim | parse-evaluate | 0 | 0 | 0 | 2,954 | 18 | 0 | 1,920 | 324,480 | 2,920 | 0 | 3,740 |
| refusal-text-unichar | evaluate | 0 | 0 | 0 | 1,352 | 17 | 0 | 720 | 150,400 | 1,464 | 0 | 3,684 |
| refusal-text-unichar | parse-evaluate | 0 | 0 | 0 | 1,581 | 17 | 0 | 1,040 | 187,200 | 1,892 | 0 | 3,732 |
| refusal-text-unicode | evaluate | 0 | 0 | 0 | 2,437 | 21 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,688 |
| refusal-text-unicode | parse-evaluate | 0 | 0 | 0 | 2,952 | 21 | 0 | 1,920 | 324,720 | 2,923 | 0 | 3,744 |
| refusal-text-upper | evaluate | 0 | 0 | 0 | 2,445 | 19 | 0 | 1,200 | 214,400 | 1,576 | 0 | 3,776 |
| refusal-text-upper | parse-evaluate | 0 | 0 | 0 | 3,016 | 19 | 0 | 1,920 | 324,560 | 2,921 | 0 | 3,700 |
| rept-growth | evaluate | 8,192 | 8,192,000 | 16,384 | 26,443 | 8,210 | 0 | 6,000 | 8,688,000 | 8,640 | 8,192 | 3,652 |
| rept-growth | parse-evaluate | 8,192 | 8,192,000 | 16,384 | 26,668 | 8,210 | 0 | 10,000 | 9,151,000 | 9,071 | 8,192 | 3,584 |
| resource-text-concatenate | evaluate | 0 | 0 | 0 | 508 | 15 | 0 | 400 | 33,920 | 360 | 0 | 3,776 |
| resource-text-concatenate | parse-evaluate | 0 | 0 | 0 | 907 | 15 | 0 | 960 | 144,160 | 1,706 | 0 | 3,764 |
| resource-text-len | evaluate | 0 | 0 | 0 | 384 | 6 | 0 | 320 | 23,680 | 296 | 0 | 3,744 |
| resource-text-len | parse-evaluate | 0 | 0 | 0 | 738 | 6 | 0 | 880 | 132,640 | 1,626 | 0 | 3,768 |
| resource-text-search | evaluate | 0 | 0 | 0 | 653 | 12 | 0 | 480 | 64,640 | 744 | 0 | 3,704 |
| resource-text-search | parse-evaluate | 0 | 0 | 0 | 1,056 | 12 | 0 | 1,040 | 174,160 | 2,081 | 0 | 3,668 |
| resource-text-substitute | evaluate | 0 | 0 | 0 | 542 | 15 | 0 | 400 | 33,920 | 360 | 0 | 3,800 |
| resource-text-substitute | parse-evaluate | 0 | 0 | 0 | 934 | 15 | 0 | 960 | 144,080 | 1,705 | 0 | 3,664 |
| search-worstcase-exact | evaluate | 4,229 | 0 | 4,229 | 2,995 | 16,421 | 0 | 5,000 | 496,000 | 448 | 0 | 3,540 |
| search-worstcase-exact | parse-evaluate | 4,229 | 0 | 4,229 | 5,553 | 16,421 | 0 | 9,000 | 9,161,000 | 9,081 | 0 | 3,620 |
| search-worstcase-find | evaluate | 4,229 | 0 | 4,229 | 9,735 | 8,482 | 0 | 5,000 | 496,000 | 448 | 0 | 3,552 |
| search-worstcase-find | parse-evaluate | 4,229 | 0 | 4,229 | 11,092 | 8,482 | 0 | 9,000 | 5,191,000 | 5,111 | 0 | 3,500 |
| search-worstcase-search | evaluate | 4,229 | 0 | 4,229 | 97,353 | 17,079 | 0 | 8,000 | 4,220,000 | 4,172 | 0 | 3,592 |
| search-worstcase-search | parse-evaluate | 4,229 | 0 | 4,229 | 98,581 | 17,079 | 0 | 12,000 | 8,917,000 | 8,837 | 0 | 3,656 |
| search-worstcase-substitute | evaluate | 4,229 | 3,970,000 | 8,199 | 14,305 | 12,463 | 0 | 7,000 | 4,818,000 | 4,594 | 3,970 | 3,508 |
| search-worstcase-substitute | parse-evaluate | 4,229 | 3,970,000 | 8,199 | 15,756 | 12,463 | 0 | 11,000 | 9,523,000 | 9,267 | 3,970 | 3,596 |
| text-format-fraction-six | evaluate | 31 | 3 | 34 | 1,590 | 60 | 0 | 6 | 499 | 451 | 3 | 3,540 |
| text-format-fraction-six | parse-evaluate | 31 | 3 | 34 | 1,860 | 60 | 0 | 10 | 988 | 908 | 3 | 3,640 |
| tiny-text-asc | evaluate | 12 | 11,000 | 23 | 548 | 30 | 0 | 4,000 | 400,000 | 400 | 0 | 3,592 |
| tiny-text-asc | parse-evaluate | 12 | 11,000 | 23 | 735 | 30 | 0 | 8,000 | 867,000 | 835 | 0 | 3,616 |
| tiny-text-char | evaluate | 12 | 1,000 | 13 | 626 | 12 | 0 | 5,000 | 401,000 | 401 | 1 | 3,620 |
| tiny-text-char | parse-evaluate | 12 | 1,000 | 13 | 813 | 12 | 0 | 9,000 | 858,000 | 826 | 1 | 3,616 |
| tiny-text-clean | evaluate | 12 | 11,000 | 23 | 582 | 32 | 0 | 4,000 | 400,000 | 400 | 0 | 3,684 |
| tiny-text-clean | parse-evaluate | 12 | 11,000 | 23 | 789 | 32 | 0 | 8,000 | 869,000 | 837 | 0 | 3,652 |
| tiny-text-code | evaluate | 12 | 0 | 12 | 512 | 31 | 0 | 4,000 | 400,000 | 400 | 0 | 3,564 |
| tiny-text-code | parse-evaluate | 12 | 0 | 12 | 702 | 31 | 0 | 8,000 | 868,000 | 836 | 0 | 3,600 |
| tiny-text-concatenate | evaluate | 12 | 16,000 | 28 | 903 | 67 | 0 | 6,000 | 512,000 | 464 | 16 | 3,588 |
| tiny-text-concatenate | parse-evaluate | 12 | 16,000 | 28 | 1,134 | 67 | 0 | 10,000 | 995,000 | 915 | 16 | 3,540 |
| tiny-text-dollar | evaluate | 12 | 6,000 | 18 | 1,099 | 22 | 0 | 6,000 | 502,000 | 454 | 6 | 3,596 |
| tiny-text-dollar | parse-evaluate | 12 | 6,000 | 18 | 1,320 | 22 | 0 | 10,000 | 963,000 | 883 | 6 | 3,552 |
| tiny-text-exact | evaluate | 12 | 0 | 12 | 723 | 57 | 0 | 5,000 | 496,000 | 448 | 0 | 3,592 |
| tiny-text-exact | parse-evaluate | 12 | 0 | 12 | 967 | 57 | 0 | 9,000 | 979,000 | 899 | 0 | 3,552 |
| tiny-text-find | evaluate | 12 | 0 | 12 | 749 | 36 | 0 | 5,000 | 496,000 | 448 | 0 | 3,536 |
| tiny-text-find | parse-evaluate | 12 | 0 | 12 | 976 | 36 | 0 | 9,000 | 968,000 | 888 | 0 | 3,528 |
| tiny-text-fixed | evaluate | 12 | 5,000 | 17 | 1,234 | 26 | 0 | 7,000 | 853,000 | 629 | 5 | 3,620 |
| tiny-text-fixed | parse-evaluate | 12 | 5,000 | 17 | 1,484 | 26 | 0 | 11,000 | 1,320,000 | 1,064 | 5 | 3,660 |
| tiny-text-jis | evaluate | 12 | 27,000 | 39 | 796 | 67 | 0 | 5,000 | 427,000 | 427 | 27 | 3,556 |
| tiny-text-jis | parse-evaluate | 12 | 27,000 | 39 | 987 | 67 | 0 | 9,000 | 894,000 | 862 | 27 | 3,528 |
| tiny-text-left | evaluate | 12 | 2,000 | 14 | 707 | 35 | 0 | 5,000 | 496,000 | 448 | 0 | 3,552 |
| tiny-text-left | parse-evaluate | 12 | 2,000 | 14 | 932 | 35 | 0 | 9,000 | 966,000 | 886 | 0 | 3,540 |
| tiny-text-len | evaluate | 12 | 0 | 12 | 511 | 30 | 0 | 4,000 | 400,000 | 400 | 0 | 3,620 |
| tiny-text-len | parse-evaluate | 12 | 0 | 12 | 700 | 30 | 0 | 8,000 | 867,000 | 835 | 0 | 3,544 |
| tiny-text-lower | evaluate | 12 | 11,000 | 23 | 898 | 43 | 0 | 5,000 | 411,000 | 411 | 11 | 3,612 |
| tiny-text-lower | parse-evaluate | 12 | 11,000 | 23 | 1,120 | 43 | 0 | 9,000 | 880,000 | 848 | 11 | 3,672 |
| tiny-text-mid | evaluate | 12 | 2,000 | 14 | 866 | 38 | 0 | 6,000 | 848,000 | 624 | 0 | 3,588 |
| tiny-text-mid | parse-evaluate | 12 | 2,000 | 14 | 1,118 | 38 | 0 | 10,000 | 1,319,000 | 1,063 | 0 | 3,604 |
| tiny-text-proper | evaluate | 12 | 11,000 | 23 | 923 | 44 | 0 | 5,000 | 411,000 | 411 | 11 | 3,596 |
| tiny-text-proper | parse-evaluate | 12 | 11,000 | 23 | 1,145 | 44 | 0 | 9,000 | 881,000 | 849 | 11 | 3,628 |
| tiny-text-replace | evaluate | 12 | 11,000 | 23 | 1,205 | 58 | 0 | 8,000 | 1,051,000 | 731 | 11 | 3,616 |
| tiny-text-replace | parse-evaluate | 12 | 11,000 | 23 | 1,571 | 58 | 0 | 13,000 | 2,298,000 | 1,562 | 11 | 3,652 |
| tiny-text-rept | evaluate | 12 | 22,000 | 34 | 849 | 57 | 0 | 6,000 | 518,000 | 470 | 22 | 3,536 |
| tiny-text-rept | parse-evaluate | 12 | 22,000 | 34 | 1,075 | 57 | 0 | 10,000 | 988,000 | 908 | 22 | 3,624 |
| tiny-text-right | evaluate | 12 | 2,000 | 14 | 735 | 36 | 0 | 5,000 | 496,000 | 448 | 0 | 3,532 |
| tiny-text-right | parse-evaluate | 12 | 2,000 | 14 | 967 | 36 | 0 | 9,000 | 967,000 | 887 | 0 | 3,620 |
| tiny-text-search | evaluate | 12 | 0 | 12 | 1,043 | 46 | 0 | 8,000 | 524,000 | 476 | 0 | 3,580 |
| tiny-text-search | parse-evaluate | 12 | 0 | 12 | 1,267 | 46 | 0 | 12,000 | 998,000 | 918 | 0 | 3,668 |
| tiny-text-substitute | evaluate | 12 | 11,000 | 23 | 1,139 | 58 | 0 | 7,000 | 859,000 | 635 | 11 | 3,664 |
| tiny-text-substitute | parse-evaluate | 12 | 11,000 | 23 | 1,423 | 58 | 0 | 11,000 | 1,341,000 | 1,085 | 11 | 3,652 |
| tiny-text-t | evaluate | 12 | 11,000 | 23 | 494 | 28 | 0 | 4,000 | 400,000 | 400 | 0 | 3,500 |
| tiny-text-t | parse-evaluate | 12 | 11,000 | 23 | 675 | 28 | 0 | 8,000 | 865,000 | 833 | 0 | 3,592 |
| tiny-text-text | evaluate | 12 | 5,000 | 17 | 1,440 | 25 | 0 | 6,000 | 501,000 | 453 | 5 | 3,536 |
| tiny-text-text | parse-evaluate | 12 | 5,000 | 17 | 1,634 | 25 | 0 | 10,000 | 965,000 | 885 | 5 | 3,656 |
| tiny-text-trim | evaluate | 12 | 11,000 | 23 | 520 | 31 | 0 | 4,000 | 400,000 | 400 | 0 | 3,596 |
| tiny-text-trim | parse-evaluate | 12 | 11,000 | 23 | 719 | 31 | 0 | 8,000 | 868,000 | 836 | 0 | 3,576 |
| tiny-text-unichar | evaluate | 12 | 1,000 | 13 | 622 | 15 | 0 | 5,000 | 401,000 | 401 | 1 | 3,536 |
| tiny-text-unichar | parse-evaluate | 12 | 1,000 | 13 | 827 | 15 | 0 | 9,000 | 861,000 | 829 | 1 | 3,612 |
| tiny-text-unicode | evaluate | 12 | 0 | 12 | 517 | 34 | 0 | 4,000 | 400,000 | 400 | 0 | 3,548 |
| tiny-text-unicode | parse-evaluate | 12 | 0 | 12 | 715 | 34 | 0 | 8,000 | 871,000 | 839 | 0 | 3,516 |
| tiny-text-upper | evaluate | 12 | 11,000 | 23 | 894 | 43 | 0 | 5,000 | 411,000 | 411 | 11 | 3,540 |
| tiny-text-upper | parse-evaluate | 12 | 11,000 | 23 | 1,096 | 43 | 0 | 9,000 | 880,000 | 848 | 11 | 3,664 |

The resolver is an immutable borrowing fixture. Each child validates one typed text, number, or logical result (or typed failure) before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity.
