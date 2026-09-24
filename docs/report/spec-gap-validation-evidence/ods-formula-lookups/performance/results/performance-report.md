# ODS lookup evaluator performance profile

The baseline is committed `635fd2e1348b621426b50909cbd5765c91837306`. The profile uses three warmups and fifteen fresh child processes in both evaluator phases. It covers nine lookup functions in 88 candidate lanes, 33 matched controls (including SUMIFS and direct/projected ROWS/ISREF metadata), and reports exact elapsed_ns_p50 / repeat floats before cross-sample medians.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline bytes/repeat | candidate bytes/repeat | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 46,292.750 | 46,330.250 | +0.1% | 2,311 | 2,311 | 88 | 88 | 3,632 | 4,168 |
| array-control-16x16-arithmetic | parse-evaluate | 59,260.250 | 59,720.500 | +0.8% | 2,311 | 2,311 | 352 | 352 | 3,688 | 4,104 |
| array-control-16x16-sin | evaluate | 128,593.250 | 130,278.250 | +1.3% | 2,311 | 2,311 | 2,144 | 2,144 | 4,216 | 4,420 |
| array-control-16x16-sin | parse-evaluate | 140,768.000 | 143,930.750 | +2.2% | 2,311 | 2,311 | 2,412 | 2,412 | 4,224 | 4,468 |
| array-control-4x4-arithmetic | evaluate | 4,549.762 | 4,427.775 | -2.7% | 151 | 151 | 1,120 | 1,120 | 3,440 | 3,864 |
| array-control-4x4-arithmetic | parse-evaluate | 5,202.650 | 5,202.525 | -0.0% | 151 | 151 | 2,240 | 2,240 | 3,460 | 3,876 |
| array-control-4x4-sin | evaluate | 9,989.925 | 10,000.925 | +0.1% | 151 | 151 | 3,840 | 3,840 | 3,956 | 4,192 |
| array-control-4x4-sin | parse-evaluate | 10,849.938 | 10,830.550 | -0.2% | 151 | 151 | 5,040 | 5,040 | 4,144 | 4,448 |
| concat-borrowed-literals | evaluate | 715.424 | 718.274 | +0.4% | 21 | 21 | 6,000 | 6,000 | 3,408 | 3,652 |
| concat-borrowed-literals | parse-evaluate | 859.034 | 872.145 | +1.5% | 21 | 21 | 9,000 | 9,000 | 3,424 | 3,616 |
| concat-growth-chain | evaluate | 1,894.689 | 1,918.250 | +1.2% | 20 | 20 | 10,000 | 10,000 | 3,380 | 3,612 |
| concat-growth-chain | parse-evaluate | 2,264.961 | 2,298.532 | +1.5% | 20 | 20 | 17,000 | 17,000 | 3,396 | 3,640 |
| concat-owned-left | evaluate | 1,067.585 | 1,072.775 | +0.5% | 24 | 24 | 8,000 | 8,000 | 3,464 | 3,632 |
| concat-owned-left | parse-evaluate | 1,304.227 | 1,287.647 | -1.3% | 24 | 24 | 13,000 | 13,000 | 3,444 | 3,604 |
| concat-owned-right | evaluate | 1,075.976 | 1,087.505 | +1.1% | 24 | 24 | 8,000 | 8,000 | 3,472 | 3,640 |
| concat-owned-right | parse-evaluate | 1,295.957 | 1,311.527 | +1.2% | 24 | 24 | 13,000 | 13,000 | 3,412 | 3,644 |
| database-control-dstdev | evaluate | 3,980.000 | 3,920.000 | -1.5% | 30 | 30 | 20 | 20 | 3,696 | 3,920 |
| database-control-dstdev | parse-evaluate | 4,570.000 | 4,640.000 | +1.5% | 30 | 30 | 29 | 29 | 3,724 | 3,908 |
| database-control-dsum | evaluate | 3,960.000 | 4,050.000 | +2.3% | 28 | 28 | 20 | 20 | 3,760 | 3,896 |
| database-control-dsum | parse-evaluate | 4,650.000 | 4,680.000 | +0.6% | 28 | 28 | 29 | 29 | 3,780 | 3,884 |
| database-control-dvar | evaluate | 3,970.000 | 3,940.000 | -0.8% | 28 | 28 | 20 | 20 | 3,784 | 3,880 |
| database-control-dvar | parse-evaluate | 4,560.000 | 4,590.000 | +0.7% | 28 | 28 | 29 | 29 | 3,800 | 3,924 |
| literal-aggregate-4x1-sum | evaluate | 1,880.000 | 1,850.000 | -1.6% | 15 | 15 | 10 | 10 | 3,756 | 3,876 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,570.000 | 2,520.000 | -1.9% | 15 | 15 | 23 | 23 | 3,748 | 3,948 |
| reference-aggregate-64x4-sum | evaluate | 10,397.750 | 10,372.500 | -0.2% | 16 | 16 | 32 | 32 | 3,764 | 3,880 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,795.000 | 10,682.500 | -1.0% | 16 | 16 | 60 | 60 | 3,760 | 3,912 |
| reference-array-16x4-arithmetic | evaluate | 9,561.550 | 9,517.550 | -0.5% | 16 | 16 | 800 | 800 | 3,456 | 3,884 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,941.300 | 9,902.163 | -0.4% | 16 | 16 | 1,280 | 1,280 | 3,452 | 3,848 |
| reference-conditional-256x4-sumifs | evaluate | 84,480.500 | 95,125.500 | +12.6% | 52 | 52 | 42 | 42 | 3,724 | 3,904 |
| reference-conditional-256x4-sumifs | parse-evaluate | 88,865.500 | 87,565.500 | -1.5% | 52 | 52 | 68 | 68 | 3,776 | 3,896 |
| reference-control-average | evaluate | 10,425.000 | 10,412.500 | -0.1% | 20 | 20 | 32 | 32 | 3,692 | 3,880 |
| reference-control-average | parse-evaluate | 10,845.000 | 10,845.000 | +0.0% | 20 | 20 | 60 | 60 | 3,656 | 3,956 |
| reference-control-counta | evaluate | 8,745.000 | 8,752.500 | +0.1% | 19 | 19 | 32 | 32 | 3,648 | 3,928 |
| reference-control-counta | parse-evaluate | 9,165.000 | 9,152.500 | -0.1% | 19 | 19 | 60 | 60 | 3,628 | 3,920 |
| reference-control-isref-descriptor | evaluate | 1,600.000 | 1,580.000 | -1.2% | 18 | 18 | 10 | 10 | 3,720 | 3,940 |
| reference-control-isref-descriptor | parse-evaluate | 2,040.000 | 2,000.000 | -2.0% | 18 | 18 | 17 | 17 | 3,756 | 3,904 |
| reference-control-isref-projected | evaluate | 4,760.000 | 4,880.000 | +2.5% | 47 | 47 | 26 | 26 | 3,776 | 3,936 |
| reference-control-isref-projected | parse-evaluate | 5,580.000 | 5,810.000 | +4.1% | 47 | 47 | 38 | 38 | 3,752 | 3,936 |
| reference-control-row-projected | evaluate | 5,750.000 | 5,700.000 | -0.9% | 37 | 37 | 32 | 32 | 3,748 | 3,876 |
| reference-control-row-projected | parse-evaluate | 6,740.000 | 6,770.000 | +0.4% | 37 | 37 | 46 | 46 | 3,776 | 3,864 |
| reference-control-rows-descriptor | evaluate | 1,640.000 | 1,590.000 | -3.0% | 17 | 17 | 10 | 10 | 3,788 | 3,912 |
| reference-control-rows-descriptor | parse-evaluate | 2,040.000 | 2,070.000 | +1.5% | 17 | 17 | 17 | 17 | 3,780 | 3,908 |
| reference-control-rows-projected | evaluate | 4,990.000 | 4,980.000 | -0.2% | 40 | 40 | 29 | 29 | 3,668 | 3,920 |
| reference-control-rows-projected | parse-evaluate | 5,870.000 | 5,940.000 | +1.2% | 40 | 40 | 43 | 43 | 3,768 | 3,928 |
| representative-median | evaluate | 1,643.548 | 1,595.318 | -2.9% | 16 | 16 | 9,000 | 9,000 | 3,540 | 3,668 |
| representative-median | parse-evaluate | 1,886.519 | 1,874.220 | -0.7% | 16 | 16 | 14,000 | 14,000 | 3,536 | 3,668 |
| representative-percentrank | evaluate | 963.075 | 974.675 | +1.2% | 19 | 19 | 7,000 | 7,000 | 3,536 | 3,672 |
| representative-percentrank | parse-evaluate | 1,218.366 | 1,225.586 | +0.6% | 19 | 19 | 11,000 | 11,000 | 3,536 | 3,748 |
| representative-rank | evaluate | 925.635 | 929.714 | +0.4% | 12 | 12 | 7,000 | 7,000 | 3,548 | 3,700 |
| representative-rank | parse-evaluate | 1,141.365 | 1,150.926 | +0.8% | 12 | 12 | 11,000 | 11,000 | 3,544 | 3,696 |
| scalar-aggregate-sum | evaluate | 578.203 | 577.973 | -0.0% | 10 | 10 | 4,000 | 4,000 | 3,392 | 3,752 |
| scalar-aggregate-sum | parse-evaluate | 748.453 | 740.383 | -1.1% | 10 | 10 | 8,000 | 8,000 | 3,456 | 3,704 |
| scalar-control-arithmetic | evaluate | 571.192 | 570.003 | -0.2% | 10 | 10 | 5,000 | 5,000 | 3,452 | 3,664 |
| scalar-control-arithmetic | parse-evaluate | 703.333 | 711.473 | +1.2% | 10 | 10 | 8,000 | 8,000 | 3,516 | 3,636 |
| scalar-control-average | evaluate | 1,033.895 | 1,020.645 | -1.3% | 23 | 23 | 6,000 | 6,000 | 3,488 | 3,808 |
| scalar-control-average | parse-evaluate | 1,323.606 | 1,254.426 | -5.2% | 23 | 23 | 10,000 | 10,000 | 3,536 | 3,664 |
| scalar-control-counta | evaluate | 860.754 | 821.184 | -4.6% | 24 | 24 | 6,000 | 6,000 | 3,448 | 3,720 |
| scalar-control-counta | parse-evaluate | 1,115.365 | 1,076.705 | -3.5% | 24 | 24 | 10,000 | 10,000 | 3,420 | 3,712 |
| scalar-control-imsum | evaluate | 1,420.427 | 1,417.337 | -0.2% | 41 | 41 | 8,000 | 8,000 | 3,408 | 3,812 |
| scalar-control-imsum | parse-evaluate | 1,914.920 | 1,921.470 | +0.3% | 41 | 41 | 17,000 | 17,000 | 3,408 | 3,724 |
| scalar-control-sin | evaluate | 465.992 | 465.062 | -0.2% | 10 | 10 | 4,000 | 4,000 | 3,816 | 4,016 |
| scalar-control-sin | parse-evaluate | 652.824 | 649.853 | -0.5% | 10 | 10 | 8,000 | 8,000 | 3,832 | 4,020 |
| scalar-control-stdev | evaluate | 856.734 | 854.214 | -0.3% | 21 | 21 | 6,000 | 6,000 | 3,416 | 3,760 |
| scalar-control-stdev | parse-evaluate | 1,091.496 | 1,084.735 | -0.6% | 21 | 21 | 10,000 | 10,000 | 3,428 | 3,668 |
| scalar-control-var | evaluate | 858.815 | 843.524 | -1.8% | 19 | 19 | 6,000 | 6,000 | 3,416 | 3,780 |
| scalar-control-var | parse-evaluate | 1,092.615 | 1,053.355 | -3.6% | 19 | 19 | 10,000 | 10,000 | 3,520 | 3,704 |

## Lookup workloads

The candidate matrix covers ADDRESS A1/R1C1 formatting, lazy CHOOSE branches, exact linear and sorted approximate binary lookup scaling, selected-value-only reads, zero-read INDEX/OFFSET/INDIRECT descriptor probes, matrix output/storage, typed shape/resource/cancellation failures, lazy projected consumers, and projected dynamic INDIRECT descriptor scaling at 8/64/256 coordinates. Per-case read intervals are in case-matrix.json and are checked during preflight and retained verification.

| case | phase | input bytes | output bytes p50 | bytes/repeat | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| lookup-address-a1-absolute | evaluate | 22 | 128 | 26 | 1,874.062 | 30 | 0 | 288 | 81,792 | 1,980 | 4 | 3,920 |
| lookup-address-a1-absolute | parse-evaluate | 22 | 128 | 26 | 2,161.594 | 30 | 0 | 448 | 121,408 | 2,802 | 4 | 3,884 |
| lookup-address-a1-column-absolute | evaluate | 22 | 96 | 25 | 1,889.375 | 29 | 0 | 288 | 81,760 | 1,979 | 3 | 3,884 |
| lookup-address-a1-column-absolute | parse-evaluate | 22 | 96 | 25 | 2,154.375 | 29 | 0 | 448 | 121,376 | 2,801 | 3 | 3,892 |
| lookup-address-a1-relative | evaluate | 22 | 64 | 24 | 1,876.250 | 28 | 0 | 288 | 81,728 | 1,978 | 2 | 3,864 |
| lookup-address-a1-relative | parse-evaluate | 22 | 64 | 24 | 2,163.781 | 28 | 0 | 448 | 121,344 | 2,800 | 2 | 3,960 |
| lookup-address-a1-row-absolute | evaluate | 22 | 96 | 25 | 1,870.312 | 29 | 0 | 288 | 81,760 | 1,979 | 3 | 3,888 |
| lookup-address-a1-row-absolute | parse-evaluate | 22 | 96 | 25 | 2,170.312 | 29 | 0 | 448 | 121,376 | 2,801 | 3 | 3,868 |
| lookup-address-invalid | evaluate | 22 | 0 | 22 | 1,647.844 | 26 | 0 | 256 | 81,664 | 1,976 | 0 | 3,892 |
| lookup-address-invalid | parse-evaluate | 22 | 0 | 22 | 1,941.562 | 26 | 0 | 416 | 121,280 | 2,798 | 0 | 3,864 |
| lookup-address-r1c1-absolute | evaluate | 23 | 128 | 27 | 1,893.438 | 31 | 0 | 288 | 81,792 | 1,980 | 4 | 3,876 |
| lookup-address-r1c1-absolute | parse-evaluate | 23 | 128 | 27 | 2,198.156 | 31 | 0 | 448 | 121,440 | 2,803 | 4 | 3,912 |
| lookup-address-r1c1-relative | evaluate | 23 | 256 | 31 | 1,907.531 | 35 | 0 | 288 | 81,920 | 1,984 | 8 | 3,888 |
| lookup-address-r1c1-relative | parse-evaluate | 23 | 256 | 31 | 2,185.000 | 35 | 0 | 448 | 121,568 | 2,807 | 8 | 3,928 |
| lookup-address-sheet | evaluate | 29 | 288 | 38 | 2,118.438 | 45 | 0 | 320 | 137,248 | 2,945 | 9 | 3,896 |
| lookup-address-sheet | parse-evaluate | 29 | 288 | 38 | 2,453.750 | 45 | 0 | 512 | 179,392 | 3,782 | 9 | 3,888 |
| lookup-cancel-indirect | evaluate | 23 | 0 | 23 | 625.000 | 9 | 0 | 14 | 4,390 | 3,930 | 0 | 3,880 |
| lookup-cancel-indirect | parse-evaluate | 23 | 0 | 23 | 897.500 | 9 | 0 | 34 | 6,402 | 4,369 | 0 | 3,888 |
| lookup-cancel-vlookup | evaluate | 29 | 0 | 29 | 607.500 | 6 | 0 | 12 | 5,400 | 4,824 | 0 | 3,892 |
| lookup-cancel-vlookup | parse-evaluate | 29 | 0 | 29 | 1,215.000 | 6 | 0 | 44 | 13,972 | 6,551 | 0 | 3,884 |
| lookup-choose-first | evaluate | 19 | 0 | 19 | 586.562 | 14 | 0 | 128 | 20,224 | 632 | 0 | 3,768 |
| lookup-choose-first | parse-evaluate | 19 | 0 | 19 | 874.062 | 14 | 0 | 288 | 59,744 | 1,451 | 0 | 3,868 |
| lookup-choose-invalid-index | evaluate | 14 | 0 | 14 | 505.656 | 10 | 0 | 128 | 20,224 | 632 | 0 | 3,832 |
| lookup-choose-invalid-index | parse-evaluate | 14 | 0 | 14 | 720.312 | 10 | 0 | 256 | 35,008 | 1,062 | 0 | 3,712 |
| lookup-choose-lazy-if | evaluate | 46 | 0 | 46 | 678.781 | 14 | 0 | 128 | 20,224 | 632 | 0 | 3,876 |
| lookup-choose-lazy-if | parse-evaluate | 46 | 0 | 46 | 1,401.594 | 14 | 0 | 480 | 93,472 | 2,409 | 0 | 3,892 |
| lookup-choose-middle-lazy | evaluate | 51 | 0 | 51 | 594.688 | 14 | 0 | 128 | 20,224 | 632 | 0 | 3,860 |
| lookup-choose-middle-lazy | parse-evaluate | 51 | 0 | 51 | 1,397.219 | 14 | 0 | 544 | 93,728 | 2,417 | 0 | 3,824 |
| lookup-choose-position-sensitive-index | evaluate | 20 | 0 | 20 | 3,063.875 | 35 | 0 | 168 | 15,040 | 1,464 | 176 | 3,824 |
| lookup-choose-position-sensitive-index | parse-evaluate | 20 | 0 | 20 | 3,478.750 | 35 | 0 | 240 | 26,208 | 2,316 | 176 | 3,796 |
| lookup-choose-projected | evaluate | 39 | 0 | 39 | 5,918.750 | 64 | 0 | 280 | 33,088 | 2,408 | 176 | 3,936 |
| lookup-choose-projected | parse-evaluate | 39 | 0 | 39 | 6,571.250 | 64 | 0 | 376 | 57,976 | 4,111 | 176 | 3,952 |
| lookup-choose-selected-error | evaluate | 17 | 0 | 17 | 578.781 | 16 | 0 | 128 | 20,224 | 632 | 0 | 3,868 |
| lookup-choose-selected-error | parse-evaluate | 17 | 0 | 17 | 793.469 | 16 | 0 | 256 | 35,104 | 1,065 | 0 | 3,864 |
| lookup-choose-selected-index | evaluate | 38 | 0 | 38 | 1,848.156 | 31 | 1 | 288 | 49,408 | 1,480 | 0 | 3,900 |
| lookup-choose-selected-index | parse-evaluate | 38 | 0 | 38 | 2,548.156 | 31 | 1 | 640 | 171,488 | 4,015 | 0 | 3,928 |
| lookup-choose-selected-reference | evaluate | 33 | 0 | 33 | 1,553.750 | 20 | 1 | 288 | 49,408 | 1,480 | 0 | 3,920 |
| lookup-choose-selected-reference | parse-evaluate | 33 | 0 | 33 | 2,220.000 | 20 | 1 | 640 | 121,216 | 3,244 | 0 | 3,856 |
| lookup-hlookup-approx-256-mid | evaluate | 34 | 0 | 34 | 2,830.000 | 37 | 9 | 12 | 5,400 | 4,824 | 0 | 3,916 |
| lookup-hlookup-approx-256-mid | parse-evaluate | 34 | 0 | 34 | 4,010.000 | 37 | 9 | 20 | 7,549 | 6,557 | 0 | 3,924 |
| lookup-hlookup-approx-64-mid | evaluate | 33 | 0 | 33 | 2,532.500 | 34 | 7 | 48 | 21,600 | 4,824 | 0 | 3,900 |
| lookup-hlookup-approx-64-mid | parse-evaluate | 33 | 0 | 33 | 3,000.000 | 34 | 7 | 80 | 30,192 | 6,556 | 0 | 3,912 |
| lookup-hlookup-approx-8-mid | evaluate | 31 | 0 | 31 | 2,285.625 | 30 | 4 | 384 | 172,800 | 4,824 | 0 | 3,932 |
| lookup-hlookup-approx-8-mid | parse-evaluate | 31 | 0 | 31 | 2,769.094 | 30 | 4 | 640 | 241,440 | 6,553 | 0 | 3,864 |
| lookup-hlookup-exact-256-end | evaluate | 32 | 0 | 32 | 13,900.000 | 283 | 257 | 12 | 5,400 | 4,824 | 0 | 3,876 |
| lookup-hlookup-exact-256-end | parse-evaluate | 32 | 0 | 32 | 15,140.000 | 283 | 257 | 20 | 7,547 | 6,555 | 0 | 3,932 |
| lookup-hlookup-exact-64-end | evaluate | 31 | 0 | 31 | 5,167.500 | 90 | 65 | 48 | 21,600 | 4,824 | 0 | 3,912 |
| lookup-hlookup-exact-64-end | parse-evaluate | 31 | 0 | 31 | 5,655.000 | 90 | 65 | 80 | 30,184 | 6,554 | 0 | 3,932 |
| lookup-hlookup-exact-8-start | evaluate | 29 | 0 | 29 | 2,215.625 | 26 | 2 | 384 | 172,800 | 4,824 | 0 | 3,944 |
| lookup-hlookup-exact-8-start | parse-evaluate | 29 | 0 | 29 | 2,675.969 | 26 | 2 | 640 | 241,376 | 6,551 | 0 | 3,924 |
| lookup-hlookup-invalid-row | evaluate | 29 | 0 | 29 | 2,021.875 | 24 | 0 | 384 | 172,800 | 4,824 | 0 | 3,912 |
| lookup-hlookup-invalid-row | parse-evaluate | 29 | 0 | 29 | 2,495.625 | 24 | 0 | 640 | 241,376 | 6,551 | 0 | 3,912 |
| lookup-hlookup-miss | evaluate | 31 | 0 | 31 | 2,484.062 | 34 | 8 | 384 | 172,800 | 4,824 | 0 | 3,904 |
| lookup-hlookup-miss | parse-evaluate | 31 | 0 | 31 | 2,946.594 | 34 | 8 | 640 | 241,440 | 6,553 | 0 | 3,904 |
| lookup-index-consumer | evaluate | 27 | 0 | 27 | 2,883.781 | 34 | 1 | 512 | 191,744 | 5,032 | 0 | 3,912 |
| lookup-index-consumer | parse-evaluate | 27 | 0 | 27 | 3,481.594 | 34 | 1 | 800 | 261,280 | 6,757 | 0 | 3,920 |
| lookup-index-descriptor | evaluate | 29 | 0 | 29 | 2,751.875 | 37 | 0 | 512 | 191,744 | 5,032 | 0 | 3,888 |
| lookup-index-descriptor | parse-evaluate | 29 | 0 | 29 | 3,332.500 | 37 | 0 | 800 | 261,344 | 6,759 | 0 | 3,932 |
| lookup-index-descriptor-row | evaluate | 29 | 0 | 29 | 2,723.125 | 37 | 0 | 512 | 191,744 | 5,032 | 0 | 3,924 |
| lookup-index-descriptor-row | parse-evaluate | 29 | 0 | 29 | 3,315.656 | 37 | 0 | 800 | 261,344 | 6,759 | 0 | 3,888 |
| lookup-index-invalid | evaluate | 22 | 0 | 22 | 2,161.875 | 27 | 0 | 416 | 172,800 | 4,824 | 0 | 3,860 |
| lookup-index-invalid | parse-evaluate | 22 | 0 | 22 | 2,621.250 | 27 | 0 | 640 | 216,576 | 6,160 | 0 | 3,940 |
| lookup-index-lazy-consumer | evaluate | 39 | 0 | 39 | 3,158.156 | 45 | 1 | 512 | 191,744 | 5,032 | 0 | 3,912 |
| lookup-index-lazy-consumer | parse-evaluate | 39 | 0 | 39 | 3,854.719 | 45 | 1 | 864 | 264,736 | 6,801 | 0 | 3,880 |
| lookup-index-literal | evaluate | 20 | 0 | 20 | 1,977.844 | 30 | 0 | 384 | 163,328 | 4,480 | 0 | 3,884 |
| lookup-index-literal | parse-evaluate | 20 | 0 | 20 | 2,560.656 | 30 | 0 | 736 | 258,176 | 6,100 | 0 | 3,920 |
| lookup-index-row-consumer | evaluate | 26 | 0 | 26 | 2,985.312 | 37 | 4 | 512 | 191,744 | 5,032 | 0 | 3,936 |
| lookup-index-row-consumer | parse-evaluate | 26 | 0 | 26 | 3,578.469 | 37 | 4 | 800 | 261,248 | 6,756 | 0 | 3,904 |
| lookup-index-union-list | evaluate | 40 | 0 | 40 | 3,464.094 | 44 | 0 | 640 | 220,928 | 5,384 | 0 | 3,904 |
| lookup-index-union-list | parse-evaluate | 40 | 0 | 40 | 4,260.031 | 44 | 0 | 1,024 | 292,992 | 7,156 | 0 | 3,888 |
| lookup-indirect-consumer-a1 | evaluate | 20 | 0 | 20 | 2,285.000 | 28 | 1 | 416 | 140,224 | 3,929 | 0 | 3,884 |
| lookup-indirect-consumer-a1 | parse-evaluate | 20 | 0 | 20 | 2,568.125 | 28 | 1 | 576 | 156,224 | 4,365 | 0 | 3,872 |
| lookup-indirect-consumer-r1c1 | evaluate | 30 | 0 | 30 | 2,748.781 | 43 | 1 | 448 | 148,896 | 4,057 | 0 | 3,872 |
| lookup-indirect-consumer-r1c1 | parse-evaluate | 30 | 0 | 30 | 3,062.531 | 43 | 1 | 608 | 165,216 | 4,503 | 0 | 3,928 |
| lookup-indirect-consumer-r1c1-relative | evaluate | 34 | 0 | 34 | 2,763.156 | 51 | 1 | 448 | 149,024 | 4,057 | 0 | 3,880 |
| lookup-indirect-consumer-r1c1-relative | parse-evaluate | 34 | 0 | 34 | 3,088.125 | 51 | 1 | 608 | 165,472 | 4,507 | 0 | 3,868 |
| lookup-indirect-descriptor-a1 | evaluate | 22 | 0 | 22 | 2,262.531 | 31 | 0 | 448 | 148,416 | 4,057 | 0 | 3,932 |
| lookup-indirect-descriptor-a1 | parse-evaluate | 22 | 0 | 22 | 2,581.875 | 31 | 0 | 608 | 164,480 | 4,495 | 0 | 3,912 |
| lookup-indirect-descriptor-r1c1 | evaluate | 32 | 0 | 32 | 2,597.188 | 46 | 0 | 448 | 148,896 | 4,057 | 0 | 3,880 |
| lookup-indirect-descriptor-r1c1 | parse-evaluate | 32 | 0 | 32 | 2,944.094 | 46 | 0 | 608 | 165,280 | 4,505 | 0 | 3,932 |
| lookup-indirect-invalid-isref | evaluate | 29 | 0 | 29 | 1,865.312 | 43 | 0 | 352 | 137,376 | 3,717 | 0 | 3,912 |
| lookup-indirect-invalid-isref | parse-evaluate | 29 | 0 | 29 | 2,192.812 | 43 | 0 | 512 | 153,664 | 4,162 | 0 | 3,936 |
| lookup-indirect-lazy-if | evaluate | 40 | 0 | 40 | 681.875 | 14 | 0 | 128 | 20,224 | 632 | 0 | 3,920 |
| lookup-indirect-lazy-if | parse-evaluate | 40 | 0 | 40 | 1,214.375 | 14 | 0 | 384 | 64,512 | 1,504 | 0 | 3,920 |
| lookup-indirect-projected-dynamic-256 | evaluate | 1,821 | 0 | 1,821 | 834,372.875 | 13,835 | 512 | 124,608 | 45,153,024 | 275,336 | 0 | 4,424 |
| lookup-indirect-projected-dynamic-256 | parse-evaluate | 1,821 | 0 | 1,821 | 873,716.531 | 13,835 | 512 | 133,792 | 52,437,696 | 384,550 | 0 | 4,484 |
| lookup-indirect-projected-dynamic-64 | evaluate | 477 | 0 | 477 | 210,166.656 | 3,467 | 128 | 32,128 | 11,299,584 | 69,512 | 0 | 4,164 |
| lookup-indirect-projected-dynamic-64 | parse-evaluate | 477 | 0 | 477 | 218,498.281 | 3,467 | 128 | 34,976 | 13,134,528 | 97,510 | 0 | 4,176 |
| lookup-indirect-projected-dynamic-8 | evaluate | 85 | 0 | 85 | 29,425.781 | 443 | 16 | 4,768 | 1,425,664 | 12,169 | 0 | 3,928 |
| lookup-indirect-projected-dynamic-8 | parse-evaluate | 85 | 0 | 85 | 31,207.969 | 443 | 16 | 5,536 | 1,671,104 | 16,479 | 0 | 4,152 |
| lookup-indirect-range-sum | evaluate | 23 | 0 | 23 | 3,200.000 | 51 | 16 | 14 | 4,390 | 3,930 | 0 | 3,884 |
| lookup-indirect-range-sum | parse-evaluate | 23 | 0 | 23 | 3,480.000 | 51 | 16 | 19 | 4,893 | 4,369 | 0 | 3,876 |
| lookup-indirect-sheet-consumer | evaluate | 25 | 0 | 25 | 2,242.219 | 39 | 1 | 416 | 140,192 | 3,933 | 0 | 3,856 |
| lookup-indirect-sheet-consumer | parse-evaluate | 25 | 0 | 25 | 2,546.594 | 39 | 1 | 576 | 156,352 | 4,374 | 0 | 3,936 |
| lookup-lazy-index-projection | evaluate | 49 | 0 | 49 | 6,852.500 | 87 | 1 | 304 | 59,328 | 5,752 | 176 | 3,920 |
| lookup-lazy-index-projection | parse-evaluate | 49 | 0 | 49 | 7,823.750 | 87 | 1 | 432 | 91,736 | 8,363 | 176 | 3,956 |
| lookup-lookup-approx-256 | evaluate | 38 | 0 | 38 | 2,747.812 | 36 | 9 | 448 | 171,264 | 4,904 | 0 | 3,952 |
| lookup-lookup-approx-256 | parse-evaluate | 38 | 0 | 38 | 3,292.219 | 36 | 9 | 736 | 215,616 | 6,258 | 0 | 3,904 |
| lookup-lookup-approx-64 | evaluate | 35 | 0 | 35 | 2,646.906 | 33 | 7 | 448 | 171,264 | 4,904 | 0 | 3,852 |
| lookup-lookup-approx-64 | parse-evaluate | 35 | 0 | 35 | 3,184.062 | 33 | 7 | 736 | 215,520 | 6,255 | 0 | 3,888 |
| lookup-lookup-approx-8 | evaluate | 32 | 0 | 32 | 2,515.656 | 29 | 4 | 448 | 171,264 | 4,904 | 0 | 3,876 |
| lookup-lookup-approx-8 | parse-evaluate | 32 | 0 | 32 | 3,080.031 | 29 | 4 | 736 | 215,424 | 6,252 | 0 | 3,920 |
| lookup-lookup-approx-last | evaluate | 36 | 0 | 36 | 2,738.781 | 34 | 9 | 448 | 171,264 | 4,904 | 0 | 3,936 |
| lookup-lookup-approx-last | parse-evaluate | 36 | 0 | 36 | 3,266.906 | 34 | 9 | 736 | 215,552 | 6,256 | 0 | 3,872 |
| lookup-lookup-descending-result | evaluate | 32 | 0 | 32 | 2,560.938 | 30 | 5 | 448 | 171,264 | 4,904 | 0 | 3,892 |
| lookup-lookup-descending-result | parse-evaluate | 32 | 0 | 32 | 3,120.938 | 30 | 5 | 736 | 215,424 | 6,252 | 0 | 3,892 |
| lookup-lookup-lazy-if | evaluate | 45 | 0 | 45 | 2,821.906 | 40 | 4 | 448 | 171,264 | 4,904 | 0 | 3,872 |
| lookup-lookup-lazy-if | parse-evaluate | 45 | 0 | 45 | 3,534.719 | 40 | 4 | 832 | 243,488 | 6,681 | 0 | 3,956 |
| lookup-lookup-literal-text | evaluate | 37 | 0 | 37 | 2,659.719 | 52 | 0 | 448 | 222,976 | 5,528 | 0 | 3,896 |
| lookup-lookup-literal-text | parse-evaluate | 37 | 0 | 37 | 3,469.688 | 52 | 0 | 960 | 326,560 | 7,229 | 0 | 3,912 |
| lookup-lookup-short-result | evaluate | 21 | 0 | 21 | 2,165.312 | 32 | 0 | 416 | 168,192 | 4,584 | 0 | 3,948 |
| lookup-lookup-short-result | parse-evaluate | 21 | 0 | 21 | 2,863.469 | 32 | 0 | 832 | 268,192 | 6,269 | 0 | 3,952 |
| lookup-match-approx-256-mid | evaluate | 27 | 0 | 27 | 2,600.000 | 31 | 8 | 11 | 4,952 | 4,504 | 0 | 3,888 |
| lookup-match-approx-256-mid | parse-evaluate | 27 | 0 | 27 | 2,970.000 | 31 | 8 | 18 | 6,325 | 5,845 | 0 | 3,900 |
| lookup-match-approx-64-mid | evaluate | 25 | 0 | 25 | 2,205.000 | 28 | 6 | 44 | 19,808 | 4,504 | 0 | 3,888 |
| lookup-match-approx-64-mid | parse-evaluate | 25 | 0 | 25 | 2,815.000 | 28 | 6 | 72 | 25,292 | 5,843 | 0 | 3,956 |
| lookup-match-approx-8-mid | evaluate | 23 | 0 | 23 | 2,005.000 | 24 | 3 | 352 | 158,464 | 4,504 | 0 | 3,884 |
| lookup-match-approx-8-mid | parse-evaluate | 23 | 0 | 23 | 2,434.688 | 24 | 3 | 576 | 202,272 | 5,841 | 0 | 3,900 |
| lookup-match-descending | evaluate | 32 | 0 | 32 | 2,865.969 | 50 | 0 | 448 | 341,760 | 7,496 | 0 | 3,932 |
| lookup-match-descending | parse-evaluate | 32 | 0 | 32 | 3,870.656 | 50 | 0 | 1,088 | 554,752 | 10,856 | 0 | 3,928 |
| lookup-match-exact-256-end | evaluate | 25 | 0 | 25 | 13,740.000 | 277 | 256 | 11 | 4,952 | 4,504 | 0 | 3,884 |
| lookup-match-exact-256-end | parse-evaluate | 25 | 0 | 25 | 14,130.000 | 277 | 256 | 18 | 6,323 | 5,843 | 0 | 3,884 |
| lookup-match-exact-64-end | evaluate | 23 | 0 | 23 | 4,867.500 | 84 | 64 | 44 | 19,808 | 4,504 | 0 | 3,912 |
| lookup-match-exact-64-end | parse-evaluate | 23 | 0 | 23 | 5,402.500 | 84 | 64 | 72 | 25,284 | 5,841 | 0 | 3,880 |
| lookup-match-exact-8-end | evaluate | 21 | 0 | 21 | 2,241.875 | 27 | 8 | 352 | 158,464 | 4,504 | 0 | 3,912 |
| lookup-match-exact-8-end | parse-evaluate | 21 | 0 | 21 | 2,664.375 | 27 | 8 | 576 | 202,208 | 5,839 | 0 | 3,932 |
| lookup-match-exact-miss | evaluate | 23 | 0 | 23 | 2,230.312 | 29 | 8 | 352 | 158,464 | 4,504 | 0 | 3,952 |
| lookup-match-exact-miss | parse-evaluate | 23 | 0 | 23 | 2,685.938 | 29 | 8 | 576 | 202,272 | 5,841 | 0 | 3,884 |
| lookup-match-nested-munit-key | evaluate | 63 | 0 | 63 | 10,363.750 | 141 | 5 | 408 | 87,936 | 5,736 | 176 | 3,896 |
| lookup-match-nested-munit-key | parse-evaluate | 63 | 0 | 63 | 11,453.875 | 141 | 5 | 560 | 120,728 | 8,363 | 176 | 3,916 |
| lookup-match-position-sensitive-key | evaluate | 25 | 0 | 25 | 3,157.500 | 32 | 5 | 136 | 48,320 | 5,368 | 176 | 3,952 |
| lookup-match-position-sensitive-key | parse-evaluate | 25 | 0 | 25 | 3,822.500 | 32 | 5 | 232 | 66,712 | 7,123 | 176 | 3,928 |
| lookup-match-projected-invariant | evaluate | 43 | 0 | 43 | 5,530.000 | 70 | 2 | 248 | 53,088 | 5,480 | 176 | 3,912 |
| lookup-match-projected-invariant | parse-evaluate | 43 | 0 | 43 | 6,383.750 | 70 | 2 | 368 | 85,192 | 8,085 | 176 | 3,936 |
| lookup-offset-consumer | evaluate | 23 | 0 | 23 | 2,826.906 | 31 | 1 | 512 | 191,744 | 5,032 | 0 | 3,892 |
| lookup-offset-consumer | parse-evaluate | 23 | 0 | 23 | 3,320.938 | 31 | 1 | 768 | 261,120 | 6,752 | 0 | 3,928 |
| lookup-offset-descriptor | evaluate | 25 | 0 | 25 | 2,721.594 | 34 | 0 | 512 | 191,744 | 5,032 | 0 | 3,952 |
| lookup-offset-descriptor | parse-evaluate | 25 | 0 | 25 | 3,153.750 | 34 | 0 | 768 | 261,184 | 6,754 | 0 | 3,876 |
| lookup-offset-invalid | evaluate | 19 | 0 | 19 | 2,224.688 | 24 | 0 | 416 | 172,800 | 4,824 | 0 | 3,876 |
| lookup-offset-invalid | parse-evaluate | 19 | 0 | 19 | 2,648.125 | 24 | 0 | 640 | 241,024 | 6,540 | 0 | 3,932 |
| lookup-offset-large-descriptor | evaluate | 31 | 0 | 31 | 3,280.000 | 46 | 0 | 576 | 269,568 | 6,440 | 0 | 3,948 |
| lookup-offset-large-descriptor | parse-evaluate | 31 | 0 | 31 | 3,879.094 | 46 | 0 | 896 | 344,064 | 8,216 | 0 | 3,852 |
| lookup-offset-lazy-if | evaluate | 36 | 0 | 36 | 3,130.000 | 42 | 1 | 512 | 191,744 | 5,032 | 0 | 3,880 |
| lookup-offset-lazy-if | parse-evaluate | 36 | 0 | 36 | 3,738.469 | 42 | 1 | 832 | 264,608 | 6,797 | 0 | 3,876 |
| lookup-offset-range-sum | evaluate | 27 | 0 | 27 | 3,600.000 | 44 | 4 | 17 | 7,912 | 6,184 | 0 | 3,916 |
| lookup-offset-range-sum | parse-evaluate | 27 | 0 | 27 | 4,280.000 | 44 | 4 | 27 | 10,236 | 7,956 | 0 | 3,880 |
| lookup-offset-union-isref | evaluate | 33 | 0 | 33 | 2,938.438 | 38 | 0 | 544 | 208,128 | 4,984 | 0 | 3,864 |
| lookup-offset-union-isref | parse-evaluate | 33 | 0 | 33 | 3,609.094 | 38 | 0 | 864 | 279,904 | 6,747 | 0 | 3,880 |
| lookup-offset-zero-size | evaluate | 22 | 0 | 22 | 2,487.531 | 32 | 0 | 416 | 228,096 | 5,784 | 0 | 3,912 |
| lookup-offset-zero-size | parse-evaluate | 22 | 0 | 22 | 2,881.250 | 32 | 0 | 672 | 298,720 | 7,511 | 0 | 3,912 |
| lookup-output-index-row | evaluate | 23 | 0 | 23 | 3,435.000 | 38 | 4 | 136 | 50,496 | 5,032 | 352 | 3,904 |
| lookup-output-index-row | parse-evaluate | 23 | 0 | 23 | 3,942.500 | 38 | 4 | 208 | 68,104 | 6,785 | 352 | 3,912 |
| lookup-output-offset-matrix | evaluate | 24 | 0 | 24 | 3,961.250 | 45 | 4 | 152 | 69,952 | 6,440 | 352 | 3,936 |
| lookup-output-offset-matrix | parse-evaluate | 24 | 0 | 24 | 4,515.125 | 45 | 4 | 224 | 88,264 | 8,209 | 352 | 3,944 |
| lookup-resource-match | evaluate | 26 | 0 | 26 | 937.500 | 15 | 0 | 28 | 13,024 | 3,192 | 0 | 3,948 |
| lookup-resource-match | parse-evaluate | 26 | 0 | 26 | 1,375.000 | 15 | 0 | 56 | 18,516 | 4,533 | 0 | 3,860 |
| lookup-resource-vlookup | evaluate | 29 | 0 | 29 | 1,100.000 | 18 | 0 | 32 | 14,048 | 3,320 | 0 | 3,880 |
| lookup-resource-vlookup | parse-evaluate | 29 | 0 | 29 | 1,800.000 | 18 | 0 | 64 | 22,620 | 5,047 | 0 | 3,888 |
| lookup-shape-refusal-union | evaluate | 33 | 0 | 33 | 2,468.125 | 26 | 0 | 480 | 179,456 | 4,728 | 0 | 3,872 |
| lookup-shape-refusal-union | parse-evaluate | 33 | 0 | 33 | 3,077.188 | 26 | 0 | 832 | 250,272 | 6,493 | 0 | 3,876 |
| lookup-vlookup-approx-256-mid | evaluate | 31 | 0 | 31 | 2,820.000 | 37 | 9 | 12 | 5,400 | 4,824 | 0 | 3,880 |
| lookup-vlookup-approx-256-mid | parse-evaluate | 31 | 0 | 31 | 4,120.000 | 37 | 9 | 20 | 7,545 | 6,553 | 0 | 3,884 |
| lookup-vlookup-approx-64-mid | evaluate | 29 | 0 | 29 | 2,570.000 | 34 | 7 | 48 | 21,600 | 4,824 | 0 | 3,928 |
| lookup-vlookup-approx-64-mid | parse-evaluate | 29 | 0 | 29 | 3,040.000 | 34 | 7 | 80 | 30,172 | 6,551 | 0 | 3,864 |
| lookup-vlookup-approx-8-mid | evaluate | 27 | 0 | 27 | 2,301.594 | 30 | 4 | 384 | 172,800 | 4,824 | 0 | 3,956 |
| lookup-vlookup-approx-8-mid | parse-evaluate | 27 | 0 | 27 | 2,766.562 | 30 | 4 | 640 | 241,312 | 6,549 | 0 | 3,920 |
| lookup-vlookup-exact-256-end | evaluate | 29 | 0 | 29 | 13,960.000 | 283 | 257 | 12 | 5,400 | 4,824 | 0 | 3,892 |
| lookup-vlookup-exact-256-end | parse-evaluate | 29 | 0 | 29 | 15,250.000 | 283 | 257 | 20 | 7,543 | 6,551 | 0 | 3,888 |
| lookup-vlookup-exact-64-end | evaluate | 27 | 0 | 27 | 5,172.500 | 90 | 65 | 48 | 21,600 | 4,824 | 0 | 3,912 |
| lookup-vlookup-exact-64-end | parse-evaluate | 27 | 0 | 27 | 5,675.000 | 90 | 65 | 80 | 30,164 | 6,549 | 0 | 3,876 |
| lookup-vlookup-exact-8-start | evaluate | 25 | 0 | 25 | 2,196.281 | 26 | 2 | 384 | 172,800 | 4,824 | 0 | 3,924 |
| lookup-vlookup-exact-8-start | parse-evaluate | 25 | 0 | 25 | 2,670.031 | 26 | 2 | 640 | 241,248 | 6,547 | 0 | 3,912 |
| lookup-vlookup-exact-miss | evaluate | 27 | 0 | 27 | 2,490.312 | 34 | 8 | 384 | 172,800 | 4,824 | 0 | 3,956 |
| lookup-vlookup-exact-miss | parse-evaluate | 27 | 0 | 27 | 2,954.094 | 34 | 8 | 640 | 241,312 | 6,549 | 0 | 3,912 |
| lookup-vlookup-invalid-column | evaluate | 25 | 0 | 25 | 2,038.438 | 24 | 0 | 384 | 172,800 | 4,824 | 0 | 3,912 |
| lookup-vlookup-invalid-column | parse-evaluate | 25 | 0 | 25 | 2,494.406 | 24 | 0 | 640 | 241,248 | 6,547 | 0 | 3,892 |

The resolver is an immutable borrowing fixture. Each child validates exact text, number, logical, error, failure, and matrix coordinates before timing the evaluator and drop path. The profile does not measure save, recalculation, native producer acceptance, cold filesystem state, or cross-platform bit identity. retained-files.json is a flat relative-path to SHA256 map and excludes itself.
