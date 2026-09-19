# ODS dispersion evaluator performance profile

The baseline is committed `55e147bfa0676ce6ecdc609efc682b98568b8a5f`. It contributes matched arithmetic, SIN, IMSUM, DSUM, SUM, SUMIFS, AVERAGE, COUNTA, DVAR, and DSTDEV controls. The eight dispersion reducers are candidate-only evidence; each cell is the p50 across fifteen fresh child processes, with time, work, and resolver reads normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,917 | 46,147 | +0.5% | 88 | 88 | 3,460 | 3,388 |
| array-control-16x16-arithmetic | parse-evaluate | 59,585 | 58,085 | -2.5% | 352 | 352 | 3,440 | 3,424 |
| array-control-16x16-sin | evaluate | 126,540 | 126,555 | +0.0% | 2,144 | 2,144 | 3,596 | 3,620 |
| array-control-16x16-sin | parse-evaluate | 139,518 | 139,355 | -0.1% | 2,412 | 2,412 | 3,628 | 3,596 |
| array-control-4x4-arithmetic | evaluate | 4,338 | 4,276 | -1.4% | 1,120 | 1,120 | 3,168 | 3,168 |
| array-control-4x4-arithmetic | parse-evaluate | 5,217 | 5,121 | -1.8% | 2,240 | 2,240 | 3,076 | 3,160 |
| array-control-4x4-sin | evaluate | 9,717 | 9,689 | -0.3% | 3,840 | 3,840 | 3,352 | 3,452 |
| array-control-4x4-sin | parse-evaluate | 10,499 | 10,503 | +0.0% | 5,040 | 5,040 | 3,360 | 3,456 |
| database-control-dstdev | evaluate | 3,830 | 3,790 | -1.0% | 20 | 20 | 3,184 | 3,176 |
| database-control-dstdev | parse-evaluate | 4,490 | 4,400 | -2.0% | 29 | 29 | 3,152 | 3,172 |
| database-control-dsum | evaluate | 3,920 | 3,890 | -0.8% | 20 | 20 | 3,196 | 3,160 |
| database-control-dsum | parse-evaluate | 4,450 | 4,460 | +0.2% | 29 | 29 | 3,208 | 3,172 |
| database-control-dvar | evaluate | 3,820 | 3,800 | -0.5% | 20 | 20 | 3,168 | 3,148 |
| database-control-dvar | parse-evaluate | 4,380 | 4,350 | -0.7% | 29 | 29 | 3,172 | 3,176 |
| literal-aggregate-4x1-sum | evaluate | 1,760 | 1,770 | +0.6% | 10 | 10 | 3,040 | 3,156 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,470 | 2,390 | -3.2% | 23 | 23 | 3,036 | 3,164 |
| reference-aggregate-64x4-sum | evaluate | 10,205 | 10,237 | +0.3% | 32 | 32 | 3,160 | 3,144 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,580 | 10,615 | +0.3% | 60 | 60 | 3,196 | 3,148 |
| reference-array-16x4-arithmetic | evaluate | 9,201 | 9,348 | +1.6% | 800 | 800 | 3,184 | 3,164 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,684 | 9,628 | -0.6% | 1,280 | 1,280 | 3,172 | 3,164 |
| reference-conditional-256x4-sumifs | evaluate | 78,785 | 79,320 | +0.7% | 42 | 42 | 3,152 | 3,168 |
| reference-conditional-256x4-sumifs | parse-evaluate | 79,530 | 79,805 | +0.3% | 68 | 68 | 3,216 | 3,164 |
| reference-control-average | evaluate | 10,087 | 10,130 | +0.4% | 32 | 32 | 3,040 | 3,164 |
| reference-control-average | parse-evaluate | 10,520 | 10,600 | +0.8% | 60 | 60 | 3,172 | 3,172 |
| reference-control-counta | evaluate | 8,547 | 8,530 | -0.2% | 32 | 32 | 3,120 | 3,168 |
| reference-control-counta | parse-evaluate | 8,945 | 8,937 | -0.1% | 60 | 60 | 3,132 | 3,156 |
| scalar-aggregate-sum | evaluate | 569 | 574 | +0.9% | 4,000 | 4,000 | 2,992 | 3,092 |
| scalar-aggregate-sum | parse-evaluate | 751 | 765 | +1.9% | 8,000 | 8,000 | 2,980 | 2,972 |
| scalar-control-arithmetic | evaluate | 568 | 569 | +0.2% | 5,000 | 5,000 | 2,988 | 2,924 |
| scalar-control-arithmetic | parse-evaluate | 702 | 703 | +0.1% | 8,000 | 8,000 | 2,956 | 2,904 |
| scalar-control-imsum | evaluate | 1,373 | 1,405 | +2.3% | 8,000 | 8,000 | 3,000 | 2,976 |
| scalar-control-imsum | parse-evaluate | 1,919 | 1,938 | +1.0% | 17,000 | 17,000 | 3,020 | 2,944 |
| scalar-control-sin | evaluate | 456 | 461 | +1.1% | 4,000 | 4,000 | 3,288 | 3,252 |
| scalar-control-sin | parse-evaluate | 639 | 642 | +0.5% | 8,000 | 8,000 | 3,288 | 3,236 |

## Candidate dispersion reducer workloads

The ordinary rows cover scalar and inline-array inputs, rectangular references, ordered reference lists, one 3-D reference, mixed reference/scalar arguments, empty selections, formula errors, and typed reference-cell resource refusal. Nested rows use an outer array `IF` and a scalar dispersion reducer at 64 rows; VAR, VARA, and STDEVP add 256-row and 1024-row scaling rows. Resolver reads expose whether the reducer is projected once or rebuilt per output cell. VAR, VARP, and STDEVP list rows are intentional zero-read shape refusals; STDEV and the A variants admit ordered lists.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 3d-dispersion-stdev | evaluate | 1,580 | 19 | 3 | 9 | 1,640 | 1,528 | 0 | 3,168 |
| 3d-dispersion-stdev | parse-evaluate | 2,190 | 19 | 3 | 18 | 3,025 | 2,881 | 0 | 3,160 |
| 3d-dispersion-stdeva | evaluate | 1,640 | 20 | 3 | 9 | 1,640 | 1,528 | 0 | 3,160 |
| 3d-dispersion-stdeva | parse-evaluate | 2,230 | 20 | 3 | 18 | 3,026 | 2,882 | 0 | 3,156 |
| 3d-dispersion-stdevp | evaluate | 1,630 | 20 | 3 | 9 | 1,640 | 1,528 | 0 | 3,148 |
| 3d-dispersion-stdevp | parse-evaluate | 2,220 | 20 | 3 | 18 | 3,026 | 2,882 | 0 | 3,132 |
| 3d-dispersion-stdevpa | evaluate | 1,650 | 21 | 3 | 9 | 1,640 | 1,528 | 0 | 3,160 |
| 3d-dispersion-stdevpa | parse-evaluate | 2,270 | 21 | 3 | 18 | 3,027 | 2,883 | 0 | 3,108 |
| 3d-dispersion-var | evaluate | 1,590 | 17 | 3 | 9 | 1,640 | 1,528 | 0 | 3,180 |
| 3d-dispersion-var | parse-evaluate | 2,190 | 17 | 3 | 18 | 3,023 | 2,879 | 0 | 3,160 |
| 3d-dispersion-vara | evaluate | 1,610 | 18 | 3 | 9 | 1,640 | 1,528 | 0 | 3,168 |
| 3d-dispersion-vara | parse-evaluate | 2,180 | 18 | 3 | 18 | 3,024 | 2,880 | 0 | 3,160 |
| 3d-dispersion-varp | evaluate | 1,630 | 18 | 3 | 9 | 1,640 | 1,528 | 0 | 3,160 |
| 3d-dispersion-varp | parse-evaluate | 2,200 | 18 | 3 | 18 | 3,024 | 2,880 | 0 | 3,160 |
| 3d-dispersion-varpa | evaluate | 1,620 | 19 | 3 | 9 | 1,640 | 1,528 | 0 | 3,160 |
| 3d-dispersion-varpa | parse-evaluate | 2,170 | 19 | 3 | 18 | 3,025 | 2,881 | 0 | 3,176 |
| array-dispersion-stdev | evaluate | 1,900 | 29 | 0 | 11 | 5,000 | 4,376 | 0 | 3,108 |
| array-dispersion-stdev | parse-evaluate | 2,690 | 29 | 0 | 24 | 8,121 | 6,057 | 0 | 3,140 |
| array-dispersion-stdeva | evaluate | 2,040 | 36 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| array-dispersion-stdeva | parse-evaluate | 2,630 | 36 | 0 | 19 | 6,369 | 5,233 | 0 | 3,160 |
| array-dispersion-stdevp | evaluate | 1,920 | 30 | 0 | 11 | 5,000 | 4,376 | 0 | 3,164 |
| array-dispersion-stdevp | parse-evaluate | 2,730 | 30 | 0 | 24 | 8,122 | 6,058 | 0 | 3,180 |
| array-dispersion-stdevpa | evaluate | 2,080 | 37 | 0 | 11 | 5,000 | 4,376 | 0 | 3,164 |
| array-dispersion-stdevpa | parse-evaluate | 2,590 | 37 | 0 | 19 | 6,370 | 5,234 | 0 | 3,160 |
| array-dispersion-var | evaluate | 1,880 | 27 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| array-dispersion-var | parse-evaluate | 2,710 | 27 | 0 | 24 | 8,119 | 6,055 | 0 | 3,160 |
| array-dispersion-vara | evaluate | 2,060 | 34 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| array-dispersion-vara | parse-evaluate | 2,460 | 34 | 0 | 19 | 6,367 | 5,231 | 0 | 3,160 |
| array-dispersion-varp | evaluate | 1,930 | 28 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| array-dispersion-varp | parse-evaluate | 2,700 | 28 | 0 | 24 | 8,120 | 6,056 | 0 | 3,160 |
| array-dispersion-varpa | evaluate | 2,090 | 35 | 0 | 11 | 5,000 | 4,376 | 0 | 3,164 |
| array-dispersion-varpa | parse-evaluate | 2,450 | 35 | 0 | 19 | 6,368 | 5,232 | 0 | 3,136 |
| empty-dispersion-stdev | evaluate | 8,342 | 268 | 256 | 32 | 5,664 | 1,416 | 0 | 3,128 |
| empty-dispersion-stdev | parse-evaluate | 8,757 | 268 | 256 | 60 | 11,120 | 2,748 | 0 | 3,160 |
| empty-dispersion-stdeva | evaluate | 8,315 | 269 | 256 | 32 | 5,664 | 1,416 | 0 | 3,160 |
| empty-dispersion-stdeva | parse-evaluate | 8,797 | 269 | 256 | 60 | 11,124 | 2,749 | 0 | 3,160 |
| empty-dispersion-stdevp | evaluate | 8,322 | 269 | 256 | 32 | 5,664 | 1,416 | 0 | 3,168 |
| empty-dispersion-stdevp | parse-evaluate | 8,725 | 269 | 256 | 60 | 11,124 | 2,749 | 0 | 3,172 |
| empty-dispersion-stdevpa | evaluate | 8,432 | 270 | 256 | 32 | 5,664 | 1,416 | 0 | 3,188 |
| empty-dispersion-stdevpa | parse-evaluate | 8,810 | 270 | 256 | 60 | 11,128 | 2,750 | 0 | 3,172 |
| empty-dispersion-var | evaluate | 8,332 | 266 | 256 | 32 | 5,664 | 1,416 | 0 | 3,188 |
| empty-dispersion-var | parse-evaluate | 8,697 | 266 | 256 | 60 | 11,112 | 2,746 | 0 | 3,160 |
| empty-dispersion-vara | evaluate | 8,352 | 267 | 256 | 32 | 5,664 | 1,416 | 0 | 3,188 |
| empty-dispersion-vara | parse-evaluate | 8,755 | 267 | 256 | 60 | 11,116 | 2,747 | 0 | 3,184 |
| empty-dispersion-varp | evaluate | 8,325 | 267 | 256 | 32 | 5,664 | 1,416 | 0 | 3,172 |
| empty-dispersion-varp | parse-evaluate | 8,737 | 267 | 256 | 60 | 11,116 | 2,747 | 0 | 3,160 |
| empty-dispersion-varpa | evaluate | 8,315 | 268 | 256 | 32 | 5,664 | 1,416 | 0 | 3,172 |
| empty-dispersion-varpa | parse-evaluate | 8,735 | 268 | 256 | 60 | 11,120 | 2,748 | 0 | 3,156 |
| error-dispersion-stdev | evaluate | 3,584 | 76 | 64 | 640 | 113,280 | 1,416 | 0 | 3,212 |
| error-dispersion-stdev | parse-evaluate | 3,972 | 76 | 64 | 1,200 | 222,400 | 2,748 | 0 | 3,172 |
| error-dispersion-stdeva | evaluate | 3,574 | 77 | 64 | 640 | 113,280 | 1,416 | 0 | 3,148 |
| error-dispersion-stdeva | parse-evaluate | 3,980 | 77 | 64 | 1,200 | 222,480 | 2,749 | 0 | 3,168 |
| error-dispersion-stdevp | evaluate | 3,584 | 77 | 64 | 640 | 113,280 | 1,416 | 0 | 3,160 |
| error-dispersion-stdevp | parse-evaluate | 3,989 | 77 | 64 | 1,200 | 222,480 | 2,749 | 0 | 3,176 |
| error-dispersion-stdevpa | evaluate | 3,599 | 78 | 64 | 640 | 113,280 | 1,416 | 0 | 3,164 |
| error-dispersion-stdevpa | parse-evaluate | 4,054 | 78 | 64 | 1,200 | 222,560 | 2,750 | 0 | 3,180 |
| error-dispersion-var | evaluate | 3,586 | 74 | 64 | 640 | 113,280 | 1,416 | 0 | 3,160 |
| error-dispersion-var | parse-evaluate | 3,974 | 74 | 64 | 1,200 | 222,240 | 2,746 | 0 | 3,168 |
| error-dispersion-vara | evaluate | 3,580 | 75 | 64 | 640 | 113,280 | 1,416 | 0 | 3,128 |
| error-dispersion-vara | parse-evaluate | 3,977 | 75 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,172 |
| error-dispersion-varp | evaluate | 3,582 | 75 | 64 | 640 | 113,280 | 1,416 | 0 | 3,148 |
| error-dispersion-varp | parse-evaluate | 3,970 | 75 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,160 |
| error-dispersion-varpa | evaluate | 3,576 | 76 | 64 | 640 | 113,280 | 1,416 | 0 | 3,164 |
| error-dispersion-varpa | parse-evaluate | 3,978 | 76 | 64 | 1,200 | 222,400 | 2,748 | 0 | 3,184 |
| list-dispersion-stdev | evaluate | 11,950 | 274 | 256 | 48 | 7,776 | 1,576 | 0 | 3,176 |
| list-dispersion-stdev | parse-evaluate | 12,517 | 274 | 256 | 84 | 13,288 | 2,922 | 0 | 3,184 |
| list-dispersion-stdeva | evaluate | 11,532 | 403 | 256 | 48 | 7,776 | 1,576 | 0 | 3,160 |
| list-dispersion-stdeva | parse-evaluate | 12,087 | 403 | 256 | 84 | 13,292 | 2,923 | 0 | 3,196 |
| list-dispersion-stdevp | evaluate | 1,827 | 19 | 0 | 48 | 7,776 | 1,576 | 0 | 3,148 |
| list-dispersion-stdevp | parse-evaluate | 2,402 | 19 | 0 | 84 | 13,292 | 2,923 | 0 | 3,184 |
| list-dispersion-stdevpa | evaluate | 11,565 | 404 | 256 | 48 | 7,776 | 1,576 | 0 | 3,164 |
| list-dispersion-stdevpa | parse-evaluate | 12,130 | 404 | 256 | 84 | 13,296 | 2,924 | 0 | 3,172 |
| list-dispersion-var | evaluate | 1,807 | 16 | 0 | 48 | 7,776 | 1,576 | 0 | 3,132 |
| list-dispersion-var | parse-evaluate | 2,377 | 16 | 0 | 84 | 13,280 | 2,920 | 0 | 3,168 |
| list-dispersion-vara | evaluate | 11,520 | 401 | 256 | 48 | 7,776 | 1,576 | 0 | 3,160 |
| list-dispersion-vara | parse-evaluate | 12,120 | 401 | 256 | 84 | 13,284 | 2,921 | 0 | 3,172 |
| list-dispersion-varp | evaluate | 1,810 | 17 | 0 | 48 | 7,776 | 1,576 | 0 | 3,164 |
| list-dispersion-varp | parse-evaluate | 2,362 | 17 | 0 | 84 | 13,284 | 2,921 | 0 | 3,164 |
| list-dispersion-varpa | evaluate | 11,545 | 402 | 256 | 48 | 7,776 | 1,576 | 0 | 3,184 |
| list-dispersion-varpa | parse-evaluate | 12,122 | 402 | 256 | 84 | 13,288 | 2,922 | 0 | 3,160 |
| mixed-dispersion-stdev | evaluate | 3,657 | 89 | 64 | 800 | 200,320 | 2,056 | 0 | 3,168 |
| mixed-dispersion-stdev | parse-evaluate | 4,111 | 89 | 64 | 1,360 | 310,160 | 3,397 | 0 | 3,160 |
| mixed-dispersion-stdeva | evaluate | 4,011 | 122 | 64 | 800 | 200,320 | 2,056 | 0 | 3,172 |
| mixed-dispersion-stdeva | parse-evaluate | 4,416 | 122 | 64 | 1,360 | 310,240 | 3,398 | 0 | 3,188 |
| mixed-dispersion-stdevp | evaluate | 3,652 | 90 | 64 | 800 | 200,320 | 2,056 | 0 | 3,172 |
| mixed-dispersion-stdevp | parse-evaluate | 4,056 | 90 | 64 | 1,360 | 310,240 | 3,398 | 0 | 3,164 |
| mixed-dispersion-stdevpa | evaluate | 3,986 | 123 | 64 | 800 | 200,320 | 2,056 | 0 | 3,164 |
| mixed-dispersion-stdevpa | parse-evaluate | 4,407 | 123 | 64 | 1,360 | 310,320 | 3,399 | 0 | 3,164 |
| mixed-dispersion-var | evaluate | 3,648 | 87 | 64 | 800 | 200,320 | 2,056 | 0 | 3,164 |
| mixed-dispersion-var | parse-evaluate | 4,046 | 87 | 64 | 1,360 | 310,000 | 3,395 | 0 | 3,156 |
| mixed-dispersion-vara | evaluate | 3,985 | 120 | 64 | 800 | 200,320 | 2,056 | 0 | 3,160 |
| mixed-dispersion-vara | parse-evaluate | 4,389 | 120 | 64 | 1,360 | 310,080 | 3,396 | 0 | 3,164 |
| mixed-dispersion-varp | evaluate | 3,668 | 88 | 64 | 800 | 200,320 | 2,056 | 0 | 3,160 |
| mixed-dispersion-varp | parse-evaluate | 4,091 | 88 | 64 | 1,360 | 310,080 | 3,396 | 0 | 3,172 |
| mixed-dispersion-varpa | evaluate | 3,983 | 121 | 64 | 800 | 200,320 | 2,056 | 0 | 3,172 |
| mixed-dispersion-varpa | parse-evaluate | 4,396 | 121 | 64 | 1,360 | 310,160 | 3,397 | 0 | 3,168 |
| nested-projected-dispersion-1024-stdevp | evaluate | 1,032,295 | 24,616 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,196 |
| nested-projected-dispersion-1024-stdevp | parse-evaluate | 1,033,265 | 24,616 | 8,192 | 4,174 | 1,061,993 | 724,553 | 0 | 4,228 |
| nested-projected-dispersion-1024-var | evaluate | 1,020,535 | 24,613 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,184 |
| nested-projected-dispersion-1024-var | parse-evaluate | 1,035,695 | 24,613 | 8,192 | 4,174 | 1,061,990 | 724,550 | 0 | 4,204 |
| nested-projected-dispersion-1024-vara | evaluate | 1,018,355 | 26,662 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,232 |
| nested-projected-dispersion-1024-vara | parse-evaluate | 1,025,325 | 26,662 | 8,192 | 4,174 | 1,061,991 | 724,551 | 0 | 4,232 |
| nested-projected-dispersion-256-stdevp | evaluate | 270,456 | 6,184 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,416 |
| nested-projected-dispersion-256-stdevp | parse-evaluate | 271,506 | 6,184 | 2,048 | 2,188 | 538,830 | 183,879 | 0 | 3,404 |
| nested-projected-dispersion-256-var | evaluate | 271,221 | 6,181 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,460 |
| nested-projected-dispersion-256-var | parse-evaluate | 271,301 | 6,181 | 2,048 | 2,188 | 538,824 | 183,876 | 0 | 3,432 |
| nested-projected-dispersion-256-vara | evaluate | 270,431 | 6,694 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,452 |
| nested-projected-dispersion-256-vara | parse-evaluate | 271,926 | 6,694 | 2,048 | 2,188 | 538,826 | 183,877 | 0 | 3,440 |
| nested-projected-dispersion-64-stdev | evaluate | 71,195 | 1,575 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,448 |
| nested-projected-dispersion-64-stdev | parse-evaluate | 71,630 | 1,575 | 512 | 1,272 | 285,072 | 48,708 | 0 | 3,448 |
| nested-projected-dispersion-64-stdeva | evaluate | 70,252 | 1,704 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,412 |
| nested-projected-dispersion-64-stdeva | parse-evaluate | 71,607 | 1,704 | 512 | 1,272 | 285,076 | 48,709 | 0 | 3,444 |
| nested-projected-dispersion-64-stdevp | evaluate | 70,535 | 1,576 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,440 |
| nested-projected-dispersion-64-stdevp | parse-evaluate | 71,225 | 1,576 | 512 | 1,272 | 285,076 | 48,709 | 0 | 3,428 |
| nested-projected-dispersion-64-stdevpa | evaluate | 70,357 | 1,705 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,448 |
| nested-projected-dispersion-64-stdevpa | parse-evaluate | 71,030 | 1,705 | 512 | 1,272 | 285,080 | 48,710 | 0 | 3,416 |
| nested-projected-dispersion-64-var | evaluate | 70,600 | 1,573 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,420 |
| nested-projected-dispersion-64-var | parse-evaluate | 72,125 | 1,573 | 512 | 1,272 | 285,064 | 48,706 | 0 | 3,448 |
| nested-projected-dispersion-64-vara | evaluate | 70,240 | 1,702 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,448 |
| nested-projected-dispersion-64-vara | parse-evaluate | 70,998 | 1,702 | 512 | 1,272 | 285,068 | 48,707 | 0 | 3,416 |
| nested-projected-dispersion-64-varp | evaluate | 70,837 | 1,574 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,424 |
| nested-projected-dispersion-64-varp | parse-evaluate | 71,437 | 1,574 | 512 | 1,272 | 285,068 | 48,707 | 0 | 3,464 |
| nested-projected-dispersion-64-varpa | evaluate | 70,065 | 1,703 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,432 |
| nested-projected-dispersion-64-varpa | parse-evaluate | 71,150 | 1,703 | 512 | 1,272 | 285,072 | 48,708 | 0 | 3,416 |
| reference-dispersion-stdev | evaluate | 41,565 | 1,036 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,172 |
| reference-dispersion-stdev | parse-evaluate | 41,945 | 1,036 | 1,024 | 30 | 5,562 | 2,749 | 0 | 3,144 |
| reference-dispersion-stdeva | evaluate | 39,730 | 1,549 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,160 |
| reference-dispersion-stdeva | parse-evaluate | 40,305 | 1,549 | 1,024 | 30 | 5,564 | 2,750 | 0 | 3,168 |
| reference-dispersion-stdevp | evaluate | 41,605 | 1,037 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,176 |
| reference-dispersion-stdevp | parse-evaluate | 41,970 | 1,037 | 1,024 | 30 | 5,564 | 2,750 | 0 | 3,188 |
| reference-dispersion-stdevpa | evaluate | 39,805 | 1,550 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,160 |
| reference-dispersion-stdevpa | parse-evaluate | 40,460 | 1,550 | 1,024 | 30 | 5,566 | 2,751 | 0 | 3,160 |
| reference-dispersion-var | evaluate | 41,460 | 1,034 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,168 |
| reference-dispersion-var | parse-evaluate | 42,025 | 1,034 | 1,024 | 30 | 5,558 | 2,747 | 0 | 3,160 |
| reference-dispersion-vara | evaluate | 39,875 | 1,547 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,196 |
| reference-dispersion-vara | parse-evaluate | 40,195 | 1,547 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,160 |
| reference-dispersion-varp | evaluate | 41,485 | 1,035 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,160 |
| reference-dispersion-varp | parse-evaluate | 41,980 | 1,035 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,172 |
| reference-dispersion-varpa | evaluate | 39,830 | 1,548 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,172 |
| reference-dispersion-varpa | parse-evaluate | 40,165 | 1,548 | 1,024 | 30 | 5,562 | 2,749 | 0 | 3,196 |
| resource-dispersion-stdev | evaluate | 392 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,192 |
| resource-dispersion-stdev | parse-evaluate | 790 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,168 |
| resource-dispersion-stdeva | evaluate | 390 | 9 | 0 | 16 | 1,184 | 296 | 0 | 3,168 |
| resource-dispersion-stdeva | parse-evaluate | 785 | 9 | 0 | 44 | 6,644 | 1,629 | 0 | 3,160 |
| resource-dispersion-stdevp | evaluate | 397 | 9 | 0 | 16 | 1,184 | 296 | 0 | 3,180 |
| resource-dispersion-stdevp | parse-evaluate | 787 | 9 | 0 | 44 | 6,644 | 1,629 | 0 | 3,180 |
| resource-dispersion-stdevpa | evaluate | 392 | 10 | 0 | 16 | 1,184 | 296 | 0 | 3,168 |
| resource-dispersion-stdevpa | parse-evaluate | 792 | 10 | 0 | 44 | 6,648 | 1,630 | 0 | 3,196 |
| resource-dispersion-var | evaluate | 390 | 6 | 0 | 16 | 1,184 | 296 | 0 | 3,196 |
| resource-dispersion-var | parse-evaluate | 770 | 6 | 0 | 44 | 6,632 | 1,626 | 0 | 3,188 |
| resource-dispersion-vara | evaluate | 395 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,184 |
| resource-dispersion-vara | parse-evaluate | 782 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,188 |
| resource-dispersion-varp | evaluate | 395 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,164 |
| resource-dispersion-varp | parse-evaluate | 795 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,168 |
| resource-dispersion-varpa | evaluate | 392 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,108 |
| resource-dispersion-varpa | parse-evaluate | 787 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,164 |
| scalar-dispersion-stdev | evaluate | 677 | 23 | 0 | 5,000 | 496,000 | 448 | 0 | 3,072 |
| scalar-dispersion-stdev | parse-evaluate | 918 | 23 | 0 | 9,000 | 960,000 | 880 | 0 | 3,100 |
| scalar-dispersion-stdeva | evaluate | 671 | 24 | 0 | 5,000 | 496,000 | 448 | 0 | 2,976 |
| scalar-dispersion-stdeva | parse-evaluate | 903 | 24 | 0 | 9,000 | 961,000 | 881 | 0 | 2,980 |
| scalar-dispersion-stdevp | evaluate | 682 | 24 | 0 | 5,000 | 496,000 | 448 | 0 | 2,928 |
| scalar-dispersion-stdevp | parse-evaluate | 911 | 24 | 0 | 9,000 | 961,000 | 881 | 0 | 2,968 |
| scalar-dispersion-stdevpa | evaluate | 677 | 25 | 0 | 5,000 | 496,000 | 448 | 0 | 3,024 |
| scalar-dispersion-stdevpa | parse-evaluate | 903 | 25 | 0 | 9,000 | 962,000 | 882 | 0 | 3,032 |
| scalar-dispersion-var | evaluate | 662 | 21 | 0 | 5,000 | 496,000 | 448 | 0 | 3,052 |
| scalar-dispersion-var | parse-evaluate | 882 | 21 | 0 | 9,000 | 958,000 | 878 | 0 | 2,972 |
| scalar-dispersion-vara | evaluate | 669 | 22 | 0 | 5,000 | 496,000 | 448 | 0 | 2,980 |
| scalar-dispersion-vara | parse-evaluate | 889 | 22 | 0 | 9,000 | 959,000 | 879 | 0 | 2,976 |
| scalar-dispersion-varp | evaluate | 667 | 22 | 0 | 5,000 | 496,000 | 448 | 0 | 2,956 |
| scalar-dispersion-varp | parse-evaluate | 889 | 22 | 0 | 9,000 | 959,000 | 879 | 0 | 2,940 |
| scalar-dispersion-varpa | evaluate | 681 | 23 | 0 | 5,000 | 496,000 | 448 | 0 | 3,072 |
| scalar-dispersion-varpa | parse-evaluate | 913 | 23 | 0 | 9,000 | 960,000 | 880 | 0 | 2,980 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed finite result or formula error before timing. The profile does not measure save, recalculation, cache publication, native producer acceptance, cold filesystem state, wildcard/regular-expression host profiles, or cross-platform bit identity.

## Matched-control threshold review

No matched-control p50 metric exceeded ±5% between baseline and candidate across time, allocator calls, requested or released bytes, peak live bytes, result-live budget, evaluator work, resolver reads, or RSS. Candidate-only dispersion rows have no baseline comparison; their absolute measurements are retained above. The machine-readable review is in `threshold-review.json`.
