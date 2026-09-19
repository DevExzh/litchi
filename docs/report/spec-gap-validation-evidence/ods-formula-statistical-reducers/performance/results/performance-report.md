# ODS statistical reducer evaluator performance profile

The baseline is committed `f7fe857007b7bcb65e0512c5b5caf9135fe74a5a`. It contributes matched arithmetic, SIN, IMSUM, DSUM, SUM, and SUMIFS controls. The nine statistical reducers are candidate-only evidence; the baseline has no corresponding valid path for those rows. Each cell is the p50 across fifteen fresh child processes; time, work, and resolver reads are normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,880 | 45,937 | +0.1% | 88 | 88 | 3,320 | 3,392 |
| array-control-16x16-arithmetic | parse-evaluate | 58,420 | 58,907 | +0.8% | 352 | 352 | 3,320 | 3,416 |
| array-control-16x16-sin | evaluate | 124,808 | 126,975 | +1.7% | 2,144 | 2,144 | 3,604 | 3,628 |
| array-control-16x16-sin | parse-evaluate | 138,408 | 139,050 | +0.5% | 2,412 | 2,412 | 3,628 | 3,640 |
| array-control-4x4-arithmetic | evaluate | 4,345 | 4,371 | +0.6% | 1,120 | 1,120 | 3,116 | 3,068 |
| array-control-4x4-arithmetic | parse-evaluate | 5,192 | 5,265 | +1.4% | 2,240 | 2,240 | 3,060 | 3,128 |
| array-control-4x4-sin | evaluate | 9,607 | 9,718 | +1.2% | 3,840 | 3,840 | 3,340 | 3,348 |
| array-control-4x4-sin | parse-evaluate | 10,463 | 10,556 | +0.9% | 5,040 | 5,040 | 3,368 | 3,348 |
| database-control-dsum | evaluate | 3,930 | 3,820 | -2.8% | 20 | 20 | 3,104 | 3,168 |
| database-control-dsum | parse-evaluate | 4,460 | 4,470 | +0.2% | 29 | 29 | 3,064 | 3,184 |
| literal-aggregate-4x1-sum | evaluate | 1,780 | 1,760 | -1.1% | 10 | 10 | 3,100 | 3,168 |
| literal-aggregate-4x1-sum | parse-evaluate | 2,540 | 2,500 | -1.6% | 23 | 23 | 3,056 | 3,120 |
| reference-aggregate-64x4-sum | evaluate | 10,297 | 10,145 | -1.5% | 32 | 32 | 3,088 | 3,172 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,642 | 10,540 | -1.0% | 60 | 60 | 3,056 | 3,164 |
| reference-array-16x4-arithmetic | evaluate | 9,162 | 9,300 | +1.5% | 800 | 800 | 3,064 | 3,144 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,473 | 9,643 | +1.8% | 1,280 | 1,280 | 3,064 | 3,144 |
| reference-conditional-256x4-sumifs | evaluate | 78,225 | 78,515 | +0.4% | 42 | 42 | 3,056 | 3,172 |
| reference-conditional-256x4-sumifs | parse-evaluate | 79,355 | 79,680 | +0.4% | 68 | 68 | 3,172 | 3,156 |
| scalar-aggregate-sum | evaluate | 595 | 578 | -2.9% | 4,000 | 4,000 | 2,952 | 2,904 |
| scalar-aggregate-sum | parse-evaluate | 773 | 742 | -4.0% | 8,000 | 8,000 | 3,032 | 2,908 |
| scalar-control-arithmetic | evaluate | 559 | 559 | +0.0% | 5,000 | 5,000 | 3,020 | 2,908 |
| scalar-control-arithmetic | parse-evaluate | 690 | 703 | +1.9% | 8,000 | 8,000 | 2,924 | 2,920 |
| scalar-control-imsum | evaluate | 1,423 | 1,375 | -3.4% | 8,000 | 8,000 | 2,968 | 2,908 |
| scalar-control-imsum | parse-evaluate | 2,046 | 1,935 | -5.4% | 17,000 | 17,000 | 2,952 | 2,924 |
| scalar-control-sin | evaluate | 470 | 458 | -2.6% | 4,000 | 4,000 | 3,188 | 3,208 |
| scalar-control-sin | parse-evaluate | 661 | 637 | -3.6% | 8,000 | 8,000 | 3,204 | 3,232 |

## Candidate statistical reducer workloads

The ordinary rows cover scalar and inline-array inputs, rectangular references, ordered reference lists, one 3-D reference, empty selections, formula errors, and typed reference-cell resource refusal. All nine reducers have a 64-row nested scalar projection; `AVERAGE`, `COUNTA`, and `COUNTBLANK` also have 256-row and 1024-row nested projections. Their resolver reads expose whether the reducer is projected once or rebuilt per output cell.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 3d-statistical-average | evaluate | 1,730 | 21 | 3 | 9 | 1,640 | 1,528 | 0 | 3,172 |
| 3d-statistical-average | parse-evaluate | 2,360 | 21 | 3 | 18 | 3,027 | 2,883 | 0 | 3,156 |
| 3d-statistical-averagea | evaluate | 1,780 | 22 | 3 | 9 | 1,640 | 1,528 | 0 | 3,140 |
| 3d-statistical-averagea | parse-evaluate | 2,350 | 22 | 3 | 18 | 3,028 | 2,884 | 0 | 3,164 |
| 3d-statistical-count | evaluate | 1,580 | 19 | 3 | 9 | 1,640 | 1,528 | 0 | 3,168 |
| 3d-statistical-count | parse-evaluate | 2,150 | 19 | 3 | 18 | 3,025 | 2,881 | 0 | 3,156 |
| 3d-statistical-counta | evaluate | 1,550 | 20 | 3 | 9 | 1,640 | 1,528 | 0 | 3,160 |
| 3d-statistical-counta | parse-evaluate | 2,161 | 20 | 3 | 18 | 3,026 | 2,882 | 0 | 3,180 |
| 3d-statistical-countblank | evaluate | 1,580 | 23 | 3 | 9 | 1,640 | 1,528 | 0 | 3,192 |
| 3d-statistical-countblank | parse-evaluate | 2,170 | 23 | 3 | 18 | 3,030 | 2,886 | 0 | 3,176 |
| 3d-statistical-max | evaluate | 1,550 | 17 | 3 | 9 | 1,640 | 1,528 | 0 | 3,164 |
| 3d-statistical-max | parse-evaluate | 2,180 | 17 | 3 | 18 | 3,023 | 2,879 | 0 | 3,152 |
| 3d-statistical-maxa | evaluate | 1,570 | 18 | 3 | 9 | 1,640 | 1,528 | 0 | 3,176 |
| 3d-statistical-maxa | parse-evaluate | 2,110 | 18 | 3 | 18 | 3,024 | 2,880 | 0 | 3,164 |
| 3d-statistical-min | evaluate | 1,550 | 17 | 3 | 9 | 1,640 | 1,528 | 0 | 3,136 |
| 3d-statistical-min | parse-evaluate | 2,100 | 17 | 3 | 18 | 3,023 | 2,879 | 0 | 3,172 |
| 3d-statistical-mina | evaluate | 1,590 | 18 | 3 | 9 | 1,640 | 1,528 | 0 | 3,152 |
| 3d-statistical-mina | parse-evaluate | 2,100 | 18 | 3 | 18 | 3,024 | 2,880 | 0 | 3,184 |
| empty-statistical-average | evaluate | 8,312 | 270 | 256 | 32 | 5,664 | 1,416 | 0 | 3,108 |
| empty-statistical-average | parse-evaluate | 8,692 | 270 | 256 | 60 | 11,128 | 2,750 | 0 | 3,168 |
| empty-statistical-averagea | evaluate | 8,285 | 271 | 256 | 32 | 5,664 | 1,416 | 0 | 3,164 |
| empty-statistical-averagea | parse-evaluate | 8,830 | 271 | 256 | 60 | 11,132 | 2,751 | 0 | 3,192 |
| empty-statistical-count | evaluate | 8,300 | 268 | 256 | 32 | 5,664 | 1,416 | 0 | 3,164 |
| empty-statistical-count | parse-evaluate | 8,685 | 268 | 256 | 60 | 11,120 | 2,748 | 0 | 3,164 |
| empty-statistical-counta | evaluate | 8,270 | 269 | 256 | 32 | 5,664 | 1,416 | 0 | 3,156 |
| empty-statistical-counta | parse-evaluate | 8,690 | 269 | 256 | 60 | 11,124 | 2,749 | 0 | 3,184 |
| empty-statistical-countblank | evaluate | 8,537 | 272 | 256 | 32 | 5,664 | 1,416 | 0 | 3,160 |
| empty-statistical-countblank | parse-evaluate | 8,890 | 272 | 256 | 60 | 11,140 | 2,753 | 0 | 3,168 |
| empty-statistical-max | evaluate | 8,365 | 266 | 256 | 32 | 5,664 | 1,416 | 0 | 3,188 |
| empty-statistical-max | parse-evaluate | 8,692 | 266 | 256 | 60 | 11,112 | 2,746 | 0 | 3,160 |
| empty-statistical-maxa | evaluate | 8,327 | 267 | 256 | 32 | 5,664 | 1,416 | 0 | 3,160 |
| empty-statistical-maxa | parse-evaluate | 8,792 | 267 | 256 | 60 | 11,116 | 2,747 | 0 | 3,168 |
| empty-statistical-min | evaluate | 8,402 | 266 | 256 | 32 | 5,664 | 1,416 | 0 | 3,164 |
| empty-statistical-min | parse-evaluate | 8,695 | 266 | 256 | 60 | 11,112 | 2,746 | 0 | 3,168 |
| empty-statistical-mina | evaluate | 8,312 | 267 | 256 | 32 | 5,664 | 1,416 | 0 | 3,168 |
| empty-statistical-mina | parse-evaluate | 8,647 | 267 | 256 | 60 | 11,116 | 2,747 | 0 | 3,176 |
| error-statistical-average | evaluate | 3,334 | 78 | 64 | 640 | 113,280 | 1,416 | 0 | 3,164 |
| error-statistical-average | parse-evaluate | 3,726 | 78 | 64 | 1,200 | 222,560 | 2,750 | 0 | 3,160 |
| error-statistical-averagea | evaluate | 3,326 | 79 | 64 | 640 | 113,280 | 1,416 | 0 | 3,180 |
| error-statistical-averagea | parse-evaluate | 3,707 | 79 | 64 | 1,200 | 222,640 | 2,751 | 0 | 3,144 |
| error-statistical-count | evaluate | 3,148 | 76 | 64 | 640 | 113,280 | 1,416 | 0 | 3,196 |
| error-statistical-count | parse-evaluate | 3,483 | 76 | 64 | 1,200 | 222,400 | 2,748 | 0 | 3,168 |
| error-statistical-counta | evaluate | 2,945 | 77 | 64 | 640 | 113,280 | 1,416 | 0 | 3,188 |
| error-statistical-counta | parse-evaluate | 3,348 | 77 | 64 | 1,200 | 222,480 | 2,749 | 0 | 3,160 |
| error-statistical-countblank | evaluate | 2,927 | 80 | 64 | 640 | 113,280 | 1,416 | 0 | 3,180 |
| error-statistical-countblank | parse-evaluate | 3,331 | 80 | 64 | 1,200 | 222,800 | 2,753 | 0 | 3,180 |
| error-statistical-max | evaluate | 3,113 | 74 | 64 | 640 | 113,280 | 1,416 | 0 | 3,184 |
| error-statistical-max | parse-evaluate | 3,515 | 74 | 64 | 1,200 | 222,240 | 2,746 | 0 | 3,184 |
| error-statistical-maxa | evaluate | 3,113 | 75 | 64 | 640 | 113,280 | 1,416 | 0 | 3,164 |
| error-statistical-maxa | parse-evaluate | 3,489 | 75 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,172 |
| error-statistical-min | evaluate | 3,147 | 74 | 64 | 640 | 113,280 | 1,416 | 0 | 3,188 |
| error-statistical-min | parse-evaluate | 3,517 | 74 | 64 | 1,200 | 222,240 | 2,746 | 0 | 3,172 |
| error-statistical-mina | evaluate | 3,115 | 75 | 64 | 640 | 113,280 | 1,416 | 0 | 3,196 |
| error-statistical-mina | parse-evaluate | 3,499 | 75 | 64 | 1,200 | 222,320 | 2,747 | 0 | 3,164 |
| list-statistical-average | evaluate | 1,845 | 20 | 0 | 48 | 7,776 | 1,576 | 0 | 3,148 |
| list-statistical-average | parse-evaluate | 2,435 | 20 | 0 | 84 | 13,296 | 2,924 | 0 | 3,168 |
| list-statistical-averagea | evaluate | 11,045 | 405 | 256 | 48 | 7,776 | 1,576 | 0 | 3,164 |
| list-statistical-averagea | parse-evaluate | 11,605 | 405 | 256 | 84 | 13,300 | 2,925 | 0 | 3,156 |
| list-statistical-count | evaluate | 9,865 | 274 | 256 | 48 | 7,776 | 1,576 | 0 | 3,160 |
| list-statistical-count | parse-evaluate | 10,587 | 274 | 256 | 84 | 13,288 | 2,922 | 0 | 3,164 |
| list-statistical-counta | evaluate | 9,295 | 275 | 256 | 48 | 7,776 | 1,576 | 0 | 3,184 |
| list-statistical-counta | parse-evaluate | 9,855 | 275 | 256 | 84 | 13,292 | 2,923 | 0 | 3,184 |
| list-statistical-countblank | evaluate | 9,240 | 278 | 256 | 48 | 7,776 | 1,576 | 0 | 3,172 |
| list-statistical-countblank | parse-evaluate | 9,837 | 278 | 256 | 84 | 13,308 | 2,927 | 0 | 3,184 |
| list-statistical-max | evaluate | 9,995 | 272 | 256 | 48 | 7,776 | 1,576 | 0 | 3,172 |
| list-statistical-max | parse-evaluate | 10,550 | 272 | 256 | 84 | 13,280 | 2,920 | 0 | 3,164 |
| list-statistical-maxa | evaluate | 10,365 | 401 | 256 | 48 | 7,776 | 1,576 | 0 | 3,160 |
| list-statistical-maxa | parse-evaluate | 11,022 | 401 | 256 | 84 | 13,284 | 2,921 | 0 | 3,152 |
| list-statistical-min | evaluate | 10,005 | 272 | 256 | 48 | 7,776 | 1,576 | 0 | 3,176 |
| list-statistical-min | parse-evaluate | 10,622 | 272 | 256 | 84 | 13,280 | 2,920 | 0 | 3,140 |
| list-statistical-mina | evaluate | 10,370 | 401 | 256 | 48 | 7,776 | 1,576 | 0 | 3,164 |
| list-statistical-mina | parse-evaluate | 10,937 | 401 | 256 | 84 | 13,284 | 2,921 | 0 | 3,156 |
| literal-statistical-average | evaluate | 2,100 | 31 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| literal-statistical-average | parse-evaluate | 2,900 | 31 | 0 | 24 | 8,123 | 6,059 | 0 | 3,188 |
| literal-statistical-averagea | evaluate | 2,250 | 38 | 0 | 11 | 5,000 | 4,376 | 0 | 3,172 |
| literal-statistical-averagea | parse-evaluate | 3,230 | 38 | 0 | 19 | 6,371 | 5,235 | 0 | 3,188 |
| literal-statistical-count | evaluate | 1,890 | 29 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| literal-statistical-count | parse-evaluate | 2,630 | 29 | 0 | 24 | 8,121 | 6,057 | 0 | 3,156 |
| literal-statistical-counta | evaluate | 1,990 | 35 | 0 | 11 | 5,000 | 4,376 | 0 | 3,164 |
| literal-statistical-counta | parse-evaluate | 2,980 | 35 | 0 | 19 | 6,369 | 5,233 | 0 | 3,176 |
| literal-statistical-countblank | evaluate | 1,790 | 29 | 0 | 11 | 5,000 | 4,376 | 0 | 3,164 |
| literal-statistical-countblank | parse-evaluate | 2,540 | 29 | 0 | 24 | 8,126 | 6,062 | 0 | 3,160 |
| literal-statistical-max | evaluate | 1,890 | 27 | 0 | 11 | 5,000 | 4,376 | 0 | 3,176 |
| literal-statistical-max | parse-evaluate | 2,620 | 27 | 0 | 24 | 8,119 | 6,055 | 0 | 3,176 |
| literal-statistical-maxa | evaluate | 2,010 | 34 | 0 | 11 | 5,000 | 4,376 | 0 | 3,192 |
| literal-statistical-maxa | parse-evaluate | 3,130 | 34 | 0 | 19 | 6,367 | 5,231 | 0 | 3,164 |
| literal-statistical-min | evaluate | 1,910 | 27 | 0 | 11 | 5,000 | 4,376 | 0 | 3,144 |
| literal-statistical-min | parse-evaluate | 2,680 | 27 | 0 | 24 | 8,119 | 6,055 | 0 | 3,168 |
| literal-statistical-mina | evaluate | 2,060 | 34 | 0 | 11 | 5,000 | 4,376 | 0 | 3,160 |
| literal-statistical-mina | parse-evaluate | 2,990 | 34 | 0 | 19 | 6,367 | 5,231 | 0 | 3,176 |
| nested-projected-statistical-1024-average | evaluate | 1,001,975 | 24,617 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,220 |
| nested-projected-statistical-1024-average | parse-evaluate | 1,000,725 | 24,617 | 8,192 | 4,174 | 1,061,994 | 724,554 | 0 | 4,220 |
| nested-projected-statistical-1024-counta | evaluate | 975,564 | 24,616 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,208 |
| nested-projected-statistical-1024-counta | parse-evaluate | 979,945 | 24,616 | 8,192 | 4,174 | 1,061,993 | 724,553 | 0 | 4,184 |
| nested-projected-statistical-1024-countblank | evaluate | 970,795 | 24,619 | 8,192 | 4,161 | 1,059,704 | 722,776 | 0 | 4,220 |
| nested-projected-statistical-1024-countblank | parse-evaluate | 982,935 | 24,619 | 8,192 | 4,174 | 1,061,997 | 724,557 | 0 | 4,200 |
| nested-projected-statistical-256-average | evaluate | 262,406 | 6,185 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,424 |
| nested-projected-statistical-256-average | parse-evaluate | 265,636 | 6,185 | 2,048 | 2,188 | 538,832 | 183,880 | 0 | 3,420 |
| nested-projected-statistical-256-counta | evaluate | 256,291 | 6,184 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,384 |
| nested-projected-statistical-256-counta | parse-evaluate | 258,811 | 6,184 | 2,048 | 2,188 | 538,830 | 183,879 | 0 | 3,396 |
| nested-projected-statistical-256-countblank | evaluate | 258,086 | 6,187 | 2,048 | 2,162 | 534,256 | 182,104 | 0 | 3,424 |
| nested-projected-statistical-256-countblank | parse-evaluate | 259,281 | 6,187 | 2,048 | 2,188 | 538,838 | 183,883 | 0 | 3,416 |
| nested-projected-statistical-64-average | evaluate | 69,380 | 1,577 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,428 |
| nested-projected-statistical-64-average | parse-evaluate | 69,642 | 1,577 | 512 | 1,272 | 285,080 | 48,710 | 0 | 3,408 |
| nested-projected-statistical-64-averagea | evaluate | 68,675 | 1,706 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,388 |
| nested-projected-statistical-64-averagea | parse-evaluate | 70,030 | 1,706 | 512 | 1,272 | 285,084 | 48,711 | 0 | 3,424 |
| nested-projected-statistical-64-count | evaluate | 67,970 | 1,575 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,432 |
| nested-projected-statistical-64-count | parse-evaluate | 68,450 | 1,575 | 512 | 1,272 | 285,072 | 48,708 | 0 | 3,420 |
| nested-projected-statistical-64-counta | evaluate | 67,897 | 1,576 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,424 |
| nested-projected-statistical-64-counta | parse-evaluate | 68,110 | 1,576 | 512 | 1,272 | 285,076 | 48,709 | 0 | 3,448 |
| nested-projected-statistical-64-countblank | evaluate | 67,013 | 1,579 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,436 |
| nested-projected-statistical-64-countblank | parse-evaluate | 68,257 | 1,579 | 512 | 1,272 | 285,092 | 48,713 | 0 | 3,420 |
| nested-projected-statistical-64-max | evaluate | 68,007 | 1,573 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,428 |
| nested-projected-statistical-64-max | parse-evaluate | 68,830 | 1,573 | 512 | 1,272 | 285,064 | 48,706 | 0 | 3,452 |
| nested-projected-statistical-64-maxa | evaluate | 68,563 | 1,702 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,424 |
| nested-projected-statistical-64-maxa | parse-evaluate | 68,805 | 1,702 | 512 | 1,272 | 285,068 | 48,707 | 0 | 3,424 |
| nested-projected-statistical-64-min | evaluate | 67,685 | 1,573 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,424 |
| nested-projected-statistical-64-min | parse-evaluate | 68,525 | 1,573 | 512 | 1,272 | 285,064 | 48,706 | 0 | 3,432 |
| nested-projected-statistical-64-mina | evaluate | 68,837 | 1,702 | 512 | 1,220 | 275,936 | 46,936 | 0 | 3,424 |
| nested-projected-statistical-64-mina | parse-evaluate | 69,622 | 1,702 | 512 | 1,272 | 285,068 | 48,707 | 0 | 3,424 |
| reference-statistical-average | evaluate | 36,735 | 1,038 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,172 |
| reference-statistical-average | parse-evaluate | 37,220 | 1,038 | 1,024 | 30 | 5,566 | 2,751 | 0 | 3,156 |
| reference-statistical-averagea | evaluate | 37,210 | 1,551 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,176 |
| reference-statistical-averagea | parse-evaluate | 37,370 | 1,551 | 1,024 | 30 | 5,568 | 2,752 | 0 | 3,172 |
| reference-statistical-count | evaluate | 33,105 | 1,036 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,164 |
| reference-statistical-count | parse-evaluate | 33,435 | 1,036 | 1,024 | 30 | 5,562 | 2,749 | 0 | 3,164 |
| reference-statistical-counta | evaluate | 30,685 | 1,037 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,184 |
| reference-statistical-counta | parse-evaluate | 31,130 | 1,037 | 1,024 | 30 | 5,564 | 2,750 | 0 | 3,172 |
| reference-statistical-countblank | evaluate | 31,540 | 1,040 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,148 |
| reference-statistical-countblank | parse-evaluate | 32,005 | 1,040 | 1,024 | 30 | 5,572 | 2,754 | 0 | 3,172 |
| reference-statistical-max | evaluate | 33,625 | 1,034 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,152 |
| reference-statistical-max | parse-evaluate | 34,105 | 1,034 | 1,024 | 30 | 5,558 | 2,747 | 0 | 3,176 |
| reference-statistical-maxa | evaluate | 35,050 | 1,547 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,192 |
| reference-statistical-maxa | parse-evaluate | 35,510 | 1,547 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,168 |
| reference-statistical-min | evaluate | 33,665 | 1,034 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,176 |
| reference-statistical-min | parse-evaluate | 34,015 | 1,034 | 1,024 | 30 | 5,558 | 2,747 | 0 | 3,188 |
| reference-statistical-mina | evaluate | 35,020 | 1,547 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,168 |
| reference-statistical-mina | parse-evaluate | 35,695 | 1,547 | 1,024 | 30 | 5,560 | 2,748 | 0 | 3,164 |
| resource-statistical-average | evaluate | 392 | 10 | 0 | 16 | 1,184 | 296 | 0 | 3,156 |
| resource-statistical-average | parse-evaluate | 787 | 10 | 0 | 44 | 6,648 | 1,630 | 0 | 3,172 |
| resource-statistical-averagea | evaluate | 397 | 11 | 0 | 16 | 1,184 | 296 | 0 | 3,176 |
| resource-statistical-averagea | parse-evaluate | 790 | 11 | 0 | 44 | 6,652 | 1,631 | 0 | 3,164 |
| resource-statistical-count | evaluate | 387 | 8 | 0 | 16 | 1,184 | 296 | 0 | 3,164 |
| resource-statistical-count | parse-evaluate | 787 | 8 | 0 | 44 | 6,640 | 1,628 | 0 | 3,148 |
| resource-statistical-counta | evaluate | 397 | 9 | 0 | 16 | 1,184 | 296 | 0 | 3,172 |
| resource-statistical-counta | parse-evaluate | 795 | 9 | 0 | 44 | 6,644 | 1,629 | 0 | 3,192 |
| resource-statistical-countblank | evaluate | 392 | 13 | 0 | 16 | 1,184 | 296 | 0 | 3,140 |
| resource-statistical-countblank | parse-evaluate | 802 | 13 | 0 | 44 | 6,660 | 1,633 | 0 | 3,164 |
| resource-statistical-max | evaluate | 387 | 6 | 0 | 16 | 1,184 | 296 | 0 | 3,180 |
| resource-statistical-max | parse-evaluate | 772 | 6 | 0 | 44 | 6,632 | 1,626 | 0 | 3,168 |
| resource-statistical-maxa | evaluate | 385 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,184 |
| resource-statistical-maxa | parse-evaluate | 782 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,176 |
| resource-statistical-min | evaluate | 387 | 6 | 0 | 16 | 1,184 | 296 | 0 | 3,164 |
| resource-statistical-min | parse-evaluate | 777 | 6 | 0 | 44 | 6,632 | 1,626 | 0 | 3,156 |
| resource-statistical-mina | evaluate | 385 | 7 | 0 | 16 | 1,184 | 296 | 0 | 3,156 |
| resource-statistical-mina | parse-evaluate | 772 | 7 | 0 | 44 | 6,636 | 1,627 | 0 | 3,164 |
| scalar-statistical-average | evaluate | 677 | 18 | 0 | 4,000 | 400,000 | 400 | 0 | 2,908 |
| scalar-statistical-average | parse-evaluate | 871 | 18 | 0 | 8,000 | 862,000 | 830 | 0 | 2,904 |
| scalar-statistical-averagea | evaluate | 668 | 19 | 0 | 4,000 | 400,000 | 400 | 0 | 2,908 |
| scalar-statistical-averagea | parse-evaluate | 872 | 19 | 0 | 8,000 | 863,000 | 831 | 0 | 2,908 |
| scalar-statistical-count | evaluate | 485 | 16 | 0 | 4,000 | 400,000 | 400 | 0 | 2,912 |
| scalar-statistical-count | parse-evaluate | 668 | 16 | 0 | 8,000 | 860,000 | 828 | 0 | 2,900 |
| scalar-statistical-counta | evaluate | 474 | 17 | 0 | 4,000 | 400,000 | 400 | 0 | 2,904 |
| scalar-statistical-counta | parse-evaluate | 671 | 17 | 0 | 8,000 | 861,000 | 829 | 0 | 2,900 |
| scalar-statistical-countblank | evaluate | 1,200 | 15 | 1 | 8 | 1,416 | 1,416 | 0 | 3,156 |
| scalar-statistical-countblank | parse-evaluate | 1,590 | 15 | 1 | 14 | 2,779 | 2,747 | 0 | 3,180 |
| scalar-statistical-max | evaluate | 469 | 14 | 0 | 4,000 | 400,000 | 400 | 0 | 2,916 |
| scalar-statistical-max | parse-evaluate | 650 | 14 | 0 | 8,000 | 858,000 | 826 | 0 | 2,900 |
| scalar-statistical-maxa | evaluate | 472 | 15 | 0 | 4,000 | 400,000 | 400 | 0 | 2,944 |
| scalar-statistical-maxa | parse-evaluate | 654 | 15 | 0 | 8,000 | 859,000 | 827 | 0 | 2,912 |
| scalar-statistical-min | evaluate | 471 | 14 | 0 | 4,000 | 400,000 | 400 | 0 | 2,912 |
| scalar-statistical-min | parse-evaluate | 652 | 14 | 0 | 8,000 | 858,000 | 826 | 0 | 2,952 |
| scalar-statistical-mina | evaluate | 471 | 15 | 0 | 4,000 | 400,000 | 400 | 0 | 2,908 |
| scalar-statistical-mina | parse-evaluate | 655 | 15 | 0 | 8,000 | 859,000 | 827 | 0 | 2,928 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed finite result or formula error before timing. The profile does not measure save, recalculation, cache publication, native producer acceptance, cold filesystem state, wildcard/regular-expression host profiles, or cross-platform bit identity.
