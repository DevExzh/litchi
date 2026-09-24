# ODS aggregate evaluator performance profile

The baseline is committed `2aeeb8d2f`. It contributes only matched controls: arithmetic, SIN, IMSUM, DSUM, and their literal/reference projections. The seven new aggregate functions have no valid baseline implementation, so their rows are candidate-only evidence. Each cell is the p50 across fifteen fresh child processes; time and work are normalized by the fixed repeat count.

The 16x16 literal parse-evaluate rows repeat classification of a large literal AST on each fixed repeat. They therefore include parser/classifier work for this harness fixture; compare them with the corresponding evaluate rows and streamed-reference rows only as workload-specific observations. Their repeat count is four, as recorded in the raw receipts.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 46,067 | 45,642 | -0.9% | 88 | 88 | 3,168 | 3,212 |
| array-control-16x16-arithmetic | parse-evaluate | 60,675 | 59,870 | -1.3% | 352 | 352 | 3,176 | 3,208 |
| array-control-16x16-sin | evaluate | 125,868 | 127,813 | +1.5% | 2,144 | 2,144 | 3,464 | 3,500 |
| array-control-16x16-sin | parse-evaluate | 140,940 | 140,095 | -0.6% | 2,412 | 2,412 | 3,484 | 3,468 |
| array-control-4x4-arithmetic | evaluate | 4,422 | 4,498 | +1.7% | 1,120 | 1,120 | 2,924 | 2,956 |
| array-control-4x4-arithmetic | parse-evaluate | 5,245 | 5,285 | +0.8% | 2,240 | 2,240 | 2,916 | 2,984 |
| array-control-4x4-sin | evaluate | 9,622 | 9,696 | +0.8% | 3,840 | 3,840 | 3,208 | 3,508 |
| array-control-4x4-sin | parse-evaluate | 10,522 | 10,591 | +0.7% | 5,040 | 5,040 | 3,220 | 3,504 |
| database-control-dsum | evaluate | 3,910 | 3,900 | -0.3% | 20 | 20 | 2,892 | 2,924 |
| database-control-dsum | parse-evaluate | 4,460 | 4,460 | +0.0% | 29 | 29 | 2,944 | 2,892 |
| reference-array-arithmetic | evaluate | 3,387 | 3,414 | +0.8% | 800 | 800 | 2,920 | 2,948 |
| reference-array-arithmetic | parse-evaluate | 3,707 | 3,718 | +0.3% | 1,280 | 1,280 | 2,920 | 2,900 |
| reference-array-sin | evaluate | 8,487 | 8,566 | +0.9% | 3,440 | 3,440 | 3,264 | 3,252 |
| reference-array-sin | parse-evaluate | 8,890 | 9,030 | +1.6% | 4,000 | 4,000 | 3,256 | 3,252 |
| reference-scalar-arithmetic | evaluate | 772 | 764 | -1.0% | 5,000 | 5,000 | 2,960 | 2,932 |
| reference-scalar-arithmetic | parse-evaluate | 1,027 | 1,036 | +0.9% | 10,000 | 10,000 | 2,908 | 2,952 |
| reference-scalar-sin | evaluate | 1,526 | 1,504 | -1.4% | 11,000 | 11,000 | 3,260 | 3,224 |
| reference-scalar-sin | parse-evaluate | 1,828 | 1,817 | -0.6% | 17,000 | 17,000 | 3,252 | 3,248 |
| scalar-control-arithmetic | evaluate | 579 | 577 | -0.3% | 5,000 | 5,000 | 2,904 | 2,832 |
| scalar-control-arithmetic | parse-evaluate | 721 | 721 | +0.0% | 8,000 | 8,000 | 2,880 | 2,880 |
| scalar-control-imsum | evaluate | 1,447 | 1,430 | -1.2% | 8,000 | 8,000 | 2,912 | 2,852 |
| scalar-control-imsum | parse-evaluate | 2,010 | 1,996 | -0.7% | 17,000 | 17,000 | 2,908 | 2,880 |
| scalar-control-sin | evaluate | 450 | 454 | +0.9% | 4,000 | 4,000 | 3,152 | 3,132 |
| scalar-control-sin | parse-evaluate | 634 | 634 | +0.0% | 8,000 | 8,000 | 3,116 | 3,160 |

## Candidate aggregate workloads

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| literal-aggregate-16x16-product | evaluate | 25,957 | 2,829 | 0 | 88 | 615,072 | 88,632 | 0 | 3,188 |
| literal-aggregate-16x16-product | parse-evaluate | 36,872 | 2,829 | 0 | 356 | 1,063,628 | 144,195 | 0 | 3,188 |
| literal-aggregate-16x16-sum | evaluate | 26,742 | 2,825 | 0 | 88 | 615,072 | 88,632 | 0 | 3,204 |
| literal-aggregate-16x16-sum | parse-evaluate | 37,002 | 2,825 | 0 | 356 | 1,063,612 | 144,191 | 0 | 3,164 |
| literal-aggregate-16x16-sumproduct | evaluate | 64,177 | 5,910 | 0 | 104 | 1,108,320 | 162,744 | 0 | 3,200 |
| literal-aggregate-16x16-sumproduct | parse-evaluate | 91,045 | 5,910 | 0 | 584 | 2,007,328 | 273,864 | 0 | 3,432 |
| literal-aggregate-16x16-sumsq | evaluate | 26,637 | 2,827 | 0 | 88 | 615,072 | 88,632 | 0 | 3,144 |
| literal-aggregate-16x16-sumsq | parse-evaluate | 37,157 | 2,827 | 0 | 356 | 1,063,620 | 144,193 | 0 | 3,160 |
| literal-aggregate-16x16-sumx2my2 | evaluate | 62,805 | 5,908 | 0 | 104 | 1,108,320 | 162,744 | 0 | 3,172 |
| literal-aggregate-16x16-sumx2my2 | parse-evaluate | 92,255 | 5,908 | 0 | 584 | 2,007,320 | 273,862 | 0 | 3,468 |
| literal-aggregate-16x16-sumx2py2 | evaluate | 64,197 | 5,908 | 0 | 104 | 1,108,320 | 162,744 | 0 | 3,184 |
| literal-aggregate-16x16-sumx2py2 | parse-evaluate | 92,315 | 5,908 | 0 | 584 | 2,007,320 | 273,862 | 0 | 3,456 |
| literal-aggregate-16x16-sumxmy2 | evaluate | 63,687 | 5,907 | 0 | 104 | 1,108,320 | 162,744 | 0 | 3,188 |
| literal-aggregate-16x16-sumxmy2 | parse-evaluate | 92,370 | 5,907 | 0 | 584 | 2,007,316 | 273,861 | 0 | 3,460 |
| literal-aggregate-4x4-product | evaluate | 3,133 | 189 | 0 | 1,120 | 781,440 | 6,072 | 0 | 3,140 |
| literal-aggregate-4x4-product | parse-evaluate | 4,061 | 189 | 0 | 2,320 | 1,308,400 | 9,555 | 0 | 3,168 |
| literal-aggregate-4x4-sum | evaluate | 3,242 | 185 | 0 | 1,120 | 781,440 | 6,072 | 0 | 3,116 |
| literal-aggregate-4x4-sum | parse-evaluate | 4,102 | 185 | 0 | 2,320 | 1,308,080 | 9,551 | 0 | 3,156 |
| literal-aggregate-4x4-sumproduct | evaluate | 6,296 | 390 | 0 | 1,440 | 1,584,000 | 12,984 | 0 | 3,184 |
| literal-aggregate-4x4-sumproduct | parse-evaluate | 7,669 | 390 | 0 | 3,200 | 2,668,160 | 19,944 | 0 | 3,208 |
| literal-aggregate-4x4-sumproduct-k3 | evaluate | 8,566 | 571 | 0 | 1,600 | 1,715,840 | 14,584 | 0 | 3,180 |
| literal-aggregate-4x4-sumproduct-k3 | parse-evaluate | 10,370 | 571 | 0 | 3,760 | 2,824,480 | 21,690 | 0 | 3,168 |
| literal-aggregate-4x4-sumsq | evaluate | 3,270 | 187 | 0 | 1,120 | 781,440 | 6,072 | 0 | 2,900 |
| literal-aggregate-4x4-sumsq | parse-evaluate | 4,137 | 187 | 0 | 2,320 | 1,308,240 | 9,553 | 0 | 3,184 |
| literal-aggregate-4x4-sumx2my2 | evaluate | 6,210 | 388 | 0 | 1,440 | 1,584,000 | 12,984 | 0 | 3,168 |
| literal-aggregate-4x4-sumx2my2 | parse-evaluate | 7,621 | 388 | 0 | 3,200 | 2,668,000 | 19,942 | 0 | 3,204 |
| literal-aggregate-4x4-sumx2py2 | evaluate | 6,222 | 388 | 0 | 1,440 | 1,584,000 | 12,984 | 0 | 3,172 |
| literal-aggregate-4x4-sumx2py2 | parse-evaluate | 7,617 | 388 | 0 | 3,200 | 2,668,000 | 19,942 | 0 | 3,216 |
| literal-aggregate-4x4-sumxmy2 | evaluate | 6,242 | 387 | 0 | 1,440 | 1,584,000 | 12,984 | 0 | 3,208 |
| literal-aggregate-4x4-sumxmy2 | parse-evaluate | 7,516 | 387 | 0 | 3,200 | 2,667,920 | 19,941 | 0 | 3,208 |
| nested-if-projected-1024-sumproduct | evaluate | 1,842,219 | 41,020 | 12,288 | 4,155 | 1,775,720 | 1,445,416 | 0 | 4,216 |
| nested-if-projected-1024-sumproduct | parse-evaluate | 1,842,609 | 41,020 | 12,288 | 4,171 | 1,779,566 | 1,447,982 | 0 | 4,192 |
| nested-if-projected-256-sumproduct | evaluate | 457,782 | 10,300 | 3,072 | 2,158 | 897,232 | 364,072 | 0 | 3,420 |
| nested-if-projected-256-sumproduct | parse-evaluate | 462,397 | 10,300 | 3,072 | 2,190 | 904,918 | 366,635 | 0 | 3,448 |
| nested-if-projected-64-sumproduct | evaluate | 111,543 | 2,620 | 768 | 1,228 | 467,360 | 93,736 | 0 | 3,208 |
| nested-if-projected-64-sumproduct | parse-evaluate | 112,685 | 2,620 | 768 | 1,292 | 482,720 | 96,296 | 0 | 3,200 |
| nested-projected-1024-sumproduct | evaluate | 1,029,725 | 24,612 | 12,288 | 20 | 726,312 | 724,152 | 0 | 3,792 |
| nested-projected-1024-sumproduct | parse-evaluate | 1,041,285 | 24,612 | 12,288 | 34 | 728,582 | 725,942 | 0 | 3,756 |
| nested-projected-256-sumproduct | evaluate | 252,911 | 6,180 | 3,072 | 40 | 371,280 | 183,480 | 0 | 3,180 |
| nested-projected-256-sumproduct | parse-evaluate | 256,721 | 6,180 | 3,072 | 68 | 375,814 | 185,267 | 0 | 3,204 |
| nested-projected-64-sumproduct | evaluate | 63,015 | 1,572 | 768 | 80 | 201,888 | 48,312 | 0 | 3,208 |
| nested-projected-64-sumproduct | parse-evaluate | 63,845 | 1,572 | 768 | 136 | 210,944 | 50,096 | 0 | 3,164 |
| reference-aggregate-1024x4-product | evaluate | 134,451 | 4,110 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,172 |
| reference-aggregate-1024x4-product | parse-evaluate | 134,731 | 4,110 | 4,096 | 15 | 2,784 | 2,752 | 0 | 3,204 |
| reference-aggregate-1024x4-sum | evaluate | 144,691 | 4,106 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,164 |
| reference-aggregate-1024x4-sum | parse-evaluate | 145,001 | 4,106 | 4,096 | 15 | 2,780 | 2,748 | 0 | 3,204 |
| reference-aggregate-1024x4-sumproduct | evaluate | 349,302 | 12,312 | 8,192 | 13 | 4,376 | 4,312 | 0 | 3,188 |
| reference-aggregate-1024x4-sumproduct | parse-evaluate | 349,992 | 12,312 | 8,192 | 22 | 5,762 | 5,666 | 0 | 3,196 |
| reference-aggregate-1024x4-sumproduct-k3 | evaluate | 510,393 | 16,414 | 12,288 | 18 | 5,992 | 5,416 | 0 | 3,204 |
| reference-aggregate-1024x4-sumproduct-k3 | parse-evaluate | 523,053 | 16,414 | 12,288 | 29 | 7,393 | 6,785 | 0 | 3,156 |
| reference-aggregate-1024x4-sumsq | evaluate | 157,291 | 4,108 | 4,096 | 8 | 1,416 | 1,416 | 0 | 3,160 |
| reference-aggregate-1024x4-sumsq | parse-evaluate | 157,921 | 4,108 | 4,096 | 15 | 2,782 | 2,750 | 0 | 3,200 |
| reference-aggregate-1024x4-sumx2my2 | evaluate | 346,482 | 12,310 | 8,192 | 13 | 4,376 | 4,312 | 0 | 3,192 |
| reference-aggregate-1024x4-sumx2my2 | parse-evaluate | 339,912 | 12,310 | 8,192 | 22 | 5,760 | 5,664 | 0 | 3,156 |
| reference-aggregate-1024x4-sumx2py2 | evaluate | 341,331 | 12,310 | 8,192 | 13 | 4,376 | 4,312 | 0 | 3,204 |
| reference-aggregate-1024x4-sumx2py2 | parse-evaluate | 342,941 | 12,310 | 8,192 | 22 | 5,760 | 5,664 | 0 | 3,196 |
| reference-aggregate-1024x4-sumxmy2 | evaluate | 344,591 | 12,309 | 8,192 | 13 | 4,376 | 4,312 | 0 | 3,188 |
| reference-aggregate-1024x4-sumxmy2 | parse-evaluate | 348,432 | 12,309 | 8,192 | 22 | 5,759 | 5,663 | 0 | 3,176 |
| reference-aggregate-16x4-product | evaluate | 3,154 | 78 | 64 | 640 | 113,280 | 1,416 | 0 | 2,900 |
| reference-aggregate-16x4-product | parse-evaluate | 3,553 | 78 | 64 | 1,200 | 222,560 | 2,750 | 0 | 2,944 |
| reference-aggregate-16x4-sum | evaluate | 3,423 | 74 | 64 | 640 | 113,280 | 1,416 | 0 | 2,968 |
| reference-aggregate-16x4-sum | parse-evaluate | 3,793 | 74 | 64 | 1,200 | 222,240 | 2,746 | 0 | 2,948 |
| reference-aggregate-16x4-sumproduct | evaluate | 7,607 | 216 | 128 | 1,040 | 350,080 | 4,312 | 0 | 2,888 |
| reference-aggregate-16x4-sumproduct | parse-evaluate | 8,099 | 216 | 128 | 1,760 | 460,640 | 5,662 | 0 | 2,924 |
| reference-aggregate-16x4-sumsq | evaluate | 3,474 | 76 | 64 | 640 | 113,280 | 1,416 | 0 | 2,900 |
| reference-aggregate-16x4-sumsq | parse-evaluate | 3,854 | 76 | 64 | 1,200 | 222,400 | 2,748 | 0 | 2,924 |
| reference-aggregate-16x4-sumx2my2 | evaluate | 7,333 | 214 | 128 | 1,040 | 350,080 | 4,312 | 0 | 2,888 |
| reference-aggregate-16x4-sumx2my2 | parse-evaluate | 7,830 | 214 | 128 | 1,760 | 460,480 | 5,660 | 0 | 2,888 |
| reference-aggregate-16x4-sumx2py2 | evaluate | 7,387 | 214 | 128 | 1,040 | 350,080 | 4,312 | 0 | 2,908 |
| reference-aggregate-16x4-sumx2py2 | parse-evaluate | 7,956 | 214 | 128 | 1,760 | 460,480 | 5,660 | 0 | 2,952 |
| reference-aggregate-16x4-sumxmy2 | evaluate | 7,432 | 213 | 128 | 1,040 | 350,080 | 4,312 | 0 | 2,900 |
| reference-aggregate-16x4-sumxmy2 | parse-evaluate | 7,984 | 213 | 128 | 1,760 | 460,400 | 5,659 | 0 | 2,956 |
| reference-aggregate-256x4-product | evaluate | 34,440 | 1,038 | 1,024 | 16 | 2,832 | 1,416 | 0 | 2,948 |
| reference-aggregate-256x4-product | parse-evaluate | 34,860 | 1,038 | 1,024 | 30 | 5,566 | 2,751 | 0 | 2,940 |
| reference-aggregate-256x4-sum | evaluate | 36,985 | 1,034 | 1,024 | 16 | 2,832 | 1,416 | 0 | 3,084 |
| reference-aggregate-256x4-sum | parse-evaluate | 37,390 | 1,034 | 1,024 | 30 | 5,558 | 2,747 | 0 | 2,968 |
| reference-aggregate-256x4-sumproduct | evaluate | 88,555 | 3,096 | 2,048 | 26 | 8,752 | 4,312 | 0 | 2,908 |
| reference-aggregate-256x4-sumproduct | parse-evaluate | 89,285 | 3,096 | 2,048 | 44 | 11,520 | 5,664 | 0 | 2,952 |
| reference-aggregate-256x4-sumsq | evaluate | 40,425 | 1,036 | 1,024 | 16 | 2,832 | 1,416 | 0 | 2,908 |
| reference-aggregate-256x4-sumsq | parse-evaluate | 40,775 | 1,036 | 1,024 | 30 | 5,562 | 2,749 | 0 | 2,912 |
| reference-aggregate-256x4-sumx2my2 | evaluate | 85,430 | 3,094 | 2,048 | 26 | 8,752 | 4,312 | 0 | 2,952 |
| reference-aggregate-256x4-sumx2my2 | parse-evaluate | 86,400 | 3,094 | 2,048 | 44 | 11,516 | 5,662 | 0 | 2,900 |
| reference-aggregate-256x4-sumx2py2 | evaluate | 87,435 | 3,094 | 2,048 | 26 | 8,752 | 4,312 | 0 | 2,880 |
| reference-aggregate-256x4-sumx2py2 | parse-evaluate | 87,540 | 3,094 | 2,048 | 44 | 11,516 | 5,662 | 0 | 2,920 |
| reference-aggregate-256x4-sumxmy2 | evaluate | 87,615 | 3,093 | 2,048 | 26 | 8,752 | 4,312 | 0 | 2,924 |
| reference-aggregate-256x4-sumxmy2 | parse-evaluate | 88,175 | 3,093 | 2,048 | 44 | 11,514 | 5,661 | 0 | 2,924 |
| reference-aggregate-64x4-product | evaluate | 9,437 | 270 | 256 | 32 | 5,664 | 1,416 | 0 | 2,928 |
| reference-aggregate-64x4-product | parse-evaluate | 9,860 | 270 | 256 | 60 | 11,128 | 2,750 | 0 | 2,916 |
| reference-aggregate-64x4-sum | evaluate | 10,122 | 266 | 256 | 32 | 5,664 | 1,416 | 0 | 2,904 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,462 | 266 | 256 | 60 | 11,112 | 2,746 | 0 | 2,952 |
| reference-aggregate-64x4-sumproduct | evaluate | 23,665 | 792 | 512 | 52 | 17,504 | 4,312 | 0 | 2,936 |
| reference-aggregate-64x4-sumproduct | parse-evaluate | 24,255 | 792 | 512 | 88 | 23,032 | 5,662 | 0 | 2,948 |
| reference-aggregate-64x4-sumsq | evaluate | 10,717 | 268 | 256 | 32 | 5,664 | 1,416 | 0 | 2,952 |
| reference-aggregate-64x4-sumsq | parse-evaluate | 11,155 | 268 | 256 | 60 | 11,120 | 2,748 | 0 | 2,940 |
| reference-aggregate-64x4-sumx2my2 | evaluate | 23,032 | 790 | 512 | 52 | 17,504 | 4,312 | 0 | 2,952 |
| reference-aggregate-64x4-sumx2my2 | parse-evaluate | 23,485 | 790 | 512 | 88 | 23,024 | 5,660 | 0 | 3,060 |
| reference-aggregate-64x4-sumx2py2 | evaluate | 23,420 | 790 | 512 | 52 | 17,504 | 4,312 | 0 | 2,936 |
| reference-aggregate-64x4-sumx2py2 | parse-evaluate | 23,910 | 790 | 512 | 88 | 23,024 | 5,660 | 0 | 2,916 |
| reference-aggregate-64x4-sumxmy2 | evaluate | 23,757 | 789 | 512 | 52 | 17,504 | 4,312 | 0 | 2,948 |
| reference-aggregate-64x4-sumxmy2 | parse-evaluate | 24,125 | 789 | 512 | 88 | 23,020 | 5,659 | 0 | 2,952 |
| scalar-aggregate-product | evaluate | 478 | 18 | 0 | 4,000 | 400,000 | 400 | 0 | 2,868 |
| scalar-aggregate-product | parse-evaluate | 675 | 18 | 0 | 8,000 | 862,000 | 830 | 0 | 2,872 |
| scalar-aggregate-sum | evaluate | 549 | 14 | 0 | 4,000 | 400,000 | 400 | 0 | 2,816 |
| scalar-aggregate-sum | parse-evaluate | 740 | 14 | 0 | 8,000 | 858,000 | 826 | 0 | 2,868 |
| scalar-aggregate-sumproduct | evaluate | 791 | 30 | 0 | 5,000 | 496,000 | 448 | 0 | 2,876 |
| scalar-aggregate-sumproduct | parse-evaluate | 1,037 | 30 | 0 | 9,000 | 967,000 | 887 | 0 | 2,920 |
| scalar-aggregate-sumsq | evaluate | 614 | 16 | 0 | 4,000 | 400,000 | 400 | 0 | 2,860 |
| scalar-aggregate-sumsq | parse-evaluate | 807 | 16 | 0 | 8,000 | 860,000 | 828 | 0 | 2,868 |
| scalar-aggregate-sumx2my2 | evaluate | 796 | 28 | 0 | 5,000 | 496,000 | 448 | 0 | 2,872 |
| scalar-aggregate-sumx2my2 | parse-evaluate | 1,032 | 28 | 0 | 9,000 | 965,000 | 885 | 0 | 2,872 |
| scalar-aggregate-sumx2py2 | evaluate | 785 | 28 | 0 | 5,000 | 496,000 | 448 | 0 | 2,864 |
| scalar-aggregate-sumx2py2 | parse-evaluate | 1,025 | 28 | 0 | 9,000 | 965,000 | 885 | 0 | 2,864 |
| scalar-aggregate-sumxmy2 | evaluate | 779 | 27 | 0 | 5,000 | 496,000 | 448 | 0 | 2,872 |
| scalar-aggregate-sumxmy2 | parse-evaluate | 1,011 | 27 | 0 | 9,000 | 964,000 | 884 | 0 | 2,872 |

## Nested projection scaling

The nested rows use `SUMPRODUCT(range+SUM(range);range)` and `SUM(IF(range;SUMPRODUCT(range+1;range);0))`. Their work and resolver-read columns are retained to expose whether invariant inner and branch aggregates are projected once per outer evaluation. This is a bounded evaluator scaling observation, not a whole-workbook recalculation claim.

| case | elements | work/repeat | reference reads | time ns/repeat |
| --- | ---: | ---: | ---: | ---: |
| nested-projected-64-sumproduct | 256 | 1,572 | 768 | 63,015 |
| nested-if-projected-64-sumproduct | 256 | 2,620 | 768 | 111,543 |
| nested-projected-256-sumproduct | 1024 | 6,180 | 3,072 | 252,911 |
| nested-if-projected-256-sumproduct | 1024 | 10,300 | 3,072 | 457,782 |
| nested-projected-1024-sumproduct | 4096 | 24,612 | 12,288 | 1,029,725 |
| nested-if-projected-1024-sumproduct | 4096 | 41,020 | 12,288 | 1,842,219 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed result and does not independently establish libm accuracy. The profile does not measure save, cache publication, native producer acceptance, cold filesystem state, or cross-platform bit identity.
