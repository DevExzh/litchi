# ODS conditional aggregate evaluator performance profile

The baseline is committed `5b125e9ea5870a56b2648ad02917894a5549c571`. It contributes nine matched arithmetic, SIN, IMSUM, DSUM, and SUM controls. Three array controls and all conditional aggregate rows are candidate-only evidence; the baseline has no corresponding valid path for those rows. Each cell is the p50 across fifteen fresh child processes; time, work, and resolver reads are normalized by the fixed repeat count.

## Matched controls

| case | phase | baseline ns/repeat | candidate ns/repeat | delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-4x4-arithmetic | evaluate | 4,502 | 4,428 | -1.6% | 1,120 | 1,120 | 2,972 | 2,896 |
| array-control-4x4-arithmetic | parse-evaluate | 5,234 | 5,165 | -1.3% | 2,240 | 2,240 | 2,972 | 2,932 |
| database-control-dsum | evaluate | 3,850 | 3,910 | +1.6% | 20 | 20 | 3,012 | 2,984 |
| database-control-dsum | parse-evaluate | 4,480 | 4,550 | +1.6% | 29 | 29 | 3,020 | 2,972 |
| literal-aggregate-4x4-sum | evaluate | 1,850 | 1,840 | -0.5% | 10 | 10 | 2,956 | 2,984 |
| literal-aggregate-4x4-sum | parse-evaluate | 2,520 | 2,470 | -2.0% | 23 | 23 | 3,016 | 3,152 |
| reference-aggregate-64x4-sum | evaluate | 10,195 | 10,287 | +0.9% | 32 | 32 | 2,952 | 2,984 |
| reference-aggregate-64x4-sum | parse-evaluate | 10,610 | 10,647 | +0.3% | 60 | 60 | 2,972 | 2,996 |
| reference-array-16x4-arithmetic | evaluate | 9,283 | 9,399 | +1.2% | 800 | 800 | 2,904 | 2,980 |
| reference-array-16x4-arithmetic | parse-evaluate | 9,646 | 9,644 | -0.0% | 1,280 | 1,280 | 2,912 | 2,976 |
| scalar-aggregate-sum | evaluate | 551 | 575 | +4.4% | 4,000 | 4,000 | 2,924 | 2,924 |
| scalar-aggregate-sum | parse-evaluate | 742 | 753 | +1.5% | 8,000 | 8,000 | 2,924 | 2,924 |
| scalar-control-arithmetic | evaluate | 560 | 560 | +0.0% | 5,000 | 5,000 | 2,940 | 2,956 |
| scalar-control-arithmetic | parse-evaluate | 699 | 702 | +0.4% | 8,000 | 8,000 | 2,928 | 2,924 |
| scalar-control-imsum | evaluate | 1,425 | 1,380 | -3.2% | 8,000 | 8,000 | 2,912 | 2,996 |
| scalar-control-imsum | parse-evaluate | 2,006 | 1,956 | -2.5% | 17,000 | 17,000 | 2,936 | 2,924 |
| scalar-control-sin | evaluate | 459 | 455 | -0.9% | 4,000 | 4,000 | 3,168 | 3,236 |
| scalar-control-sin | parse-evaluate | 646 | 647 | +0.2% | 8,000 | 8,000 | 3,200 | 3,284 |

## Candidate conditional aggregate workloads

The ordinary rows cover omitted and explicit destination ranges, one and two criteria, exact text criteria, empty selections, and mismatched geometry. Nested rows use an outer array `IF` and a scalar conditional reducer at 64, 256, and 1024 rows; their resolver reads expose whether the reducer is projected once or rebuilt per output cell.

| case | phase | time ns/repeat | work/repeat | reference reads | alloc calls | requested bytes | peak live bytes | result-live budget | RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 46,090 | 2,825 | 0 | 88 | 703,648 | 110,776 | 22,528 | 3,240 |
| array-control-16x16-arithmetic | parse-evaluate | 59,770 | 2,825 | 0 | 352 | 1,152,060 | 166,335 | 22,528 | 3,240 |
| array-control-16x16-sin | evaluate | 125,863 | 3,336 | 0 | 2,144 | 1,279,328 | 111,896 | 22,528 | 3,548 |
| array-control-16x16-sin | parse-evaluate | 139,888 | 3,336 | 0 | 2,412 | 1,727,868 | 167,455 | 22,528 | 3,596 |
| array-control-4x4-sin | evaluate | 9,620 | 216 | 0 | 3,840 | 1,624,960 | 8,216 | 1,408 | 3,276 |
| array-control-4x4-sin | parse-evaluate | 10,449 | 216 | 0 | 5,040 | 2,151,600 | 11,695 | 1,408 | 3,304 |
| error-conditional-averageif-empty | evaluate | 5,957 | 102 | 64 | 1,200 | 447,360 | 5,016 | 0 | 2,988 |
| error-conditional-averageif-empty | parse-evaluate | 6,515 | 102 | 64 | 1,920 | 558,400 | 6,372 | 0 | 2,992 |
| error-conditional-averageifs-empty | evaluate | 6,256 | 117 | 64 | 1,680 | 640,000 | 6,656 | 0 | 2,976 |
| error-conditional-averageifs-empty | parse-evaluate | 7,104 | 117 | 64 | 2,720 | 819,520 | 8,420 | 0 | 3,108 |
| error-conditional-constant-range | evaluate | 1,710 | 25 | 0 | 10 | 4,056 | 3,816 | 0 | 3,108 |
| error-conditional-constant-range | parse-evaluate | 2,171 | 25 | 0 | 19 | 5,447 | 4,663 | 0 | 3,036 |
| error-conditional-countifs-mismatch | evaluate | 2,768 | 41 | 0 | 1,360 | 469,120 | 5,288 | 0 | 3,100 |
| error-conditional-countifs-mismatch | parse-evaluate | 3,399 | 41 | 0 | 2,160 | 641,520 | 7,027 | 0 | 3,024 |
| error-conditional-sumifs-mismatch | evaluate | 2,586 | 34 | 0 | 1,360 | 453,760 | 5,096 | 0 | 3,012 |
| error-conditional-sumifs-mismatch | parse-evaluate | 3,195 | 34 | 0 | 2,080 | 564,400 | 6,447 | 0 | 3,028 |
| nested-conditional-1024-sumifs | evaluate | 1,224,086 | 27,731 | 11,264 | 4,190 | 1,091,728 | 728,176 | 0 | 3,772 |
| nested-conditional-1024-sumifs | parse-evaluate | 1,234,146 | 27,731 | 11,264 | 4,209 | 1,095,679 | 730,775 | 0 | 3,988 |
| nested-conditional-256-sumifs | evaluate | 296,391 | 6,995 | 2,816 | 2,220 | 561,440 | 187,504 | 0 | 3,368 |
| nested-conditional-256-sumifs | parse-evaluate | 300,236 | 6,995 | 2,816 | 2,258 | 569,334 | 190,099 | 0 | 3,428 |
| nested-conditional-64-sumifs | evaluate | 83,917 | 1,811 | 704 | 1,336 | 311,872 | 52,336 | 0 | 3,324 |
| nested-conditional-64-sumifs | parse-evaluate | 84,483 | 1,811 | 704 | 1,412 | 327,644 | 54,927 | 0 | 3,252 |
| reference-conditional-1024x4-averageifs | evaluate | 300,771 | 7,219 | 7,168 | 21 | 8,000 | 6,656 | 0 | 2,980 |
| reference-conditional-1024x4-averageifs | parse-evaluate | 301,422 | 7,219 | 7,168 | 34 | 10,249 | 8,425 | 0 | 3,028 |
| reference-conditional-1024x4-countif | evaluate | 243,421 | 4,121 | 4,096 | 10 | 3,976 | 3,912 | 0 | 3,000 |
| reference-conditional-1024x4-countif | parse-evaluate | 243,531 | 4,121 | 4,096 | 17 | 5,350 | 5,254 | 0 | 2,924 |
| reference-conditional-16x1-countif-text | evaluate | 2,820 | 69 | 16 | 10 | 3,976 | 3,912 | 0 | 3,000 |
| reference-conditional-16x1-countif-text | parse-evaluate | 3,280 | 69 | 16 | 17 | 5,346 | 5,250 | 0 | 2,980 |
| reference-conditional-16x4-sumif-implicit | evaluate | 5,850 | 87 | 64 | 800 | 318,080 | 3,912 | 0 | 3,112 |
| reference-conditional-16x4-sumif-implicit | parse-evaluate | 6,179 | 87 | 64 | 1,360 | 427,680 | 5,250 | 0 | 3,028 |
| reference-conditional-256x4-countifs | evaluate | 67,565 | 1,577 | 1,536 | 34 | 11,728 | 5,288 | 0 | 2,996 |
| reference-conditional-256x4-countifs | parse-evaluate | 68,375 | 1,577 | 1,536 | 54 | 16,044 | 7,030 | 0 | 3,136 |
| reference-conditional-256x4-sumifs | evaluate | 78,220 | 1,839 | 1,792 | 42 | 16,000 | 6,656 | 0 | 3,028 |
| reference-conditional-256x4-sumifs | parse-evaluate | 78,755 | 1,839 | 1,792 | 68 | 20,484 | 8,418 | 0 | 3,160 |
| reference-conditional-2x4-sumif-anchor-clip | evaluate | 3,320 | 42 | 10 | 15 | 5,592 | 5,016 | 0 | 3,016 |
| reference-conditional-2x4-sumif-anchor-clip | parse-evaluate | 4,560 | 42 | 10 | 23 | 6,977 | 6,369 | 0 | 2,980 |
| reference-conditional-3d-sumif | evaluate | 4,220 | 49 | 5 | 17 | 6,040 | 5,240 | 0 | 2,996 |
| reference-conditional-3d-sumif | parse-evaluate | 4,480 | 49 | 5 | 30 | 7,465 | 6,633 | 0 | 2,996 |
| reference-conditional-64x4-averageif-explicit | evaluate | 24,227 | 420 | 384 | 60 | 22,368 | 5,016 | 0 | 3,028 |
| reference-conditional-64x4-averageif-explicit | parse-evaluate | 24,547 | 420 | 384 | 96 | 27,916 | 6,371 | 0 | 3,004 |
| reference-conditional-64x4-averageif-implicit | evaluate | 17,735 | 283 | 256 | 40 | 15,904 | 3,912 | 0 | 2,996 |
| reference-conditional-64x4-averageif-implicit | parse-evaluate | 18,430 | 283 | 256 | 68 | 21,400 | 5,254 | 0 | 2,980 |
| reference-conditional-64x4-sumif-explicit | evaluate | 24,237 | 416 | 384 | 60 | 22,368 | 5,016 | 0 | 2,996 |
| reference-conditional-64x4-sumif-explicit | parse-evaluate | 24,367 | 416 | 384 | 96 | 27,900 | 6,367 | 0 | 3,000 |
| reference-conditional-6x4-sumif-list | evaluate | 4,280 | 54 | 24 | 14 | 4,632 | 4,136 | 0 | 2,976 |
| reference-conditional-6x4-sumif-list | parse-evaluate | 5,720 | 54 | 24 | 25 | 6,847 | 5,903 | 0 | 3,028 |
| resource-conditional-reference-cells | evaluate | 632 | 10 | 0 | 480 | 228,480 | 2,792 | 0 | 2,984 |
| resource-conditional-reference-cells | parse-evaluate | 1,016 | 10 | 0 | 1,040 | 338,080 | 4,130 | 0 | 2,944 |

The resolver is an immutable borrowing fixture. Direct f64 fixture arithmetic validates one untimed finite result or formula error before timing. The profile does not measure save, recalculation, cache publication, native producer acceptance, cold filesystem state, wildcard/regular-expression host profiles, or cross-platform bit identity.
