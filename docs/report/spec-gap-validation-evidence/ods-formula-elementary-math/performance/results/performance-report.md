# ODS elementary-math performance profile

The baseline timing lane contains matched arithmetic/ROUND/trigonometric controls. The candidate lane contains all named scalar, array, and synthetic local-reference cases. New elementary-math cases have no baseline timing comparison because the baseline does not implement them. Time deltas are candidate minus baseline; positive values are slower. Samples are p50 across fresh child processes, and each row's time is normalized by the harness repeat count.

## Matched controls

| case | phase | baseline time ns/repeat | candidate time ns/repeat | time delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,632 | 45,802 | +0.4% | 88 | 88 | 3,160 | 3,196 |
| array-control-16x16-arithmetic | parse-evaluate | 59,715 | 60,165 | +0.8% | 352 | 352 | 3,160 | 3,204 |
| array-control-16x16-round | evaluate | 172,635 | 171,558 | -0.6% | 2,144 | 2,144 | 3,112 | 3,204 |
| array-control-16x16-round | parse-evaluate | 187,948 | 185,123 | -1.5% | 2,412 | 2,412 | 3,180 | 3,212 |
| array-control-16x16-sin | evaluate | 125,175 | 125,390 | +0.2% | 2,144 | 2,144 | 3,432 | 3,460 |
| array-control-16x16-sin | parse-evaluate | 138,940 | 140,390 | +1.0% | 2,412 | 2,412 | 3,404 | 3,476 |
| array-control-4x4-arithmetic | evaluate | 4,378 | 4,368 | -0.2% | 1,120 | 1,120 | 2,892 | 2,936 |
| array-control-4x4-arithmetic | parse-evaluate | 5,246 | 5,152 | -1.8% | 2,240 | 2,240 | 2,884 | 2,948 |
| array-control-4x4-round | evaluate | 12,786 | 12,717 | -0.5% | 3,840 | 3,840 | 2,904 | 2,928 |
| array-control-4x4-round | parse-evaluate | 13,699 | 13,657 | -0.3% | 5,040 | 5,040 | 2,932 | 3,200 |
| array-control-4x4-sin | evaluate | 9,665 | 9,637 | -0.3% | 3,840 | 3,840 | 3,184 | 3,244 |
| array-control-4x4-sin | parse-evaluate | 10,581 | 10,628 | +0.4% | 5,040 | 5,040 | 3,212 | 3,500 |
| reference-array-arithmetic | evaluate | 3,153 | 3,128 | -0.8% | 800 | 800 | 2,892 | 2,964 |
| reference-array-arithmetic | parse-evaluate | 3,481 | 3,471 | -0.3% | 1,280 | 1,280 | 2,904 | 2,960 |
| reference-array-round | evaluate | 11,760 | 11,597 | -1.4% | 3,520 | 3,520 | 2,908 | 2,964 |
| reference-array-round | parse-evaluate | 12,223 | 12,273 | +0.4% | 4,080 | 4,080 | 2,916 | 2,960 |
| reference-array-sin | evaluate | 8,258 | 8,326 | +0.8% | 3,440 | 3,440 | 3,208 | 3,216 |
| reference-array-sin | parse-evaluate | 8,687 | 8,741 | +0.6% | 4,000 | 4,000 | 3,192 | 3,244 |
| reference-scalar-arithmetic | evaluate | 759 | 736 | -3.0% | 5,000 | 5,000 | 2,932 | 2,948 |
| reference-scalar-arithmetic | parse-evaluate | 993 | 994 | +0.1% | 10,000 | 10,000 | 2,880 | 2,952 |
| reference-scalar-round | evaluate | 1,943 | 1,863 | -4.1% | 12,000 | 12,000 | 2,904 | 2,964 |
| reference-scalar-round | parse-evaluate | 2,226 | 2,197 | -1.3% | 18,000 | 18,000 | 2,896 | 2,976 |
| reference-scalar-sin | evaluate | 1,495 | 1,509 | +0.9% | 11,000 | 11,000 | 3,204 | 3,216 |
| reference-scalar-sin | parse-evaluate | 1,793 | 1,804 | +0.6% | 17,000 | 17,000 | 3,204 | 3,248 |
| scalar-control-arithmetic | evaluate | 571 | 565 | -1.1% | 5,000 | 5,000 | 2,892 | 2,932 |
| scalar-control-arithmetic | parse-evaluate | 708 | 708 | +0.0% | 8,000 | 8,000 | 2,880 | 2,928 |
| scalar-control-round | evaluate | 724 | 742 | +2.5% | 5,000 | 5,000 | 2,868 | 2,928 |
| scalar-control-round | parse-evaluate | 1,001 | 964 | -3.7% | 9,000 | 9,000 | 2,872 | 2,940 |
| scalar-control-sin | evaluate | 442 | 449 | +1.6% | 4,000 | 4,000 | 3,116 | 3,080 |
| scalar-control-sin | parse-evaluate | 627 | 630 | +0.5% | 8,000 | 8,000 | 3,156 | 3,076 |

## Candidate elementary-math and representative workloads

| case | phase | supported | time ns/repeat | alloc calls | requested bytes | result-live memory | RSS KiB |
| --- | --- | :---: | ---: | ---: | ---: | ---: | ---: |
| array-elementary-16x16-abs | evaluate | yes | 146,125 | 2,144 | 1,279,328 | 22,528 | 3,216 |
| array-elementary-16x16-abs | parse-evaluate | yes | 166,020 | 2,428 | 2,158,332 | 22,528 | 3,208 |
| array-elementary-16x16-exp | evaluate | yes | 126,653 | 2,144 | 1,279,328 | 22,528 | 3,240 |
| array-elementary-16x16-exp | parse-evaluate | yes | 139,900 | 2,412 | 1,730,940 | 22,528 | 3,244 |
| array-elementary-16x16-ln | evaluate | yes | 125,893 | 2,144 | 1,279,328 | 22,528 | 3,500 |
| array-elementary-16x16-ln | parse-evaluate | yes | 140,015 | 2,412 | 1,730,936 | 22,528 | 3,500 |
| array-elementary-16x16-sign | evaluate | yes | 146,263 | 2,144 | 1,279,328 | 22,528 | 3,208 |
| array-elementary-16x16-sign | parse-evaluate | yes | 167,140 | 2,428 | 2,158,336 | 22,528 | 3,196 |
| array-elementary-16x16-sqrt | evaluate | yes | 126,188 | 2,144 | 1,279,328 | 22,528 | 3,196 |
| array-elementary-16x16-sqrt | parse-evaluate | yes | 139,138 | 2,412 | 1,730,944 | 22,528 | 3,216 |
| array-elementary-4x4-abs | evaluate | yes | 10,935 | 3,840 | 1,624,960 | 1,408 | 2,936 |
| array-elementary-4x4-abs | parse-evaluate | yes | 11,947 | 5,200 | 2,689,200 | 1,408 | 3,216 |
| array-elementary-4x4-exp | evaluate | yes | 9,677 | 3,840 | 1,624,960 | 1,408 | 2,984 |
| array-elementary-4x4-exp | parse-evaluate | yes | 10,579 | 5,040 | 2,155,440 | 1,408 | 3,244 |
| array-elementary-4x4-ln | evaluate | yes | 9,614 | 3,840 | 1,624,960 | 1,408 | 3,260 |
| array-elementary-4x4-ln | parse-evaluate | yes | 10,557 | 5,040 | 2,155,360 | 1,408 | 3,484 |
| array-elementary-4x4-sign | evaluate | yes | 11,026 | 3,840 | 1,624,960 | 1,408 | 2,944 |
| array-elementary-4x4-sign | parse-evaluate | yes | 11,955 | 5,200 | 2,689,280 | 1,408 | 3,184 |
| array-elementary-4x4-sqrt | evaluate | yes | 9,637 | 3,840 | 1,624,960 | 1,408 | 2,936 |
| array-elementary-4x4-sqrt | parse-evaluate | yes | 10,659 | 5,040 | 2,155,520 | 1,408 | 3,192 |
| reference-elementary-abs | evaluate | yes | 1,481 | 11,000 | 2,152,000 | 0 | 2,960 |
| reference-elementary-abs | parse-evaluate | yes | 1,814 | 17,000 | 3,508,000 | 0 | 2,912 |
| reference-elementary-array-abs | evaluate | yes | 8,231 | 3,440 | 1,069,440 | 1,408 | 2,936 |
| reference-elementary-array-abs | parse-evaluate | yes | 8,581 | 4,000 | 1,178,320 | 1,408 | 2,952 |
| reference-elementary-array-ln | evaluate | yes | 8,306 | 3,440 | 1,069,440 | 1,408 | 3,244 |
| reference-elementary-array-ln | parse-evaluate | yes | 8,732 | 4,000 | 1,178,240 | 1,408 | 3,264 |
| reference-elementary-array-sqrt | evaluate | yes | 8,257 | 3,440 | 1,069,440 | 1,408 | 2,960 |
| reference-elementary-array-sqrt | parse-evaluate | yes | 8,782 | 4,000 | 1,178,400 | 1,408 | 2,932 |
| reference-elementary-ln | evaluate | yes | 1,487 | 11,000 | 2,152,000 | 0 | 3,272 |
| reference-elementary-ln | parse-evaluate | yes | 1,790 | 17,000 | 3,507,000 | 0 | 3,256 |
| reference-elementary-sqrt | evaluate | yes | 1,487 | 11,000 | 2,152,000 | 0 | 2,948 |
| reference-elementary-sqrt | parse-evaluate | yes | 1,813 | 17,000 | 3,509,000 | 0 | 2,940 |
| scalar-elementary-abs | evaluate | yes | 604 | 5,000 | 496,000 | 0 | 2,912 |
| scalar-elementary-abs | parse-evaluate | yes | 814 | 9,000 | 962,000 | 0 | 2,948 |
| scalar-elementary-exp | evaluate | yes | 449 | 4,000 | 400,000 | 0 | 2,968 |
| scalar-elementary-exp | parse-evaluate | yes | 634 | 8,000 | 865,000 | 0 | 2,968 |
| scalar-elementary-ln | evaluate | yes | 445 | 4,000 | 400,000 | 0 | 3,076 |
| scalar-elementary-ln | parse-evaluate | yes | 634 | 8,000 | 864,000 | 0 | 3,076 |
| scalar-elementary-log | evaluate | yes | 621 | 5,000 | 496,000 | 0 | 3,076 |
| scalar-elementary-log | parse-evaluate | yes | 834 | 9,000 | 965,000 | 0 | 3,076 |
| scalar-elementary-log10 | evaluate | yes | 472 | 4,000 | 400,000 | 0 | 3,076 |
| scalar-elementary-log10 | parse-evaluate | yes | 664 | 8,000 | 867,000 | 0 | 3,080 |
| scalar-elementary-mod | evaluate | yes | 621 | 5,000 | 496,000 | 0 | 2,928 |
| scalar-elementary-mod | parse-evaluate | yes | 832 | 9,000 | 965,000 | 0 | 2,948 |
| scalar-elementary-power | evaluate | yes | 644 | 5,000 | 496,000 | 0 | 3,076 |
| scalar-elementary-power | parse-evaluate | yes | 871 | 9,000 | 968,000 | 0 | 3,064 |
| scalar-elementary-quotient | evaluate | yes | 619 | 5,000 | 496,000 | 0 | 2,948 |
| scalar-elementary-quotient | parse-evaluate | yes | 847 | 9,000 | 970,000 | 0 | 2,908 |
| scalar-elementary-sign | evaluate | yes | 611 | 5,000 | 496,000 | 0 | 2,924 |
| scalar-elementary-sign | parse-evaluate | yes | 819 | 9,000 | 963,000 | 0 | 2,956 |
| scalar-elementary-sqrt | evaluate | yes | 447 | 4,000 | 400,000 | 0 | 2,948 |
| scalar-elementary-sqrt | parse-evaluate | yes | 637 | 8,000 | 866,000 | 0 | 2,948 |
| scalar-elementary-sqrtpi | evaluate | yes | 458 | 4,000 | 400,000 | 0 | 2,920 |
| scalar-elementary-sqrtpi | parse-evaluate | yes | 647 | 8,000 | 868,000 | 0 | 2,928 |

The synthetic local-reference rows use the profile's immutable resolver. They do not measure the production worksheet adapter or full-workbook recalculation. The direct f64 oracle checks evaluator projection and is not an independent libm-accuracy implementation.
