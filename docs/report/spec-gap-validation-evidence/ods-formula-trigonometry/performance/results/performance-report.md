# ODS trigonometry performance profile

The baseline timing lane contains the ten matched arithmetic/ROUND controls. The candidate lane contains all named scalar, array, and synthetic local-reference cases. New trigonometric cases have no baseline timing comparison because the baseline does not implement them. Time deltas are candidate minus baseline; positive values are slower. Samples are p50 across fresh child processes, and each row's time is normalized by the harness repeat count.

## Matched controls

| case | phase | baseline time ns/repeat | candidate time ns/repeat | time delta | baseline alloc calls | candidate alloc calls | baseline RSS KiB | candidate RSS KiB |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-arithmetic | evaluate | 45,435 | 45,105 | -0.7% | 88 | 88 | 3,120 | 3,200 |
| array-control-16x16-arithmetic | parse-evaluate | 60,510 | 60,905 | +0.7% | 352 | 352 | 3,140 | 3,164 |
| array-control-16x16-round | evaluate | 178,255 | 172,320 | -3.3% | 2,144 | 2,144 | 3,124 | 3,208 |
| array-control-16x16-round | parse-evaluate | 196,061 | 189,780 | -3.2% | 2,412 | 2,412 | 3,136 | 3,164 |
| array-control-4x4-arithmetic | evaluate | 4,400 | 4,492 | +2.1% | 1,120 | 1,120 | 2,892 | 2,940 |
| array-control-4x4-arithmetic | parse-evaluate | 5,265 | 5,334 | +1.3% | 2,240 | 2,240 | 2,884 | 2,940 |
| array-control-4x4-round | evaluate | 13,135 | 12,855 | -2.1% | 3,840 | 3,840 | 2,880 | 2,936 |
| array-control-4x4-round | parse-evaluate | 14,029 | 13,705 | -2.3% | 5,040 | 5,040 | 2,904 | 2,904 |
| reference-array-arithmetic | evaluate | 3,175 | 3,175 | +0.0% | 800 | 800 | 2,888 | 2,952 |
| reference-array-arithmetic | parse-evaluate | 3,513 | 3,509 | -0.1% | 1,280 | 1,280 | 2,864 | 2,940 |
| reference-array-round | evaluate | 12,155 | 11,804 | -2.9% | 3,520 | 3,520 | 2,900 | 2,896 |
| reference-array-round | parse-evaluate | 12,532 | 12,134 | -3.2% | 4,080 | 4,080 | 2,892 | 2,936 |
| reference-scalar-arithmetic | evaluate | 737 | 758 | +2.8% | 5,000 | 5,000 | 2,880 | 2,948 |
| reference-scalar-arithmetic | parse-evaluate | 985 | 1,018 | +3.4% | 10,000 | 10,000 | 2,880 | 2,944 |
| reference-scalar-round | evaluate | 1,911 | 1,880 | -1.6% | 12,000 | 12,000 | 2,892 | 2,900 |
| reference-scalar-round | parse-evaluate | 2,224 | 2,192 | -1.4% | 18,000 | 18,000 | 2,880 | 2,904 |
| scalar-control-arithmetic | evaluate | 585 | 580 | -0.9% | 5,000 | 5,000 | 2,836 | 2,680 |
| scalar-control-arithmetic | parse-evaluate | 730 | 713 | -2.3% | 8,000 | 8,000 | 2,704 | 2,680 |
| scalar-control-round | evaluate | 743 | 742 | -0.1% | 5,000 | 5,000 | 2,844 | 2,732 |
| scalar-control-round | parse-evaluate | 980 | 978 | -0.2% | 9,000 | 9,000 | 2,840 | 2,664 |

## Candidate trigonometric and representative workloads

| case | phase | supported | time ns/repeat | alloc calls | requested bytes | result-live memory | RSS KiB |
| --- | --- | :---: | ---: | ---: | ---: | ---: | ---: |
| array-control-16x16-acosh | evaluate | yes | 133,760 | 2,144 | 1,279,328 | 22,528 | 3,488 |
| array-control-16x16-acosh | parse-evaluate | yes | 150,013 | 2,412 | 1,730,948 | 22,528 | 3,524 |
| array-control-16x16-asin | evaluate | yes | 126,195 | 2,144 | 1,279,328 | 22,528 | 3,460 |
| array-control-16x16-asin | parse-evaluate | yes | 142,893 | 2,412 | 1,730,944 | 22,528 | 3,340 |
| array-control-16x16-cos | evaluate | yes | 124,158 | 2,144 | 1,279,328 | 22,528 | 3,492 |
| array-control-16x16-cos | parse-evaluate | yes | 140,933 | 2,412 | 1,730,940 | 22,528 | 3,460 |
| array-control-16x16-sin | evaluate | yes | 124,460 | 2,144 | 1,279,328 | 22,528 | 3,492 |
| array-control-16x16-sin | parse-evaluate | yes | 142,235 | 2,412 | 1,730,940 | 22,528 | 3,464 |
| array-control-16x16-tanh | evaluate | yes | 128,323 | 2,144 | 1,279,328 | 22,528 | 3,420 |
| array-control-16x16-tanh | parse-evaluate | yes | 145,868 | 2,412 | 1,730,944 | 22,528 | 3,460 |
| array-control-4x4-acosh | evaluate | yes | 10,235 | 3,840 | 1,624,960 | 1,408 | 3,236 |
| array-control-4x4-acosh | parse-evaluate | yes | 11,026 | 5,040 | 2,155,600 | 1,408 | 3,268 |
| array-control-4x4-asin | evaluate | yes | 9,738 | 3,840 | 1,624,960 | 1,408 | 3,164 |
| array-control-4x4-asin | parse-evaluate | yes | 10,629 | 5,040 | 2,155,520 | 1,408 | 3,144 |
| array-control-4x4-cos | evaluate | yes | 9,653 | 3,840 | 1,624,960 | 1,408 | 3,196 |
| array-control-4x4-cos | parse-evaluate | yes | 10,564 | 5,040 | 2,155,440 | 1,408 | 3,220 |
| array-control-4x4-sin | evaluate | yes | 9,787 | 3,840 | 1,624,960 | 1,408 | 3,212 |
| array-control-4x4-sin | parse-evaluate | yes | 10,577 | 5,040 | 2,155,440 | 1,408 | 3,172 |
| array-control-4x4-tanh | evaluate | yes | 9,848 | 3,840 | 1,624,960 | 1,408 | 3,176 |
| array-control-4x4-tanh | parse-evaluate | yes | 10,770 | 5,040 | 2,155,520 | 1,408 | 3,212 |
| reference-array-acosh | evaluate | yes | 8,797 | 3,440 | 1,069,440 | 1,408 | 3,256 |
| reference-array-acosh | parse-evaluate | yes | 9,235 | 4,000 | 1,178,480 | 1,408 | 3,268 |
| reference-array-sin | evaluate | yes | 8,266 | 3,440 | 1,069,440 | 1,408 | 3,220 |
| reference-array-sin | parse-evaluate | yes | 8,702 | 4,000 | 1,178,320 | 1,408 | 3,156 |
| reference-array-tanh | evaluate | yes | 8,490 | 3,440 | 1,069,440 | 1,408 | 3,240 |
| reference-array-tanh | parse-evaluate | yes | 8,869 | 4,000 | 1,178,400 | 1,408 | 3,208 |
| reference-scalar-acosh | evaluate | yes | 1,532 | 11,000 | 2,152,000 | 0 | 3,292 |
| reference-scalar-acosh | parse-evaluate | yes | 1,856 | 17,000 | 3,510,000 | 0 | 3,256 |
| reference-scalar-atanh | evaluate | yes | 1,521 | 11,000 | 2,152,000 | 0 | 3,260 |
| reference-scalar-atanh | parse-evaluate | yes | 1,855 | 17,000 | 3,510,000 | 0 | 3,212 |
| reference-scalar-sin | evaluate | yes | 1,502 | 11,000 | 2,152,000 | 0 | 3,212 |
| reference-scalar-sin | parse-evaluate | yes | 1,808 | 17,000 | 3,508,000 | 0 | 3,212 |
| scalar-trig-acos | evaluate | yes | 460 | 4,000 | 400,000 | 0 | 2,956 |
| scalar-trig-acos | parse-evaluate | yes | 650 | 8,000 | 866,000 | 0 | 2,980 |
| scalar-trig-acosh | evaluate | yes | 490 | 4,000 | 400,000 | 0 | 3,260 |
| scalar-trig-acosh | parse-evaluate | yes | 676 | 8,000 | 867,000 | 0 | 3,192 |
| scalar-trig-acot | evaluate | yes | 465 | 4,000 | 400,000 | 0 | 2,996 |
| scalar-trig-acot | parse-evaluate | yes | 655 | 8,000 | 866,000 | 0 | 3,008 |
| scalar-trig-acoth | evaluate | yes | 486 | 4,000 | 400,000 | 0 | 3,160 |
| scalar-trig-acoth | parse-evaluate | yes | 675 | 8,000 | 867,000 | 0 | 3,256 |
| scalar-trig-asin | evaluate | yes | 464 | 4,000 | 400,000 | 0 | 2,980 |
| scalar-trig-asin | parse-evaluate | yes | 647 | 8,000 | 866,000 | 0 | 2,980 |
| scalar-trig-asinh | evaluate | yes | 488 | 4,000 | 400,000 | 0 | 3,012 |
| scalar-trig-asinh | parse-evaluate | yes | 675 | 8,000 | 867,000 | 0 | 3,004 |
| scalar-trig-atan | evaluate | yes | 461 | 4,000 | 400,000 | 0 | 3,028 |
| scalar-trig-atan | parse-evaluate | yes | 650 | 8,000 | 866,000 | 0 | 2,972 |
| scalar-trig-atan2 | evaluate | yes | 644 | 5,000 | 496,000 | 0 | 2,980 |
| scalar-trig-atan2 | parse-evaluate | yes | 851 | 9,000 | 961,000 | 0 | 3,012 |
| scalar-trig-atanh | evaluate | yes | 480 | 4,000 | 400,000 | 0 | 3,012 |
| scalar-trig-atanh | parse-evaluate | yes | 675 | 8,000 | 867,000 | 0 | 2,996 |
| scalar-trig-cos | evaluate | yes | 459 | 4,000 | 400,000 | 0 | 3,008 |
| scalar-trig-cos | parse-evaluate | yes | 640 | 8,000 | 865,000 | 0 | 2,980 |
| scalar-trig-cosh | evaluate | yes | 467 | 4,000 | 400,000 | 0 | 3,012 |
| scalar-trig-cosh | parse-evaluate | yes | 652 | 8,000 | 866,000 | 0 | 3,012 |
| scalar-trig-cot | evaluate | yes | 464 | 4,000 | 400,000 | 0 | 2,988 |
| scalar-trig-cot | parse-evaluate | yes | 648 | 8,000 | 865,000 | 0 | 3,000 |
| scalar-trig-coth | evaluate | yes | 470 | 4,000 | 400,000 | 0 | 2,976 |
| scalar-trig-coth | parse-evaluate | yes | 655 | 8,000 | 866,000 | 0 | 2,988 |
| scalar-trig-csc | evaluate | yes | 461 | 4,000 | 400,000 | 0 | 2,996 |
| scalar-trig-csc | parse-evaluate | yes | 643 | 8,000 | 865,000 | 0 | 2,968 |
| scalar-trig-csch | evaluate | yes | 471 | 4,000 | 400,000 | 0 | 3,000 |
| scalar-trig-csch | parse-evaluate | yes | 661 | 8,000 | 866,000 | 0 | 2,988 |
| scalar-trig-degrees | evaluate | yes | 457 | 4,000 | 400,000 | 0 | 2,664 |
| scalar-trig-degrees | parse-evaluate | yes | 649 | 8,000 | 869,000 | 0 | 2,680 |
| scalar-trig-pi | evaluate | yes | 392 | 4,000 | 400,000 | 0 | 2,636 |
| scalar-trig-pi | parse-evaluate | yes | 491 | 6,000 | 789,000 | 0 | 2,668 |
| scalar-trig-radians | evaluate | yes | 455 | 4,000 | 400,000 | 0 | 2,716 |
| scalar-trig-radians | parse-evaluate | yes | 651 | 8,000 | 870,000 | 0 | 2,664 |
| scalar-trig-sec | evaluate | yes | 456 | 4,000 | 400,000 | 0 | 2,984 |
| scalar-trig-sec | parse-evaluate | yes | 643 | 8,000 | 865,000 | 0 | 3,016 |
| scalar-trig-sech | evaluate | yes | 473 | 4,000 | 400,000 | 0 | 2,992 |
| scalar-trig-sech | parse-evaluate | yes | 657 | 8,000 | 866,000 | 0 | 3,012 |
| scalar-trig-sin | evaluate | yes | 460 | 4,000 | 400,000 | 0 | 2,968 |
| scalar-trig-sin | parse-evaluate | yes | 643 | 8,000 | 865,000 | 0 | 3,004 |
| scalar-trig-sinh | evaluate | yes | 471 | 4,000 | 400,000 | 0 | 3,004 |
| scalar-trig-sinh | parse-evaluate | yes | 654 | 8,000 | 866,000 | 0 | 3,004 |
| scalar-trig-tan | evaluate | yes | 465 | 4,000 | 400,000 | 0 | 3,004 |
| scalar-trig-tan | parse-evaluate | yes | 643 | 8,000 | 865,000 | 0 | 2,980 |
| scalar-trig-tanh | evaluate | yes | 465 | 4,000 | 400,000 | 0 | 2,984 |
| scalar-trig-tanh | parse-evaluate | yes | 655 | 8,000 | 866,000 | 0 | 2,976 |

The synthetic local-reference rows use the profile's immutable resolver. They do not measure the production worksheet adapter or full-workbook recalculation. The direct f64 oracle checks evaluator projection and is not an independent libm-accuracy implementation.
