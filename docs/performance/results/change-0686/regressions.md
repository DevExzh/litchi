# Native paired regression review triggers

Every row below exceeds 5% p50 in both candidate/baseline pairs.
Single-leg changes and all tail metrics remain in the full comparison.

| Case | Source | Phase | A1 | B1 | B2 | A2 | Paired change % |
|---|---|---|---:|---:|---:|---:|---:|
| 54016-stored-524288 | owned | q2 | 552573 | 597478 | 598173 | 551908 | +8.13 / +8.38 |
| 54016-stored-524288 | file | q2 | 564252 | 607828 | 613068 | 571678 | +7.72 / +7.24 |
| Plan1-stored-32768 | owned | q2 | 33910 | 35860 | 35775 | 33905.5 | +5.75 / +5.51 |
| formula-refusal-2097152 | owned | q1 | 4040 | 4390 | 4330 | 3990 | +8.66 / +8.52 |
| formula-refusal-2097152 | owned | q3 | 4410 | 4780 | 4740 | 4430 | +8.39 / +7.00 |
| formula-refusal-2097152 | owned | q8 | 4300 | 4690 | 4670 | 4310 | +9.07 / +8.35 |
| formula-refusal-2097152 | owned | q3-to-q8-mean | 4325 | 4725.92 | 4693.42 | 4336.75 | +9.27 / +8.22 |
| formula-refusal-2097152 | file | q1 | 5540 | 5930 | 5940 | 5520 | +7.04 / +7.61 |
| formula-refusal-2097152 | file | q2 | 6140 | 6500 | 6540 | 6170 | +5.86 / +6.00 |
| formula-refusal-2097152 | file | q3 | 5920 | 6335 | 6350 | 5950 | +7.01 / +6.72 |
| formula-refusal-2097152 | file | q8 | 5820 | 6240 | 6230 | 5790.5 | +7.22 / +7.59 |
| formula-refusal-2097152 | file | q3-to-q8-mean | 5860.83 | 6269.25 | 6275.83 | 5840 | +6.97 / +7.46 |
| synthetic-70000-default | owned | q2 | 1.07225e+06 | 1.14422e+06 | 1.14493e+06 | 1.07204e+06 | +6.71 / +6.80 |
| synthetic-70000-default | file | q2 | 1.07604e+06 | 1.1537e+06 | 1.15519e+06 | 1.07706e+06 | +7.22 / +7.25 |
| synthetic-100000-default | owned | q2 | 1.44974e+06 | 1.55227e+06 | 1.55245e+06 | 1.4505e+06 | +7.07 / +7.03 |
| synthetic-100000-default | file | q2 | 1.45243e+06 | 1.56076e+06 | 1.56044e+06 | 1.45468e+06 | +7.46 / +7.27 |
