# Final native review triggers

Every row below exceeds 5% in both paired p50 comparisons. Units are ns.
Full A/A, means, p95 and p99 remain in `comparison.json`; p99 of 30 is the sample maximum.

| Route | Phase | A1 | B1 | B2 | A2 | B1/A1 % | B2/A2 % |
|---|---|---:|---:|---:|---:|---:|---:|
| Simple-missing-owned-native-prepared | q1 | 455 | 520 | 520 | 470 | 14.29 | 10.64 |
| Simple-missing-owned-native-prepared | q2-build-trigger | 885 | 935 | 950 | 840 | 5.65 | 13.10 |
| Simple-missing-owned-native-prepared | queries-2 | 1340 | 1460 | 1485 | 1310 | 8.96 | 13.36 |
| Simple-missing-owned-native-prepared | queries-3 | 1425 | 1555 | 1575 | 1400 | 9.12 | 12.50 |
| Simple-stored-owned-native-prepared | q1 | 805 | 870 | 880 | 785 | 8.07 | 12.10 |
| Simple-stored-owned-native-prepared | q2-build-trigger | 860 | 930 | 940 | 840 | 8.14 | 11.90 |
| Simple-stored-owned-native-prepared | q3-warm-repeat-q1 | 540 | 625 | 610 | 535 | 15.74 | 14.02 |
| Simple-stored-owned-native-prepared | queries-2 | 1655 | 1800 | 1840 | 1625 | 8.76 | 13.23 |
| Simple-stored-owned-native-prepared | queries-3 | 2200 | 2435 | 2440 | 2160 | 10.68 | 12.96 |
| formula-refusal-file-native-prepared | q1 | 5545 | 5845 | 5960 | 5635 | 5.41 | 5.77 |
| formula-refusal-file-native-prepared | q2-build-trigger | 6100 | 6550 | 6610 | 6155 | 7.38 | 7.39 |
| formula-refusal-file-native-prepared | q3-warm-repeat-q1 | 5980 | 6385 | 6485 | 6055 | 6.77 | 7.10 |
| formula-refusal-file-native-prepared | queries-2 | 11655 | 12405.5 | 12605 | 11780 | 6.44 | 7.00 |
| formula-refusal-file-native-prepared | queries-3 | 17660 | 18785 | 19135 | 17820 | 6.37 | 7.38 |
| formula-refusal-owned-native-prepared | q1 | 4030 | 4450 | 4450 | 4050 | 10.42 | 9.88 |
| formula-refusal-owned-native-prepared | q2-build-trigger | 4550 | 5030 | 4995 | 4550 | 10.55 | 9.78 |
| formula-refusal-owned-native-prepared | q3-warm-repeat-q1 | 4395 | 4825 | 4795 | 4440 | 9.78 | 8.00 |
| formula-refusal-owned-native-prepared | queries-2 | 8615 | 9480 | 9470 | 8580 | 10.04 | 10.37 |
| formula-refusal-owned-native-prepared | queries-3 | 13005 | 14300 | 14280 | 13035 | 9.96 | 9.55 |
| formula-refusal-owned-native-visit | visit | 15955.5 | 16795 | 16815 | 15840 | 5.26 | 6.16 |
