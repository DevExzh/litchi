# 0517 baseline/candidate native comparison

Each record compares the same API route and workload in one baseline campaign and its matching candidate campaign. Native rows were validated through `capture.validate`; profile and hardware lanes are excluded.
All phase metrics are in nanoseconds and use nearest-rank p50/p95/p99 plus arithmetic mean over 60 measured rows. RSS is the one whole-child GNU time maximum in KiB.
The p50 intervals use a fixed seed `5339665` and 4000 unpaired route bootstrap iterations. Each route is resampled within each internal repeat, then candidate median / baseline median is calculated. The two internal repeats are not independent processes, so intervals are descriptive only. RSS has no phase-local bootstrap because each child contributes one high-water observation.

## Pair summary

A positive delta is a candidate regression. Adverse flags are directional and use 5% for p50/mean/RSS, 10% for p95, and 15% for p99.

| Baseline | Candidate | Records | Flagged records | Flags |
| --- | --- | ---: | ---: | ---: |
| r1 | after-r1 | 24 | 20 | 50 |
| r2 | after-r2 | 24 | 19 | 43 |

## p50 route values

Each cell is candidate / baseline followed by the absolute p50 values. This table covers all 24 cases in both campaign pairs.

| Pair | Workload | elapsed p50 ratio | publish p50 ratio | RSS ratio |
| --- | --- | ---: | ---: | ---: |
| r1 | p128-k1-file-batch | 0.8864 (651163→577222 ns) | 0.8075 (354722→286441 ns) | 0.9988 (6516→6508 KiB) |
| r1 | p128-k1-file-repeated | 0.8933 (649263→579982 ns) | 0.8096 (353852→286482 ns) | 1.0321 (6604→6816 KiB) |
| r1 | p128-k1-owned-batch | 0.8748 (600853→525643 ns) | 0.7813 (309511→241810 ns) | 1.0121 (7288→7376 KiB) |
| r1 | p128-k1-owned-repeated | 0.9509 (599822→570352 ns) | 0.9221 (309801→285671 ns) | 0.9704 (7292→7076 KiB) |
| r1 | p128-k32-file-batch | 0.8985 (875463→786563 ns) | 0.8291 (399082→330882 ns) | 1.0006 (6692→6696 KiB) |
| r1 | p128-k32-file-repeated | 0.9643 (6101035→5883004 ns) | 0.8305 (404902→336251 ns) | 0.9905 (6724→6660 KiB) |
| r1 | p128-k32-owned-batch | 0.9034 (796633→719712 ns) | 0.8089 (360551→291641 ns) | 1.0101 (7124→7196 KiB) |
| r1 | p128-k32-owned-repeated | 0.9614 (5630692→5413152 ns) | 0.8939 (327552→292792 ns) | 0.9770 (7308→7140 KiB) |
| r1 | p128-k8-file-batch | 0.8938 (686703→613742 ns) | 0.8084 (353362→285671 ns) | 0.9866 (6880→6788 KiB) |
| r1 | p128-k8-file-repeated | 0.9151 (1622777→1484976 ns) | 0.7345 (395142→290251 ns) | 1.0304 (6580→6780 KiB) |
| r1 | p128-k8-owned-batch | 0.8827 (633223→558963 ns) | 0.7836 (311251→243881 ns) | 1.0000 (7296→7296 KiB) |
| r1 | p128-k8-owned-repeated | 0.9042 (1553277→1404526 ns) | 0.6915 (354742→245301 ns) | 1.0183 (7204→7336 KiB) |
| r1 | p512-k1-file-batch | 0.8726 (2306679→2012859 ns) | 0.7805 (1225605→956624 ns) | 1.0000 (6488→6488 KiB) |
| r1 | p512-k1-file-repeated | 0.8801 (2292759→2017928 ns) | 0.7904 (1215045→960344 ns) | 1.0072 (6640→6688 KiB) |
| r1 | p512-k1-owned-batch | 0.8724 (2251449→1964198 ns) | 0.7773 (1174245→912774 ns) | 1.0210 (7048→7196 KiB) |
| r1 | p512-k1-owned-repeated | 0.8761 (2245869→1967718 ns) | 0.7771 (1177835→915344 ns) | 1.0017 (6960→6972 KiB) |
| r1 | p512-k32-file-batch | 0.8873 (2441110→2166009 ns) | 0.7824 (1224715→958264 ns) | 1.0086 (6944→7004 KiB) |
| r1 | p512-k32-file-repeated | 0.9771 (17542400→17141330 ns) | 0.7950 (1211855→963433 ns) | 0.9931 (7004→6956 KiB) |
| r1 | p512-k32-owned-batch | 0.8856 (2387739→2114599 ns) | 0.7767 (1188695→923243 ns) | 0.9721 (7604→7392 KiB) |
| r1 | p512-k32-owned-repeated | 0.9689 (17125228→16591778 ns) | 0.7824 (1179145→922584 ns) | 1.0227 (7392→7560 KiB) |
| r1 | p512-k8-file-batch | 0.8630 (2304639→1988798 ns) | 0.7568 (1204575→911603 ns) | 0.9965 (6956→6932 KiB) |
| r1 | p512-k8-file-repeated | 0.9366 (5438882→5093991 ns) | 0.7944 (1210084→961334 ns) | 1.0048 (6720→6752 KiB) |
| r1 | p512-k8-owned-batch | 0.8778 (2278859→2000298 ns) | 0.7806 (1175195→917353 ns) | 1.0082 (7316→7376 KiB) |
| r1 | p512-k8-owned-repeated | 0.9288 (5415962→5030421 ns) | 0.7884 (1169815→922334 ns) | 0.9843 (7136→7024 KiB) |
| r2 | p128-k1-file-batch | 0.8887 (650743→578302 ns) | 0.8182 (353912→289581 ns) | 0.9921 (6564→6512 KiB) |
| r2 | p128-k1-file-repeated | 0.8939 (649503→580572 ns) | 0.8057 (356322→287072 ns) | 1.0068 (6516→6560 KiB) |
| r2 | p128-k1-owned-batch | 0.9504 (600813→571012 ns) | 0.9128 (312322→285072 ns) | 0.9731 (7280→7084 KiB) |
| r2 | p128-k1-owned-repeated | 0.8748 (603442→527862 ns) | 0.7743 (315322→244151 ns) | 0.9928 (7196→7144 KiB) |
| r2 | p128-k32-file-batch | 0.9068 (866934→786154 ns) | 0.8264 (402222→332381 ns) | 0.9649 (6832→6592 KiB) |
| r2 | p128-k32-file-repeated | 0.9813 (6028874→5916224 ns) | 0.8364 (403042→337121 ns) | 0.9875 (6728→6644 KiB) |
| r2 | p128-k32-owned-batch | 0.8991 (797914→717403 ns) | 0.8050 (361812→291251 ns) | 0.9944 (7204→7164 KiB) |
| r2 | p128-k32-owned-repeated | 0.9630 (5640303→5431352 ns) | 0.7984 (366312→292471 ns) | 0.9989 (7148→7140 KiB) |
| r2 | p128-k8-file-batch | 0.8954 (685113→613472 ns) | 0.8128 (352681→286672 ns) | 0.9866 (6872→6780 KiB) |
| r2 | p128-k8-file-repeated | 0.9336 (1637077→1528436 ns) | 0.8281 (398842→330301 ns) | 0.9958 (6660→6632 KiB) |
| r2 | p128-k8-owned-batch | 0.8850 (633743→560833 ns) | 0.7818 (311722→243711 ns) | 0.9995 (7296→7292 KiB) |
| r2 | p128-k8-owned-repeated | 0.9449 (1533756→1449206 ns) | 0.8999 (324021→291571 ns) | 0.9972 (7216→7196 KiB) |
| r2 | p512-k1-file-batch | 0.8834 (2294219→2026658 ns) | 0.7904 (1219095→963584 ns) | 1.0168 (6436→6544 KiB) |
| r2 | p512-k1-file-repeated | 0.8591 (2356209→2024158 ns) | 0.7857 (1226135→963414 ns) | 1.0423 (6436→6708 KiB) |
| r2 | p512-k1-owned-batch | 0.8770 (2246179→1969808 ns) | 0.7813 (1177785→920194 ns) | 1.0160 (7008→7120 KiB) |
| r2 | p512-k1-owned-repeated | 0.8715 (2256539→1966548 ns) | 0.7776 (1178904→916734 ns) | 0.9994 (6960→6956 KiB) |
| r2 | p512-k32-file-batch | 0.8659 (2496110→2161289 ns) | 0.7679 (1251815→961294 ns) | 0.9875 (7044→6956 KiB) |
| r2 | p512-k32-file-repeated | 0.9677 (17498841→16933799 ns) | 0.7842 (1204825→944814 ns) | 1.0023 (6964→6980 KiB) |
| r2 | p512-k32-owned-batch | 0.8884 (2380250→2114709 ns) | 0.7807 (1180925→921904 ns) | 1.0272 (7344→7544 KiB) |
| r2 | p512-k32-owned-repeated | 0.9700 (17161439→16646447 ns) | 0.7829 (1186475→928943 ns) | 0.9904 (7472→7400 KiB) |
| r2 | p512-k8-file-batch | 0.8659 (2312320→2002158 ns) | 0.7538 (1213285→914584 ns) | 1.0058 (6956→6996 KiB) |
| r2 | p512-k8-file-repeated | 0.9343 (5501072→5139780 ns) | 0.7977 (1217105→970854 ns) | 0.9976 (6708→6692 KiB) |
| r2 | p512-k8-owned-batch | 0.8949 (2268259→2029948 ns) | 0.7896 (1172594→925883 ns) | 1.0187 (7276→7412 KiB) |
| r2 | p512-k8-owned-repeated | 0.9266 (5421372→5023291 ns) | 0.7836 (1173885→919814 ns) | 1.0061 (7240→7284 KiB) |

## Identity and counters

Output identities are checked before comparison. Work and source-read counters are reported per case in JSON; this table shows whether any changed.

| Pair | Identity matches | Work changed | Source read calls changed | Source bytes changed |
| --- | ---: | ---: | ---: | ---: |
| r1 | 24/24 | 24 | 0 | 0 |
| r2 | 24/24 | 24 | 0 | 0 |

## Adverse flags

### r1 → after-r1

| Workload | Flags |
| --- | --- |
| p128-k1-file-repeated | commit_ns.p50 +10.00% (300→330 ns; >5%); commit_ns.mean +8.09% (307.0→331.8333333333333 ns; >5%) |
| p128-k1-owned-batch | commit_ns.p50 +13.33% (300→340 ns; >5%) |
| p128-k1-owned-repeated | commit_ns.p50 +10.00% (300→330 ns; >5%); commit_ns.p95 +18.75% (320→380 ns; >10%); commit_ns.mean +12.01% (299.6666666666667→335.6666666666667 ns; >5%) |
| p128-k32-file-batch | open_ns.p99 +42.12% (16430→23350 ns; >15%); drop_ns.p99 +351.07% (2800→12630 ns; >15%); drop_ns.mean +5.33% (2666.8333333333335→2809.016666666667 ns; >5%) |
| p128-k32-owned-batch | commit_ns.p99 +16.31% (1410→1640 ns; >15%) |
| p128-k32-owned-repeated | open_ns.p95 +94.44% (14200→27610 ns; >10%); open_ns.p99 +39.28% (20370→28371 ns; >15%); open_ns.mean +11.08% (13729.016666666666→15250.45 ns; >5%); commit_ns.p50 +6.90% (1160→1240 ns; >5%); commit_ns.p95 +12.68% (1270→1431 ns; >10%); commit_ns.mean +7.03% (1171.6833333333334→1254.0166666666667 ns; >5%) |
| p128-k8-file-batch | open_ns.p99 +65.42% (16480→27261 ns; >15%); commit_ns.p99 +62.22% (450→730 ns; >15%); commit_ns.mean +5.13% (389.5→409.5 ns; >5%) |
| p128-k8-file-repeated | open_ns.p99 +42.27% (17200→24470 ns; >15%); commit_ns.p99 +1058.82% (510→5910 ns; >15%); commit_ns.mean +18.59% (438.5→520.0166666666667 ns; >5%) |
| p128-k8-owned-batch | commit_ns.p50 +5.26% (380→400 ns; >5%); drop_ns.p99 +43.94% (1320→1900 ns; >15%) |
| p128-k8-owned-repeated | open_ns.p99 +41.09% (14260→20120 ns; >15%) |
| p512-k1-file-repeated | commit_ns.p50 +5.13% (390→410 ns; >5%); commit_ns.p99 +55.38% (650→1010 ns; >15%); commit_ns.mean +7.02% (391.5→419.0 ns; >5%) |
| p512-k1-owned-batch | commit_ns.p50 +9.38% (320→350 ns; >5%); commit_ns.mean +6.26% (332.8333333333333→353.6666666666667 ns; >5%) |
| p512-k1-owned-repeated | open_ns.p99 +28.39% (29520→37901 ns; >15%) |
| p512-k32-file-batch | commit_ns.p99 +23.53% (1530→1890 ns; >15%) |
| p512-k32-file-repeated | open_ns.p99 +49.59% (17200→25730 ns; >15%) |
| p512-k32-owned-batch | open_ns.p50 +49.08% (14730→21960 ns; >5%) |
| p512-k32-owned-repeated | open_ns.p50 +15.62% (14340→16580 ns; >5%); drop_ns.p99 +181.52% (3030→8530 ns; >15%); drop_ns.mean +5.73% (2822.516666666667→2984.2 ns; >5%) |
| p512-k8-file-batch | commit_ns.p99 +24.00% (500→620 ns; >15%); drop_ns.p99 +38.85% (1390→1930 ns; >15%) |
| p512-k8-file-repeated | open_ns.p50 +86.24% (16130→30040 ns; >5%); open_ns.p95 +90.67% (16710→31861 ns; >10%); open_ns.p99 +118.80% (17130→37481 ns; >15%); open_ns.mean +76.76% (16044.733333333334→28360.35 ns; >5%); commit_ns.p50 +12.20% (410→460 ns; >5%); commit_ns.mean +9.85% (417.8333333333333→459.0 ns; >5%) |
| p512-k8-owned-batch | open_ns.p99 +18.04% (29100→34350 ns; >15%); commit_ns.p50 +9.52% (420→460 ns; >5%); commit_ns.p99 +1107.55% (530→6400 ns; >15%); commit_ns.mean +31.63% (424.3333333333333→558.5333333333333 ns; >5%); drop_ns.p99 +21.13% (1420→1720 ns; >15%) |

### r2 → after-r2

| Workload | Flags |
| --- | --- |
| p128-k1-file-repeated | commit_ns.p50 +6.67% (300→320 ns; >5%); commit_ns.mean +6.30% (299.0→317.8333333333333 ns; >5%) |
| p128-k1-owned-batch | open_ns.p50 +5.33% (12560→13230 ns; >5%); open_ns.p99 +29.33% (16610→21481 ns; >15%); open_ns.mean +5.09% (12720.333333333334→13367.433333333332 ns; >5%); commit_ns.p50 +20.00% (300→360 ns; >5%); commit_ns.mean +16.14% (309.8333333333333→359.8333333333333 ns; >5%) |
| p128-k1-owned-repeated | commit_ns.p50 +10.34% (290→320 ns; >5%); commit_ns.mean +5.92% (301.3333333333333→319.1666666666667 ns; >5%) |
| p128-k32-owned-batch | open_ns.p99 +25.48% (27710→34770 ns; >15%) |
| p128-k32-owned-repeated | commit_ns.mean +5.50% (1210.8666666666666→1277.5 ns; >5%) |
| p128-k8-file-batch | commit_ns.p50 +7.69% (390→420 ns; >5%); commit_ns.mean +6.54% (397.8333333333333→423.8333333333333 ns; >5%) |
| p128-k8-file-repeated | commit_ns.p50 +6.98% (430→460 ns; >5%); commit_ns.mean +7.13% (430.3333333333333→461.0 ns; >5%) |
| p128-k8-owned-batch | commit_ns.p50 +15.79% (380→440 ns; >5%); commit_ns.p95 +14.63% (410→470 ns; >10%); commit_ns.mean +15.63% (379.5→438.8333333333333 ns; >5%) |
| p128-k8-owned-repeated | open_ns.p95 +97.09% (14080→27750 ns; >10%); open_ns.p99 +53.28% (18580→28480 ns; >15%); open_ns.mean +24.50% (13501.683333333332→16809.683333333334 ns; >5%); drop_ns.p50 +5.60% (1250→1320 ns; >5%); drop_ns.mean +5.56% (1252.8333333333333→1322.5 ns; >5%) |
| p512-k1-file-repeated | open_ns.p99 +77.29% (31750→56290 ns; >15%) |
| p512-k1-owned-batch | open_ns.p99 +19.53% (31080→37150 ns; >15%); commit_ns.p50 +8.82% (340→370 ns; >5%); commit_ns.mean +6.47% (350.5→373.1666666666667 ns; >5%) |
| p512-k1-owned-repeated | commit_ns.p50 +6.45% (310→330 ns; >5%) |
| p512-k32-file-batch | open_ns.p99 +27.70% (16680→21300 ns; >15%) |
| p512-k32-file-repeated | commit_ns.p99 +16.89% (1480→1730 ns; >15%) |
| p512-k32-owned-batch | open_ns.p50 +71.69% (13880→23830 ns; >5%); open_ns.p99 +31.49% (29280→38500 ns; >15%); drop_ns.p99 +204.14% (2900→8820 ns; >15%) |
| p512-k32-owned-repeated | commit_ns.p99 +22.45% (1470→1800 ns; >15%); drop_ns.p99 +208.21% (3410→10510 ns; >15%) |
| p512-k8-file-batch | commit_ns.p99 +90.00% (500→950 ns; >15%) |
| p512-k8-file-repeated | commit_ns.p99 +310.14% (690→2830 ns; >15%); commit_ns.mean +5.02% (498.5→523.5 ns; >5%); drop_ns.p99 +162.36% (1780→4670 ns; >15%); drop_ns.mean +5.27% (1318.5→1388.0 ns; >5%) |
| p512-k8-owned-batch | commit_ns.p50 +6.98% (430→460 ns; >5%); commit_ns.p99 +82.35% (510→930 ns; >15%); commit_ns.mean +8.40% (434.6666666666667→471.1666666666667 ns; >5%) |

The JSON contains all 48 case comparisons, every phase/statistic, RSS, bindings, output identities, and bootstrap metadata.
