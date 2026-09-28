# 0816 finite-budget delayed-source scaling

This report is an offline replay of the retained native, observer, and qualification receipts.
The measurements use warm in-memory bytes and finite hierarchical execution budgets.
Width-normalized efficiency, source controls and Amdahl fits are descriptive diagnostics;
they do not infer active worker counts or establish a causal decomposition.

- Native reports/samples: 432 / 12960
- Observer reports/samples: 144 / 288
- Qualification reports/samples: 72 / 72
- Scaling rows: 72; bootstrap: 10000 resamples, seed 816816

| Route | Shape | State | Source | Width | p50 ns | Speedup | Efficiency | CPU/wall | RSS KiB | Flags |
|---|---|---|---|---:|---:|---:|---:|---:|---:|---|
| cfb | large | fresh | local | 1 | 1.21883e+06 | 1 | 1 | 1.00054 | 36684 | spread |
| cfb | large | fresh | local | 2 | 1.04908e+06 | 1.15849 | 0.579247 | 1.89632 | 36834 | spread,tail,vs-width1 |
| cfb | large | fresh | local | 4 | 677118 | 1.81899 | 0.454748 | 3.53467 | 36748 | spread,tail |
| cfb | large | fresh | local | 8 | 582818 | 2.11472 | 0.26434 | 5.07038 | 36922 | spread,tail |
| cfb | large | fresh | capped | 1 | 1.22385e+06 | 1 | 1 | 1.00036 | 36620 | spread,tail |
| cfb | large | fresh | capped | 2 | 1.07919e+06 | 1.13239 | 0.566195 | 1.89934 | 36918 | spread,tail,vs-width1 |
| cfb | large | fresh | capped | 4 | 684724 | 1.79869 | 0.449673 | 3.52387 | 36712 | spread,tail |
| cfb | large | fresh | capped | 8 | 579988 | 2.11057 | 0.263821 | 5.19352 | 37052 | spread,tail |
| cfb | large | fresh | delayed | 1 | 4.00476e+07 | 1 | 1 | 0.0421048 | 36682 | - |
| cfb | large | fresh | delayed | 2 | 2.08832e+07 | 1.91795 | 0.958974 | 0.136519 | 36792 | - |
| cfb | large | fresh | delayed | 4 | 1.05695e+07 | 3.79143 | 0.947857 | 0.293905 | 36692 | - |
| cfb | large | fresh | delayed | 8 | 5.44092e+06 | 7.36375 | 0.920468 | 0.640796 | 36662 | - |
| cfb | large | primed | local | 1 | 1.15723e+06 | 1 | 1 | 1.00045 | 36652 | spread,tail |
| cfb | large | primed | local | 2 | 1.02321e+06 | 1.12629 | 0.563145 | 1.86484 | 36908 | spread,tail,vs-width1 |
| cfb | large | primed | local | 4 | 648223 | 1.80584 | 0.45146 | 3.33718 | 36620 | spread,tail |
| cfb | large | primed | local | 8 | 501462 | 2.29595 | 0.286994 | 5.32389 | 37180 | spread,tail |
| cfb | large | primed | capped | 1 | 1.16735e+06 | 1 | 1 | 1.00074 | 36606 | spread,tail |
| cfb | large | primed | capped | 2 | 1.03296e+06 | 1.14104 | 0.570518 | 1.8543 | 36820 | spread,tail,vs-width1 |
| cfb | large | primed | capped | 4 | 599698 | 1.96716 | 0.49179 | 3.45462 | 36710 | spread,tail |
| cfb | large | primed | capped | 8 | 521038 | 2.24312 | 0.28039 | 5.36698 | 36984 | spread,tail |
| cfb | large | primed | delayed | 1 | 4.00171e+07 | 1 | 1 | 0.0407846 | 36620 | spread |
| cfb | large | primed | delayed | 2 | 2.07736e+07 | 1.92754 | 0.96377 | 0.129195 | 36668 | - |
| cfb | large | primed | delayed | 4 | 1.05062e+07 | 3.81095 | 0.952736 | 0.269763 | 36684 | spread |
| cfb | large | primed | delayed | 8 | 5.34657e+06 | 7.48992 | 0.93624 | 0.592242 | 36844 | - |
| cfb | mixed | fresh | local | 1 | 1.17242e+06 | 1 | 1 | 1.00056 | 35596 | spread,tail |
| cfb | mixed | fresh | local | 2 | 1.17872e+06 | 0.997693 | 0.498847 | 1.00054 | 35598 | spread,tail,negative |
| cfb | mixed | fresh | local | 4 | 1.17451e+06 | 1.0003 | 0.250075 | 1.00043 | 35596 | spread,tail,vs-width1 |
| cfb | mixed | fresh | local | 8 | 1.17414e+06 | 0.996558 | 0.12457 | 1.00056 | 35612 | spread,tail,negative |
| cfb | mixed | fresh | capped | 1 | 1.18464e+06 | 1 | 1 | 1.00051 | 35654 | spread |
| cfb | mixed | fresh | capped | 2 | 1.17893e+06 | 1.00418 | 0.502092 | 1.00055 | 35658 | spread,vs-width1 |
| cfb | mixed | fresh | capped | 4 | 1.17769e+06 | 1.00187 | 0.250467 | 1.00046 | 35598 | spread,vs-width1 |
| cfb | mixed | fresh | capped | 8 | 1.1792e+06 | 1.00147 | 0.125184 | 1.0004 | 35596 | spread,tail,vs-width1 |
| cfb | mixed | fresh | delayed | 1 | 3.91144e+07 | 1 | 1 | 0.0416177 | 35654 | spread |
| cfb | mixed | fresh | delayed | 2 | 3.91133e+07 | 0.999838 | 0.499919 | 0.0418597 | 35644 | spread,negative |
| cfb | mixed | fresh | delayed | 4 | 3.91053e+07 | 1.00021 | 0.250052 | 0.0413664 | 35596 | spread |
| cfb | mixed | fresh | delayed | 8 | 3.9113e+07 | 1.00027 | 0.125033 | 0.0416786 | 35646 | spread |
| parts | large | fresh | local | 1 | 777309 | 1 | 1 | 1.00047 | 29066 | - |
| parts | large | fresh | local | 2 | 1.35996e+06 | 0.571682 | 0.285841 | 1.81853 | 37400 | spread,tail,negative,vs-width1 |
| parts | large | fresh | local | 4 | 816080 | 0.951872 | 0.237968 | 3.23938 | 37314 | spread,tail,negative,vs-width1 |
| parts | large | fresh | local | 8 | 785309 | 0.989491 | 0.123686 | 3.8086 | 37748 | spread,tail,negative,vs-width1 |
| parts | large | fresh | capped | 1 | 774904 | 1 | 1 | 1.00048 | 29036 | - |
| parts | large | fresh | capped | 2 | 1.39582e+06 | 0.55481 | 0.277405 | 1.8 | 37276 | spread,tail,negative,vs-width1 |
| parts | large | fresh | capped | 4 | 824049 | 0.941273 | 0.235318 | 3.19306 | 37356 | spread,tail,negative,vs-width1 |
| parts | large | fresh | capped | 8 | 785209 | 0.986909 | 0.123364 | 3.85731 | 37784 | spread,tail,negative,vs-width1 |
| parts | large | fresh | delayed | 1 | 1.05563e+07 | 1 | 1 | 0.086481 | 28876 | spread |
| parts | large | fresh | delayed | 2 | 6.41032e+06 | 1.64839 | 0.824195 | 0.43453 | 37160 | vs-width1 |
| parts | large | fresh | delayed | 4 | 3.34246e+06 | 3.1591 | 0.789776 | 0.86415 | 37428 | spread,vs-width1 |
| parts | large | fresh | delayed | 8 | 1.86038e+06 | 5.6744 | 0.7093 | 1.67898 | 37742 | spread,tail,vs-width1 |
| parts | large | primed | local | 1 | 2995 | 1 | 1 | 1.13129 | 29048 | spread,tail |
| parts | large | primed | local | 2 | 194146 | 0.0154654 | 0.00773268 | 1.36282 | 37362 | spread,tail,negative,vs-width1 |
| parts | large | primed | local | 4 | 181671 | 0.0165163 | 0.00412908 | 1.97213 | 37388 | spread,tail,negative,vs-width1 |
| parts | large | primed | local | 8 | 234101 | 0.0127393 | 0.00159241 | 2.63319 | 37822 | spread,tail,negative,vs-width1 |
| parts | large | primed | capped | 1 | 3010 | 1 | 1 | 1.1073 | 29066 | spread,tail |
| parts | large | primed | capped | 2 | 189311 | 0.0159345 | 0.00796724 | 1.36807 | 37458 | spread,tail,negative,vs-width1 |
| parts | large | primed | capped | 4 | 181616 | 0.0162854 | 0.00407135 | 1.95161 | 37348 | spread,tail,negative,vs-width1 |
| parts | large | primed | capped | 8 | 238317 | 0.012749 | 0.00159362 | 2.6173 | 37750 | spread,tail,negative,vs-width1 |
| parts | large | primed | delayed | 1 | 3035 | 1 | 1 | 1.11515 | 28970 | spread,tail |
| parts | large | primed | delayed | 2 | 215102 | 0.0140431 | 0.00702154 | 1.33291 | 37340 | spread,tail,negative,vs-width1 |
| parts | large | primed | delayed | 4 | 201316 | 0.0147942 | 0.00369854 | 1.8931 | 37674 | spread,tail,negative,vs-width1 |
| parts | large | primed | delayed | 8 | 247712 | 0.0119077 | 0.00148846 | 2.56938 | 37742 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | local | 1 | 759499 | 1 | 1 | 1.00051 | 28012 | - |
| parts | mixed | fresh | local | 2 | 758108 | 1.00142 | 0.500712 | 1.00048 | 28238 | - |
| parts | mixed | fresh | local | 4 | 758664 | 1.0001 | 0.250025 | 1.0005 | 28158 | - |
| parts | mixed | fresh | local | 8 | 759634 | 0.999023 | 0.124878 | 1.00036 | 28076 | spread,negative |
| parts | mixed | fresh | capped | 1 | 759504 | 1 | 1 | 1.00026 | 28232 | - |
| parts | mixed | fresh | capped | 2 | 758339 | 1.00189 | 0.500946 | 1.00047 | 28236 | spread |
| parts | mixed | fresh | capped | 4 | 757939 | 1.00351 | 0.250877 | 1.00049 | 28202 | - |
| parts | mixed | fresh | capped | 8 | 759444 | 1.00039 | 0.125049 | 1.0005 | 28234 | spread,tail,vs-width1 |
| parts | mixed | fresh | delayed | 1 | 1.05311e+07 | 1 | 1 | 0.0835079 | 28086 | - |
| parts | mixed | fresh | delayed | 2 | 1.0531e+07 | 0.999889 | 0.499945 | 0.0835928 | 28202 | negative |
| parts | mixed | fresh | delayed | 4 | 1.05291e+07 | 1.00006 | 0.250016 | 0.0833311 | 28204 | - |
| parts | mixed | fresh | delayed | 8 | 1.05305e+07 | 0.999988 | 0.124998 | 0.0834394 | 28262 | negative |

Observer source counters are retained in `observer.csv` and are never pooled into native timing.
