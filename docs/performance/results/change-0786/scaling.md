# 0786 finite execution-budget scaling

This report is an offline replay of the retained native, observer, and qualification receipts.
The measurements use warm in-memory bytes and finite hierarchical execution budgets.
Requested-width efficiency and Amdahl fits are descriptive; they do not infer active worker counts
or establish a causal decomposition.

- Native reports/samples: 720 / 21600
- Observer reports/samples: 240 / 480
- Qualification reports/samples: 120 / 120
- Scaling rows: 120; bootstrap: 10000 resamples, seed 786078

| Route | Shape | State | Floor | Width | p50 ns | Speedup | Efficiency | CPU/wall | RSS KiB | Flags |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|
| cfb | large | fresh | 0 | 1 | 1.23367e+06 | 1 | 1 | 1.0005 | 36664 | spread,tail |
| cfb | large | fresh | 0 | 2 | 1.2482e+06 | 0.989792 | 0.494896 | 1.85101 | 36636 | spread,tail,negative,vs-width1 |
| cfb | large | fresh | 0 | 4 | 827383 | 1.49378 | 0.373444 | 3.3397 | 36668 | spread,tail |
| cfb | large | fresh | 0 | 8 | 643298 | 1.91937 | 0.239922 | 5.79555 | 36660 | spread,tail |
| cfb | large | fresh | 0 | 32 | 828873 | 1.49718 | 0.0467869 | 10.4003 | 44200 | spread,tail,vs-width1 |
| cfb | large | fresh | 65536 | 1 | 1.24371e+06 | 1 | 1 | 1.00041 | 36640 | spread,tail |
| cfb | large | fresh | 65536 | 2 | 1.22537e+06 | 1.00648 | 0.50324 | 1.86142 | 36756 | spread,tail,vs-width1 |
| cfb | large | fresh | 65536 | 4 | 807358 | 1.55104 | 0.38776 | 3.38133 | 36634 | spread,tail |
| cfb | large | fresh | 65536 | 8 | 585112 | 2.10564 | 0.263205 | 5.97758 | 36654 | spread,tail |
| cfb | large | fresh | 65536 | 32 | 815824 | 1.52969 | 0.0478029 | 10.2087 | 44076 | spread,rss-spread,tail,vs-width1 |
| cfb | large | primed | 0 | 1 | 1.16297e+06 | 1 | 1 | 1.00054 | 36680 | spread,tail |
| cfb | large | primed | 0 | 2 | 1.20921e+06 | 0.973085 | 0.486543 | 1.81633 | 36702 | spread,tail,negative,vs-width1 |
| cfb | large | primed | 0 | 4 | 774238 | 1.52052 | 0.380129 | 3.23256 | 36616 | spread,tail |
| cfb | large | primed | 0 | 8 | 487627 | 2.3949 | 0.299362 | 6.11322 | 36650 | spread,tail |
| cfb | large | primed | 0 | 32 | 366806 | 3.16425 | 0.0988827 | 18.9333 | 44090 | spread,tail,vs-width1 |
| cfb | large | primed | 65536 | 1 | 1.17032e+06 | 1 | 1 | 1.00086 | 36616 | spread,tail |
| cfb | large | primed | 65536 | 2 | 1.18066e+06 | 1.00016 | 0.500079 | 1.82178 | 36674 | spread,tail,vs-width1 |
| cfb | large | primed | 65536 | 4 | 735853 | 1.59043 | 0.397607 | 3.30941 | 36666 | spread,tail |
| cfb | large | primed | 65536 | 8 | 499482 | 2.3556 | 0.29445 | 6.04215 | 36634 | spread,tail |
| cfb | large | primed | 65536 | 32 | 370722 | 3.19134 | 0.0997292 | 18.7975 | 43792 | spread,tail,vs-width1 |
| cfb | mixed | fresh | 0 | 1 | 1.1903e+06 | 1 | 1 | 1.00049 | 35592 | spread,tail |
| cfb | mixed | fresh | 0 | 2 | 1.20492e+06 | 0.98798 | 0.49399 | 1.83182 | 35602 | spread,tail,negative,vs-width1 |
| cfb | mixed | fresh | 0 | 4 | 769743 | 1.52382 | 0.380956 | 3.3278 | 35616 | spread,tail |
| cfb | mixed | fresh | 0 | 8 | 593248 | 2.02017 | 0.252521 | 5.86227 | 35626 | spread,tail |
| cfb | mixed | fresh | 0 | 32 | 812018 | 1.46887 | 0.0459022 | 10.3947 | 43028 | spread,tail,vs-width1 |
| cfb | mixed | fresh | 65536 | 1 | 1.15964e+06 | 1 | 1 | 1.00055 | 35656 | spread,tail |
| cfb | mixed | fresh | 65536 | 2 | 1.1742e+06 | 0.9891 | 0.49455 | 1.00046 | 35648 | spread,tail,negative,vs-width1 |
| cfb | mixed | fresh | 65536 | 4 | 1.16848e+06 | 0.994549 | 0.248637 | 1.00052 | 35654 | spread,negative |
| cfb | mixed | fresh | 65536 | 8 | 1.17868e+06 | 0.991529 | 0.123941 | 1.00053 | 35592 | negative |
| cfb | mixed | fresh | 65536 | 32 | 1.1761e+06 | 0.993977 | 0.0310618 | 1.00051 | 35624 | spread,tail,negative,vs-width1 |
| cfb | small | fresh | 0 | 1 | 6320 | 1 | 1 | 1.03502 | 3850 | spread,tail |
| cfb | small | fresh | 0 | 2 | 133025 | 0.0476158 | 0.0238079 | 2.16364 | 3834 | spread,rss-spread,tail,negative,vs-width1 |
| cfb | small | fresh | 0 | 4 | 128840 | 0.049228 | 0.012307 | 3.27578 | 3830 | spread,rss-spread,tail,negative,vs-width1 |
| cfb | small | fresh | 0 | 8 | 171781 | 0.0368222 | 0.00460278 | 4.37068 | 3826 | spread,tail,negative,vs-width1 |
| cfb | small | fresh | 0 | 32 | 810914 | 0.00777122 | 0.000242851 | 6.58684 | 4506 | spread,rss-spread,tail,negative,vs-width1 |
| cfb | small | fresh | 65536 | 1 | 6345 | 1 | 1 | 1.03417 | 3826 | spread,rss-spread,tail |
| cfb | small | fresh | 65536 | 2 | 6340 | 1.00099 | 0.500493 | 1.03329 | 3834 | spread,tail,vs-width1 |
| cfb | small | fresh | 65536 | 4 | 6345 | 0.998435 | 0.249609 | 1.03419 | 3880 | spread,tail,negative,vs-width1 |
| cfb | small | fresh | 65536 | 8 | 6315 | 1.00475 | 0.125594 | 1.03332 | 3836 | spread,tail |
| cfb | small | fresh | 65536 | 32 | 6340 | 1.00237 | 0.0313239 | 1.03588 | 3890 | spread,tail,vs-width1 |
| opc | large | fresh | 0 | 1 | 842658 | 1 | 1 | 1.00046 | 28744 | - |
| opc | large | fresh | 0 | 2 | 1.54606e+06 | 0.544114 | 0.272057 | 1.83943 | 37116 | spread,tail,negative,vs-width1 |
| opc | large | fresh | 0 | 4 | 1.05147e+06 | 0.80236 | 0.20059 | 3.15016 | 37202 | spread,tail,negative,vs-width1 |
| opc | large | fresh | 0 | 8 | 818464 | 1.02651 | 0.128314 | 5.25154 | 37050 | spread,tail,vs-width1 |
| opc | large | fresh | 0 | 32 | 1.27337e+06 | 0.662657 | 0.020708 | 9.7326 | 46454 | spread,tail,negative,vs-width1 |
| opc | large | fresh | 65536 | 1 | 841858 | 1 | 1 | 1.00032 | 28802 | spread |
| opc | large | fresh | 65536 | 2 | 1.52872e+06 | 0.551111 | 0.275556 | 1.84993 | 37096 | spread,tail,negative,vs-width1 |
| opc | large | fresh | 65536 | 4 | 1.06714e+06 | 0.788523 | 0.197131 | 3.13026 | 37168 | spread,tail,negative,vs-width1 |
| opc | large | fresh | 65536 | 8 | 826268 | 1.01908 | 0.127386 | 5.22017 | 37000 | spread,tail,vs-width1 |
| opc | large | fresh | 65536 | 32 | 1.26942e+06 | 0.663687 | 0.0207402 | 9.6803 | 46396 | spread,tail,negative,vs-width1 |
| opc | large | primed | 0 | 1 | 839208 | 1 | 1 | 1.00046 | 28802 | spread |
| opc | large | primed | 0 | 2 | 1.49458e+06 | 0.563568 | 0.281784 | 1.80734 | 37340 | spread,tail,negative,vs-width1 |
| opc | large | primed | 0 | 4 | 966124 | 0.864959 | 0.21624 | 3.0955 | 37532 | spread,tail,negative,vs-width1 |
| opc | large | primed | 0 | 8 | 676448 | 1.23771 | 0.154714 | 5.34592 | 37774 | spread,tail,vs-width1 |
| opc | large | primed | 0 | 32 | 534842 | 1.56772 | 0.0489911 | 15.2221 | 46658 | spread,tail,vs-width1 |
| opc | large | primed | 65536 | 1 | 840599 | 1 | 1 | 1.00053 | 28754 | spread |
| opc | large | primed | 65536 | 2 | 1.57985e+06 | 0.532881 | 0.26644 | 1.78859 | 37358 | spread,tail,negative,vs-width1 |
| opc | large | primed | 65536 | 4 | 993500 | 0.850128 | 0.212532 | 3.06234 | 37508 | spread,tail,negative,vs-width1 |
| opc | large | primed | 65536 | 8 | 641058 | 1.30997 | 0.163746 | 5.46466 | 37758 | spread,tail,vs-width1 |
| opc | large | primed | 65536 | 32 | 552548 | 1.51969 | 0.0474903 | 14.9957 | 46876 | spread,tail,vs-width1 |
| opc | mixed | fresh | 0 | 1 | 824094 | 1 | 1 | 1.00047 | 28072 | spread |
| opc | mixed | fresh | 0 | 2 | 1.54838e+06 | 0.532028 | 0.266014 | 1.79839 | 36108 | spread,tail,negative,vs-width1 |
| opc | mixed | fresh | 0 | 4 | 1.04008e+06 | 0.793281 | 0.19832 | 3.10069 | 36162 | spread,tail,negative,vs-width1 |
| opc | mixed | fresh | 0 | 8 | 799818 | 1.0316 | 0.12895 | 5.25371 | 35944 | spread,tail,vs-width1 |
| opc | mixed | fresh | 0 | 32 | 1.27552e+06 | 0.647124 | 0.0202226 | 9.62927 | 45704 | spread,tail,negative,vs-width1 |
| opc | mixed | fresh | 65536 | 1 | 823214 | 1 | 1 | 1.00049 | 27982 | spread |
| opc | mixed | fresh | 65536 | 2 | 823688 | 1.0018 | 0.500901 | 1.0005 | 28090 | spread,vs-width1 |
| opc | mixed | fresh | 65536 | 4 | 823178 | 1.0006 | 0.25015 | 1.00048 | 27968 | - |
| opc | mixed | fresh | 65536 | 8 | 824514 | 0.998096 | 0.124762 | 1.00045 | 27988 | spread,negative,vs-width1 |
| opc | mixed | fresh | 65536 | 32 | 823009 | 1.00075 | 0.0312734 | 1.00043 | 28040 | spread,vs-width1 |
| opc | small | fresh | 0 | 1 | 289721 | 1 | 1 | 1.00083 | 4208 | tail |
| opc | small | fresh | 0 | 2 | 393266 | 0.738873 | 0.369437 | 1.80304 | 4244 | spread,rss-spread,tail,negative,vs-width1 |
| opc | small | fresh | 0 | 4 | 338516 | 0.85697 | 0.214243 | 2.71963 | 4202 | spread,tail,negative,vs-width1 |
| opc | small | fresh | 0 | 8 | 398142 | 0.729249 | 0.0911562 | 4.09782 | 4110 | spread,rss-spread,tail,negative,vs-width1 |
| opc | small | fresh | 0 | 32 | 1.04141e+06 | 0.278885 | 0.00871516 | 8.07798 | 5862 | spread,rss-spread,tail,negative,vs-width1 |
| opc | small | fresh | 65536 | 1 | 290296 | 1 | 1 | 1.00038 | 4194 | spread,rss-spread,tail |
| opc | small | fresh | 65536 | 2 | 290266 | 1.00033 | 0.500165 | 1.00083 | 4256 | spread,rss-spread,tail |
| opc | small | fresh | 65536 | 4 | 291096 | 0.998847 | 0.249712 | 1.00084 | 4198 | spread,rss-spread,tail,negative |
| opc | small | fresh | 65536 | 8 | 290292 | 1.00002 | 0.125002 | 1.00093 | 4198 | spread,rss-spread,tail,vs-width1 |
| opc | small | fresh | 65536 | 32 | 290431 | 0.998536 | 0.0312043 | 1.00085 | 4326 | spread,rss-spread,negative,vs-width1 |
| parts | large | fresh | 0 | 1 | 776834 | 1 | 1 | 1.0005 | 28984 | spread |
| parts | large | fresh | 0 | 2 | 1.43711e+06 | 0.540508 | 0.270254 | 1.82595 | 37196 | spread,tail,negative,vs-width1 |
| parts | large | fresh | 0 | 4 | 901974 | 0.861718 | 0.21543 | 3.16295 | 37198 | spread,tail,negative,vs-width1 |
| parts | large | fresh | 0 | 8 | 648553 | 1.19766 | 0.149708 | 5.1702 | 37032 | spread,tail,vs-width1 |
| parts | large | fresh | 0 | 32 | 794848 | 0.97734 | 0.0305419 | 6.42223 | 38010 | spread,tail,negative,vs-width1 |
| parts | large | fresh | 65536 | 1 | 775518 | 1 | 1 | 1.00048 | 28964 | spread |
| parts | large | fresh | 65536 | 2 | 1.48612e+06 | 0.521315 | 0.260657 | 1.80057 | 37042 | spread,tail,negative,vs-width1 |
| parts | large | fresh | 65536 | 4 | 903234 | 0.857473 | 0.214368 | 3.16612 | 37044 | spread,tail,negative,vs-width1 |
| parts | large | fresh | 65536 | 8 | 662468 | 1.17042 | 0.146303 | 5.09143 | 37124 | spread,tail,vs-width1 |
| parts | large | fresh | 65536 | 32 | 813864 | 0.951667 | 0.0297396 | 6.32469 | 38520 | spread,rss-spread,tail,negative,vs-width1 |
| parts | large | primed | 0 | 1 | 3045 | 1 | 1 | 1.09945 | 28790 | spread,tail |
| parts | large | primed | 0 | 2 | 228391 | 0.0131621 | 0.00658104 | 1.45489 | 37028 | spread,tail,negative,vs-width1 |
| parts | large | primed | 0 | 4 | 232126 | 0.0130654 | 0.00326636 | 2.02471 | 37058 | spread,tail,negative,vs-width1 |
| parts | large | primed | 0 | 8 | 295286 | 0.0101122 | 0.00126402 | 2.6137 | 37092 | spread,tail,negative,vs-width1 |
| parts | large | primed | 0 | 32 | 578022 | 0.00524438 | 0.000163887 | 2.47624 | 37956 | spread,tail,negative,vs-width1 |
| parts | large | primed | 65536 | 1 | 3010 | 1 | 1 | 1.11385 | 28802 | spread,tail |
| parts | large | primed | 65536 | 2 | 225306 | 0.0132396 | 0.00661979 | 1.46458 | 37096 | spread,tail,negative,vs-width1 |
| parts | large | primed | 65536 | 4 | 232921 | 0.0130175 | 0.00325437 | 2.02621 | 36988 | spread,tail,negative,vs-width1 |
| parts | large | primed | 65536 | 8 | 301091 | 0.01002 | 0.0012525 | 2.61391 | 36974 | spread,tail,negative,vs-width1 |
| parts | large | primed | 65536 | 32 | 574118 | 0.00523829 | 0.000163697 | 2.45777 | 37854 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 0 | 1 | 759248 | 1 | 1 | 1.00055 | 28244 | spread,tail |
| parts | mixed | fresh | 0 | 2 | 1.51669e+06 | 0.501913 | 0.250956 | 1.75625 | 36132 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 0 | 4 | 930964 | 0.817696 | 0.204424 | 3.06175 | 36126 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 0 | 8 | 663333 | 1.1476 | 0.14345 | 4.97928 | 35940 | spread,tail,vs-width1 |
| parts | mixed | fresh | 0 | 32 | 768454 | 0.986434 | 0.0308261 | 6.61467 | 36538 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 65536 | 1 | 757874 | 1 | 1 | 1.0005 | 28200 | spread |
| parts | mixed | fresh | 65536 | 2 | 758908 | 0.998745 | 0.499372 | 1.00047 | 28212 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 65536 | 4 | 759959 | 0.997086 | 0.249272 | 1.00047 | 28232 | spread,tail,negative,vs-width1 |
| parts | mixed | fresh | 65536 | 8 | 758418 | 0.996706 | 0.124588 | 1.00046 | 28036 | negative |
| parts | mixed | fresh | 65536 | 32 | 757813 | 0.99955 | 0.0312359 | 1.00047 | 28196 | negative |
| parts | small | fresh | 0 | 1 | 235911 | 1 | 1 | 1.00098 | 4198 | spread,rss-spread,tail |
| parts | small | fresh | 0 | 2 | 330206 | 0.714424 | 0.357212 | 1.77044 | 4174 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 0 | 4 | 284826 | 0.827552 | 0.206888 | 2.66838 | 4240 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 0 | 8 | 323056 | 0.730536 | 0.091317 | 3.56785 | 4198 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 0 | 32 | 606078 | 0.389176 | 0.0121618 | 3.76738 | 4808 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 65536 | 1 | 235456 | 1 | 1 | 1.00108 | 4364 | spread,rss-spread,tail |
| parts | small | fresh | 65536 | 2 | 235861 | 0.998727 | 0.499363 | 1.00115 | 4286 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 65536 | 4 | 236061 | 0.997737 | 0.249434 | 1.00111 | 4260 | spread,rss-spread,tail,negative,vs-width1 |
| parts | small | fresh | 65536 | 8 | 235791 | 0.999217 | 0.124902 | 1.00103 | 4270 | tail,negative |
| parts | small | fresh | 65536 | 32 | 235586 | 1.00083 | 0.031276 | 1.00103 | 4366 | tail,vs-width1 |

Observer source counters are retained in `observer.csv` and are never pooled into native timing.
