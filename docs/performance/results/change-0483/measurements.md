# Change 0483 measurements

This document independently recomputes the formal 0483 measurements from the 24 retained raw reports and their GNU `time -v` receipts (30 samples per process, 720 samples total). Percentiles use linear interpolation at `(n - 1) * q`; 95% intervals are the normal approximation `mean ± 1.96 × SEM`. The raw derivation and every comparison remain in [measurement-review.json](measurement-review.json).

For each instrumentation lane, materialized and bounded are route selections in the same executable, using a fresh process per capture. Every timed lifecycle appends one plain paragraph at the tail. The normal and allocator lanes use separate binaries, so their elapsed values are reported as separate lanes.

## Corpus and protocol

| paragraphs | source XML bytes | source archive bytes | materialized output bytes | bounded output bytes |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 3,313 | 2,144 | 2,149 | 2,147 |
| 8,192 | 401,585 | 23,419 | 23,482 | 23,563 |
| 131,072 | 6,422,705 | 344,002 | 344,068 | 344,172 |

The report-level source and output hashes are retained in the JSON review and were required to be constant for each count and route across the eight corresponding reports. The protocol records `phase_samples: false`, so this matrix has no phase retention or phase I/O rows.

## Total lifecycle

`elapsed` is the timed total lifecycle. `output_throughput` uses sink accepted bytes; `source_requested_throughput` uses logical source requested bytes, so the two rates describe different byte domains.

| lane | count | repeat | elapsed mean ms | 95% CI ms | p50 ms | p95 ms | p99 ms | GNU time max RSS KiB | process RSS delta mean | process VmHWM mean |
| :--- | ---: | ---: | ---: | :--- | ---: | ---: | ---: | ---: | ---: | ---: |
| allocator/materialized/a1 | 131,072 | 1 | 241.567767 | [241.174851, 241.960684] | 241.444034 | 243.668811 | 244.240817 | 102528 | 0.000000 | 104988672.000000 |
| allocator/materialized/a1 | 64 | 1 | 0.160888 | [0.159112, 0.162663] | 0.159200 | 0.171594 | 0.177253 | 5692 | 0.000000 | 6242304.000000 |
| allocator/materialized/a1 | 8,192 | 1 | 14.715415 | [14.688396, 14.742434] | 14.707497 | 14.872083 | 14.913817 | 11084 | 0.000000 | 11350016.000000 |
| normal/materialized/a1 | 131,072 | 1 | 243.769324 | [243.272014, 244.266634] | 243.355312 | 246.945179 | 247.344233 | 102500 | 0.000000 | 104960000.000000 |
| normal/materialized/a1 | 64 | 1 | 0.159472 | [0.157799, 0.161146] | 0.157810 | 0.170176 | 0.174126 | 5824 | 0.000000 | 6180864.000000 |
| normal/materialized/a1 | 8,192 | 1 | 14.913749 | [14.877525, 14.949973] | 14.891958 | 15.028709 | 15.279308 | 11172 | 0.000000 | 11440128.000000 |
| allocator/materialized/a2 | 131,072 | 2 | 239.776372 | [239.506316, 240.046427] | 239.636589 | 241.180058 | 241.498821 | 102532 | 0.000000 | 104992768.000000 |
| allocator/materialized/a2 | 64 | 2 | 0.160956 | [0.159494, 0.162417] | 0.159601 | 0.167867 | 0.175350 | 5884 | 0.000000 | 6004736.000000 |
| allocator/materialized/a2 | 8,192 | 2 | 14.692533 | [14.679086, 14.705981] | 14.692677 | 14.742386 | 14.768912 | 11284 | 0.000000 | 11554816.000000 |
| normal/materialized/a2 | 131,072 | 2 | 240.863305 | [240.424499, 241.302110] | 240.658082 | 242.926152 | 243.336311 | 102716 | 0.000000 | 105181184.000000 |
| normal/materialized/a2 | 64 | 2 | 0.157844 | [0.156445, 0.159242] | 0.156091 | 0.165958 | 0.168497 | 5912 | 0.000000 | 6266880.000000 |
| normal/materialized/a2 | 8,192 | 2 | 14.921802 | [14.885279, 14.958324] | 14.965684 | 15.042167 | 15.096313 | 11316 | 0.000000 | 11587584.000000 |
| allocator/bounded/b1 | 131,072 | 1 | 483.957975 | [483.628571, 484.287379] | 483.943814 | 485.813819 | 486.234523 | 102812 | 0.000000 | 105279488.000000 |
| allocator/bounded/b1 | 64 | 1 | 0.320107 | [0.318222, 0.321992] | 0.317876 | 0.328928 | 0.330684 | 5724 | 0.000000 | 6049792.000000 |
| allocator/bounded/b1 | 8,192 | 1 | 30.441892 | [30.419633, 30.464152] | 30.453489 | 30.523035 | 30.555090 | 11192 | 0.000000 | 11460608.000000 |
| normal/bounded/b1 | 131,072 | 1 | 481.701301 | [481.261890, 482.140711] | 481.574457 | 483.871386 | 484.459898 | 102676 | 0.000000 | 105140224.000000 |
| normal/bounded/b1 | 64 | 1 | 0.315602 | [0.313725, 0.317479] | 0.313852 | 0.325379 | 0.327604 | 5632 | 0.000000 | 5959680.000000 |
| normal/bounded/b1 | 8,192 | 1 | 30.233109 | [30.214651, 30.251566] | 30.235757 | 30.321004 | 30.343634 | 11108 | 0.000000 | 11374592.000000 |
| allocator/bounded/b2 | 131,072 | 2 | 482.482510 | [482.131764, 482.833256] | 482.669613 | 483.953476 | 484.369002 | 102628 | 0.000000 | 105091072.000000 |
| allocator/bounded/b2 | 64 | 2 | 0.320162 | [0.318166, 0.322159] | 0.316967 | 0.331195 | 0.332604 | 5692 | 0.000000 | 6258688.000000 |
| allocator/bounded/b2 | 8,192 | 2 | 30.287702 | [30.266539, 30.308865] | 30.289634 | 30.363191 | 30.406573 | 11016 | 0.000000 | 11280384.000000 |
| normal/bounded/b2 | 131,072 | 2 | 485.276936 | [484.875654, 485.678218] | 485.194436 | 487.309210 | 487.874399 | 102500 | 0.000000 | 104960000.000000 |
| normal/bounded/b2 | 64 | 2 | 0.314569 | [0.312679, 0.316459] | 0.311766 | 0.323568 | 0.324085 | 5904 | 0.000000 | 6103040.000000 |
| normal/bounded/b2 | 8,192 | 2 | 30.469390 | [30.448522, 30.490258] | 30.460080 | 30.552895 | 30.575430 | 11048 | 682.666667 | 11313152.000000 |

### Throughput

| lane | count | repeat | metric | mean | 95% CI | p50 | p95 | p99 |
| :--- | ---: | ---: | :--- | ---: | :--- | ---: | ---: | ---: |
| allocator/materialized/a1 | 131,072 | 1 | output_throughput (bytes/s) | 1424340.870636 | [1422033.354662, 1426648.386611] | 1425042.511217 | 1432008.796439 | 1433847.318150 |
| allocator/materialized/a1 | 131,072 | 1 | source_requested_throughput (bytes/s) | 4279012.768622 | [4272080.516293, 4285945.020951] | 4281120.641158 | 4302048.793983 | 4307572.091136 |
| allocator/materialized/a1 | 64 | 1 | output_throughput (bytes/s) | 13368614.117260 | [13230841.408311, 13506386.826209] | 13498744.570765 | 13611991.741175 | 13615502.933113 |
| allocator/materialized/a1 | 64 | 1 | source_requested_throughput (bytes/s) | 50245833.515639 | [49728015.846872, 50763651.184407] | 50734927.825997 | 51160566.446474 | 51173763.234411 |
| allocator/materialized/a1 | 8,192 | 1 | output_throughput (bytes/s) | 1595781.987363 | [1592865.564764, 1598698.409962] | 1596600.709647 | 1606905.640150 | 1610113.769132 |
| allocator/materialized/a1 | 8,192 | 1 | source_requested_throughput (bytes/s) | 4886292.328396 | [4877362.227991, 4895222.428801] | 4888799.260074 | 4920353.008180 | 4930176.314971 |
| normal/materialized/a1 | 131,072 | 1 | output_throughput (bytes/s) | 1411493.116023 | [1408636.805379, 1414349.426667] | 1413850.383684 | 1419679.080369 | 1421389.288059 |
| normal/materialized/a1 | 131,072 | 1 | source_requested_throughput (bytes/s) | 4240415.472728 | [4231834.528398, 4248996.417058] | 4247497.189349 | 4265007.792364 | 4270145.607819 |
| normal/materialized/a1 | 64 | 1 | output_throughput (bytes/s) | 13486218.814445 | [13353330.371350, 13619107.257540] | 13617599.175183 | 13731152.946619 | 13780780.031160 |
| normal/materialized/a1 | 64 | 1 | source_requested_throughput (bytes/s) | 50687849.867042 | [50188389.673984, 51187310.060100] | 51181641.944139 | 51608432.922216 | 51794955.938425 |
| normal/materialized/a1 | 8,192 | 1 | output_throughput (bytes/s) | 1574588.851403 | [1570845.729718, 1578331.973089] | 1576824.295957 | 1582566.517101 | 1582811.705210 |
| normal/materialized/a1 | 8,192 | 1 | source_requested_throughput (bytes/s) | 4821398.841394 | [4809937.384302, 4832860.298486] | 4828243.783659 | 4845826.493169 | 4846577.260371 |
| allocator/materialized/a2 | 131,072 | 2 | output_throughput (bytes/s) | 1434967.452480 | [1433354.065921, 1436580.839040] | 1435790.760105 | 1440965.657301 | 1442524.146161 |
| allocator/materialized/a2 | 131,072 | 2 | source_requested_throughput (bytes/s) | 4310937.204924 | [4306090.260046, 4315784.149802] | 4313410.590272 | 4328957.045221 | 4333639.066126 |
| allocator/materialized/a2 | 64 | 2 | output_throughput (bytes/s) | 13359261.288605 | [13245975.295035, 13472547.282175] | 13464873.523037 | 13570903.409304 | 13580750.464591 |
| allocator/materialized/a2 | 64 | 2 | source_requested_throughput (bytes/s) | 50210680.980951 | [49784896.443927, 50636465.517974] | 50607623.753174 | 51006136.266612 | 51043146.348301 |
| allocator/materialized/a2 | 8,192 | 2 | output_throughput (bytes/s) | 1598236.873316 | [1596773.693290, 1599700.053341] | 1598211.088707 | 1604929.794099 | 1606499.762535 |
| allocator/materialized/a2 | 8,192 | 2 | source_requested_throughput (bytes/s) | 4893809.201309 | [4889328.936843, 4898289.465775] | 4893730.248711 | 4914302.957809 | 4919110.208916 |
| normal/materialized/a2 | 131,072 | 2 | output_throughput (bytes/s) | 1428514.075559 | [1425911.366953, 1431116.784164] | 1429696.639274 | 1440239.807295 | 1444453.237539 |
| normal/materialized/a2 | 131,072 | 2 | source_requested_throughput (bytes/s) | 4291549.934069 | [4283730.862395, 4299369.005743] | 4295102.598561 | 4326776.442593 | 4339434.453177 |
| normal/materialized/a2 | 64 | 2 | output_throughput (bytes/s) | 13622498.109894 | [13506502.435720, 13738493.784067] | 13767609.976362 | 13867337.544538 | 13888726.171397 |
| normal/materialized/a2 | 64 | 2 | source_requested_throughput (bytes/s) | 51200054.552634 | [50764085.701867, 51636023.403402] | 51745456.388586 | 52120281.687869 | 52200670.677698 |
| normal/materialized/a2 | 8,192 | 2 | output_throughput (bytes/s) | 1573741.995785 | [1569875.451181, 1577608.540389] | 1569056.270650 | 1594020.870594 | 1597766.384612 |
| normal/materialized/a2 | 8,192 | 2 | source_requested_throughput (bytes/s) | 4818805.765306 | [4806966.386629, 4830645.143984] | 4804458.051796 | 4880899.780148 | 4892368.562575 |
| allocator/bounded/b1 | 131,072 | 1 | output_throughput (bytes/s) | 711163.406335 | [710679.984919, 711646.827751] | 711181.733312 | 713190.693086 | 713577.907755 |
| allocator/bounded/b1 | 131,072 | 1 | source_requested_throughput (bytes/s) | 4250220.659792 | [4247331.523947, 4253109.795637] | 4250330.189750 | 4262336.603830 | 4264650.766470 |
| allocator/bounded/b1 | 64 | 1 | output_throughput (bytes/s) | 6708866.400782 | [6669847.631053, 6747885.170511] | 6754206.084385 | 6819057.447824 | 6823935.987632 |
| allocator/bounded/b1 | 64 | 1 | source_requested_throughput (bytes/s) | 18039257.443742 | [17934341.115077, 18144173.772408] | 18161169.876644 | 18335546.644754 | 18348664.395250 |
| allocator/bounded/b1 | 8,192 | 1 | output_throughput (bytes/s) | 774035.166303 | [773468.801153, 774601.531452] | 773737.288869 | 776638.563387 | 777115.887364 |
| allocator/bounded/b1 | 8,192 | 1 | source_requested_throughput (bytes/s) | 4382892.415804 | [4379685.432934, 4386099.398675] | 4381205.716282 | 4397633.876957 | 4400336.673590 |
| normal/bounded/b1 | 131,072 | 1 | output_throughput (bytes/s) | 714497.053034 | [713845.980100, 715148.125967] | 714680.832030 | 717087.744468 | 717675.482867 |
| normal/bounded/b1 | 131,072 | 1 | source_requested_throughput (bytes/s) | 4270143.976915 | [4266252.882955, 4274035.070875] | 4271242.319829 | 4285627.071464 | 4289139.650800 |
| normal/bounded/b1 | 64 | 1 | output_throughput (bytes/s) | 6804673.091226 | [6764711.883199, 6844634.299253] | 6840809.786358 | 6943220.584628 | 6945666.530048 |
| normal/bounded/b1 | 64 | 1 | source_requested_throughput (bytes/s) | 18296869.005890 | [18189418.584866, 18404319.426914] | 18394035.815858 | 18669404.953451 | 18675981.778279 |
| normal/bounded/b1 | 8,192 | 1 | output_throughput (bytes/s) | 779379.532383 | [778903.865285, 779855.199480] | 779309.085877 | 781227.681401 | 781939.790301 |
| normal/bounded/b1 | 8,192 | 1 | source_requested_throughput (bytes/s) | 4413154.324538 | [4410460.909813, 4415847.739263] | 4412755.428635 | 4423619.273246 | 4427651.514718 |
| allocator/bounded/b2 | 131,072 | 2 | output_throughput (bytes/s) | 713338.547759 | [712820.047770, 713857.047748] | 713059.183041 | 715320.132814 | 715870.021012 |
| allocator/bounded/b2 | 131,072 | 2 | source_requested_throughput (bytes/s) | 4263220.247419 | [4260121.466821, 4266319.028016] | 4261550.642817 | 4275063.058320 | 4278349.428454 |
| allocator/bounded/b2 | 64 | 2 | output_throughput (bytes/s) | 6707907.947134 | [6666733.090462, 6749082.803806] | 6773577.026844 | 6817343.749030 | 6824367.829898 |
| allocator/bounded/b2 | 64 | 2 | source_requested_throughput (bytes/s) | 18036680.288218 | [17925966.525960, 18147394.050476] | 18213255.787598 | 18330938.734583 | 18349825.562179 |
| allocator/bounded/b2 | 8,192 | 2 | output_throughput (bytes/s) | 777975.383766 | [777431.770555, 778518.996976] | 777922.947848 | 780505.639152 | 780958.362139 |
| allocator/bounded/b2 | 8,192 | 2 | source_requested_throughput (bytes/s) | 4405203.481227 | [4402125.328808, 4408281.633645] | 4404906.568379 | 4419530.785237 | 4422094.281357 |
| normal/bounded/b2 | 131,072 | 2 | output_throughput (bytes/s) | 709231.676763 | [708645.669417, 709817.684108] | 709348.633025 | 711395.025528 | 712478.854753 |
| normal/bounded/b2 | 131,072 | 2 | source_requested_throughput (bytes/s) | 4238675.806859 | [4235173.573047, 4242178.040671] | 4239374.788162 | 4251604.916449 | 4258082.349516 |
| normal/bounded/b2 | 64 | 2 | output_throughput (bytes/s) | 6827041.717694 | [6786568.055261, 6867515.380126] | 6886579.794888 | 6930199.348429 | 6931179.246837 |
| normal/bounded/b2 | 64 | 2 | source_requested_throughput (bytes/s) | 18357015.294013 | [18248186.950640, 18465843.637386] | 18517105.335767 | 18634392.565664 | 18637027.383320 |
| normal/bounded/b2 | 8,192 | 2 | output_throughput (bytes/s) | 773336.233319 | [772806.502772, 773865.963865] | 773569.871207 | 775496.143915 | 776411.573961 |
| normal/bounded/b2 | 8,192 | 2 | source_requested_throughput (bytes/s) | 4378934.781568 | [4375935.238272, 4381934.324863] | 4380257.731446 | 4391165.047300 | 4396348.573296 |

### In-process RSS observers

| lane | count | repeat | metric | mean | 95% CI | p50 | p95 | p99 |
| :--- | ---: | ---: | :--- | ---: | :--- | ---: | ---: | ---: |
| allocator/materialized/a1 | 131,072 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 131,072 | 1 | process.peak_rss_bytes (bytes) | 104988672.000000 | [104988672.000000, 104988672.000000] | 104988672.000000 | 104988672.000000 | 104988672.000000 |
| allocator/materialized/a1 | 64 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 64 | 1 | process.peak_rss_bytes (bytes) | 6242304.000000 | [6242304.000000, 6242304.000000] | 6242304.000000 | 6242304.000000 | 6242304.000000 |
| allocator/materialized/a1 | 8,192 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 8,192 | 1 | process.peak_rss_bytes (bytes) | 11350016.000000 | [11350016.000000, 11350016.000000] | 11350016.000000 | 11350016.000000 | 11350016.000000 |
| normal/materialized/a1 | 131,072 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a1 | 131,072 | 1 | process.peak_rss_bytes (bytes) | 104960000.000000 | [104960000.000000, 104960000.000000] | 104960000.000000 | 104960000.000000 | 104960000.000000 |
| normal/materialized/a1 | 64 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a1 | 64 | 1 | process.peak_rss_bytes (bytes) | 6180864.000000 | [6180864.000000, 6180864.000000] | 6180864.000000 | 6180864.000000 | 6180864.000000 |
| normal/materialized/a1 | 8,192 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a1 | 8,192 | 1 | process.peak_rss_bytes (bytes) | 11440128.000000 | [11440128.000000, 11440128.000000] | 11440128.000000 | 11440128.000000 | 11440128.000000 |
| allocator/materialized/a2 | 131,072 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 131,072 | 2 | process.peak_rss_bytes (bytes) | 104992768.000000 | [104992768.000000, 104992768.000000] | 104992768.000000 | 104992768.000000 | 104992768.000000 |
| allocator/materialized/a2 | 64 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 64 | 2 | process.peak_rss_bytes (bytes) | 6004736.000000 | [6004736.000000, 6004736.000000] | 6004736.000000 | 6004736.000000 | 6004736.000000 |
| allocator/materialized/a2 | 8,192 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 8,192 | 2 | process.peak_rss_bytes (bytes) | 11554816.000000 | [11554816.000000, 11554816.000000] | 11554816.000000 | 11554816.000000 | 11554816.000000 |
| normal/materialized/a2 | 131,072 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a2 | 131,072 | 2 | process.peak_rss_bytes (bytes) | 105181184.000000 | [105181184.000000, 105181184.000000] | 105181184.000000 | 105181184.000000 | 105181184.000000 |
| normal/materialized/a2 | 64 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a2 | 64 | 2 | process.peak_rss_bytes (bytes) | 6266880.000000 | [6266880.000000, 6266880.000000] | 6266880.000000 | 6266880.000000 | 6266880.000000 |
| normal/materialized/a2 | 8,192 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/materialized/a2 | 8,192 | 2 | process.peak_rss_bytes (bytes) | 11587584.000000 | [11587584.000000, 11587584.000000] | 11587584.000000 | 11587584.000000 | 11587584.000000 |
| allocator/bounded/b1 | 131,072 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 131,072 | 1 | process.peak_rss_bytes (bytes) | 105279488.000000 | [105279488.000000, 105279488.000000] | 105279488.000000 | 105279488.000000 | 105279488.000000 |
| allocator/bounded/b1 | 64 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 64 | 1 | process.peak_rss_bytes (bytes) | 6049792.000000 | [6049792.000000, 6049792.000000] | 6049792.000000 | 6049792.000000 | 6049792.000000 |
| allocator/bounded/b1 | 8,192 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 8,192 | 1 | process.peak_rss_bytes (bytes) | 11460608.000000 | [11460608.000000, 11460608.000000] | 11460608.000000 | 11460608.000000 | 11460608.000000 |
| normal/bounded/b1 | 131,072 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/bounded/b1 | 131,072 | 1 | process.peak_rss_bytes (bytes) | 105140224.000000 | [105140224.000000, 105140224.000000] | 105140224.000000 | 105140224.000000 | 105140224.000000 |
| normal/bounded/b1 | 64 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/bounded/b1 | 64 | 1 | process.peak_rss_bytes (bytes) | 5959680.000000 | [5959680.000000, 5959680.000000] | 5959680.000000 | 5959680.000000 | 5959680.000000 |
| normal/bounded/b1 | 8,192 | 1 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/bounded/b1 | 8,192 | 1 | process.peak_rss_bytes (bytes) | 11374592.000000 | [11374592.000000, 11374592.000000] | 11374592.000000 | 11374592.000000 | 11374592.000000 |
| allocator/bounded/b2 | 131,072 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 131,072 | 2 | process.peak_rss_bytes (bytes) | 105091072.000000 | [105091072.000000, 105091072.000000] | 105091072.000000 | 105091072.000000 | 105091072.000000 |
| allocator/bounded/b2 | 64 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 64 | 2 | process.peak_rss_bytes (bytes) | 6258688.000000 | [6258688.000000, 6258688.000000] | 6258688.000000 | 6258688.000000 | 6258688.000000 |
| allocator/bounded/b2 | 8,192 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 8,192 | 2 | process.peak_rss_bytes (bytes) | 11280384.000000 | [11280384.000000, 11280384.000000] | 11280384.000000 | 11280384.000000 | 11280384.000000 |
| normal/bounded/b2 | 131,072 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/bounded/b2 | 131,072 | 2 | process.peak_rss_bytes (bytes) | 104960000.000000 | [104960000.000000, 104960000.000000] | 104960000.000000 | 104960000.000000 | 104960000.000000 |
| normal/bounded/b2 | 64 | 2 | process.rss_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| normal/bounded/b2 | 64 | 2 | process.peak_rss_bytes (bytes) | 6103040.000000 | [6103040.000000, 6103040.000000] | 6103040.000000 | 6103040.000000 | 6103040.000000 |
| normal/bounded/b2 | 8,192 | 2 | process.rss_bytes (bytes) | 682.666667 | [-655.360000, 2020.693333] | 0.000000 | 0.000000 | 14540.800000 |
| normal/bounded/b2 | 8,192 | 2 | process.peak_rss_bytes (bytes) | 11313152.000000 | [11313152.000000, 11313152.000000] | 11313152.000000 | 11313152.000000 | 11313152.000000 |

### Source read and sink write dimensions

| lane | count | repeat | metric | mean | 95% CI | p50 | p95 | p99 |
| :--- | ---: | ---: | :--- | ---: | :--- | ---: | ---: | ---: |
| allocator/materialized/a1 | 131,072 | 1 | source_reads.calls (calls) | 62.000000 | [62.000000, 62.000000] | 62.000000 | 62.000000 | 62.000000 |
| allocator/materialized/a1 | 131,072 | 1 | source_reads.requested_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| allocator/materialized/a1 | 131,072 | 1 | source_reads.returned_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| allocator/materialized/a1 | 131,072 | 1 | sink.accepted_bytes (bytes) | 344068.000000 | [344068.000000, 344068.000000] | 344068.000000 | 344068.000000 | 344068.000000 |
| allocator/materialized/a1 | 131,072 | 1 | sink.write_calls (calls) | 39.000000 | [39.000000, 39.000000] | 39.000000 | 39.000000 | 39.000000 |
| allocator/materialized/a1 | 131,072 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/materialized/a1 | 64 | 1 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| allocator/materialized/a1 | 64 | 1 | source_reads.requested_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| allocator/materialized/a1 | 64 | 1 | source_reads.returned_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| allocator/materialized/a1 | 64 | 1 | sink.accepted_bytes (bytes) | 2149.000000 | [2149.000000, 2149.000000] | 2149.000000 | 2149.000000 | 2149.000000 |
| allocator/materialized/a1 | 64 | 1 | sink.write_calls (calls) | 19.000000 | [19.000000, 19.000000] | 19.000000 | 19.000000 | 19.000000 |
| allocator/materialized/a1 | 64 | 1 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| allocator/materialized/a1 | 8,192 | 1 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| allocator/materialized/a1 | 8,192 | 1 | source_reads.requested_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| allocator/materialized/a1 | 8,192 | 1 | source_reads.returned_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| allocator/materialized/a1 | 8,192 | 1 | sink.accepted_bytes (bytes) | 23482.000000 | [23482.000000, 23482.000000] | 23482.000000 | 23482.000000 | 23482.000000 |
| allocator/materialized/a1 | 8,192 | 1 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| allocator/materialized/a1 | 8,192 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/materialized/a1 | 131,072 | 1 | source_reads.calls (calls) | 62.000000 | [62.000000, 62.000000] | 62.000000 | 62.000000 | 62.000000 |
| normal/materialized/a1 | 131,072 | 1 | source_reads.requested_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| normal/materialized/a1 | 131,072 | 1 | source_reads.returned_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| normal/materialized/a1 | 131,072 | 1 | sink.accepted_bytes (bytes) | 344068.000000 | [344068.000000, 344068.000000] | 344068.000000 | 344068.000000 | 344068.000000 |
| normal/materialized/a1 | 131,072 | 1 | sink.write_calls (calls) | 39.000000 | [39.000000, 39.000000] | 39.000000 | 39.000000 | 39.000000 |
| normal/materialized/a1 | 131,072 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/materialized/a1 | 64 | 1 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| normal/materialized/a1 | 64 | 1 | source_reads.requested_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| normal/materialized/a1 | 64 | 1 | source_reads.returned_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| normal/materialized/a1 | 64 | 1 | sink.accepted_bytes (bytes) | 2149.000000 | [2149.000000, 2149.000000] | 2149.000000 | 2149.000000 | 2149.000000 |
| normal/materialized/a1 | 64 | 1 | sink.write_calls (calls) | 19.000000 | [19.000000, 19.000000] | 19.000000 | 19.000000 | 19.000000 |
| normal/materialized/a1 | 64 | 1 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| normal/materialized/a1 | 8,192 | 1 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| normal/materialized/a1 | 8,192 | 1 | source_reads.requested_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| normal/materialized/a1 | 8,192 | 1 | source_reads.returned_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| normal/materialized/a1 | 8,192 | 1 | sink.accepted_bytes (bytes) | 23482.000000 | [23482.000000, 23482.000000] | 23482.000000 | 23482.000000 | 23482.000000 |
| normal/materialized/a1 | 8,192 | 1 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| normal/materialized/a1 | 8,192 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/materialized/a2 | 131,072 | 2 | source_reads.calls (calls) | 62.000000 | [62.000000, 62.000000] | 62.000000 | 62.000000 | 62.000000 |
| allocator/materialized/a2 | 131,072 | 2 | source_reads.requested_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| allocator/materialized/a2 | 131,072 | 2 | source_reads.returned_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| allocator/materialized/a2 | 131,072 | 2 | sink.accepted_bytes (bytes) | 344068.000000 | [344068.000000, 344068.000000] | 344068.000000 | 344068.000000 | 344068.000000 |
| allocator/materialized/a2 | 131,072 | 2 | sink.write_calls (calls) | 39.000000 | [39.000000, 39.000000] | 39.000000 | 39.000000 | 39.000000 |
| allocator/materialized/a2 | 131,072 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/materialized/a2 | 64 | 2 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| allocator/materialized/a2 | 64 | 2 | source_reads.requested_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| allocator/materialized/a2 | 64 | 2 | source_reads.returned_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| allocator/materialized/a2 | 64 | 2 | sink.accepted_bytes (bytes) | 2149.000000 | [2149.000000, 2149.000000] | 2149.000000 | 2149.000000 | 2149.000000 |
| allocator/materialized/a2 | 64 | 2 | sink.write_calls (calls) | 19.000000 | [19.000000, 19.000000] | 19.000000 | 19.000000 | 19.000000 |
| allocator/materialized/a2 | 64 | 2 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| allocator/materialized/a2 | 8,192 | 2 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| allocator/materialized/a2 | 8,192 | 2 | source_reads.requested_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| allocator/materialized/a2 | 8,192 | 2 | source_reads.returned_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| allocator/materialized/a2 | 8,192 | 2 | sink.accepted_bytes (bytes) | 23482.000000 | [23482.000000, 23482.000000] | 23482.000000 | 23482.000000 | 23482.000000 |
| allocator/materialized/a2 | 8,192 | 2 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| allocator/materialized/a2 | 8,192 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/materialized/a2 | 131,072 | 2 | source_reads.calls (calls) | 62.000000 | [62.000000, 62.000000] | 62.000000 | 62.000000 | 62.000000 |
| normal/materialized/a2 | 131,072 | 2 | source_reads.requested_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| normal/materialized/a2 | 131,072 | 2 | source_reads.returned_bytes (bytes) | 1033651.000000 | [1033651.000000, 1033651.000000] | 1033651.000000 | 1033651.000000 | 1033651.000000 |
| normal/materialized/a2 | 131,072 | 2 | sink.accepted_bytes (bytes) | 344068.000000 | [344068.000000, 344068.000000] | 344068.000000 | 344068.000000 | 344068.000000 |
| normal/materialized/a2 | 131,072 | 2 | sink.write_calls (calls) | 39.000000 | [39.000000, 39.000000] | 39.000000 | 39.000000 | 39.000000 |
| normal/materialized/a2 | 131,072 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/materialized/a2 | 64 | 2 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| normal/materialized/a2 | 64 | 2 | source_reads.requested_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| normal/materialized/a2 | 64 | 2 | source_reads.returned_bytes (bytes) | 8077.000000 | [8077.000000, 8077.000000] | 8077.000000 | 8077.000000 | 8077.000000 |
| normal/materialized/a2 | 64 | 2 | sink.accepted_bytes (bytes) | 2149.000000 | [2149.000000, 2149.000000] | 2149.000000 | 2149.000000 | 2149.000000 |
| normal/materialized/a2 | 64 | 2 | sink.write_calls (calls) | 19.000000 | [19.000000, 19.000000] | 19.000000 | 19.000000 | 19.000000 |
| normal/materialized/a2 | 64 | 2 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| normal/materialized/a2 | 8,192 | 2 | source_reads.calls (calls) | 42.000000 | [42.000000, 42.000000] | 42.000000 | 42.000000 | 42.000000 |
| normal/materialized/a2 | 8,192 | 2 | source_reads.requested_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| normal/materialized/a2 | 8,192 | 2 | source_reads.returned_bytes (bytes) | 71902.000000 | [71902.000000, 71902.000000] | 71902.000000 | 71902.000000 | 71902.000000 |
| normal/materialized/a2 | 8,192 | 2 | sink.accepted_bytes (bytes) | 23482.000000 | [23482.000000, 23482.000000] | 23482.000000 | 23482.000000 | 23482.000000 |
| normal/materialized/a2 | 8,192 | 2 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| normal/materialized/a2 | 8,192 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/bounded/b1 | 131,072 | 1 | source_reads.calls (calls) | 119.000000 | [119.000000, 119.000000] | 119.000000 | 119.000000 | 119.000000 |
| allocator/bounded/b1 | 131,072 | 1 | source_reads.requested_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| allocator/bounded/b1 | 131,072 | 1 | source_reads.returned_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| allocator/bounded/b1 | 131,072 | 1 | sink.accepted_bytes (bytes) | 344172.000000 | [344172.000000, 344172.000000] | 344172.000000 | 344172.000000 | 344172.000000 |
| allocator/bounded/b1 | 131,072 | 1 | sink.write_calls (calls) | 50.000000 | [50.000000, 50.000000] | 50.000000 | 50.000000 | 50.000000 |
| allocator/bounded/b1 | 131,072 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/bounded/b1 | 64 | 1 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| allocator/bounded/b1 | 64 | 1 | source_reads.requested_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| allocator/bounded/b1 | 64 | 1 | source_reads.returned_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| allocator/bounded/b1 | 64 | 1 | sink.accepted_bytes (bytes) | 2147.000000 | [2147.000000, 2147.000000] | 2147.000000 | 2147.000000 | 2147.000000 |
| allocator/bounded/b1 | 64 | 1 | sink.write_calls (calls) | 18.000000 | [18.000000, 18.000000] | 18.000000 | 18.000000 | 18.000000 |
| allocator/bounded/b1 | 64 | 1 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| allocator/bounded/b1 | 8,192 | 1 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| allocator/bounded/b1 | 8,192 | 1 | source_reads.requested_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| allocator/bounded/b1 | 8,192 | 1 | source_reads.returned_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| allocator/bounded/b1 | 8,192 | 1 | sink.accepted_bytes (bytes) | 23563.000000 | [23563.000000, 23563.000000] | 23563.000000 | 23563.000000 | 23563.000000 |
| allocator/bounded/b1 | 8,192 | 1 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| allocator/bounded/b1 | 8,192 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/bounded/b1 | 131,072 | 1 | source_reads.calls (calls) | 119.000000 | [119.000000, 119.000000] | 119.000000 | 119.000000 | 119.000000 |
| normal/bounded/b1 | 131,072 | 1 | source_reads.requested_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| normal/bounded/b1 | 131,072 | 1 | source_reads.returned_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| normal/bounded/b1 | 131,072 | 1 | sink.accepted_bytes (bytes) | 344172.000000 | [344172.000000, 344172.000000] | 344172.000000 | 344172.000000 | 344172.000000 |
| normal/bounded/b1 | 131,072 | 1 | sink.write_calls (calls) | 50.000000 | [50.000000, 50.000000] | 50.000000 | 50.000000 | 50.000000 |
| normal/bounded/b1 | 131,072 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/bounded/b1 | 64 | 1 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| normal/bounded/b1 | 64 | 1 | source_reads.requested_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| normal/bounded/b1 | 64 | 1 | source_reads.returned_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| normal/bounded/b1 | 64 | 1 | sink.accepted_bytes (bytes) | 2147.000000 | [2147.000000, 2147.000000] | 2147.000000 | 2147.000000 | 2147.000000 |
| normal/bounded/b1 | 64 | 1 | sink.write_calls (calls) | 18.000000 | [18.000000, 18.000000] | 18.000000 | 18.000000 | 18.000000 |
| normal/bounded/b1 | 64 | 1 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| normal/bounded/b1 | 8,192 | 1 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| normal/bounded/b1 | 8,192 | 1 | source_reads.requested_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| normal/bounded/b1 | 8,192 | 1 | source_reads.returned_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| normal/bounded/b1 | 8,192 | 1 | sink.accepted_bytes (bytes) | 23563.000000 | [23563.000000, 23563.000000] | 23563.000000 | 23563.000000 | 23563.000000 |
| normal/bounded/b1 | 8,192 | 1 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| normal/bounded/b1 | 8,192 | 1 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/bounded/b2 | 131,072 | 2 | source_reads.calls (calls) | 119.000000 | [119.000000, 119.000000] | 119.000000 | 119.000000 | 119.000000 |
| allocator/bounded/b2 | 131,072 | 2 | source_reads.requested_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| allocator/bounded/b2 | 131,072 | 2 | source_reads.returned_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| allocator/bounded/b2 | 131,072 | 2 | sink.accepted_bytes (bytes) | 344172.000000 | [344172.000000, 344172.000000] | 344172.000000 | 344172.000000 | 344172.000000 |
| allocator/bounded/b2 | 131,072 | 2 | sink.write_calls (calls) | 50.000000 | [50.000000, 50.000000] | 50.000000 | 50.000000 | 50.000000 |
| allocator/bounded/b2 | 131,072 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| allocator/bounded/b2 | 64 | 2 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| allocator/bounded/b2 | 64 | 2 | source_reads.requested_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| allocator/bounded/b2 | 64 | 2 | source_reads.returned_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| allocator/bounded/b2 | 64 | 2 | sink.accepted_bytes (bytes) | 2147.000000 | [2147.000000, 2147.000000] | 2147.000000 | 2147.000000 | 2147.000000 |
| allocator/bounded/b2 | 64 | 2 | sink.write_calls (calls) | 18.000000 | [18.000000, 18.000000] | 18.000000 | 18.000000 | 18.000000 |
| allocator/bounded/b2 | 64 | 2 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| allocator/bounded/b2 | 8,192 | 2 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| allocator/bounded/b2 | 8,192 | 2 | source_reads.requested_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| allocator/bounded/b2 | 8,192 | 2 | source_reads.returned_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| allocator/bounded/b2 | 8,192 | 2 | sink.accepted_bytes (bytes) | 23563.000000 | [23563.000000, 23563.000000] | 23563.000000 | 23563.000000 | 23563.000000 |
| allocator/bounded/b2 | 8,192 | 2 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| allocator/bounded/b2 | 8,192 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/bounded/b2 | 131,072 | 2 | source_reads.calls (calls) | 119.000000 | [119.000000, 119.000000] | 119.000000 | 119.000000 | 119.000000 |
| normal/bounded/b2 | 131,072 | 2 | source_reads.requested_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| normal/bounded/b2 | 131,072 | 2 | source_reads.returned_bytes (bytes) | 2056921.000000 | [2056921.000000, 2056921.000000] | 2056921.000000 | 2056921.000000 | 2056921.000000 |
| normal/bounded/b2 | 131,072 | 2 | sink.accepted_bytes (bytes) | 344172.000000 | [344172.000000, 344172.000000] | 344172.000000 | 344172.000000 | 344172.000000 |
| normal/bounded/b2 | 131,072 | 2 | sink.write_calls (calls) | 50.000000 | [50.000000, 50.000000] | 50.000000 | 50.000000 | 50.000000 |
| normal/bounded/b2 | 131,072 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |
| normal/bounded/b2 | 64 | 2 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| normal/bounded/b2 | 64 | 2 | source_reads.requested_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| normal/bounded/b2 | 64 | 2 | source_reads.returned_bytes (bytes) | 5773.000000 | [5773.000000, 5773.000000] | 5773.000000 | 5773.000000 | 5773.000000 |
| normal/bounded/b2 | 64 | 2 | sink.accepted_bytes (bytes) | 2147.000000 | [2147.000000, 2147.000000] | 2147.000000 | 2147.000000 | 2147.000000 |
| normal/bounded/b2 | 64 | 2 | sink.write_calls (calls) | 18.000000 | [18.000000, 18.000000] | 18.000000 | 18.000000 | 18.000000 |
| normal/bounded/b2 | 64 | 2 | sink.largest_write (events) | 877.000000 | [877.000000, 877.000000] | 877.000000 | 877.000000 | 877.000000 |
| normal/bounded/b2 | 8,192 | 2 | source_reads.calls (calls) | 59.000000 | [59.000000, 59.000000] | 59.000000 | 59.000000 | 59.000000 |
| normal/bounded/b2 | 8,192 | 2 | source_reads.requested_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| normal/bounded/b2 | 8,192 | 2 | source_reads.returned_bytes (bytes) | 133423.000000 | [133423.000000, 133423.000000] | 133423.000000 | 133423.000000 | 133423.000000 |
| normal/bounded/b2 | 8,192 | 2 | sink.accepted_bytes (bytes) | 23563.000000 | [23563.000000, 23563.000000] | 23563.000000 | 23563.000000 | 23563.000000 |
| normal/bounded/b2 | 8,192 | 2 | sink.write_calls (calls) | 20.000000 | [20.000000, 20.000000] | 20.000000 | 20.000000 | 20.000000 |
| normal/bounded/b2 | 8,192 | 2 | sink.largest_write (events) | 16384.000000 | [16384.000000, 16384.000000] | 16384.000000 | 16384.000000 | 16384.000000 |

### Read and write histogram shapes

The histogram counters are per sample; the table shows the first sample for each lane and the full vectors are retained in the JSON review.

| lane | count | repeat | source request histogram | sink write histogram |
| :--- | ---: | ---: | :--- | :--- |
| allocator/materialized/a1 | 131,072 | 1 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_4097_to_16384=3, bytes_16385_to_65536=20 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=21 |
| allocator/materialized/a1 | 64 | 1 | bytes_0=3, bytes_1_to_512=36, bytes_513_to_4096=3 | bytes_1_to_512=18, bytes_513_to_4096=1 |
| allocator/materialized/a1 | 8,192 | 1 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_16385_to_65536=3 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=2 |
| normal/materialized/a1 | 131,072 | 1 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_4097_to_16384=3, bytes_16385_to_65536=20 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=21 |
| normal/materialized/a1 | 64 | 1 | bytes_0=3, bytes_1_to_512=36, bytes_513_to_4096=3 | bytes_1_to_512=18, bytes_513_to_4096=1 |
| normal/materialized/a1 | 8,192 | 1 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_16385_to_65536=3 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=2 |
| allocator/materialized/a2 | 131,072 | 2 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_4097_to_16384=3, bytes_16385_to_65536=20 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=21 |
| allocator/materialized/a2 | 64 | 2 | bytes_0=3, bytes_1_to_512=36, bytes_513_to_4096=3 | bytes_1_to_512=18, bytes_513_to_4096=1 |
| allocator/materialized/a2 | 8,192 | 2 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_16385_to_65536=3 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=2 |
| normal/materialized/a2 | 131,072 | 2 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_4097_to_16384=3, bytes_16385_to_65536=20 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=21 |
| normal/materialized/a2 | 64 | 2 | bytes_0=3, bytes_1_to_512=36, bytes_513_to_4096=3 | bytes_1_to_512=18, bytes_513_to_4096=1 |
| normal/materialized/a2 | 8,192 | 2 | bytes_0=3, bytes_1_to_512=35, bytes_513_to_4096=1, bytes_16385_to_65536=3 | bytes_1_to_512=17, bytes_513_to_4096=1, bytes_4097_to_16384=2 |
| allocator/bounded/b1 | 131,072 | 1 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_4097_to_16384=6, bytes_16385_to_65536=60 | bytes_1_to_512=16, bytes_513_to_4096=5, bytes_4097_to_16384=29 |
| allocator/bounded/b1 | 64 | 1 | bytes_0=8, bytes_1_to_512=50, bytes_513_to_4096=1 | bytes_1_to_512=17, bytes_513_to_4096=1 |
| allocator/bounded/b1 | 8,192 | 1 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_16385_to_65536=6 | bytes_1_to_512=16, bytes_513_to_4096=3, bytes_4097_to_16384=1 |
| normal/bounded/b1 | 131,072 | 1 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_4097_to_16384=6, bytes_16385_to_65536=60 | bytes_1_to_512=16, bytes_513_to_4096=5, bytes_4097_to_16384=29 |
| normal/bounded/b1 | 64 | 1 | bytes_0=8, bytes_1_to_512=50, bytes_513_to_4096=1 | bytes_1_to_512=17, bytes_513_to_4096=1 |
| normal/bounded/b1 | 8,192 | 1 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_16385_to_65536=6 | bytes_1_to_512=16, bytes_513_to_4096=3, bytes_4097_to_16384=1 |
| allocator/bounded/b2 | 131,072 | 2 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_4097_to_16384=6, bytes_16385_to_65536=60 | bytes_1_to_512=16, bytes_513_to_4096=5, bytes_4097_to_16384=29 |
| allocator/bounded/b2 | 64 | 2 | bytes_0=8, bytes_1_to_512=50, bytes_513_to_4096=1 | bytes_1_to_512=17, bytes_513_to_4096=1 |
| allocator/bounded/b2 | 8,192 | 2 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_16385_to_65536=6 | bytes_1_to_512=16, bytes_513_to_4096=3, bytes_4097_to_16384=1 |
| normal/bounded/b2 | 131,072 | 2 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_4097_to_16384=6, bytes_16385_to_65536=60 | bytes_1_to_512=16, bytes_513_to_4096=5, bytes_4097_to_16384=29 |
| normal/bounded/b2 | 64 | 2 | bytes_0=8, bytes_1_to_512=50, bytes_513_to_4096=1 | bytes_1_to_512=17, bytes_513_to_4096=1 |
| normal/bounded/b2 | 8,192 | 2 | bytes_0=8, bytes_1_to_512=44, bytes_513_to_4096=1, bytes_16385_to_65536=6 | bytes_1_to_512=16, bytes_513_to_4096=3, bytes_4097_to_16384=1 |

### Allocator counters and peaks

| lane | count | repeat | metric | mean | 95% CI | p50 | p95 | p99 |
| :--- | ---: | ---: | :--- | ---: | :--- | ---: | ---: | ---: |
| allocator/materialized/a1 | 131,072 | 1 | allocation.allocation_calls (calls) | 508228.000000 | [508228.000000, 508228.000000] | 508228.000000 | 508228.000000 | 508228.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.reallocation_calls (calls) | 507973.000000 | [507973.000000, 507973.000000] | 507973.000000 | 507973.000000 | 507973.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.allocated_bytes (bytes) | 549263355169.000000 | [549263355169.000000, 549263355169.000000] | 549263355169.000000 | 549263355169.000000 | 549263355169.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.deallocated_bytes (bytes) | 549263355169.000000 | [549263355169.000000, 549263355169.000000] | 549263355169.000000 | 549263355169.000000 | 549263355169.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.live_bytes_before (events) | 1057142.000000 | [1056940.383408, 1057343.616592] | 1057142.000000 | 1057977.200000 | 1058051.440000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.live_bytes_after (events) | 1057142.000000 | [1056940.383408, 1057343.616592] | 1057142.000000 | 1057977.200000 | 1058051.440000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.peak_live_bytes_before (events) | 97081148.000000 | [97081148.000000, 97081148.000000] | 97081148.000000 | 97081148.000000 | 97081148.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.peak_live_bytes_after (events) | 97081148.000000 | [97081148.000000, 97081148.000000] | 97081148.000000 | 97081148.000000 | 97081148.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.region_peak_live_bytes (bytes) | 36428432.000000 | [36428230.383408, 36428633.616592] | 36428432.000000 | 36429267.200000 | 36429341.440000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.incremental_peak_live_bytes (bytes) | 35371290.000000 | [35371290.000000, 35371290.000000] | 35371290.000000 | 35371290.000000 | 35371290.000000 |
| allocator/materialized/a1 | 131,072 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.allocation_calls (calls) | 322.000000 | [322.000000, 322.000000] | 322.000000 | 322.000000 | 322.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.reallocation_calls (calls) | 67.000000 | [67.000000, 67.000000] | 67.000000 | 67.000000 | 67.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.allocated_bytes (bytes) | 1032137.000000 | [1032137.000000, 1032137.000000] | 1032137.000000 | 1032137.000000 | 1032137.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.deallocated_bytes (bytes) | 1032137.000000 | [1032137.000000, 1032137.000000] | 1032137.000000 | 1032137.000000 | 1032137.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.live_bytes_before (events) | 31330.000000 | [31128.383408, 31531.616592] | 31330.000000 | 32165.200000 | 32239.440000 |
| allocator/materialized/a1 | 64 | 1 | allocation.live_bytes_after (events) | 31330.000000 | [31128.383408, 31531.616592] | 31330.000000 | 32165.200000 | 32239.440000 |
| allocator/materialized/a1 | 64 | 1 | allocation.peak_live_bytes_before (events) | 667545.000000 | [667545.000000, 667545.000000] | 667545.000000 | 667545.000000 | 667545.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.peak_live_bytes_after (events) | 667545.000000 | [667545.000000, 667545.000000] | 667545.000000 | 667545.000000 | 667545.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.region_peak_live_bytes (bytes) | 538116.000000 | [537914.383408, 538317.616592] | 538116.000000 | 538951.200000 | 539025.440000 |
| allocator/materialized/a1 | 64 | 1 | allocation.incremental_peak_live_bytes (bytes) | 506786.000000 | [506786.000000, 506786.000000] | 506786.000000 | 506786.000000 | 506786.000000 |
| allocator/materialized/a1 | 64 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.allocation_calls (calls) | 16708.000000 | [16708.000000, 16708.000000] | 16708.000000 | 16708.000000 | 16708.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.reallocation_calls (calls) | 16453.000000 | [16453.000000, 16453.000000] | 16453.000000 | 16453.000000 | 16453.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.allocated_bytes (bytes) | 1614459169.000000 | [1614459169.000000, 1614459169.000000] | 1614459169.000000 | 1614459169.000000 | 1614459169.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.deallocated_bytes (bytes) | 1614459169.000000 | [1614459169.000000, 1614459169.000000] | 1614459169.000000 | 1614459169.000000 | 1614459169.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.live_bytes_before (events) | 95364.000000 | [95162.383408, 95565.616592] | 95364.000000 | 96199.200000 | 96273.440000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.live_bytes_after (events) | 95364.000000 | [95162.383408, 95565.616592] | 95364.000000 | 96199.200000 | 96273.440000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.peak_live_bytes_before (events) | 6528442.000000 | [6528442.000000, 6528442.000000] | 6528442.000000 | 6528442.000000 | 6528442.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.peak_live_bytes_after (events) | 6528442.000000 | [6528442.000000, 6528442.000000] | 6528442.000000 | 6528442.000000 | 6528442.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.region_peak_live_bytes (bytes) | 2765214.000000 | [2765012.383408, 2765415.616592] | 2765214.000000 | 2766049.200000 | 2766123.440000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.incremental_peak_live_bytes (bytes) | 2669850.000000 | [2669850.000000, 2669850.000000] | 2669850.000000 | 2669850.000000 | 2669850.000000 |
| allocator/materialized/a1 | 8,192 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.allocation_calls (calls) | 508228.000000 | [508228.000000, 508228.000000] | 508228.000000 | 508228.000000 | 508228.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.reallocation_calls (calls) | 507973.000000 | [507973.000000, 507973.000000] | 507973.000000 | 507973.000000 | 507973.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.allocated_bytes (bytes) | 549263355169.000000 | [549263355169.000000, 549263355169.000000] | 549263355169.000000 | 549263355169.000000 | 549263355169.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.deallocated_bytes (bytes) | 549263355169.000000 | [549263355169.000000, 549263355169.000000] | 549263355169.000000 | 549263355169.000000 | 549263355169.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.live_bytes_before (events) | 1057142.000000 | [1056940.383408, 1057343.616592] | 1057142.000000 | 1057977.200000 | 1058051.440000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.live_bytes_after (events) | 1057142.000000 | [1056940.383408, 1057343.616592] | 1057142.000000 | 1057977.200000 | 1058051.440000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.peak_live_bytes_before (events) | 97081148.000000 | [97081148.000000, 97081148.000000] | 97081148.000000 | 97081148.000000 | 97081148.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.peak_live_bytes_after (events) | 97081148.000000 | [97081148.000000, 97081148.000000] | 97081148.000000 | 97081148.000000 | 97081148.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.region_peak_live_bytes (bytes) | 36428432.000000 | [36428230.383408, 36428633.616592] | 36428432.000000 | 36429267.200000 | 36429341.440000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.incremental_peak_live_bytes (bytes) | 35371290.000000 | [35371290.000000, 35371290.000000] | 35371290.000000 | 35371290.000000 | 35371290.000000 |
| allocator/materialized/a2 | 131,072 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.allocation_calls (calls) | 322.000000 | [322.000000, 322.000000] | 322.000000 | 322.000000 | 322.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.reallocation_calls (calls) | 67.000000 | [67.000000, 67.000000] | 67.000000 | 67.000000 | 67.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.allocated_bytes (bytes) | 1032137.000000 | [1032137.000000, 1032137.000000] | 1032137.000000 | 1032137.000000 | 1032137.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.deallocated_bytes (bytes) | 1032137.000000 | [1032137.000000, 1032137.000000] | 1032137.000000 | 1032137.000000 | 1032137.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.live_bytes_before (events) | 31330.000000 | [31128.383408, 31531.616592] | 31330.000000 | 32165.200000 | 32239.440000 |
| allocator/materialized/a2 | 64 | 2 | allocation.live_bytes_after (events) | 31330.000000 | [31128.383408, 31531.616592] | 31330.000000 | 32165.200000 | 32239.440000 |
| allocator/materialized/a2 | 64 | 2 | allocation.peak_live_bytes_before (events) | 667545.000000 | [667545.000000, 667545.000000] | 667545.000000 | 667545.000000 | 667545.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.peak_live_bytes_after (events) | 667545.000000 | [667545.000000, 667545.000000] | 667545.000000 | 667545.000000 | 667545.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.region_peak_live_bytes (bytes) | 538116.000000 | [537914.383408, 538317.616592] | 538116.000000 | 538951.200000 | 539025.440000 |
| allocator/materialized/a2 | 64 | 2 | allocation.incremental_peak_live_bytes (bytes) | 506786.000000 | [506786.000000, 506786.000000] | 506786.000000 | 506786.000000 | 506786.000000 |
| allocator/materialized/a2 | 64 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.allocation_calls (calls) | 16708.000000 | [16708.000000, 16708.000000] | 16708.000000 | 16708.000000 | 16708.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.deallocation_calls (calls) | 255.000000 | [255.000000, 255.000000] | 255.000000 | 255.000000 | 255.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.reallocation_calls (calls) | 16453.000000 | [16453.000000, 16453.000000] | 16453.000000 | 16453.000000 | 16453.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.allocated_bytes (bytes) | 1614459169.000000 | [1614459169.000000, 1614459169.000000] | 1614459169.000000 | 1614459169.000000 | 1614459169.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.deallocated_bytes (bytes) | 1614459169.000000 | [1614459169.000000, 1614459169.000000] | 1614459169.000000 | 1614459169.000000 | 1614459169.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.live_bytes_before (events) | 95364.000000 | [95162.383408, 95565.616592] | 95364.000000 | 96199.200000 | 96273.440000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.live_bytes_after (events) | 95364.000000 | [95162.383408, 95565.616592] | 95364.000000 | 96199.200000 | 96273.440000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.peak_live_bytes_before (events) | 6528442.000000 | [6528442.000000, 6528442.000000] | 6528442.000000 | 6528442.000000 | 6528442.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.peak_live_bytes_after (events) | 6528442.000000 | [6528442.000000, 6528442.000000] | 6528442.000000 | 6528442.000000 | 6528442.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.region_peak_live_bytes (bytes) | 2765214.000000 | [2765012.383408, 2765415.616592] | 2765214.000000 | 2766049.200000 | 2766123.440000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.incremental_peak_live_bytes (bytes) | 2669850.000000 | [2669850.000000, 2669850.000000] | 2669850.000000 | 2669850.000000 | 2669850.000000 |
| allocator/materialized/a2 | 8,192 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.live_bytes_before (events) | 1057393.000000 | [1057191.383408, 1057594.616592] | 1057393.000000 | 1058228.200000 | 1058302.440000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.live_bytes_after (events) | 1057393.000000 | [1057191.383408, 1057594.616592] | 1057393.000000 | 1058228.200000 | 1058302.440000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.peak_live_bytes_before (events) | 97081143.000000 | [97081143.000000, 97081143.000000] | 97081143.000000 | 97081143.000000 | 97081143.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.peak_live_bytes_after (events) | 97081143.000000 | [97081143.000000, 97081143.000000] | 97081143.000000 | 97081143.000000 | 97081143.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.region_peak_live_bytes (bytes) | 1667268.000000 | [1667066.383408, 1667469.616592] | 1667268.000000 | 1668103.200000 | 1668177.440000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b1 | 131,072 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.live_bytes_before (events) | 31581.000000 | [31379.383408, 31782.616592] | 31581.000000 | 32416.200000 | 32490.440000 |
| allocator/bounded/b1 | 64 | 1 | allocation.live_bytes_after (events) | 31581.000000 | [31379.383408, 31782.616592] | 31581.000000 | 32416.200000 | 32490.440000 |
| allocator/bounded/b1 | 64 | 1 | allocation.peak_live_bytes_before (events) | 667540.000000 | [667540.000000, 667540.000000] | 667540.000000 | 667540.000000 | 667540.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.peak_live_bytes_after (events) | 667540.000000 | [667540.000000, 667540.000000] | 667540.000000 | 667540.000000 | 667540.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.region_peak_live_bytes (bytes) | 641456.000000 | [641254.383408, 641657.616592] | 641456.000000 | 642291.200000 | 642365.440000 |
| allocator/bounded/b1 | 64 | 1 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b1 | 64 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.live_bytes_before (events) | 95615.000000 | [95413.383408, 95816.616592] | 95615.000000 | 96450.200000 | 96524.440000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.live_bytes_after (events) | 95615.000000 | [95413.383408, 95816.616592] | 95615.000000 | 96450.200000 | 96524.440000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.peak_live_bytes_before (events) | 6528437.000000 | [6528437.000000, 6528437.000000] | 6528437.000000 | 6528437.000000 | 6528437.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.peak_live_bytes_after (events) | 6528437.000000 | [6528437.000000, 6528437.000000] | 6528437.000000 | 6528437.000000 | 6528437.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.region_peak_live_bytes (bytes) | 705490.000000 | [705288.383408, 705691.616592] | 705490.000000 | 706325.200000 | 706399.440000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b1 | 8,192 | 1 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.live_bytes_before (events) | 1057393.000000 | [1057191.383408, 1057594.616592] | 1057393.000000 | 1058228.200000 | 1058302.440000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.live_bytes_after (events) | 1057393.000000 | [1057191.383408, 1057594.616592] | 1057393.000000 | 1058228.200000 | 1058302.440000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.peak_live_bytes_before (events) | 97081143.000000 | [97081143.000000, 97081143.000000] | 97081143.000000 | 97081143.000000 | 97081143.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.peak_live_bytes_after (events) | 97081143.000000 | [97081143.000000, 97081143.000000] | 97081143.000000 | 97081143.000000 | 97081143.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.region_peak_live_bytes (bytes) | 1667268.000000 | [1667066.383408, 1667469.616592] | 1667268.000000 | 1668103.200000 | 1668177.440000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b2 | 131,072 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.live_bytes_before (events) | 31581.000000 | [31379.383408, 31782.616592] | 31581.000000 | 32416.200000 | 32490.440000 |
| allocator/bounded/b2 | 64 | 2 | allocation.live_bytes_after (events) | 31581.000000 | [31379.383408, 31782.616592] | 31581.000000 | 32416.200000 | 32490.440000 |
| allocator/bounded/b2 | 64 | 2 | allocation.peak_live_bytes_before (events) | 667540.000000 | [667540.000000, 667540.000000] | 667540.000000 | 667540.000000 | 667540.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.peak_live_bytes_after (events) | 667540.000000 | [667540.000000, 667540.000000] | 667540.000000 | 667540.000000 | 667540.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.region_peak_live_bytes (bytes) | 641456.000000 | [641254.383408, 641657.616592] | 641456.000000 | 642291.200000 | 642365.440000 |
| allocator/bounded/b2 | 64 | 2 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b2 | 64 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.allocation_calls (calls) | 292.000000 | [292.000000, 292.000000] | 292.000000 | 292.000000 | 292.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.deallocation_calls (calls) | 251.000000 | [251.000000, 251.000000] | 251.000000 | 251.000000 | 251.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.reallocation_calls (calls) | 41.000000 | [41.000000, 41.000000] | 41.000000 | 41.000000 | 41.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.failed_allocation_calls (calls) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.allocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.deallocated_bytes (bytes) | 1956965.000000 | [1956965.000000, 1956965.000000] | 1956965.000000 | 1956965.000000 | 1956965.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.live_bytes_before (events) | 95615.000000 | [95413.383408, 95816.616592] | 95615.000000 | 96450.200000 | 96524.440000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.live_bytes_after (events) | 95615.000000 | [95413.383408, 95816.616592] | 95615.000000 | 96450.200000 | 96524.440000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.peak_live_bytes_before (events) | 6528437.000000 | [6528437.000000, 6528437.000000] | 6528437.000000 | 6528437.000000 | 6528437.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.peak_live_bytes_after (events) | 6528437.000000 | [6528437.000000, 6528437.000000] | 6528437.000000 | 6528437.000000 | 6528437.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.region_peak_live_bytes (bytes) | 705490.000000 | [705288.383408, 705691.616592] | 705490.000000 | 706325.200000 | 706399.440000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.incremental_peak_live_bytes (bytes) | 609875.000000 | [609875.000000, 609875.000000] | 609875.000000 | 609875.000000 | 609875.000000 |
| allocator/bounded/b2 | 8,192 | 2 | allocation.retention_delta_bytes (bytes) | 0.000000 | [0.000000, 0.000000] | 0.000000 | 0.000000 | 0.000000 |

The allocator `region_peak_live_bytes` values are absolute high-water endpoints. The operation incremental peak in the table is calculated per sample as `region_peak_live_bytes - live_bytes_before`; it is the reported operation quantity. Allocation bytes are requested callback accounting, including the full requested `new_size` for reallocations, and do not measure physical bytes copied.

## Repeat drift and route flags

Repeat drift is `(repeat 2 mean - repeat 1 mean) / repeat 1 mean`; route comparisons pair materialized and bounded arms within the same repeat, count, and instrumentation. All changes, including values below 5%, are in the JSON review. Flags use absolute change strictly greater than 5% and retain improvements for review. A zero baseline has no percentage; its JSON `percent` is `null` and carries an explicit zero-baseline status/flag.

### Repeat drift flags

Every metric is retained in `measurement-review.json`. The table lists every >5% review flag and every undefined zero-baseline increase; an explicit no-flag row is emitted for stable lanes.

| lane | count | metric | change | adverse |
| :--- | ---: | :--- | ---: | :---: |
| normal/materialized (a1-normal-8192-materialized → a2-normal-8192-materialized) | 8,192 | process.nonvoluntary_context_switches | 150.00% | yes |
| normal/bounded (b1-normal-8192-bounded → b2-normal-8192-bounded) | 8,192 | process.minor_faults | undefined (zero baseline) | yes |
| normal/bounded (b1-normal-8192-bounded → b2-normal-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 11.11% | yes |
| normal/bounded (b1-normal-8192-bounded → b2-normal-8192-bounded) | 8,192 | process.rss_bytes | undefined (zero baseline) | yes |
| normal/materialized (a1-normal-131072-materialized → a2-normal-131072-materialized) | 131,072 | process.nonvoluntary_context_switches | -9.41% | no |
| normal/bounded (b1-normal-131072-bounded → b2-normal-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | 6.45% | yes |
| allocator/bounded (b1-allocator-64-bounded → b2-allocator-64-bounded) | 64 | process.nonvoluntary_context_switches | -100.00% | no |
| allocator/materialized (a1-allocator-8192-materialized → a2-allocator-8192-materialized) | 8,192 | process.nonvoluntary_context_switches | 25.00% | yes |
| allocator/bounded (b1-allocator-8192-bounded → b2-allocator-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 37.50% | yes |
| allocator/materialized (a1-allocator-131072-materialized → a2-allocator-131072-materialized) | 131,072 | process.minor_faults | undefined (zero baseline) | yes |
| allocator/materialized (a1-allocator-131072-materialized → a2-allocator-131072-materialized) | 131,072 | process.nonvoluntary_context_switches | -12.50% | no |
| allocator/bounded (b1-allocator-131072-bounded → b2-allocator-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | -7.19% | no |

### Materialized/bounded route flags

Every metric is retained in `measurement-review.json`. The table lists every >5% review flag and every undefined zero-baseline increase; an explicit no-flag row is emitted for stable lanes.

| lane | count | metric | change | adverse |
| :--- | ---: | :--- | ---: | :---: |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | elapsed_ns | 97.90% | yes |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | output_throughput_bytes_per_second | -49.54% | yes |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | process.user_cpu_ticks | undefined (zero baseline) | yes |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | sink.write_calls | -5.26% | no |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | source_reads.calls | 40.48% | yes |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | source_reads.requested_bytes | -28.53% | no |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | source_reads.returned_bytes | -28.53% | no |
| normal/materialized (a1-normal-64-materialized → b1-normal-64-bounded) | 64 | source_requested_throughput_bytes_per_second | -63.90% | no |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | elapsed_ns | 99.29% | yes |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | output_throughput_bytes_per_second | -49.88% | yes |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | process.user_cpu_ticks | undefined (zero baseline) | yes |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | sink.write_calls | -5.26% | no |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | source_reads.calls | 40.48% | yes |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | source_reads.requested_bytes | -28.53% | no |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | source_reads.returned_bytes | -28.53% | no |
| normal/materialized (a2-normal-64-materialized → b2-normal-64-bounded) | 64 | source_requested_throughput_bytes_per_second | -64.15% | no |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | elapsed_ns | 102.72% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | output_throughput_bytes_per_second | -50.50% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 350.00% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | process.user_cpu_ticks | 106.82% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | source_reads.calls | 40.48% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | source_reads.requested_bytes | 85.56% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | source_reads.returned_bytes | 85.56% | yes |
| normal/materialized (a1-normal-8192-materialized → b1-normal-8192-bounded) | 8,192 | source_requested_throughput_bytes_per_second | -8.47% | no |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | elapsed_ns | 104.19% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | output_throughput_bytes_per_second | -50.86% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | process.minor_faults | undefined (zero baseline) | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 100.00% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | process.rss_bytes | undefined (zero baseline) | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | process.user_cpu_ticks | 102.22% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | source_reads.calls | 40.48% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | source_reads.requested_bytes | 85.56% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | source_reads.returned_bytes | 85.56% | yes |
| normal/materialized (a2-normal-8192-materialized → b2-normal-8192-bounded) | 8,192 | source_requested_throughput_bytes_per_second | -9.13% | no |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | elapsed_ns | 97.61% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | output_throughput_bytes_per_second | -49.38% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | 82.35% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | process.user_cpu_ticks | 97.40% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | sink.write_calls | 28.21% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | source_reads.calls | 91.94% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | source_reads.requested_bytes | 99.00% | yes |
| normal/materialized (a1-normal-131072-materialized → b1-normal-131072-bounded) | 131,072 | source_reads.returned_bytes | 99.00% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | elapsed_ns | 101.47% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | output_throughput_bytes_per_second | -50.35% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | 114.29% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | process.user_cpu_ticks | 101.38% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | sink.write_calls | 28.21% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | source_reads.calls | 91.94% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | source_reads.requested_bytes | 99.00% | yes |
| normal/materialized (a2-normal-131072-materialized → b2-normal-131072-bounded) | 131,072 | source_reads.returned_bytes | 99.00% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.allocated_bytes | 89.60% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.allocation_calls | -9.32% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.deallocated_bytes | 89.60% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.incremental_peak_live_bytes | 20.34% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.reallocation_calls | -38.81% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | allocation.region_peak_live_bytes | 19.20% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | elapsed_ns | 98.96% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | output_throughput_bytes_per_second | -49.82% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | process.nonvoluntary_context_switches | undefined (zero baseline) | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | process.user_cpu_ticks | undefined (zero baseline) | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | sink.write_calls | -5.26% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | source_reads.calls | 40.48% | yes |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | source_reads.requested_bytes | -28.53% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | source_reads.returned_bytes | -28.53% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | source_requested_throughput_bytes_per_second | -64.10% | no |
| allocator/materialized (a1-allocator-64-materialized → b1-allocator-64-bounded) | 64 | total_peak_live_bytes | 19.20% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.allocated_bytes | 89.60% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.allocation_calls | -9.32% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.deallocated_bytes | 89.60% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.incremental_peak_live_bytes | 20.34% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.reallocation_calls | -38.81% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | allocation.region_peak_live_bytes | 19.20% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | elapsed_ns | 98.91% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | output_throughput_bytes_per_second | -49.79% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | process.user_cpu_ticks | undefined (zero baseline) | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | sink.write_calls | -5.26% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | source_reads.calls | 40.48% | yes |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | source_reads.requested_bytes | -28.53% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | source_reads.returned_bytes | -28.53% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | source_requested_throughput_bytes_per_second | -64.08% | no |
| allocator/materialized (a2-allocator-64-materialized → b2-allocator-64-bounded) | 64 | total_peak_live_bytes | 19.20% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.allocated_bytes | -99.88% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.allocation_calls | -98.25% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.deallocated_bytes | -99.88% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.incremental_peak_live_bytes | -77.16% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.reallocation_calls | -99.75% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | allocation.region_peak_live_bytes | -74.49% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | elapsed_ns | 106.87% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | output_throughput_bytes_per_second | -51.49% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 100.00% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | process.user_cpu_ticks | 106.82% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | source_reads.calls | 40.48% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | source_reads.requested_bytes | 85.56% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | source_reads.returned_bytes | 85.56% | yes |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | source_requested_throughput_bytes_per_second | -10.30% | no |
| allocator/materialized (a1-allocator-8192-materialized → b1-allocator-8192-bounded) | 8,192 | total_peak_live_bytes | -74.49% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.allocated_bytes | -99.88% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.allocation_calls | -98.25% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.deallocated_bytes | -99.88% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.incremental_peak_live_bytes | -77.16% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.reallocation_calls | -99.75% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | allocation.region_peak_live_bytes | -74.49% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | elapsed_ns | 106.14% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | output_throughput_bytes_per_second | -51.32% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | process.nonvoluntary_context_switches | 120.00% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | process.user_cpu_ticks | 102.22% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | source_reads.calls | 40.48% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | source_reads.requested_bytes | 85.56% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | source_reads.returned_bytes | 85.56% | yes |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | source_requested_throughput_bytes_per_second | -9.98% | no |
| allocator/materialized (a2-allocator-8192-materialized → b2-allocator-8192-bounded) | 8,192 | total_peak_live_bytes | -74.49% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.allocated_bytes | -100.00% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.allocation_calls | -99.94% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.deallocated_bytes | -100.00% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.incremental_peak_live_bytes | -98.28% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.reallocation_calls | -99.99% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | allocation.region_peak_live_bytes | -95.42% | no |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | elapsed_ns | 100.34% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | output_throughput_bytes_per_second | -50.07% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | 89.77% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | process.user_cpu_ticks | 100.55% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | sink.write_calls | 28.21% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | source_reads.calls | 91.94% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | source_reads.requested_bytes | 99.00% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | source_reads.returned_bytes | 99.00% | yes |
| allocator/materialized (a1-allocator-131072-materialized → b1-allocator-131072-bounded) | 131,072 | total_peak_live_bytes | -95.42% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.allocated_bytes | -100.00% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.allocation_calls | -99.94% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.deallocated_bytes | -100.00% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.incremental_peak_live_bytes | -98.28% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.reallocation_calls | -99.99% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | allocation.region_peak_live_bytes | -95.42% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | elapsed_ns | 101.22% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | output_throughput_bytes_per_second | -50.29% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | process.minor_faults | -100.00% | no |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | process.nonvoluntary_context_switches | 101.30% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | process.user_cpu_ticks | 101.39% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | sink.write_calls | 28.21% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | source_reads.calls | 91.94% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | source_reads.requested_bytes | 99.00% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | source_reads.returned_bytes | 99.00% | yes |
| allocator/materialized (a2-allocator-131072-materialized → b2-allocator-131072-bounded) | 131,072 | total_peak_live_bytes | -95.42% | no |

## Scope caveats

- The materialized and bounded route arms select two paths in the same executable for each instrumentation lane; normal and allocator lanes use separate instrumentation binaries and are not a direct speed comparison.
- Each timed lifecycle appends exactly one plain paragraph at the tail. The bounded caller text is borrowed from storage prepared outside the timed region.
- requested allocation bytes are callback accounting: reallocation increments allocation_calls and reallocation_calls, and allocated_bytes includes the full requested new_size. No physical copy-byte counter is present.
- region_peak_live_bytes is an absolute live-byte high-water value. The operation peak shown here subtracts live_bytes_before; independent attribution-region peaks must not be summed.
- Total elapsed time includes source/package admission, publication, sink digest finalization, and owner drops. Corpus construction and independent source/candidate oracles remain outside the timed lifecycle.
- process.rss_bytes is a saturating RSS delta and process.peak_rss_bytes is an absolute VmHWM endpoint. GNU time RSS is a broader whole-process maximum covering setup, oracles, serialization, and teardown.
- The explicit bounded-window memory goal remains open; this one-append matrix does not establish a constant-memory or arbitrary DOCX scaling property.

The review does not turn these one-append observations into a production performance claim; the explicit bounded-window objective remains open.

Independent comparison with `summary.json`: **matched** (0 differences).
