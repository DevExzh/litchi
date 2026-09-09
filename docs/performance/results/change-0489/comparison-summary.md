# 0489 OPC candidate XML audit reuse comparison

This table retains every formal process row. Values are descriptive; a five-percent flag marks a review threshold and does not authorize a causal speedup claim.

Protocol `27fa75f85d8ea839410ab9cecc85a59acec69e00665636052e54e77022f406f8`; before attempt `formal1`; after attempt `formal1`.

The matrix contains 144 formal children and 4320 measured samples across 18 arms.

## Source custody

Before source manifest: `cb69be7b5a1f443191a27000fdfab285b707ba28517cccc3ee764d833cde0ba2` (7157 files). After source manifest: `0f5d9a455d25146823a7a8eb3620712f6e8f199abe0d40c32678d100c3994afa` (7157 files).

Changed source files: 2; all changes in the configured OPC implementation/test allowlist: `True`.

Matched candidate ZIP archive identities changed for 0 process pairs; decoded candidate content identity is checked independently and archive framing is retained per row.

## Per-process rows

| phase | arm | role | repeat | elapsed p50 (ns) | elapsed p95 (ns) | elapsed p99 (ns) | allocator operation heap p50 (B) | heap p95 (B) | heap p99 (B) | GNU time RSS (B, n=1) |
|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| before | deterministic-owned-s64-a64-short-c64 | normal | 1 | 779723.000 | 787043.500 | 788328.600 | — | — | — | 6385664 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 682597.500 | 692347.500 | 696733.590 | 989190.000 | 989190.000 | 989190.000 | 6709248 |
| before | deterministic-file-s64-a64-short-c64 | normal | 1 | 1064269.000 | 1072698.000 | 1075082.600 | — | — | — | 6307840 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 1 | 1051219.000 | 1059436.000 | 1063382.900 | 989166.000 | 989166.000 | 989166.000 | 6578176 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 674647.000 | 682039.500 | 683444.910 | — | — | — | 6447104 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 690522.000 | 698700.500 | 708898.500 | 989278.000 | 989278.000 | 989278.000 | 6451200 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12278480.000 | 12439895.500 | 12594480.310 | — | — | — | 6651904 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12299600.000 | 12334384.950 | 12336613.910 | 989278.000 | 989278.000 | 989278.000 | 6516736 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 98742383.500 | 99426874.500 | 99505094.910 | — | — | — | 17571840 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 98867994.500 | 99431657.000 | 99666057.720 | 989190.000 | 989190.000 | 989190.000 | 17596416 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 1 | 155058898.000 | 156786564.000 | 157111560.110 | — | — | — | 17567744 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 154982590.500 | 155715940.550 | 155864429.000 | 989166.000 | 989166.000 | 989166.000 | 17432576 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 99432333.000 | 100817285.050 | 100834825.900 | — | — | — | 17485824 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 99943967.500 | 100847240.100 | 101185237.100 | 989278.000 | 989278.000 | 989278.000 | 17649664 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 113395733.500 | 118769801.800 | 119665251.500 | — | — | — | 17784832 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 112414904.500 | 118008932.650 | 119817070.070 | 989278.000 | 989278.000 | 989278.000 | 17965056 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 386874452.500 | 389458113.550 | 389972757.220 | — | — | — | 75468800 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 384751434.500 | 388464711.800 | 391351265.950 | 989190.000 | 989190.000 | 989190.000 | 75677696 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 1 | 386660751.500 | 387656456.500 | 388785629.160 | — | — | — | 75345920 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 385273287.000 | 387806500.950 | 389274205.920 | 989166.000 | 989166.000 | 989166.000 | 75780096 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 384026958.000 | 386510083.500 | 388924506.970 | — | — | — | 75313152 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 386459430.000 | 387878493.700 | 388525416.400 | 989278.000 | 989278.000 | 989278.000 | 75702272 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 432465674.500 | 435072882.850 | 435156104.740 | — | — | — | 75247616 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 435669556.000 | 439264292.550 | 446812104.490 | 989278.000 | 989278.000 | 989278.000 | 75284480 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 1 | 539382.000 | 548957.050 | 561311.890 | — | — | — | 6836224 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 553937.000 | 562748.950 | 568655.200 | 9184134.000 | 9184134.000 | 9184134.000 | 6647808 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 68520425.000 | 69934354.500 | 69965751.700 | — | — | — | 20332544 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 70042119.000 | 70521617.150 | 70655238.680 | 9184134.000 | 9184134.000 | 9184134.000 | 20369408 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 383896904.000 | 387535056.650 | 389793797.990 | — | — | — | 81883136 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 384772512.000 | 386659052.800 | 386955091.660 | 9184134.000 | 9184134.000 | 9184134.000 | 81956864 |
| before | file_store-owned-s64-a64-short-c64 | normal | 1 | 3677339.500 | 3856415.000 | 4702364.650 | — | — | — | 6311936 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 1 | 3579774.000 | 3699504.550 | 3701938.890 | 795767.000 | 795767.000 | 795767.000 | 6488064 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 1 | 76959264.500 | 77882514.000 | 78032105.810 | — | — | — | 20508672 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 76866739.000 | 77487957.050 | 77591673.600 | 795773.000 | 795773.000 | 795773.000 | 20774912 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 1 | 386243318.000 | 388075607.050 | 389326505.470 | — | — | — | 82075648 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 388027967.000 | 388725146.050 | 389107919.910 | 795775.000 | 795775.000 | 795775.000 | 82259968 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 387008318.500 | 391038296.650 | 393892894.310 | 795775.000 | 795775.000 | 795775.000 | 81895424 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 2 | 387828074.500 | 388588124.950 | 388647066.210 | — | — | — | 81956864 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 77281247.000 | 79106921.600 | 79386609.600 | 795773.000 | 795773.000 | 795773.000 | 20697088 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 2 | 75535040.500 | 76398099.550 | 76787415.020 | — | — | — | 20377600 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 2 | 3698789.500 | 4254150.500 | 4373845.410 | 795767.000 | 795767.000 | 795767.000 | 6643712 |
| before | file_store-owned-s64-a64-short-c64 | normal | 2 | 3628159.000 | 3733057.550 | 3879810.700 | — | — | — | 6377472 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 384542672.500 | 390113681.550 | 391572628.410 | 9184134.000 | 9184134.000 | 9184134.000 | 83259392 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 383939221.000 | 385296331.600 | 386040512.530 | — | — | — | 81793024 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 70460041.500 | 70900592.500 | 71067813.310 | 9184134.000 | 9184134.000 | 9184134.000 | 20312064 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 68372808.000 | 68949646.000 | 69645154.430 | — | — | — | 20570112 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 560022.000 | 568232.050 | 571989.000 | 9184134.000 | 9184134.000 | 9184134.000 | 7102464 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 2 | 537837.000 | 546137.000 | 547640.710 | — | — | — | 6823936 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 434386247.000 | 436814990.250 | 439881180.330 | 989278.000 | 989278.000 | 989278.000 | 75153408 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 432930954.000 | 434376372.550 | 435107798.540 | — | — | — | 75370496 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 385061460.000 | 392753757.450 | 393272219.930 | 989278.000 | 989278.000 | 989278.000 | 75202560 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 383999466.000 | 385100492.550 | 385470842.010 | — | — | — | 75288576 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 387922061.000 | 392059592.450 | 393802289.840 | 989166.000 | 989166.000 | 989166.000 | 75300864 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 2 | 386485393.000 | 388039715.050 | 391213276.760 | — | — | — | 75374592 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 385198231.000 | 387338022.550 | 390673496.870 | 989190.000 | 989190.000 | 989190.000 | 75603968 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 384635929.500 | 386267930.650 | 388130994.090 | — | — | — | 75329536 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 110844520.000 | 111313813.000 | 111409125.300 | 989278.000 | 989278.000 | 989278.000 | 17571840 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 113313589.500 | 114890610.050 | 115171325.110 | — | — | — | 17514496 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 101177580.000 | 102068604.000 | 102195067.490 | 989278.000 | 989278.000 | 989278.000 | 17645568 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 98227758.000 | 99406083.500 | 99863622.530 | — | — | — | 17584128 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 153985508.500 | 155188783.000 | 155316489.610 | 989166.000 | 989166.000 | 989166.000 | 17584128 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 2 | 152012231.000 | 153718602.000 | 153978871.110 | — | — | — | 17657856 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 99445273.500 | 100485269.550 | 100839881.720 | 989190.000 | 989190.000 | 989190.000 | 18042880 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 97848722.000 | 98494873.500 | 98918584.820 | — | — | — | 17600512 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12352772.500 | 12449130.500 | 12453942.600 | 989278.000 | 989278.000 | 989278.000 | 6406144 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12269307.000 | 12325289.450 | 12365917.410 | — | — | — | 6574080 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 692607.500 | 697887.050 | 701997.000 | 989278.000 | 989278.000 | 989278.000 | 6709248 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 676387.500 | 686092.950 | 687836.010 | — | — | — | 6639616 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 2 | 1043089.000 | 1057395.000 | 1058606.600 | 989166.000 | 989166.000 | 989166.000 | 6774784 |
| before | deterministic-file-s64-a64-short-c64 | normal | 2 | 942559.000 | 1056844.000 | 1064324.000 | — | — | — | 6639616 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 680953.000 | 688051.450 | 692672.110 | 989190.000 | 989190.000 | 989190.000 | 6365184 |
| before | deterministic-owned-s64-a64-short-c64 | normal | 2 | 787903.000 | 795160.500 | 799490.600 | — | — | — | 6569984 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 1 | 551062.500 | 555798.500 | 559461.400 | — | — | — | 6578176 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 559957.500 | 569269.550 | 570229.590 | 792403.000 | 792403.000 | 792403.000 | 6467584 |
| after | deterministic-file-s64-a64-short-c64 | normal | 1 | 809124.000 | 815490.550 | 816828.990 | — | — | — | 6299648 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 1 | 815653.500 | 822811.000 | 828101.700 | 792379.000 | 792379.000 | 792379.000 | 6250496 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 555562.000 | 565163.500 | 566318.100 | — | — | — | 6950912 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 570022.500 | 577730.550 | 588301.090 | 792491.000 | 792491.000 | 792491.000 | 6569984 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12153878.500 | 12198663.000 | 12476803.910 | — | — | — | 6725632 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12146649.000 | 12252673.550 | 12269017.800 | 792491.000 | 792491.000 | 792491.000 | 6709248 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 77251940.500 | 77756428.950 | 77814874.210 | — | — | — | 18169856 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 78357614.500 | 79119302.050 | 79258574.710 | 792403.000 | 792403.000 | 792403.000 | 17899520 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 1 | 131920251.000 | 132179247.000 | 132214085.900 | — | — | — | 17932288 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 134565216.500 | 135262009.500 | 135450799.010 | 792379.000 | 792379.000 | 792379.000 | 17735680 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 77506345.000 | 77795742.700 | 77917340.900 | — | — | — | 17907712 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 78900489.000 | 79466320.500 | 79547281.610 | 792491.000 | 792491.000 | 792491.000 | 18149376 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 89706891.000 | 90501545.600 | 91591421.440 | — | — | — | 18239488 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 92205535.000 | 92706816.550 | 94391465.090 | 792491.000 | 792491.000 | 792491.000 | 18243584 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 304450575.500 | 305942027.200 | 306989471.380 | — | — | — | 75640832 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 303033184.500 | 304090617.550 | 304945053.550 | 792403.000 | 792403.000 | 792403.000 | 76025856 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 1 | 306353691.000 | 307692253.100 | 308036086.700 | — | — | — | 75927552 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 304065340.000 | 305595846.100 | 306451663.460 | 792379.000 | 792379.000 | 792379.000 | 75620352 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 306972266.500 | 308105624.550 | 308462251.100 | — | — | — | 75468800 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 303156260.500 | 305295612.000 | 306213885.140 | 792491.000 | 792491.000 | 792491.000 | 75460608 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 353575853.000 | 363785625.450 | 369422888.860 | — | — | — | 75665408 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 353783903.000 | 401042868.300 | 416657766.540 | 792491.000 | 792491.000 | 792491.000 | 75653120 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 1 | 429312.000 | 441855.450 | 443016.610 | — | — | — | 7008256 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 433642.000 | 442778.500 | 444697.300 | 8987347.000 | 8987347.000 | 8987347.000 | 7303168 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 52197844.000 | 52580186.050 | 52781800.610 | — | — | — | 20459520 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 53019732.000 | 53194844.000 | 53323418.700 | 8987347.000 | 8987347.000 | 8987347.000 | 20271104 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 302819975.000 | 304065415.500 | 304750022.330 | — | — | — | 81735680 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 302026613.000 | 303164461.000 | 303805348.730 | 8987347.000 | 8987347.000 | 8987347.000 | 81948672 |
| after | file_store-owned-s64-a64-short-c64 | normal | 1 | 3491029.000 | 3565583.500 | 3613307.100 | — | — | — | 6451200 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 1 | 3517899.000 | 3657707.450 | 3679395.200 | 598978.000 | 598978.000 | 598978.000 | 6713344 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 1 | 61180553.500 | 68689772.750 | 70746922.290 | — | — | — | 20779008 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 66262194.000 | 71414808.050 | 71643347.600 | 598984.000 | 598984.000 | 598984.000 | 21049344 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 1 | 307429185.000 | 324178151.650 | 325816920.260 | — | — | — | 83836928 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 308328090.000 | 311384092.500 | 312626151.260 | 598986.000 | 598986.000 | 598986.000 | 82214912 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 307207064.000 | 308701766.000 | 308949855.200 | 598986.000 | 598986.000 | 598986.000 | 82194432 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 2 | 307341946.500 | 311000316.850 | 312378498.650 | — | — | — | 81960960 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 59627801.500 | 60034418.500 | 60081491.900 | 598984.000 | 598984.000 | 598984.000 | 20799488 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 2 | 59176960.500 | 61408852.500 | 61483746.400 | — | — | — | 20713472 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 2 | 3858729.000 | 5753325.600 | 6636867.130 | 598978.000 | 598978.000 | 598978.000 | 6455296 |
| after | file_store-owned-s64-a64-short-c64 | normal | 2 | 3888859.500 | 6682262.150 | 7084446.400 | — | — | — | 6270976 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 303239948.000 | 304296081.700 | 304444995.910 | 8987347.000 | 8987347.000 | 8987347.000 | 81977344 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 306036389.500 | 307895960.050 | 308005571.190 | — | — | — | 82083840 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 53023687.000 | 53332681.500 | 53471326.610 | 8987347.000 | 8987347.000 | 8987347.000 | 20606976 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 52285334.000 | 52444600.500 | 52537456.200 | — | — | — | 20324352 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 434497.000 | 446288.500 | 450196.500 | 8987347.000 | 8987347.000 | 8987347.000 | 6766592 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 2 | 428426.500 | 437711.050 | 443989.890 | — | — | — | 6983680 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 351975384.000 | 353440839.250 | 354216174.110 | 792491.000 | 792491.000 | 792491.000 | 75689984 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 352240183.500 | 353606647.600 | 355108358.880 | — | — | — | 75558912 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 303217176.500 | 304432062.600 | 305407563.750 | 792491.000 | 792491.000 | 792491.000 | 75452416 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 303547669.000 | 304259265.000 | 305684775.160 | — | — | — | 75472896 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 306361401.000 | 307660018.550 | 307845609.400 | 792379.000 | 792379.000 | 792379.000 | 75780096 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 2 | 305777249.000 | 306955481.050 | 307539732.620 | — | — | — | 75460608 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 303889821.000 | 305384802.500 | 306365945.050 | 792403.000 | 792403.000 | 792403.000 | 75317248 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 306590571.000 | 309447023.550 | 310319812.240 | — | — | — | 75456512 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 91901802.000 | 94788120.050 | 97381845.030 | 792491.000 | 792491.000 | 792491.000 | 17723392 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 89713864.000 | 90287578.500 | 90503259.410 | — | — | — | 17747968 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 80484874.000 | 80811317.050 | 80835136.590 | 792491.000 | 792491.000 | 792491.000 | 17674240 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 78533776.500 | 79264940.550 | 79439899.000 | — | — | — | 18153472 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 132951852.000 | 134204631.000 | 134381981.300 | 792379.000 | 792379.000 | 792379.000 | 17719296 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 2 | 130379377.000 | 130799770.000 | 130889717.800 | — | — | — | 17940480 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 79306619.500 | 80130463.000 | 80162910.900 | 792403.000 | 792403.000 | 792403.000 | 17674240 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 76870285.000 | 77932624.500 | 78039726.310 | — | — | — | 17723392 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12148866.000 | 12239937.000 | 12257006.300 | 792491.000 | 792491.000 | 792491.000 | 6643712 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12143086.000 | 12344737.000 | 12670971.210 | — | — | — | 7041024 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 562977.500 | 570726.950 | 574336.100 | 792491.000 | 792491.000 | 792491.000 | 6492160 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 549102.000 | 558756.450 | 562910.800 | — | — | — | 7041024 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 2 | 818828.500 | 827112.550 | 830053.400 | 792379.000 | 792379.000 | 792379.000 | 6508544 |
| after | deterministic-file-s64-a64-short-c64 | normal | 2 | 805753.000 | 814899.000 | 817926.800 | — | — | — | 6787072 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 559362.000 | 566582.000 | 571391.200 | 792403.000 | 792403.000 | 792403.000 | 6504448 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 2 | 548442.000 | 559876.500 | 565713.000 | — | — | — | 6467584 |

## Matched before/after changes

Each row compares one arm, role, repeat, and metric. Percentages use `(after - before) / before`; lower values are favorable for elapsed time and operation heap.

| arm | role | repeat | metric | p50 change | p50 flag | p95 change | p95 flag | p99 change | p99 flag |
|---|---|---:|---|---:|---|---:|---|---:|---|
| deterministic-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -29.326% | True | -29.381% | True | -29.032% | True |
| deterministic-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 3.015% | False | 3.015% | False | 3.015% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -30.392% | True | -29.589% | True | -29.241% | True |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -1.559% | False | -1.559% | False | -1.559% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -17.967% | True | -17.777% | True | -18.157% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -3.602% | False | -3.602% | False | -3.602% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -17.856% | True | -17.654% | True | -17.509% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 2.188% | False | 2.188% | False | 2.188% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s64-a64-short-c64 | normal | 1 | elapsed_ns | -23.974% | True | -23.978% | True | -24.022% | True |
| deterministic-file-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.130% | False | -0.130% | False | -0.130% | False |
| deterministic-file-s64-a64-short-c64 | normal | 2 | elapsed_ns | -14.514% | True | -22.893% | True | -23.151% | True |
| deterministic-file-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.221% | False | 2.221% | False | 2.221% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -22.409% | True | -22.335% | True | -22.126% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -4.981% | False | -4.981% | False | -4.981% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -21.500% | True | -21.778% | True | -21.590% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -3.930% | False | -3.930% | False | -3.930% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | elapsed_ns | -17.651% | True | -17.136% | True | -17.138% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 7.814% | True | 7.814% | True | 7.814% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | elapsed_ns | -18.818% | True | -18.560% | True | -18.162% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 6.046% | True | 6.046% | True | 6.046% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -17.450% | True | -17.314% | True | -17.012% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 1.841% | False | 1.841% | False | 1.841% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -18.716% | True | -18.221% | True | -18.185% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -3.236% | False | -3.236% | False | -3.236% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | elapsed_ns | -1.015% | False | -1.939% | False | -0.934% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 1.108% | False | 1.108% | False | 1.108% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | elapsed_ns | -1.029% | False | 0.158% | False | 2.467% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 7.103% | True | 7.103% | True | 7.103% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -1.244% | False | -0.662% | False | -0.548% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 2.954% | False | 2.954% | False | 2.954% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -1.651% | False | -1.680% | False | -1.581% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 3.708% | False | 3.708% | False | 3.708% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -21.764% | True | -21.795% | True | -21.798% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 3.403% | False | 3.403% | False | 3.403% | False |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -21.440% | True | -20.876% | True | -21.107% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 0.698% | False | 0.698% | False | 0.698% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -20.745% | True | -20.428% | True | -20.476% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.723% | False | 1.723% | False | 1.723% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -20.251% | True | -20.257% | True | -20.505% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | -2.043% | False | -2.043% | False | -2.043% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -14.922% | True | -15.695% | True | -15.847% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 2.075% | False | 2.075% | False | 2.075% | False |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -14.231% | True | -14.910% | True | -14.995% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.601% | False | 1.601% | False | 1.601% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -13.174% | True | -13.135% | True | -13.097% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.739% | False | 1.739% | False | 1.739% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -13.660% | True | -13.522% | True | -13.479% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 0.769% | False | 0.769% | False | 0.769% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -22.051% | True | -22.835% | True | -22.728% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 2.413% | False | 2.413% | False | 2.413% | False |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -20.049% | True | -20.261% | True | -20.452% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 3.238% | False | 3.238% | False | 3.238% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -21.055% | True | -21.201% | True | -21.384% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 2.831% | False | 2.831% | False | 2.831% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -20.452% | True | -20.826% | True | -20.901% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 0.162% | False | 0.162% | False | 0.162% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -20.890% | True | -23.801% | True | -23.460% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 2.556% | False | 2.556% | False | 2.556% | False |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -20.827% | True | -21.414% | True | -21.419% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.333% | False | 1.333% | False | 1.333% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -17.977% | True | -21.441% | True | -21.220% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.550% | False | 1.550% | False | 1.550% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -17.089% | True | -14.846% | True | -12.591% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 0.862% | False | 0.862% | False | 0.862% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -21.305% | True | -21.444% | True | -21.279% | True |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.228% | False | 0.228% | False | 0.228% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -20.291% | True | -19.888% | True | -20.048% | True |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.169% | False | 0.169% | False | 0.169% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -21.239% | True | -21.720% | True | -22.079% | True |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.460% | False | 0.460% | False | 0.460% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -21.108% | True | -21.158% | True | -21.580% | True |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.379% | False | -0.379% | False | -0.379% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -20.769% | True | -20.628% | True | -20.770% | True |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.772% | False | 0.772% | False | 0.772% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -20.883% | True | -20.896% | True | -21.388% | True |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.114% | False | 0.114% | False | 0.114% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -21.078% | True | -21.199% | True | -21.276% | True |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.211% | False | -0.211% | False | -0.211% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -21.025% | True | -21.527% | True | -21.827% | True |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.636% | False | 0.636% | False | 0.636% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.894% | True | -19.894% | True | -19.894% | True |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -20.065% | True | -20.285% | True | -20.688% | True |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.207% | False | 0.207% | False | 0.207% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -20.951% | True | -20.992% | True | -20.698% | True |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.245% | False | 0.245% | False | 0.245% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -21.555% | True | -21.291% | True | -21.186% | True |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.319% | False | -0.319% | False | -0.319% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -21.255% | True | -22.488% | True | -22.342% | True |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.332% | False | 0.332% | False | 0.332% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -18.242% | True | -16.385% | True | -15.106% | True |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.555% | False | 0.555% | False | 0.555% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -18.638% | True | -18.594% | True | -18.386% | True |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.250% | False | 0.250% | False | 0.250% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -18.795% | True | -8.701% | True | -6.749% | True |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.490% | False | 0.490% | False | 0.490% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -18.972% | True | -19.087% | True | -19.475% | True |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.714% | False | 0.714% | False | 0.714% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -19.892% | True | -19.892% | True | -19.892% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -20.407% | True | -19.510% | True | -21.075% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 2.516% | False | 2.516% | False | 2.516% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -20.343% | True | -19.853% | True | -18.927% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.341% | False | 2.341% | False | 2.341% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -21.716% | True | -21.319% | True | -21.798% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 9.858% | True | 9.858% | True | 9.858% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -22.414% | True | -21.460% | True | -21.293% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -4.729% | False | -4.729% | False | -4.729% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -23.821% | True | -24.815% | True | -24.561% | True |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 0.624% | False | 0.624% | False | 0.624% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -23.529% | True | -23.938% | True | -24.564% | True |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -1.195% | False | -1.195% | False | -1.195% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -24.303% | True | -24.569% | True | -24.530% | True |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | -0.483% | False | -0.483% | False | -0.483% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -24.746% | True | -24.778% | True | -24.760% | True |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 1.452% | False | 1.452% | False | 1.452% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -21.119% | True | -21.539% | True | -21.818% | True |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.180% | False | -0.180% | False | -0.180% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -20.290% | True | -20.089% | True | -20.214% | True |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.356% | False | 0.356% | False | 0.356% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -21.505% | True | -21.594% | True | -21.488% | True |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.010% | False | -0.010% | False | -0.010% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -21.143% | True | -21.998% | True | -22.251% | True |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -1.540% | False | -1.540% | False | -1.540% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -2.143% | False | -2.143% | False | -2.143% | False |
| file_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -5.066% | True | -7.541% | True | -23.160% | True |
| file_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 2.206% | False | 2.206% | False | 2.206% | False |
| file_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | 7.185% | True | 79.002% | True | 82.598% | True |
| file_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -1.670% | False | -1.670% | False | -1.670% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -1.728% | False | -1.130% | False | -0.609% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 3.472% | False | 3.472% | False | 3.472% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | 4.324% | False | 35.240% | True | 51.740% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -2.836% | False | -2.836% | False | -2.836% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -20.503% | True | -11.803% | True | -9.336% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 1.318% | False | 1.318% | False | 1.318% | False |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -21.656% | True | -19.620% | True | -19.930% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.648% | False | 1.648% | False | 1.648% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -13.796% | True | -7.838% | True | -7.666% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.321% | False | 1.321% | False | 1.321% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -22.843% | True | -24.110% | True | -24.318% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 0.495% | False | 0.495% | False | 0.495% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -20.405% | True | -16.465% | True | -16.313% | True |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 2.146% | False | 2.146% | False | 2.146% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -20.753% | True | -19.967% | True | -19.624% | True |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.005% | False | 0.005% | False | 0.005% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -20.540% | True | -19.896% | True | -19.656% | True |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.055% | False | -0.055% | False | -0.055% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -20.620% | True | -21.056% | True | -21.565% | True |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.365% | False | 0.365% | False | 0.365% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -24.729% | True | -24.729% | True | -24.729% | True |

## Measurement limits

- The matrix covers three exact workloads and selected input/store arms; it is not a full provider-by-input interaction grid.
- Two process repeats support a descriptive before/after range; they do not establish a strong confidence interval.
- A five-percent flag is a review trigger, not a causal speedup claim.
- GNU time observes the whole child once and is retained with n=1; allocator operation heap is a separate instrumented metric.
