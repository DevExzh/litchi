# 0487 OPC consumed-prefix retention comparison

This table retains every formal process row. Values are descriptive; a five-percent flag marks a review threshold and does not authorize a causal speedup claim.

Protocol `cf2c061fb909377016f2ad56b0d6fd240be117c799a954fa522e83c8b79cd52c`; before attempt `formal2`; after attempt `formal1`.

The matrix contains 144 formal children and 4320 measured samples across 18 arms.

## Source custody

Before source manifest: `c4594536bc2b80c6693be6bf5617ed7e699dd117945bc13183338fb0fd437abe` (7157 files). After source manifest: `cb69be7b5a1f443191a27000fdfab285b707ba28517cccc3ee764d833cde0ba2` (7157 files).

Changed source files: 2; all changes in the configured OPC implementation/test allowlist: `True`.

Matched candidate ZIP archive identities changed for 0 process pairs; decoded candidate content identity is checked independently and archive framing is retained per row.

## Per-process rows

| phase | arm | role | repeat | elapsed p50 (ns) | elapsed p95 (ns) | elapsed p99 (ns) | allocator operation heap p50 (B) | heap p95 (B) | heap p99 (B) | GNU time RSS (B, n=1) |
|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| before | deterministic-owned-s64-a64-short-c64 | normal | 1 | 716568.000 | 764425.550 | 818109.900 | — | — | — | 6676480 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 722767.500 | 728953.000 | 733750.300 | 989190.000 | 989190.000 | 989190.000 | 6582272 |
| before | deterministic-file-s64-a64-short-c64 | normal | 1 | 1370800.000 | 1382488.500 | 1390230.700 | — | — | — | 6758400 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 1 | 1303899.500 | 1310145.950 | 1314328.110 | 989166.000 | 989166.000 | 989166.000 | 7036928 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 718467.500 | 798172.500 | 807945.400 | — | — | — | 6443008 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 728307.500 | 735038.500 | 740355.790 | 989278.000 | 989278.000 | 989278.000 | 6574080 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12290604.000 | 12430959.050 | 12794539.920 | — | — | — | 6975488 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12356119.500 | 12646127.600 | 12810386.790 | 989278.000 | 989278.000 | 989278.000 | 6565888 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 106738135.000 | 107501963.050 | 107585575.400 | — | — | — | 17842176 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 110861384.500 | 111989334.550 | 112169504.610 | 989190.000 | 989190.000 | 989190.000 | 17891328 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 1 | 239982306.500 | 241865106.500 | 242728042.440 | — | — | — | 17842176 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 244117311.000 | 246568213.100 | 247280488.730 | 989166.000 | 989166.000 | 989166.000 | 17723392 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 108282349.000 | 109049716.000 | 109137663.000 | — | — | — | 17842176 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 112395743.500 | 112828310.000 | 113010829.320 | 989278.000 | 989278.000 | 989278.000 | 17743872 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 120439742.000 | 124782617.350 | 125759627.510 | — | — | — | 17686528 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 124617882.500 | 128761474.700 | 130149314.050 | 989278.000 | 989278.000 | 989278.000 | 17920000 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 388199275.000 | 390920131.050 | 391445756.020 | — | — | — | 75644928 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 386141062.000 | 387311824.050 | 387613834.020 | 989190.000 | 989190.000 | 989190.000 | 75452416 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 1 | 387057816.500 | 390952043.250 | 392688164.930 | — | — | — | 75583488 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 388523383.500 | 390597733.600 | 391057501.310 | 989166.000 | 989166.000 | 989166.000 | 75612160 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 388244552.500 | 389419505.850 | 389574858.200 | — | — | — | 75780096 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 385330749.000 | 387069904.000 | 387512838.520 | 989278.000 | 989278.000 | 989278.000 | 75608064 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 435468771.500 | 438677359.800 | 444484214.350 | — | — | — | 75665408 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 435073544.000 | 438253005.150 | 441011132.120 | 989278.000 | 989278.000 | 989278.000 | 75714560 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 1 | 546712.000 | 555979.000 | 559038.400 | — | — | — | 7225344 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 549622.000 | 557234.500 | 558022.000 | 9184134.000 | 9184134.000 | 9184134.000 | 6836224 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 69093793.500 | 70183756.000 | 70463727.020 | — | — | — | 20643840 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 70363448.000 | 70720510.550 | 70822376.800 | 9184134.000 | 9184134.000 | 9184134.000 | 20758528 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 385660406.000 | 387424479.250 | 388346241.870 | — | — | — | 81920000 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 384269363.500 | 385435858.750 | 385454492.560 | 9184134.000 | 9184134.000 | 9184134.000 | 83607552 |
| before | file_store-owned-s64-a64-short-c64 | normal | 1 | 3581242.500 | 3665298.000 | 3677844.090 | — | — | — | 6676480 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 1 | 3602763.500 | 3690050.000 | 3765740.210 | 795767.000 | 795767.000 | 795767.000 | 6459392 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 1 | 76641216.000 | 80436847.450 | 83401834.520 | — | — | — | 20770816 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 77496794.500 | 79201548.500 | 79377168.600 | 795773.000 | 795773.000 | 795773.000 | 20619264 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 1 | 387668657.000 | 392642973.700 | 393999337.540 | — | — | — | 82370560 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 387305514.500 | 390999546.150 | 391130899.400 | 795775.000 | 795775.000 | 795775.000 | 82096128 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 387562390.500 | 390194639.250 | 391497139.340 | 795775.000 | 795775.000 | 795775.000 | 82264064 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 2 | 390841131.500 | 395599462.750 | 398848850.930 | — | — | — | 81481728 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 85003866.000 | 90294486.550 | 90492172.510 | 795773.000 | 795773.000 | 795773.000 | 20758528 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 2 | 83146679.000 | 88504482.650 | 90949902.400 | — | — | — | 20598784 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 2 | 3688143.000 | 4322832.050 | 4405458.800 | 795767.000 | 795767.000 | 795767.000 | 6860800 |
| before | file_store-owned-s64-a64-short-c64 | normal | 2 | 3553042.500 | 3658624.550 | 3704281.300 | — | — | — | 6443008 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 384365863.000 | 385545862.100 | 385693691.900 | 9184134.000 | 9184134.000 | 9184134.000 | 82055168 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 384107723.000 | 385000637.550 | 385626923.640 | — | — | — | 82149376 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 71382502.000 | 72113852.450 | 72323028.510 | 9184134.000 | 9184134.000 | 9184134.000 | 20594688 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 69713716.500 | 70371043.050 | 70497020.800 | — | — | — | 20905984 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 549647.000 | 560914.500 | 563614.600 | 9184134.000 | 9184134.000 | 9184134.000 | 7098368 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 2 | 548132.000 | 555595.500 | 557050.000 | — | — | — | 7094272 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 434599477.000 | 435559254.050 | 435711694.700 | 989278.000 | 989278.000 | 989278.000 | 75644928 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 435954629.500 | 437571183.550 | 438069103.520 | — | — | — | 75583488 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 384462439.000 | 385313657.500 | 385613953.120 | 989278.000 | 989278.000 | 989278.000 | 75489280 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 384501657.000 | 385590678.000 | 386763840.860 | — | — | — | 75624448 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 388002349.500 | 389165278.300 | 389651015.680 | 989166.000 | 989166.000 | 989166.000 | 75706368 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 2 | 387389285.500 | 388406131.100 | 389162264.210 | — | — | — | 75550720 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 385408151.000 | 387655520.750 | 389940463.300 | 989190.000 | 989190.000 | 989190.000 | 75919360 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 385679553.000 | 386937443.550 | 387226747.710 | — | — | — | 75575296 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 122605919.500 | 124938642.050 | 125399656.320 | 989278.000 | 989278.000 | 989278.000 | 17801216 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 120443942.000 | 121755067.050 | 124384492.120 | — | — | — | 17727488 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 110927837.500 | 111693356.500 | 111846731.210 | 989278.000 | 989278.000 | 989278.000 | 18178048 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 108845990.500 | 109539976.050 | 109845900.110 | — | — | — | 17674240 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 244627258.500 | 246239781.100 | 246656392.810 | 989166.000 | 989166.000 | 989166.000 | 17928192 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 2 | 238836888.500 | 240615485.500 | 241118156.310 | — | — | — | 17772544 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 111096004.500 | 112658835.100 | 112925036.000 | 989190.000 | 989190.000 | 989190.000 | 17940480 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 108125518.500 | 108964762.050 | 109259712.310 | — | — | — | 17833984 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12368390.500 | 13300576.150 | 16236970.240 | 989278.000 | 989278.000 | 989278.000 | 6930432 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12357645.000 | 13156450.100 | 15153441.590 | — | — | — | 6639616 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 733517.500 | 739796.450 | 744862.410 | 989278.000 | 989278.000 | 989278.000 | 6529024 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 717428.000 | 728343.050 | 735015.600 | — | — | — | 6791168 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 2 | 1378775.000 | 1402214.000 | 1411459.800 | 989166.000 | 989166.000 | 989166.000 | 6569984 |
| before | deterministic-file-s64-a64-short-c64 | normal | 2 | 1396820.000 | 1410900.000 | 1416227.690 | — | — | — | 6848512 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 728722.500 | 743178.000 | 749684.100 | 989190.000 | 989190.000 | 989190.000 | 6967296 |
| before | deterministic-owned-s64-a64-short-c64 | normal | 2 | 715353.000 | 762378.550 | 788362.300 | — | — | — | 7163904 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 1 | 772038.000 | 785401.500 | 791926.100 | — | — | — | 6787072 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 785887.000 | 792821.000 | 794121.700 | 989190.000 | 989190.000 | 989190.000 | 6447104 |
| after | deterministic-file-s64-a64-short-c64 | normal | 1 | 1053844.000 | 1062706.000 | 1066481.500 | — | — | — | 6500352 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 1 | 968683.500 | 1055266.500 | 1082904.200 | 989166.000 | 989166.000 | 989166.000 | 6729728 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 670728.000 | 676691.500 | 680796.100 | — | — | — | 6836224 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 779038.000 | 787834.500 | 789639.600 | 989278.000 | 989278.000 | 989278.000 | 6615040 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12391275.000 | 12454495.050 | 12467152.800 | — | — | — | 6836224 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12302440.000 | 12337375.550 | 12357026.790 | 989278.000 | 989278.000 | 989278.000 | 6905856 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 97113070.500 | 97721432.950 | 97999250.220 | — | — | — | 17752064 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 98375405.000 | 98891696.050 | 99124752.310 | 989190.000 | 989190.000 | 989190.000 | 17915904 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 1 | 153609397.500 | 154174993.600 | 154380209.700 | — | — | — | 18006016 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 154760369.500 | 156052906.050 | 156335450.820 | 989166.000 | 989166.000 | 989166.000 | 17944576 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 99714924.500 | 100912915.000 | 101007404.900 | — | — | — | 18079744 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 99780954.500 | 100105992.500 | 100402625.410 | 989278.000 | 989278.000 | 989278.000 | 18186240 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 108850066.500 | 109610001.000 | 109692326.510 | — | — | — | 17584128 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 110961782.500 | 111862323.050 | 112059626.820 | 989278.000 | 989278.000 | 989278.000 | 17752064 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 384656956.000 | 386987793.800 | 388073972.520 | — | — | — | 75706368 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 386365887.500 | 388321315.000 | 388563917.710 | 989190.000 | 989190.000 | 989190.000 | 75665408 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 1 | 389228688.500 | 390572717.000 | 390844953.310 | — | — | — | 75874304 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 390458714.500 | 393937027.100 | 394939051.840 | 989166.000 | 989166.000 | 989166.000 | 75472896 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 386700432.000 | 388354690.150 | 389142946.480 | — | — | — | 75489280 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 387733234.000 | 389167565.550 | 389715580.930 | 989278.000 | 989278.000 | 989278.000 | 75898880 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 435548588.500 | 440840027.250 | 443469577.810 | — | — | — | 75542528 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 436350367.000 | 437628667.050 | 438582449.630 | 989278.000 | 989278.000 | 989278.000 | 75636736 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 1 | 540057.500 | 554033.000 | 562315.010 | — | — | — | 7077888 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 555172.500 | 564671.000 | 576011.410 | 9184134.000 | 9184134.000 | 9184134.000 | 7397376 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 68873621.000 | 69385987.050 | 69400381.000 | — | — | — | 21045248 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 70403611.000 | 70792143.050 | 70905189.980 | 9184134.000 | 9184134.000 | 9184134.000 | 20525056 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 387778221.000 | 389211012.050 | 389673167.820 | — | — | — | 81907712 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 388784074.500 | 390178009.300 | 390425583.550 | 9184134.000 | 9184134.000 | 9184134.000 | 82092032 |
| after | file_store-owned-s64-a64-short-c64 | normal | 1 | 3556429.500 | 3624328.550 | 3656308.900 | — | — | — | 6602752 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 1 | 3572419.000 | 3669746.500 | 3724326.010 | 795765.000 | 795765.000 | 795765.000 | 6492160 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 1 | 75926434.000 | 76903615.000 | 77779568.140 | — | — | — | 20721664 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 77062213.000 | 78589013.300 | 79212570.100 | 795771.000 | 795771.000 | 795771.000 | 20475904 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 1 | 389262102.500 | 391453352.100 | 393795921.710 | — | — | — | 82329600 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 389451600.000 | 395819000.350 | 399811219.400 | 795773.000 | 795773.000 | 795773.000 | 82489344 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 389894602.500 | 391434470.050 | 392136514.860 | 795773.000 | 795773.000 | 795773.000 | 83460096 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 2 | 390454338.000 | 391518132.300 | 393000239.080 | — | — | — | 83435520 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 78553777.000 | 79791853.950 | 80104245.020 | 795771.000 | 795771.000 | 795771.000 | 20635648 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 2 | 76020087.500 | 76892506.550 | 79047517.910 | — | — | — | 20791296 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 2 | 3687289.000 | 4188738.500 | 4319759.610 | 795765.000 | 795765.000 | 795765.000 | 6250496 |
| after | file_store-owned-s64-a64-short-c64 | normal | 2 | 3565578.000 | 3635503.000 | 3643492.400 | — | — | — | 6574080 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 387398386.000 | 388651797.500 | 389554012.250 | 9184134.000 | 9184134.000 | 9184134.000 | 83845120 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 386435743.000 | 388204915.050 | 389747762.560 | — | — | — | 82350080 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 70252765.000 | 70796059.500 | 70809265.600 | 9184134.000 | 9184134.000 | 9184134.000 | 20451328 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 69250986.000 | 71001118.550 | 71107492.600 | — | — | — | 20574208 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 558202.000 | 566004.450 | 578331.410 | 9184134.000 | 9184134.000 | 9184134.000 | 6971392 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 2 | 545416.500 | 557406.500 | 560772.910 | — | — | — | 7389184 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 436362503.500 | 438358442.600 | 439869854.060 | 989278.000 | 989278.000 | 989278.000 | 75436032 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 435071238.000 | 437523372.700 | 438577517.630 | — | — | — | 75436032 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 385866133.000 | 387397406.550 | 387910511.620 | 989278.000 | 989278.000 | 989278.000 | 75689984 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 384333108.000 | 385527634.500 | 388207658.940 | — | — | — | 75366400 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 389530330.500 | 390456427.050 | 390808287.960 | 989166.000 | 989166.000 | 989166.000 | 75685888 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 2 | 388756179.500 | 390864768.950 | 391635669.960 | — | — | — | 75337728 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 389224714.000 | 390879449.500 | 391251574.920 | 989190.000 | 989190.000 | 989190.000 | 75579392 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 385571702.000 | 386299600.100 | 386449528.600 | — | — | — | 75284480 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 111416296.500 | 112324379.500 | 112437062.200 | 989278.000 | 989278.000 | 989278.000 | 17989632 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 110955179.500 | 111748786.050 | 111970118.610 | — | — | — | 17580032 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 99566527.000 | 100078497.550 | 100231460.200 | 989278.000 | 989278.000 | 989278.000 | 18251776 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 98313792.000 | 98735021.000 | 98889161.100 | — | — | — | 17563648 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 156538880.500 | 157445151.000 | 157772968.720 | 989166.000 | 989166.000 | 989166.000 | 17793024 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 2 | 153777755.500 | 154548770.100 | 155037795.510 | — | — | — | 18042880 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 98834480.000 | 99675959.650 | 100413234.920 | 989190.000 | 989190.000 | 989190.000 | 17805312 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 98898547.000 | 100246132.050 | 100578512.510 | — | — | — | 18112512 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12305151.000 | 12352376.550 | 12848025.220 | 989278.000 | 989278.000 | 989278.000 | 6963200 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12405191.500 | 13161145.600 | 13656518.110 | — | — | — | 6815744 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 775603.000 | 781931.550 | 784049.600 | 989278.000 | 989278.000 | 989278.000 | 6737920 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 750078.000 | 758934.000 | 763794.290 | — | — | — | 6447104 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 2 | 966244.000 | 1087102.500 | 1089679.900 | 989166.000 | 989166.000 | 989166.000 | 6455296 |
| after | deterministic-file-s64-a64-short-c64 | normal | 2 | 935439.000 | 945272.050 | 947880.400 | — | — | — | 6508544 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 684713.000 | 694512.000 | 698076.800 | 989190.000 | 989190.000 | 989190.000 | 6631424 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 2 | 669442.000 | 679553.450 | 753613.610 | — | — | — | 6311936 |

## Matched before/after changes

Each row compares one arm, role, repeat, and metric. Percentages use `(after - before) / before`; lower values are favorable for elapsed time and operation heap.

| arm | role | repeat | metric | p50 change | p50 flag | p95 change | p95 flag | p99 change | p99 flag |
|---|---|---:|---|---:|---|---:|---|---:|---|
| deterministic-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | 7.741% | True | 2.744% | False | -3.201% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 1.656% | False | 1.656% | False | 1.656% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -6.418% | True | -10.864% | True | -4.408% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -11.893% | True | -11.893% | True | -11.893% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | 8.733% | True | 8.762% | True | 8.228% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -2.054% | False | -2.054% | False | -2.054% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -6.039% | True | -6.548% | True | -6.884% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -4.821% | False | -4.821% | False | -4.821% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a64-short-c64 | normal | 1 | elapsed_ns | -23.122% | True | -23.131% | True | -23.287% | True |
| deterministic-file-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -3.818% | False | -3.818% | False | -3.818% | False |
| deterministic-file-s64-a64-short-c64 | normal | 2 | elapsed_ns | -33.031% | True | -33.002% | True | -33.070% | True |
| deterministic-file-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -4.964% | False | -4.964% | False | -4.964% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -25.709% | True | -19.454% | True | -17.608% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -4.366% | False | -4.366% | False | -4.366% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -29.920% | True | -22.472% | True | -22.798% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -1.746% | False | -1.746% | False | -1.746% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | elapsed_ns | -6.645% | True | -15.220% | True | -15.737% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 6.103% | True | 6.103% | True | 6.103% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | elapsed_ns | 4.551% | False | 4.200% | False | 3.915% | False |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -5.066% | True | -5.066% | True | -5.066% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | elapsed_ns | 6.966% | True | 7.183% | True | 6.657% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.623% | False | 0.623% | False | 0.623% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | elapsed_ns | 5.737% | True | 5.695% | True | 5.261% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 3.199% | False | 3.199% | False | 3.199% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | elapsed_ns | 0.819% | False | 0.189% | False | -2.559% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -1.996% | False | -1.996% | False | -1.996% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | elapsed_ns | 0.385% | False | 0.036% | False | -9.878% | True |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.653% | False | 2.653% | False | 2.653% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -0.434% | False | -2.441% | False | -3.539% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 5.178% | True | 5.178% | True | 5.178% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -0.511% | False | -7.129% | True | -20.872% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.473% | False | 0.473% | False | 0.473% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -9.017% | True | -9.098% | True | -8.910% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | -0.505% | False | -0.505% | False | -0.505% | False |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -8.534% | True | -8.001% | True | -7.945% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.562% | False | 1.562% | False | 1.562% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -11.263% | True | -11.695% | True | -11.629% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 0.137% | False | 0.137% | False | 0.137% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -11.037% | True | -11.524% | True | -11.080% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | -0.753% | False | -0.753% | False | -0.753% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -35.991% | True | -36.256% | True | -36.398% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 0.918% | False | 0.918% | False | 0.918% | False |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -35.614% | True | -35.769% | True | -35.700% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.521% | False | 1.521% | False | 1.521% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -36.604% | True | -36.710% | True | -36.778% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.248% | False | 1.248% | False | 1.248% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -36.009% | True | -36.060% | True | -36.035% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | -0.754% | False | -0.754% | False | -0.754% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -7.912% | True | -7.462% | True | -7.450% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 1.331% | False | 1.331% | False | 1.331% | False |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -9.676% | True | -9.864% | True | -9.975% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -0.626% | False | -0.626% | False | -0.626% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -11.224% | True | -11.276% | True | -11.157% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 2.493% | False | 2.493% | False | 2.493% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -10.242% | True | -10.399% | True | -10.385% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 0.406% | False | 0.406% | False | 0.406% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -9.623% | True | -12.159% | True | -12.776% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | -0.579% | False | -0.579% | False | -0.579% | False |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -7.878% | True | -8.218% | True | -9.981% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -0.832% | False | -0.832% | False | -0.832% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -10.958% | True | -13.124% | True | -13.899% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | -0.937% | False | -0.937% | False | -0.937% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -9.126% | True | -10.096% | True | -10.337% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 1.058% | False | 1.058% | False | 1.058% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -0.913% | False | -1.006% | False | -0.861% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.081% | False | 0.081% | False | 0.081% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -0.028% | False | -0.165% | False | -0.201% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.385% | False | -0.385% | False | -0.385% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 0.058% | False | 0.261% | False | 0.245% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.282% | False | 0.282% | False | 0.282% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.990% | False | 0.832% | False | 0.336% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.448% | False | -0.448% | False | -0.448% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | elapsed_ns | 0.561% | False | -0.097% | False | -0.469% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.385% | False | 0.385% | False | 0.385% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | elapsed_ns | 0.353% | False | 0.633% | False | 0.636% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.282% | False | -0.282% | False | -0.282% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 0.498% | False | 0.855% | False | 0.993% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.184% | False | -0.184% | False | -0.184% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.394% | False | 0.332% | False | 0.297% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.027% | False | -0.027% | False | -0.027% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -0.398% | False | -0.273% | False | -0.111% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.384% | False | -0.384% | False | -0.384% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -0.044% | False | -0.016% | False | 0.373% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.341% | False | -0.341% | False | -0.341% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 0.623% | False | 0.542% | False | 0.568% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.385% | False | 0.385% | False | 0.385% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.365% | False | 0.541% | False | 0.596% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.266% | False | 0.266% | False | 0.266% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | elapsed_ns | 0.018% | False | 0.493% | False | -0.228% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.162% | False | -0.162% | False | -0.162% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -0.203% | False | -0.011% | False | 0.116% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.195% | False | -0.195% | False | -0.195% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 0.293% | False | -0.142% | False | -0.551% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.103% | False | -0.103% | False | -0.103% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.406% | False | 0.643% | False | 0.954% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.276% | False | -0.276% | False | -0.276% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -1.217% | False | -0.350% | False | 0.586% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -2.041% | False | -2.041% | False | -2.041% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -0.495% | False | 0.326% | False | 0.668% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 4.157% | False | 4.157% | False | 4.157% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | 1.010% | False | 1.335% | False | 3.224% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 8.209% | True | 8.209% | True | 8.209% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | 1.556% | False | 0.907% | False | 2.611% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -1.789% | False | -1.789% | False | -1.789% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -0.319% | False | -1.137% | False | -1.509% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 1.944% | False | 1.944% | False | 1.944% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -0.664% | False | 0.895% | False | 0.866% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -1.587% | False | -1.587% | False | -1.587% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | 0.057% | False | 0.101% | False | 0.117% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | -1.125% | False | -1.125% | False | -1.125% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -1.583% | False | -1.827% | False | -2.093% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | -0.696% | False | -0.696% | False | -0.696% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | 0.549% | False | 0.461% | False | 0.342% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.015% | False | -0.015% | False | -0.015% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | 0.606% | False | 0.832% | False | 1.069% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.244% | False | 0.244% | False | 0.244% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 1.175% | False | 1.230% | False | 1.290% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -1.813% | False | -1.813% | False | -1.813% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.789% | False | 0.806% | False | 1.001% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 2.181% | False | 2.181% | False | 2.181% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| file_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -0.693% | False | -1.118% | False | -0.586% | False |
| file_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -1.104% | False | -1.104% | False | -1.104% | False |
| file_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | 0.353% | False | -0.632% | False | -1.641% | False |
| file_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.034% | False | 2.034% | False | 2.034% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -0.842% | False | -0.550% | False | -1.100% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.507% | False | 0.507% | False | 0.507% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -0.023% | False | -3.102% | False | -1.945% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -8.896% | True | -8.896% | True | -8.896% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -0.933% | False | -4.393% | False | -6.741% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | -0.237% | False | -0.237% | False | -0.237% | False |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -8.571% | True | -13.120% | True | -13.087% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 0.935% | False | 0.935% | False | 0.935% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -0.561% | False | -0.773% | False | -0.207% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | -0.695% | False | -0.695% | False | -0.695% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -7.588% | True | -11.632% | True | -11.479% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | -0.592% | False | -0.592% | False | -0.592% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | 0.411% | False | -0.303% | False | -0.052% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.050% | False | -0.050% | False | -0.050% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -0.099% | False | -1.032% | False | -1.466% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.398% | False | 2.398% | False | 2.398% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | 0.554% | False | 1.233% | False | 2.219% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.479% | False | 0.479% | False | 0.479% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | 0.602% | False | 0.318% | False | 0.163% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 1.454% | False | 1.454% | False | 1.454% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |

## Measurement limits

- The matrix covers three exact workloads and selected input/store arms; it is not a full provider-by-input interaction grid.
- Two process repeats support a descriptive before/after range; they do not establish a strong confidence interval.
- A five-percent flag is a review trigger, not a causal speedup claim.
- GNU time observes the whole child once and is retained with n=1; allocator operation heap is a separate instrumented metric.
