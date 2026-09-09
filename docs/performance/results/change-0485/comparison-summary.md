# 0485 OPC splice-audit batching comparison

This table retains every formal process row. Values are descriptive; a five-percent flag marks a review threshold and does not authorize a causal speedup claim.

Protocol `c9928d7a5d70d31c2447f434f81d908ed29c3adc4701f4ea32855005fc974ea5`; before attempt `formal1`; after attempt `formal1`.

The matrix contains 144 formal children and 4320 measured samples across 18 arms.

## Source custody

Before source manifest: `5a86d19865374a7e914ebb2f168273800b4e04de5ae287bd7ecc4fe37c29b0ae` (7157 files). After source manifest: `c4594536bc2b80c6693be6bf5617ed7e699dd117945bc13183338fb0fd437abe` (7157 files).

Changed source files: 1; all changes in the configured OPC implementation/test allowlist: `True`.

Matched candidate ZIP archive identities changed for 0 process pairs; decoded candidate content identity is checked independently and archive framing is retained per row.

## Per-process rows

| phase | arm | role | repeat | elapsed p50 (ns) | elapsed p95 (ns) | elapsed p99 (ns) | allocator operation heap p50 (B) | heap p95 (B) | heap p99 (B) | GNU time RSS (B, n=1) |
|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| before | deterministic-owned-s64-a64-short-c64 | normal | 1 | 807418.500 | 814052.500 | 815565.200 | — | — | — | 6787072 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 842614.000 | 891397.950 | 892086.500 | 989190.000 | 989190.000 | 989190.000 | 6569984 |
| before | deterministic-file-s64-a64-short-c64 | normal | 1 | 2201423.000 | 2224489.000 | 2225509.700 | — | — | — | 6897664 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 1 | 2106623.500 | 2120466.550 | 2125943.500 | 989166.000 | 989166.000 | 989166.000 | 6541312 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 808493.000 | 815631.050 | 816635.590 | — | — | — | 6602752 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 810688.000 | 817628.950 | 818026.000 | 989278.000 | 989278.000 | 989278.000 | 6553600 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12436134.000 | 12771660.000 | 12998732.620 | — | — | — | 6709248 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12474509.000 | 12945134.150 | 25514859.080 | 989278.000 | 989278.000 | 989278.000 | 6590464 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 117792042.500 | 119451264.050 | 119887538.820 | — | — | — | 17788928 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 119728765.500 | 121103485.500 | 121336144.610 | 989190.000 | 989190.000 | 989190.000 | 17793024 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 1 | 443915796.500 | 445363415.500 | 446962646.490 | — | — | — | 17793024 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 439285819.000 | 440599789.750 | 441165000.390 | 989166.000 | 989166.000 | 989166.000 | 17522688 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 118917367.500 | 119360416.200 | 119664968.220 | — | — | — | 17846272 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 120450180.000 | 121056750.500 | 121116908.910 | 989278.000 | 989278.000 | 989278.000 | 17793024 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 131074231.000 | 133017236.200 | 135001779.880 | — | — | — | 17846272 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 133043498.500 | 135657837.550 | 136286159.430 | 989278.000 | 989278.000 | 989278.000 | 17793024 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 482325351.500 | 486336379.150 | 486715703.210 | — | — | — | 75268096 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 479543336.000 | 486339865.650 | 487088885.820 | 989190.000 | 989190.000 | 989190.000 | 75526144 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 1 | 479811477.000 | 482971576.100 | 483466222.420 | — | — | — | 75452416 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 490487601.500 | 492437480.000 | 492560742.200 | 989166.000 | 989166.000 | 989166.000 | 75255808 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 480251536.500 | 482412612.000 | 482711365.640 | — | — | — | 75657216 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 482255643.500 | 484451810.550 | 484750193.210 | 989278.000 | 989278.000 | 989278.000 | 75522048 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 528228257.000 | 531036110.000 | 531131843.200 | — | — | — | 75558912 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 528464428.000 | 532573901.550 | 537696023.180 | 989278.000 | 989278.000 | 989278.000 | 75186176 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 1 | 673688.000 | 682101.950 | 682681.800 | — | — | — | 6651904 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 684887.000 | 693970.050 | 699310.200 | 9184134.000 | 9184134.000 | 9184134.000 | 6807552 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 90550720.500 | 91264266.000 | 91343183.410 | — | — | — | 20549632 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 94199267.000 | 94503156.000 | 94793572.940 | 9184134.000 | 9184134.000 | 9184134.000 | 20279296 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 481750677.000 | 483195784.550 | 484174483.350 | — | — | — | 81743872 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 479223600.000 | 483742096.550 | 483992802.100 | 9184134.000 | 9184134.000 | 9184134.000 | 81387520 |
| before | file_store-owned-s64-a64-short-c64 | normal | 1 | 4747143.000 | 5804432.100 | 6976704.760 | — | — | — | 6553600 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 1 | 4555198.000 | 6641397.550 | 7181610.120 | 795767.000 | 795767.000 | 795767.000 | 6144000 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 1 | 99297888.000 | 105358406.600 | 109715241.560 | — | — | — | 20635648 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 100899024.000 | 102802427.550 | 103158088.410 | 795773.000 | 795773.000 | 795773.000 | 20480000 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 1 | 481673404.500 | 485997097.700 | 488800913.030 | — | — | — | 82231296 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 486499891.000 | 491828688.050 | 492680725.770 | 795775.000 | 795775.000 | 795775.000 | 82096128 |
| before | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 480855126.000 | 484204664.550 | 484927944.930 | 795775.000 | 795775.000 | 795775.000 | 81920000 |
| before | file_store-owned-s131072-a64-short-c64 | normal | 2 | 482332599.000 | 484473261.050 | 485587626.260 | — | — | — | 83513344 |
| before | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 100500199.500 | 105053166.150 | 105629257.210 | 795773.000 | 795773.000 | 795773.000 | 20508672 |
| before | file_store-owned-s64-a16384-short-c64 | normal | 2 | 99754021.500 | 102099144.250 | 106918019.930 | — | — | — | 20639744 |
| before | file_store-owned-s64-a64-short-c64 | allocator | 2 | 4399902.000 | 5320029.600 | 6846761.460 | 795767.000 | 795767.000 | 795767.000 | 6578176 |
| before | file_store-owned-s64-a64-short-c64 | normal | 2 | 4060011.000 | 4799429.550 | 5250444.120 | — | — | — | 6656000 |
| before | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 483276452.500 | 486376551.750 | 487470725.130 | 9184134.000 | 9184134.000 | 9184134.000 | 83410944 |
| before | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 477528819.000 | 480821169.550 | 482666450.190 | — | — | — | 82305024 |
| before | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 94078423.500 | 94692458.550 | 94883251.210 | 9184134.000 | 9184134.000 | 9184134.000 | 20283392 |
| before | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 89934402.500 | 90526666.000 | 90642066.810 | — | — | — | 20717568 |
| before | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 690202.500 | 697571.000 | 697706.900 | 9184134.000 | 9184134.000 | 9184134.000 | 6606848 |
| before | memory_store-owned-s64-a64-short-c64 | normal | 2 | 671797.500 | 686800.000 | 697517.090 | — | — | — | 6823936 |
| before | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 528917753.500 | 535721390.050 | 538075610.720 | 989278.000 | 989278.000 | 989278.000 | 75288576 |
| before | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 525850710.000 | 528702629.800 | 533514160.150 | — | — | — | 75452416 |
| before | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 484137282.500 | 487947862.600 | 488453143.320 | 989278.000 | 989278.000 | 989278.000 | 75526144 |
| before | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 477117870.000 | 479332889.100 | 480972498.680 | — | — | — | 75542528 |
| before | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 486318236.500 | 488369514.900 | 490167006.860 | 989166.000 | 989166.000 | 989166.000 | 75325440 |
| before | deterministic-file-s131072-a64-short-c64 | normal | 2 | 478710066.000 | 481561403.600 | 482187763.620 | — | — | — | 75702272 |
| before | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 485455903.500 | 487944198.800 | 488426209.460 | 989190.000 | 989190.000 | 989190.000 | 75284480 |
| before | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 476215279.500 | 478762576.400 | 479671385.540 | — | — | — | 75722752 |
| before | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 132568405.000 | 133977665.050 | 134071000.600 | 989278.000 | 989278.000 | 989278.000 | 17514496 |
| before | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 129865936.500 | 130954025.600 | 131251187.220 | — | — | — | 17920000 |
| before | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 120100918.500 | 120793710.500 | 121054891.220 | 989278.000 | 989278.000 | 989278.000 | 17797120 |
| before | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 118566883.500 | 119774889.050 | 120344046.520 | — | — | — | 17911808 |
| before | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 445694641.500 | 446873841.000 | 447095360.470 | 989166.000 | 989166.000 | 989166.000 | 17616896 |
| before | deterministic-file-s64-a16384-short-c64 | normal | 2 | 440382764.000 | 442761271.750 | 443372814.410 | — | — | — | 17674240 |
| before | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 120322739.500 | 121497592.050 | 122446968.440 | 989190.000 | 989190.000 | 989190.000 | 17793024 |
| before | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 116774040.500 | 117404156.950 | 117650571.920 | — | — | — | 17838080 |
| before | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12472210.000 | 12558909.000 | 13027701.020 | 989278.000 | 989278.000 | 989278.000 | 6356992 |
| before | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12428125.000 | 12477920.950 | 12505519.010 | — | — | — | 6397952 |
| before | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 815048.000 | 821713.500 | 828370.800 | 989278.000 | 989278.000 | 989278.000 | 6586368 |
| before | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 803448.500 | 810203.550 | 811366.300 | — | — | — | 6377472 |
| before | deterministic-file-s64-a64-short-c64 | allocator | 2 | 2112588.500 | 2125144.000 | 2129660.010 | 989166.000 | 989166.000 | 989166.000 | 6553600 |
| before | deterministic-file-s64-a64-short-c64 | normal | 2 | 2135078.500 | 2233017.500 | 2237387.300 | — | — | — | 6627328 |
| before | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 814158.000 | 825307.000 | 836429.190 | 989190.000 | 989190.000 | 989190.000 | 6639616 |
| before | deterministic-owned-s64-a64-short-c64 | normal | 2 | 896474.000 | 905303.500 | 907682.910 | — | — | — | 6578176 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 1 | 721263.000 | 819449.050 | 832928.100 | — | — | — | 6504448 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 1 | 725703.000 | 733305.500 | 737186.500 | 989190.000 | 989190.000 | 989190.000 | 6914048 |
| after | deterministic-file-s64-a64-short-c64 | normal | 1 | 1278905.500 | 1402273.450 | 1410712.110 | — | — | — | 6561792 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 1 | 1307440.000 | 1426318.000 | 1426949.100 | 989166.000 | 989166.000 | 989166.000 | 6778880 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 1 | 728577.500 | 802380.500 | 805163.710 | — | — | — | 6766592 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 1 | 732808.000 | 741879.000 | 742725.900 | 989278.000 | 989278.000 | 989278.000 | 6582272 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 1 | 12342279.500 | 12627321.000 | 12880698.220 | — | — | — | 6574080 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 1 | 12348654.500 | 14530815.700 | 27667756.360 | 989278.000 | 989278.000 | 989278.000 | 6582272 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 1 | 107907396.000 | 109246478.500 | 109606191.220 | — | — | — | 18169856 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 1 | 110465513.000 | 111793445.050 | 111881659.500 | 989190.000 | 989190.000 | 989190.000 | 18255872 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 1 | 239498533.500 | 241838973.100 | 245204673.970 | — | — | — | 18186240 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 1 | 244364765.000 | 247047899.650 | 247453398.210 | 989166.000 | 989166.000 | 989166.000 | 17874944 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 1 | 108296360.000 | 108856304.050 | 108988705.400 | — | — | — | 17924096 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | 111327926.000 | 112222015.950 | 112577417.420 | 989278.000 | 989278.000 | 989278.000 | 17932288 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 1 | 119803351.000 | 120453381.550 | 120663846.000 | — | — | — | 17973248 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 1 | 123113845.000 | 123756893.050 | 123852276.400 | 989278.000 | 989278.000 | 989278.000 | 17772544 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 1 | 385057984.500 | 387693429.650 | 388094246.200 | — | — | — | 75599872 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 1 | 386279873.500 | 389622678.550 | 390062892.120 | 989190.000 | 989190.000 | 989190.000 | 75620352 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 1 | 389321053.000 | 393579103.000 | 393876255.410 | — | — | — | 75481088 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 1 | 388257879.500 | 389563924.700 | 391592512.280 | 989166.000 | 989166.000 | 989166.000 | 75616256 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 1 | 386893366.500 | 388677023.500 | 390356600.700 | — | — | — | 75849728 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | 385810260.500 | 387394857.650 | 388641136.720 | 989278.000 | 989278.000 | 989278.000 | 75718656 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 1 | 435677608.000 | 437109050.000 | 437656656.940 | — | — | — | 75530240 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 1 | 436008209.000 | 441986211.150 | 448308780.970 | 989278.000 | 989278.000 | 989278.000 | 75669504 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 1 | 548877.000 | 561503.500 | 570987.400 | — | — | — | 7184384 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 1 | 551377.000 | 558447.000 | 560593.600 | 9184134.000 | 9184134.000 | 9184134.000 | 7561216 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 1 | 69909337.000 | 70276346.000 | 71340446.160 | — | — | — | 20766720 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 1 | 71772368.500 | 72170795.000 | 72308342.110 | 9184134.000 | 9184134.000 | 9184134.000 | 20492288 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 1 | 384395685.000 | 385794033.550 | 386884731.650 | — | — | — | 82223104 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 1 | 385085316.000 | 386420441.500 | 386592840.910 | 9184134.000 | 9184134.000 | 9184134.000 | 82096128 |
| after | file_store-owned-s64-a64-short-c64 | normal | 1 | 3626229.000 | 3786753.000 | 3829547.600 | — | — | — | 6557696 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 1 | 3618254.000 | 3726329.500 | 3770802.700 | 795765.000 | 795765.000 | 795765.000 | 6602752 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 1 | 76174076.000 | 77556422.550 | 78957579.170 | — | — | — | 20967424 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 1 | 78790491.000 | 83673559.650 | 88530701.250 | 795771.000 | 795771.000 | 795771.000 | 20635648 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 1 | 391462886.500 | 398213028.200 | 400781633.950 | — | — | — | 82558976 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 1 | 391688189.500 | 397477598.400 | 398413412.600 | 795773.000 | 795773.000 | 795773.000 | 82550784 |
| after | file_store-owned-s131072-a64-short-c64 | allocator | 2 | 391930274.500 | 393920045.050 | 394251735.120 | 795773.000 | 795773.000 | 795773.000 | 83431424 |
| after | file_store-owned-s131072-a64-short-c64 | normal | 2 | 388556261.500 | 392053166.550 | 393385982.760 | — | — | — | 82014208 |
| after | file_store-owned-s64-a16384-short-c64 | allocator | 2 | 77612168.000 | 79017736.100 | 80061358.940 | 795771.000 | 795771.000 | 795771.000 | 20840448 |
| after | file_store-owned-s64-a16384-short-c64 | normal | 2 | 75915846.500 | 77819388.650 | 78580920.160 | — | — | — | 20508672 |
| after | file_store-owned-s64-a64-short-c64 | allocator | 2 | 3742474.000 | 4296339.050 | 4496240.110 | 795765.000 | 795765.000 | 795765.000 | 6533120 |
| after | file_store-owned-s64-a64-short-c64 | normal | 2 | 3594403.000 | 3688547.000 | 3706343.200 | — | — | — | 6524928 |
| after | memory_store-owned-s131072-a64-short-c64 | allocator | 2 | 386125775.000 | 387489546.250 | 387970821.010 | 9184134.000 | 9184134.000 | 9184134.000 | 83566592 |
| after | memory_store-owned-s131072-a64-short-c64 | normal | 2 | 384855430.500 | 386763865.100 | 387508824.930 | — | — | — | 83898368 |
| after | memory_store-owned-s64-a16384-short-c64 | allocator | 2 | 70268880.000 | 70676193.050 | 71116199.420 | 9184134.000 | 9184134.000 | 9184134.000 | 20725760 |
| after | memory_store-owned-s64-a16384-short-c64 | normal | 2 | 69194641.000 | 69527558.000 | 69606574.100 | — | — | — | 20631552 |
| after | memory_store-owned-s64-a64-short-c64 | allocator | 2 | 552487.000 | 564520.000 | 567977.000 | 9184134.000 | 9184134.000 | 9184134.000 | 7036928 |
| after | memory_store-owned-s64-a64-short-c64 | normal | 2 | 548707.000 | 560064.500 | 571814.200 | — | — | — | 7294976 |
| after | deterministic-latency-s131072-a64-short-c64 | allocator | 2 | 435401620.500 | 437394657.850 | 438448285.800 | 989278.000 | 989278.000 | 989278.000 | 75726848 |
| after | deterministic-latency-s131072-a64-short-c64 | normal | 2 | 435496487.500 | 477419822.550 | 487338054.930 | — | — | — | 75509760 |
| after | deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | 385645201.000 | 386167347.050 | 387329235.460 | 989278.000 | 989278.000 | 989278.000 | 75481088 |
| after | deterministic-short-read-s131072-a64-short-c64 | normal | 2 | 383404603.500 | 384372145.500 | 384838563.130 | — | — | — | 75497472 |
| after | deterministic-file-s131072-a64-short-c64 | allocator | 2 | 385863660.000 | 387205892.500 | 387424697.410 | 989166.000 | 989166.000 | 989166.000 | 75685888 |
| after | deterministic-file-s131072-a64-short-c64 | normal | 2 | 388864287.500 | 390484109.200 | 391196937.020 | — | — | — | 75640832 |
| after | deterministic-owned-s131072-a64-short-c64 | allocator | 2 | 386970226.500 | 388334409.000 | 388339223.490 | 989190.000 | 989190.000 | 989190.000 | 75599872 |
| after | deterministic-owned-s131072-a64-short-c64 | normal | 2 | 385006112.000 | 386224594.550 | 386563591.920 | — | — | — | 75939840 |
| after | deterministic-latency-s64-a16384-short-c64 | allocator | 2 | 123451642.000 | 124902697.000 | 125095212.110 | 989278.000 | 989278.000 | 989278.000 | 17801216 |
| after | deterministic-latency-s64-a16384-short-c64 | normal | 2 | 122461182.500 | 135897756.150 | 139943510.820 | — | — | — | 18247680 |
| after | deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | 113142803.000 | 113797677.900 | 114215494.950 | 989278.000 | 989278.000 | 989278.000 | 18239488 |
| after | deterministic-short-read-s64-a16384-short-c64 | normal | 2 | 108198798.500 | 109067398.050 | 109623819.030 | — | — | — | 17911808 |
| after | deterministic-file-s64-a16384-short-c64 | allocator | 2 | 243985883.000 | 245468968.700 | 246355219.430 | 989166.000 | 989166.000 | 989166.000 | 17911808 |
| after | deterministic-file-s64-a16384-short-c64 | normal | 2 | 238864723.000 | 240209626.100 | 240539255.110 | — | — | — | 17911808 |
| after | deterministic-owned-s64-a16384-short-c64 | allocator | 2 | 110004410.000 | 111325549.000 | 111425492.410 | 989190.000 | 989190.000 | 989190.000 | 18178048 |
| after | deterministic-owned-s64-a16384-short-c64 | normal | 2 | 106848613.500 | 108084790.000 | 108218826.110 | — | — | — | 17981440 |
| after | deterministic-latency-s64-a64-short-c64 | allocator | 2 | 12333442.500 | 12778183.550 | 13387186.130 | 989278.000 | 989278.000 | 989278.000 | 6975488 |
| after | deterministic-latency-s64-a64-short-c64 | normal | 2 | 12331102.000 | 12404074.500 | 12411216.600 | — | — | — | 6705152 |
| after | deterministic-short-read-s64-a64-short-c64 | allocator | 2 | 734403.000 | 739968.500 | 741108.690 | 989278.000 | 989278.000 | 989278.000 | 6733824 |
| after | deterministic-short-read-s64-a64-short-c64 | normal | 2 | 720068.000 | 729451.450 | 731277.610 | — | — | — | 7036928 |
| after | deterministic-file-s64-a64-short-c64 | allocator | 2 | 1403415.500 | 1422481.500 | 1424313.490 | 989166.000 | 989166.000 | 989166.000 | 6463488 |
| after | deterministic-file-s64-a64-short-c64 | normal | 2 | 1389325.000 | 1405902.000 | 1416180.690 | — | — | — | 6508544 |
| after | deterministic-owned-s64-a64-short-c64 | allocator | 2 | 731762.500 | 736585.000 | 742053.700 | 989190.000 | 989190.000 | 989190.000 | 7036928 |
| after | deterministic-owned-s64-a64-short-c64 | normal | 2 | 714053.000 | 724767.000 | 726881.400 | — | — | — | 6725632 |

## Matched before/after changes

Each row compares one arm, role, repeat, and metric. Percentages use `(after - before) / before`; lower values are favorable for elapsed time and operation heap.

| arm | role | repeat | metric | p50 change | p50 flag | p95 change | p95 flag | p99 change | p99 flag |
|---|---|---:|---|---:|---|---:|---|---:|---|
| deterministic-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -10.670% | True | 0.663% | False | 2.129% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -4.164% | False | -4.164% | False | -4.164% | False |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -20.349% | True | -19.942% | True | -19.919% | True |
| deterministic-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 2.242% | False | 2.242% | False | 2.242% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -13.875% | True | -17.735% | True | -17.364% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 5.237% | True | 5.237% | True | 5.237% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -10.120% | True | -10.750% | True | -11.283% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 5.984% | True | 5.984% | True | 5.984% | True |
| deterministic-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a64-short-c64 | normal | 1 | elapsed_ns | -41.906% | True | -36.962% | True | -36.612% | True |
| deterministic-file-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -4.869% | False | -4.869% | False | -4.869% | False |
| deterministic-file-s64-a64-short-c64 | normal | 2 | elapsed_ns | -34.929% | True | -37.040% | True | -36.704% | True |
| deterministic-file-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -1.792% | False | -1.792% | False | -1.792% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -37.937% | True | -32.736% | True | -32.879% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 3.632% | False | 3.632% | False | 3.632% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -33.569% | True | -33.064% | True | -33.120% | True |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -1.375% | False | -1.375% | False | -1.375% | False |
| deterministic-file-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | elapsed_ns | -9.885% | True | -1.625% | False | -1.405% | False |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 2.481% | False | 2.481% | False | 2.481% | False |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | elapsed_ns | -10.378% | True | -9.967% | True | -9.871% | True |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 10.340% | True | 10.340% | True | 10.340% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -9.607% | True | -9.265% | True | -9.205% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.438% | False | 0.438% | False | 0.438% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -9.895% | True | -9.948% | True | -10.534% | True |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 2.239% | False | 2.239% | False | 2.239% | False |
| deterministic-short-read-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | elapsed_ns | -0.755% | False | -1.130% | False | -0.908% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | -2.015% | False | -2.015% | False | -2.015% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | elapsed_ns | -0.781% | False | -0.592% | False | -0.754% | False |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 4.802% | False | 4.802% | False | 4.802% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -1.009% | False | 12.249% | True | 8.438% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | -0.124% | False | -0.124% | False | -0.124% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -1.113% | False | 1.746% | False | 2.759% | False |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 9.729% | True | 9.729% | True | 9.729% | True |
| deterministic-latency-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -8.392% | True | -8.543% | True | -8.576% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 2.141% | False | 2.141% | False | 2.141% | False |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -8.500% | True | -7.938% | True | -8.017% | True |
| deterministic-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 0.804% | False | 0.804% | False | 0.804% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -7.737% | True | -7.688% | True | -7.792% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 2.601% | False | 2.601% | False | 2.601% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -8.576% | True | -8.372% | True | -9.001% | True |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 2.164% | False | 2.164% | False | 2.164% | False |
| deterministic-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -46.049% | True | -45.699% | True | -45.140% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 2.210% | False | 2.210% | False | 2.210% | False |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -45.760% | True | -45.747% | True | -45.748% | True |
| deterministic-file-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.344% | False | 1.344% | False | 1.344% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -44.372% | True | -43.929% | True | -43.909% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 2.010% | False | 2.010% | False | 2.010% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -45.257% | True | -45.070% | True | -44.899% | True |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 1.674% | False | 1.674% | False | 1.674% | False |
| deterministic-file-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -8.931% | True | -8.800% | True | -8.922% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 0.436% | False | 0.436% | False | 0.436% | False |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -8.745% | True | -8.940% | True | -8.908% | True |
| deterministic-short-read-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -7.573% | True | -7.298% | True | -7.051% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 0.783% | False | 0.783% | False | 0.783% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -5.794% | True | -5.792% | True | -5.650% | True |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 2.486% | False | 2.486% | False | 2.486% | False |
| deterministic-short-read-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -8.599% | True | -9.445% | True | -10.621% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 0.711% | False | 0.711% | False | 0.711% | False |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -5.702% | True | 3.775% | False | 6.623% | True |
| deterministic-latency-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | 1.829% | False | 1.829% | False | 1.829% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -7.463% | True | -8.773% | True | -9.123% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | -0.115% | False | -0.115% | False | -0.115% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -6.877% | True | -6.773% | True | -6.695% | True |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 1.637% | False | 1.637% | False | 1.637% | False |
| deterministic-latency-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -20.166% | True | -20.283% | True | -20.263% | True |
| deterministic-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.441% | False | 0.441% | False | 0.441% | False |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -19.153% | True | -19.329% | True | -19.411% | True |
| deterministic-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.287% | False | 0.287% | False | 0.287% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -19.448% | True | -19.887% | True | -19.920% | True |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.125% | False | 0.125% | False | 0.125% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -20.287% | True | -20.414% | True | -20.492% | True |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.419% | False | 0.419% | False | 0.419% | False |
| deterministic-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -18.860% | True | -18.509% | True | -18.531% | True |
| deterministic-file-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.038% | False | 0.038% | False | 0.038% | False |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -18.768% | True | -18.913% | True | -18.870% | True |
| deterministic-file-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.081% | False | -0.081% | False | -0.081% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -20.842% | True | -20.891% | True | -20.499% | True |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.479% | False | 0.479% | False | 0.479% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -20.656% | True | -20.715% | True | -20.961% | True |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.479% | False | 0.479% | False | 0.479% | False |
| deterministic-file-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -19.439% | True | -19.431% | True | -19.133% | True |
| deterministic-short-read-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.254% | False | 0.254% | False | 0.254% | False |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -19.642% | True | -19.811% | True | -19.987% | True |
| deterministic-short-read-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -0.060% | False | -0.060% | False | -0.060% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -19.999% | True | -20.034% | True | -19.827% | True |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.260% | False | 0.260% | False | 0.260% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -20.344% | True | -20.859% | True | -20.703% | True |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.060% | False | -0.060% | False | -0.060% | False |
| deterministic-short-read-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -17.521% | True | -17.688% | True | -17.599% | True |
| deterministic-latency-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | -0.038% | False | -0.038% | False | -0.038% | False |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -17.182% | True | -9.700% | True | -8.655% | True |
| deterministic-latency-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 0.076% | False | 0.076% | False | 0.076% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -17.495% | True | -17.009% | True | -16.624% | True |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.643% | False | 0.643% | False | 0.643% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -17.681% | True | -18.354% | True | -18.515% | True |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.582% | False | 0.582% | False | 0.582% | False |
| deterministic-latency-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -18.527% | True | -17.680% | True | -16.361% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 8.005% | True | 8.005% | True | 8.005% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -18.323% | True | -18.453% | True | -18.021% | True |
| memory_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | 6.903% | True | 6.903% | True | 6.903% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -19.494% | True | -19.529% | True | -19.836% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 11.071% | True | 11.071% | True | 11.071% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -19.953% | True | -19.073% | True | -18.594% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 6.510% | True | 6.510% | True | 6.510% | True |
| memory_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -22.795% | True | -22.997% | True | -21.898% | True |
| memory_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 1.056% | False | 1.056% | False | 1.056% | False |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -23.061% | True | -23.197% | True | -23.207% | True |
| memory_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -0.415% | False | -0.415% | False | -0.415% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -23.808% | True | -23.631% | True | -23.720% | True |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 1.050% | False | 1.050% | False | 1.050% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -25.308% | True | -25.362% | True | -25.049% | True |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 2.181% | False | 2.181% | False | 2.181% | False |
| memory_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -20.209% | True | -20.158% | True | -20.094% | True |
| memory_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.586% | False | 0.586% | False | 0.586% | False |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -19.407% | True | -19.562% | True | -19.715% | True |
| memory_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | 1.936% | False | 1.936% | False | 1.936% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -19.644% | True | -20.119% | True | -20.124% | True |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.871% | False | 0.871% | False | 0.871% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -20.103% | True | -20.331% | True | -20.411% | True |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 0.187% | False | 0.187% | False | 0.187% | False |
| memory_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | 0.000% | False | 0.000% | False | 0.000% | False |
| file_store-owned-s64-a64-short-c64 | normal | 1 | elapsed_ns | -23.612% | True | -34.761% | True | -45.110% | True |
| file_store-owned-s64-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.062% | False | 0.062% | False | 0.062% | False |
| file_store-owned-s64-a64-short-c64 | normal | 2 | elapsed_ns | -11.468% | True | -23.146% | True | -29.409% | True |
| file_store-owned-s64-a64-short-c64 | normal | 2 | time_max_rss_bytes | -1.969% | False | -1.969% | False | -1.969% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | elapsed_ns | -20.569% | True | -43.892% | True | -47.494% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 7.467% | True | 7.467% | True | 7.467% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | elapsed_ns | -14.942% | True | -19.242% | True | -34.330% | True |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | time_max_rss_bytes | -0.685% | False | -0.685% | False | -0.685% | False |
| file_store-owned-s64-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | elapsed_ns | -23.287% | True | -26.388% | True | -28.034% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 1 | time_max_rss_bytes | 1.608% | False | 1.608% | False | 1.608% | False |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | elapsed_ns | -23.897% | True | -23.781% | True | -26.504% | True |
| file_store-owned-s64-a16384-short-c64 | normal | 2 | time_max_rss_bytes | -0.635% | False | -0.635% | False | -0.635% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | elapsed_ns | -21.912% | True | -18.607% | True | -14.180% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | time_max_rss_bytes | 0.760% | False | 0.760% | False | 0.760% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | elapsed_ns | -22.774% | True | -24.783% | True | -24.205% | True |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | time_max_rss_bytes | 1.618% | False | 1.618% | False | 1.618% | False |
| file_store-owned-s64-a16384-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | elapsed_ns | -18.729% | True | -18.063% | True | -18.007% | True |
| file_store-owned-s131072-a64-short-c64 | normal | 1 | time_max_rss_bytes | 0.398% | False | 0.398% | False | 0.398% | False |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | elapsed_ns | -19.442% | True | -19.076% | True | -18.988% | True |
| file_store-owned-s131072-a64-short-c64 | normal | 2 | time_max_rss_bytes | -1.795% | False | -1.795% | False | -1.795% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | elapsed_ns | -19.489% | True | -19.184% | True | -19.134% | True |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | time_max_rss_bytes | 0.554% | False | 0.554% | False | 0.554% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 1 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | elapsed_ns | -18.493% | True | -18.646% | True | -18.699% | True |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | time_max_rss_bytes | 1.845% | False | 1.845% | False | 1.845% | False |
| file_store-owned-s131072-a64-short-c64 | allocator | 2 | allocator_operation_peak_increment_bytes | -0.000% | False | -0.000% | False | -0.000% | False |

## Measurement limits

- The matrix covers three exact workloads and selected input/store arms; it is not a full provider-by-input interaction grid.
- Two process repeats support a descriptive before/after range; they do not establish a strong confidence interval.
- A five-percent flag is a review trigger, not a causal speedup claim.
- GNU time observes the whole child once and is retained with n=1; allocator operation heap is a separate instrumented metric.
