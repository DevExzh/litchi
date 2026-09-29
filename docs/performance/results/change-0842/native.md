# 0842 complete native results

Ratios are medians of six matched block ratios. Intervals are paired bootstrap endpoints; RSS is whole-child KiB. No zero-denominator ratio is invented.

| Shape | Mode | Metric | Before | After | Ratio | 95% interval |
| --- | --- | --- | ---: | ---: | ---: | --- |
| tiny | write | p50_ns | 36260.500 | 36435.000 | 1.003019 | [0.997542, 1.013260] |
| tiny | write | p95_ns | 37805.500 | 37615.000 | 1.002150 | [0.947285, 1.006558] |
| tiny | write | p99_ns | 41715.000 | 56915.000 | 1.293725 | [1.024374, 1.716561] |
| tiny | write | mean_ns | 36674.267 | 37147.483 | 1.018197 | [1.005967, 1.026481] |
| tiny | write | rss_kib | 4420.000 | 4348.000 | 1.006728 | [0.932515, 1.043612] |
| tiny | write | minor_faults | 348.500 | 351.000 | 1.010082 | [1.005755, 1.184814] |
| tiny | write | major_faults | 0.000 | 0.000 | undefined | undefined |
| tiny | read | p50_ns | 7290.000 | 7305.000 | 1.002742 | [0.999390, 1.007614] |
| tiny | read | p95_ns | 7495.000 | 7520.000 | 1.004025 | [0.958029, 1.063622] |
| tiny | read | p99_ns | 8155.000 | 14370.000 | 1.394552 | [0.940478, 3.385544] |
| tiny | read | mean_ns | 7306.683 | 7543.200 | 1.023167 | [1.001516, 1.092779] |
| tiny | read | rss_kib | 4180.000 | 4104.000 | 0.984211 | [0.961134, 1.026806] |
| tiny | read | minor_faults | 332.500 | 335.000 | 1.007530 | [1.003003, 1.012048] |
| tiny | read | major_faults | 0.000 | 0.000 | undefined | undefined |
| tiny | changed | p50_ns | 29050.000 | 29140.000 | 0.998614 | [0.990360, 1.013564] |
| tiny | changed | p95_ns | 40685.000 | 40235.000 | 1.016626 | [0.869106, 1.207760] |
| tiny | changed | p99_ns | 43130.000 | 43550.000 | 1.005438 | [0.946214, 1.173194] |
| tiny | changed | mean_ns | 30064.983 | 30084.117 | 1.000269 | [0.989143, 1.019945] |
| tiny | changed | rss_kib | 4338.000 | 4332.000 | 0.981222 | [0.946728, 1.048284] |
| tiny | changed | minor_faults | 356.500 | 359.500 | 1.009816 | [1.006995, 1.033455] |
| tiny | changed | major_faults | 0.000 | 0.000 | undefined | undefined |
| xml | write | p50_ns | 16279178.000 | 15970156.500 | 0.981007 | [0.974359, 0.983772] |
| xml | write | p95_ns | 16515563.500 | 16008606.500 | 0.969651 | [0.957473, 0.979913] |
| xml | write | p99_ns | 16547774.000 | 16020456.500 | 0.969771 | [0.944921, 0.980384] |
| xml | write | mean_ns | 16310683.300 | 15963302.683 | 0.978016 | [0.971837, 0.980704] |
| xml | write | rss_kib | 23484.000 | 23596.000 | 1.004761 | [0.997635, 1.011112] |
| xml | write | minor_faults | 53268.500 | 53270.000 | 1.000038 | [0.999972, 1.000094] |
| xml | write | major_faults | 0.000 | 0.000 | undefined | undefined |
| xml | read | p50_ns | 1306695.500 | 1042560.000 | 0.798184 | [0.794030, 0.799516] |
| xml | read | p95_ns | 1319685.500 | 1048639.500 | 0.794332 | [0.790093, 0.798151] |
| xml | read | p99_ns | 1332011.000 | 1049269.000 | 0.790205 | [0.768752, 0.795710] |
| xml | read | mean_ns | 1308073.500 | 1043388.333 | 0.796983 | [0.794629, 0.798809] |
| xml | read | rss_kib | 22456.000 | 22494.000 | 1.002213 | [0.996208, 1.005613] |
| xml | read | minor_faults | 5172.000 | 5174.500 | 1.000580 | [1.000000, 1.000677] |
| xml | read | major_faults | 0.000 | 0.000 | undefined | undefined |
| xml | changed | p50_ns | 480142.000 | 349846.500 | 0.661920 | [0.620751, 0.748841] |
| xml | changed | p95_ns | 492297.000 | 361956.000 | 0.677703 | [0.629079, 0.760582] |
| xml | changed | p99_ns | 514943.000 | 362581.500 | 0.711062 | [0.642220, 0.768077] |
| xml | changed | mean_ns | 480748.650 | 332979.450 | 0.651419 | [0.624166, 0.713583] |
| xml | changed | rss_kib | 23294.000 | 23496.000 | 1.008339 | [0.994837, 1.012592] |
| xml | changed | minor_faults | 24967.500 | 39829.500 | 1.021620 | [0.961708, 1.623205] |
| xml | changed | major_faults | 0.000 | 0.000 | undefined | undefined |
| random | write | p50_ns | 26284749.000 | 25756556.500 | 0.979640 | [0.977702, 0.981053] |
| random | write | p95_ns | 26400993.500 | 25855922.500 | 0.978867 | [0.977605, 0.980273] |
| random | write | p99_ns | 26430259.500 | 25905957.500 | 0.978640 | [0.973242, 0.980429] |
| random | write | mean_ns | 26274641.333 | 25753570.883 | 0.980182 | [0.978144, 0.980968] |
| random | write | rss_kib | 25742.000 | 25700.000 | 0.997515 | [0.993516, 1.002499] |
| random | write | minor_faults | 50853.500 | 50855.500 | 1.000039 | [0.999951, 1.011435] |
| random | write | major_faults | 0.000 | 0.000 | undefined | undefined |
| random | read | p50_ns | 1519176.500 | 1299930.500 | 0.855358 | [0.852033, 0.863678] |
| random | read | p95_ns | 1535186.000 | 1310656.000 | 0.854982 | [0.852116, 0.858012] |
| random | read | p99_ns | 1544661.000 | 1318645.500 | 0.854437 | [0.842858, 0.863020] |
| random | read | mean_ns | 1519953.350 | 1301541.283 | 0.855958 | [0.852315, 0.864105] |
| random | read | rss_kib | 24602.000 | 24700.000 | 1.003167 | [1.001707, 1.007162] |
| random | read | minor_faults | 41136.500 | 41139.000 | 1.000061 | [1.000036, 1.000085] |
| random | read | major_faults | 0.000 | 0.000 | undefined | undefined |
| random | changed | p50_ns | 14243754.000 | 13876592.500 | 0.974547 | [0.971233, 0.977969] |
| random | changed | p95_ns | 14286574.000 | 13916778.000 | 0.973567 | [0.970927, 0.977898] |
| random | changed | p99_ns | 14325290.000 | 13946028.000 | 0.972443 | [0.960448, 0.975953] |
| random | changed | mean_ns | 14245768.367 | 13879346.683 | 0.974251 | [0.971351, 0.977193] |
| random | changed | rss_kib | 26730.000 | 26856.000 | 1.004348 | [0.993298, 1.011335] |
| random | changed | minor_faults | 56401.000 | 56986.000 | 1.010480 | [0.979398, 1.031933] |
| random | changed | major_faults | 0.000 | 0.000 | undefined | undefined |
| repeated | write | p50_ns | 12328621.500 | 12030170.000 | 0.974406 | [0.971603, 0.979304] |
| repeated | write | p95_ns | 12399181.000 | 12097955.000 | 0.974360 | [0.970320, 0.975868] |
| repeated | write | p99_ns | 12443941.000 | 12117796.000 | 0.974203 | [0.950365, 0.979824] |
| repeated | write | mean_ns | 12346964.683 | 12036307.233 | 0.974310 | [0.971832, 0.977531] |
| repeated | write | rss_kib | 23360.000 | 23380.000 | 1.003605 | [0.995351, 1.005745] |
| repeated | write | minor_faults | 53226.500 | 53228.000 | 1.000038 | [0.999991, 1.000122] |
| repeated | write | major_faults | 0.000 | 0.000 | undefined | undefined |
| repeated | read | p50_ns | 1053849.000 | 828569.000 | 0.786795 | [0.782707, 0.790379] |
| repeated | read | p95_ns | 1062849.500 | 833038.000 | 0.783775 | [0.777312, 0.788172] |
| repeated | read | p99_ns | 1065814.000 | 835758.500 | 0.786781 | [0.783172, 0.813698] |
| repeated | read | mean_ns | 1054495.683 | 828615.450 | 0.786885 | [0.781968, 0.789643] |
| repeated | read | rss_kib | 22308.000 | 22258.000 | 1.003339 | [0.985336, 1.009765] |
| repeated | read | minor_faults | 5130.000 | 5132.500 | 1.000585 | [1.000195, 1.000780] |
| repeated | read | major_faults | 0.000 | 0.000 | undefined | undefined |
| repeated | changed | p50_ns | 513132.000 | 318976.000 | 0.616442 | [0.557185, 0.731619] |
| repeated | changed | p95_ns | 525512.000 | 330486.500 | 0.628360 | [0.567133, 0.720652] |
| repeated | changed | p99_ns | 537997.500 | 338456.500 | 0.627029 | [0.587565, 0.702193] |
| repeated | changed | mean_ns | 503503.883 | 314210.467 | 0.620741 | [0.569631, 0.708693] |
| repeated | changed | rss_kib | 23238.000 | 23122.000 | 0.994661 | [0.984876, 1.018777] |
| repeated | changed | minor_faults | 44078.500 | 33946.500 | 0.777942 | [0.539144, 1.757748] |
| repeated | changed | major_faults | 0.000 | 0.000 | undefined | undefined |

p50/RSS review flags: 0. Increasing p95/p99 point estimates above 5%: 2. Block-p50 spread flags: 0.

Complete block vectors and flag rows are retained in analysis.json.
