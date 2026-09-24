# XLSX ordinary worksheet SVG lifecycle uncertainty summary

This is a descriptive summary of retained raw receipts. It reports each fresh process median, each process's sample min/max range, and the range of process medians. It does not estimate a speedup, regression, causal effect, confidence interval, or scaling law.

- Processes per lane: n=3
- Minimum warmups per process: 2
- Minimum measured samples per process: 20
- Raw receipts are unchanged; this file is a derived view.

## Limitations

- The process count is n=3; between-process ranges are descriptive and are not confidence intervals.
- Within-process ranges describe the retained samples and do not model scheduler, cache, or host variation.
- This summary does not estimate a speedup, regression, causal effect, or algorithmic scaling law.

## `capture_native_fixture`

- Input bytes: 12470; SHA-256: `0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=544652.5; p2=545917; p3=546402.5; within-process sample ranges p1=[526612,573712]; p2=[529182,568482]; p3=[532492,581363]; between-process median range [544652.5,546402.5].
- `requested_alloc_bytes`: per-process medians p1=2603541; p2=2603541; p3=2603541; within-process sample ranges p1=[2603541,2603541]; p2=[2603541,2603541]; p3=[2603541,2603541]; between-process median range [2603541,2603541].
- `peak_live_delta`: per-process medians p1=207612; p2=207612; p3=207612; within-process sample ranges p1=[207612,207612]; p2=[207612,207612]; p3=[207612,207612]; between-process median range [207612,207612].
- `rss_kib`: per-process values p1=6424; p2=6444; p3=6468; range [6424,6468].

## `capture_raster_two_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=288396.5; p2=290411.5; p3=289211; within-process sample ranges p1=[285452,325461]; p2=[285782,301751]; p3=[285291,306632]; between-process median range [288396.5,290411.5].
- `requested_alloc_bytes`: per-process medians p1=2270409; p2=2270409; p3=2270409; within-process sample ranges p1=[2270409,2270409]; p2=[2270409,2270409]; p3=[2270409,2270409]; between-process median range [2270409,2270409].
- `peak_live_delta`: per-process medians p1=139014; p2=139014; p3=139014; within-process sample ranges p1=[139014,139014]; p2=[139014,139014]; p3=[139014,139014]; between-process median range [139014,139014].
- `rss_kib`: per-process values p1=6364; p2=6472; p3=6596; range [6364,6596].

## `capture_raster_two_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=318771; p2=320506.5; p3=315621.5; within-process sample ranges p1=[311312,344151]; p2=[310602,343732]; p3=[310951,326551]; between-process median range [315621.5,320506.5].
- `requested_alloc_bytes`: per-process medians p1=2529026; p2=2529026; p3=2529026; within-process sample ranges p1=[2529026,2529026]; p2=[2529026,2529026]; p3=[2529026,2529026]; between-process median range [2529026,2529026].
- `peak_live_delta`: per-process medians p1=268607; p2=268607; p3=268607; within-process sample ranges p1=[268607,268607]; p2=[268607,268607]; p3=[268607,268607]; between-process median range [268607,268607].
- `rss_kib`: per-process values p1=6908; p2=6616; p3=6756; range [6616,6908].

## `capture_raster_one_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=293481.5; p2=290531.5; p3=290751; within-process sample ranges p1=[284471,325331]; p2=[283341,302322]; p3=[285221,301842]; between-process median range [290531.5,293481.5].
- `requested_alloc_bytes`: per-process medians p1=2270409; p2=2270409; p3=2270409; within-process sample ranges p1=[2270409,2270409]; p2=[2270409,2270409]; p3=[2270409,2270409]; between-process median range [2270409,2270409].
- `peak_live_delta`: per-process medians p1=139014; p2=139014; p3=139014; within-process sample ranges p1=[139014,139014]; p2=[139014,139014]; p3=[139014,139014]; between-process median range [139014,139014].
- `rss_kib`: per-process values p1=6452; p2=6308; p3=6476; range [6308,6476].

## `capture_raster_one_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=317771.5; p2=319301.5; p3=318766; within-process sample ranges p1=[311632,340102]; p2=[311252,344522]; p3=[312961,343261]; between-process median range [317771.5,319301.5].
- `requested_alloc_bytes`: per-process medians p1=2529026; p2=2529026; p3=2529026; within-process sample ranges p1=[2529026,2529026]; p2=[2529026,2529026]; p3=[2529026,2529026]; between-process median range [2529026,2529026].
- `peak_live_delta`: per-process medians p1=268607; p2=268607; p3=268607; within-process sample ranges p1=[268607,268607]; p2=[268607,268607]; p3=[268607,268607]; between-process median range [268607,268607].
- `rss_kib`: per-process values p1=6788; p2=6764; p3=6676; range [6676,6788].

## `capture_raster_absolute_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=297681; p2=289561; p3=289871; within-process sample ranges p1=[284791,314641]; p2=[284641,301251]; p3=[282911,311911]; between-process median range [289561,297681].
- `requested_alloc_bytes`: per-process medians p1=2270409; p2=2270409; p3=2270409; within-process sample ranges p1=[2270409,2270409]; p2=[2270409,2270409]; p3=[2270409,2270409]; between-process median range [2270409,2270409].
- `peak_live_delta`: per-process medians p1=139014; p2=139014; p3=139014; within-process sample ranges p1=[139014,139014]; p2=[139014,139014]; p3=[139014,139014]; between-process median range [139014,139014].
- `rss_kib`: per-process values p1=6372; p2=6284; p3=6128; range [6128,6372].

## `capture_raster_absolute_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=319326; p2=319197; p3=318401.5; within-process sample ranges p1=[313002,338711]; p2=[313512,340012]; p3=[313521,338981]; between-process median range [318401.5,319326].
- `requested_alloc_bytes`: per-process medians p1=2529026; p2=2529026; p3=2529026; within-process sample ranges p1=[2529026,2529026]; p2=[2529026,2529026]; p3=[2529026,2529026]; between-process median range [2529026,2529026].
- `peak_live_delta`: per-process medians p1=268607; p2=268607; p3=268607; within-process sample ranges p1=[268607,268607]; p2=[268607,268607]; p3=[268607,268607]; between-process median range [268607,268607].
- `rss_kib`: per-process values p1=6600; p2=6688; p3=6580; range [6580,6688].

## `capture_attached_two_cell_small`

- Input bytes: 4419; SHA-256: `4ea46c1e4faba144528af4feae91d35e416a9120863966253173ff7f41a24578`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=258621; p2=259756.5; p3=259636; within-process sample ranges p1=[254701,277461]; p2=[255351,282711]; p3=[255091,277331]; between-process median range [258621,259756.5].
- `requested_alloc_bytes`: per-process medians p1=1750027; p2=1750027; p3=1750027; within-process sample ranges p1=[1750027,1750027]; p2=[1750027,1750027]; p3=[1750027,1750027]; between-process median range [1750027,1750027].
- `peak_live_delta`: per-process medians p1=143958; p2=143958; p3=143958; within-process sample ranges p1=[143958,143958]; p2=[143958,143958]; p3=[143958,143958]; between-process median range [143958,143958].
- `rss_kib`: per-process values p1=6288; p2=6200; p3=6316; range [6200,6316].

## `capture_attached_two_cell_large`

- Input bytes: 9976; SHA-256: `73f6b8097d319c55f2bd9ce9250b3e96150459596fb0d2a13c88d6ebfcc861c3`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=330102; p2=330746; p3=331136.5; within-process sample ranges p1=[324902,348111]; p2=[326312,354621]; p3=[324311,347261]; between-process median range [330102,331136.5].
- `requested_alloc_bytes`: per-process medians p1=2144192; p2=2144192; p3=2144192; within-process sample ranges p1=[2144192,2144192]; p2=[2144192,2144192]; p3=[2144192,2144192]; between-process median range [2144192,2144192].
- `peak_live_delta`: per-process medians p1=408587; p2=408587; p3=408587; within-process sample ranges p1=[408587,408587]; p2=[408587,408587]; p3=[408587,408587]; between-process median range [408587,408587].
- `rss_kib`: per-process values p1=6536; p2=6868; p3=6756; range [6536,6868].

## `capture_attached_one_cell_small`

- Input bytes: 4419; SHA-256: `c4c7e2149c28c765d15a9aa10bec811a7247f1107bf015b59a65e4d9593256fa`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=260296; p2=259391.5; p3=258946; within-process sample ranges p1=[255351,277861]; p2=[255182,287231]; p3=[254171,277951]; between-process median range [258946,260296].
- `requested_alloc_bytes`: per-process medians p1=1750027; p2=1750027; p3=1750027; within-process sample ranges p1=[1750027,1750027]; p2=[1750027,1750027]; p3=[1750027,1750027]; between-process median range [1750027,1750027].
- `peak_live_delta`: per-process medians p1=143958; p2=143958; p3=143958; within-process sample ranges p1=[143958,143958]; p2=[143958,143958]; p3=[143958,143958]; between-process median range [143958,143958].
- `rss_kib`: per-process values p1=6432; p2=6580; p3=6564; range [6432,6580].

## `capture_attached_one_cell_large`

- Input bytes: 9976; SHA-256: `5a68caf8869648f50e91107f800ff2d0f4028a0508014bc6adb0bd75a2451ab6`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=329861; p2=330646; p3=331876; within-process sample ranges p1=[325011,352271]; p2=[324611,343892]; p3=[324901,351222]; between-process median range [329861,331876].
- `requested_alloc_bytes`: per-process medians p1=2144192; p2=2144192; p3=2144192; within-process sample ranges p1=[2144192,2144192]; p2=[2144192,2144192]; p3=[2144192,2144192]; between-process median range [2144192,2144192].
- `peak_live_delta`: per-process medians p1=408587; p2=408587; p3=408587; within-process sample ranges p1=[408587,408587]; p2=[408587,408587]; p3=[408587,408587]; between-process median range [408587,408587].
- `rss_kib`: per-process values p1=6764; p2=6724; p3=6644; range [6644,6764].

## `capture_attached_absolute_small`

- Input bytes: 4419; SHA-256: `6e0acd00c737eafa905c2edb9448d7db5c6a6bfc09c892909b0ba036f5526e29`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=260431; p2=258496; p3=257501; within-process sample ranges p1=[254121,274421]; p2=[255231,282011]; p3=[253541,272312]; between-process median range [257501,260431].
- `requested_alloc_bytes`: per-process medians p1=1750027; p2=1750027; p3=1750027; within-process sample ranges p1=[1750027,1750027]; p2=[1750027,1750027]; p3=[1750027,1750027]; between-process median range [1750027,1750027].
- `peak_live_delta`: per-process medians p1=143958; p2=143958; p3=143958; within-process sample ranges p1=[143958,143958]; p2=[143958,143958]; p3=[143958,143958]; between-process median range [143958,143958].
- `rss_kib`: per-process values p1=6192; p2=6468; p3=6608; range [6192,6608].

## `capture_attached_absolute_large`

- Input bytes: 9967; SHA-256: `db5f2ff1e1043090ebacbc351703e16a7e84d00b353d46958a795db078419fcc`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=330537; p2=331227; p3=331496; within-process sample ranges p1=[324041,343791]; p2=[326281,349132]; p3=[326591,359122]; between-process median range [330537,331496].
- `requested_alloc_bytes`: per-process medians p1=2144183; p2=2144183; p3=2144183; within-process sample ranges p1=[2144183,2144183]; p2=[2144183,2144183]; p3=[2144183,2144183]; between-process median range [2144183,2144183].
- `peak_live_delta`: per-process medians p1=408578; p2=408578; p3=408578; within-process sample ranges p1=[408578,408578]; p2=[408578,408578]; p3=[408578,408578]; between-process median range [408578,408578].
- `rss_kib`: per-process values p1=6756; p2=6664; p3=6828; range [6664,6828].

## `clone_raster_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=290346.5; p2=292736.5; p3=290011; within-process sample ranges p1=[285792,301692]; p2=[286111,302311]; p3=[283751,304542]; between-process median range [290011,292736.5].
- `requested_alloc_bytes`: per-process medians p1=2274463; p2=2274463; p3=2274463; within-process sample ranges p1=[2274463,2274463]; p2=[2274463,2274463]; p3=[2274463,2274463]; between-process median range [2274463,2274463].
- `peak_live_delta`: per-process medians p1=142900; p2=142900; p3=142900; within-process sample ranges p1=[142900,142900]; p2=[142900,142900]; p3=[142900,142900]; between-process median range [142900,142900].
- `rss_kib`: per-process values p1=6292; p2=6132; p3=6452; range [6132,6452].

## `clone_raster_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=318631.5; p2=318682; p3=318331.5; within-process sample ranges p1=[314391,341472]; p2=[311892,340292]; p3=[312222,328742]; between-process median range [318331.5,318682].
- `requested_alloc_bytes`: per-process medians p1=2533649; p2=2533649; p3=2533649; within-process sample ranges p1=[2533649,2533649]; p2=[2533649,2533649]; p3=[2533649,2533649]; between-process median range [2533649,2533649].
- `peak_live_delta`: per-process medians p1=273062; p2=273062; p3=273062; within-process sample ranges p1=[273062,273062]; p2=[273062,273062]; p3=[273062,273062]; between-process median range [273062,273062].
- `rss_kib`: per-process values p1=6568; p2=6652; p3=6468; range [6468,6652].

## `clone_attached_small`

- Input bytes: 4415; SHA-256: `20bea6dd00843bed2b03ebe521a2b72641f7a6196ccaecb1559e7f25153f07ea`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=261111.5; p2=260251; p3=259131; within-process sample ranges p1=[256111,282012]; p2=[255461,272251]; p3=[255421,276992]; between-process median range [259131,261111.5].
- `requested_alloc_bytes`: per-process medians p1=1754606; p2=1754606; p3=1754606; within-process sample ranges p1=[1754606,1754606]; p2=[1754606,1754606]; p3=[1754606,1754606]; between-process median range [1754606,1754606].
- `peak_live_delta`: per-process medians p1=148369; p2=148369; p3=148369; within-process sample ranges p1=[148369,148369]; p2=[148369,148369]; p3=[148369,148369]; between-process median range [148369,148369].
- `rss_kib`: per-process values p1=6288; p2=6444; p3=6524; range [6288,6524].

## `clone_attached_large`

- Input bytes: 6043; SHA-256: `b07e26a3061a6d99bf2433651f87bc9eadcc8665aff06319e1c12e2114d6dcb1`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=300671.5; p2=298911.5; p3=299706; within-process sample ranges p1=[293681,316241]; p2=[293881,310251]; p3=[294251,328301]; between-process median range [298911.5,300671.5].
- `requested_alloc_bytes`: per-process medians p1=2146470; p2=2146470; p3=2146470; within-process sample ranges p1=[2146470,2146470]; p2=[2146470,2146470]; p3=[2146470,2146470]; between-process median range [2146470,2146470].
- `peak_live_delta`: per-process medians p1=410697; p2=410697; p3=410697; within-process sample ranges p1=[410697,410697]; p2=[410697,410697]; p3=[410697,410697]; between-process median range [410697,410697].
- `rss_kib`: per-process values p1=6688; p2=6712; p3=6440; range [6440,6712].

## `clone_captured_owner_small`

- Input bytes: 4414; SHA-256: `a15170a9ae51ba26b7287c3069bc29ddc5e843f6212d45bcb9e97a04474a9963`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=410; p2=410; p3=420; within-process sample ranges p1=[410,500]; p2=[410,490]; p3=[420,480]; between-process median range [410,420].
- `requested_alloc_bytes`: per-process medians p1=357; p2=357; p3=357; within-process sample ranges p1=[357,357]; p2=[357,357]; p3=[357,357]; between-process median range [357,357].
- `peak_live_delta`: per-process medians p1=357; p2=357; p3=357; within-process sample ranges p1=[357,357]; p2=[357,357]; p3=[357,357]; between-process median range [357,357].
- `rss_kib`: per-process values p1=6428; p2=6600; p3=6332; range [6332,6600].

## `clone_captured_owner_large`

- Input bytes: 7262; SHA-256: `e3af241b750e8ed984fa403e4470cb47156f821e0dfeb249268324cea674fd25`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=410; p2=410; p3=410; within-process sample ranges p1=[410,500]; p2=[410,480]; p3=[410,480]; between-process median range [410,410].
- `requested_alloc_bytes`: per-process medians p1=357; p2=357; p3=357; within-process sample ranges p1=[357,357]; p2=[357,357]; p3=[357,357]; between-process median range [357,357].
- `peak_live_delta`: per-process medians p1=357; p2=357; p3=357; within-process sample ranges p1=[357,357]; p2=[357,357]; p3=[357,357]; between-process median range [357,357].
- `rss_kib`: per-process values p1=6524; p2=6412; p3=6584; range [6412,6584].

## `inventory_shared_256`

- Input bytes: 8211; SHA-256: `e4e5b5169349be1172437b1976d55edde38b39f4473004d11ddbf32e84cdfdf2`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=2491496; p2=2457616; p3=2461340.5; within-process sample ranges p1=[2473071,2510081]; p2=[2449871,2492211]; p3=[2454231,2488291]; between-process median range [2457616,2491496].
- `requested_alloc_bytes`: per-process medians p1=3155183; p2=3155183; p3=3155183; within-process sample ranges p1=[3155183,3155183]; p2=[3155183,3155183]; p3=[3155183,3155183]; between-process median range [3155183,3155183].
- `peak_live_delta`: per-process medians p1=690565; p2=690565; p3=690565; within-process sample ranges p1=[690565,690565]; p2=[690565,690565]; p3=[690565,690565]; between-process median range [690565,690565].
- `rss_kib`: per-process values p1=6908; p2=7176; p3=7180; range [6908,7180].

## `inventory_shared_1024`

- Input bytes: 19211; SHA-256: `783b58350f586306970a8c2a12e240cfdddc410097ada6e9dc32f0ad921649db`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=9446667; p2=9454191.5; p3=9523432; within-process sample ranges p1=[9288971,9653303]; p2=[9256490,9491462]; p3=[9261741,9652202]; between-process median range [9446667,9523432].
- `requested_alloc_bytes`: per-process medians p1=9124946; p2=9124946; p3=9124946; within-process sample ranges p1=[9124946,9124946]; p2=[9124946,9124946]; p3=[9124946,9124946]; between-process median range [9124946,9124946].
- `peak_live_delta`: per-process medians p1=2730154; p2=2730154; p3=2730154; within-process sample ranges p1=[2730154,2730154]; p2=[2730154,2730154]; p3=[2730154,2730154]; between-process median range [2730154,2730154].
- `rss_kib`: per-process values p1=9988; p2=9984; p3=9904; range [9904,9988].

## `inventory_distinct_256`

- Input bytes: 136089; SHA-256: `0c27be8364b686e56cae88ac84dcf8c1d51167ef4035bc07e6791c331a790085`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=7409738; p2=7403432.5; p3=7546243; within-process sample ranges p1=[7374683,7448443]; p2=[7375043,7448652]; p3=[7368433,7655484]; between-process median range [7403432.5,7546243].
- `requested_alloc_bytes`: per-process medians p1=8681345; p2=8681345; p3=8681345; within-process sample ranges p1=[8681345,8681345]; p2=[8681345,8681345]; p3=[8681345,8681345]; between-process median range [8681345,8681345].
- `peak_live_delta`: per-process medians p1=2048896; p2=2048896; p3=2048896; within-process sample ranges p1=[2048896,2048896]; p2=[2048896,2048896]; p3=[2048896,2048896]; between-process median range [2048896,2048896].
- `rss_kib`: per-process values p1=8952; p2=8972; p3=8816; range [8816,8972].

## `inventory_distinct_1024`

- Input bytes: 539029; SHA-256: `5d9d87da0a1bd327dd1390d77f3e7f951ee30e5ab2905a5e8be7c5820e966c30`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=39979286; p2=40187917; p3=39796000.5; within-process sample ranges p1=[39585014,40295438]; p2=[39754035,40448248]; p3=[39662535,40089257]; between-process median range [39796000.5,40187917].
- `requested_alloc_bytes`: per-process medians p1=31302380; p2=31302380; p3=31302380; within-process sample ranges p1=[31302380,31302380]; p2=[31302380,31302380]; p3=[31302380,31302380]; between-process median range [31302380,31302380].
- `peak_live_delta`: per-process medians p1=7864153; p2=7864153; p3=7864153; within-process sample ranges p1=[7864153,7864153]; p2=[7864153,7864153]; p3=[7864153,7864153]; between-process median range [7864153,7864153].
- `rss_kib`: per-process values p1=17896; p2=17668; p3=17512; range [17512,17896].

## `namespace_heavy`

- Input bytes: 7856; SHA-256: `8b5326afeeaad0952fb4f96d8e520d57c4c258bcd0d5f1f0ed46e15ac0c82cb4`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1040355; p2=1046179.5; p3=1060030; within-process sample ranges p1=[1033515,1061595]; p2=[1034725,1065965]; p3=[1039435,1071665]; between-process median range [1040355,1060030].
- `requested_alloc_bytes`: per-process medians p1=3162344; p2=3162344; p3=3162344; within-process sample ranges p1=[3162344,3162344]; p2=[3162344,3162344]; p3=[3162344,3162344]; between-process median range [3162344,3162344].
- `peak_live_delta`: per-process medians p1=266738; p2=266738; p3=266738; within-process sample ranges p1=[266738,266738]; p2=[266738,266738]; p3=[266738,266738]; between-process median range [266738,266738].
- `rss_kib`: per-process values p1=6856; p2=6668; p3=6904; range [6668,6904].

## `namespace_limit_refusal`

- Input bytes: 87360; SHA-256: `63a68377e0ba64ec021d2fa59532ebbd41de86f625694f919034d4eac3a0ec9e`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=15612209; p2=15903471; p3=15969540.5; within-process sample ranges p1=[15568568,16211632]; p2=[15873940,16059981]; p3=[15939650,16127831]; between-process median range [15612209,15969540.5].
- `requested_alloc_bytes`: per-process medians p1=14689796; p2=14689796; p3=14689796; within-process sample ranges p1=[14689796,14689796]; p2=[14689796,14689796]; p3=[14689796,14689796]; between-process median range [14689796,14689796].
- `peak_live_delta`: per-process medians p1=6058545; p2=6058545; p3=6058545; within-process sample ranges p1=[6058545,6058545]; p2=[6058545,6058545]; p3=[6058545,6058545]; between-process median range [6058545,6058545].
- `rss_kib`: per-process values p1=15460; p2=15076; p3=15732; range [15076,15732].

## `attach_end_to_end_two_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1415341.5; p2=1419971; p3=1430121.5; within-process sample ranges p1=[1403476,1449996]; p2=[1406116,1451107]; p3=[1414276,1443587]; between-process median range [1415341.5,1430121.5].
- `requested_alloc_bytes`: per-process medians p1=5898875; p2=5898875; p3=5898875; within-process sample ranges p1=[5898875,5898875]; p2=[5898875,5898875]; p3=[5898875,5898875]; between-process median range [5898875,5898875].
- `peak_live_delta`: per-process medians p1=521718; p2=521718; p3=521718; within-process sample ranges p1=[521718,521718]; p2=[521718,521718]; p3=[521718,521718]; between-process median range [521718,521718].
- `rss_kib`: per-process values p1=7400; p2=7484; p3=7616; range [7400,7616].

## `attach_end_to_end_two_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1560097; p2=1552977; p3=1551707; within-process sample ranges p1=[1544837,1581377]; p2=[1535036,1571377]; p3=[1534377,1570207]; between-process median range [1551707,1560097].
- `requested_alloc_bytes`: per-process medians p1=7075630; p2=7075630; p3=7075630; within-process sample ranges p1=[7075630,7075630]; p2=[7075630,7075630]; p3=[7075630,7075630]; between-process median range [7075630,7075630].
- `peak_live_delta`: per-process medians p1=724975; p2=724975; p3=724975; within-process sample ranges p1=[724975,724975]; p2=[724975,724975]; p3=[724975,724975]; between-process median range [724975,724975].
- `rss_kib`: per-process values p1=7688; p2=7524; p3=7504; range [7504,7688].

## `attach_end_to_end_one_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1473581; p2=1415537; p3=1414661; within-process sample ranges p1=[1430336,1506777]; p2=[1400676,1437956]; p3=[1400796,1450976]; between-process median range [1414661,1473581].
- `requested_alloc_bytes`: per-process medians p1=5898875; p2=5898875; p3=5898875; within-process sample ranges p1=[5898875,5898875]; p2=[5898875,5898875]; p3=[5898875,5898875]; between-process median range [5898875,5898875].
- `peak_live_delta`: per-process medians p1=521718; p2=521718; p3=521718; within-process sample ranges p1=[521718,521718]; p2=[521718,521718]; p3=[521718,521718]; between-process median range [521718,521718].
- `rss_kib`: per-process values p1=7408; p2=7620; p3=7364; range [7364,7620].

## `attach_end_to_end_one_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1545527; p2=1557156.5; p3=1550007; within-process sample ranges p1=[1536067,1580327]; p2=[1539407,1585357]; p3=[1532437,1576647]; between-process median range [1545527,1557156.5].
- `requested_alloc_bytes`: per-process medians p1=7075628; p2=7075628; p3=7075628; within-process sample ranges p1=[7075628,7075628]; p2=[7075628,7075628]; p3=[7075628,7075628]; between-process median range [7075628,7075628].
- `peak_live_delta`: per-process medians p1=724975; p2=724975; p3=724975; within-process sample ranges p1=[724975,724975]; p2=[724975,724975]; p3=[724975,724975]; between-process median range [724975,724975].
- `rss_kib`: per-process values p1=7688; p2=7568; p3=7740; range [7568,7740].

## `attach_end_to_end_absolute_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1415896.5; p2=1431076.5; p3=1423256.5; within-process sample ranges p1=[1397656,1437666]; p2=[1414926,1452236]; p3=[1405216,1449456]; between-process median range [1415896.5,1431076.5].
- `requested_alloc_bytes`: per-process medians p1=5898865; p2=5898865; p3=5898865; within-process sample ranges p1=[5898865,5898865]; p2=[5898865,5898865]; p3=[5898865,5898865]; between-process median range [5898865,5898865].
- `peak_live_delta`: per-process medians p1=521718; p2=521718; p3=521718; within-process sample ranges p1=[521718,521718]; p2=[521718,521718]; p3=[521718,521718]; between-process median range [521718,521718].
- `rss_kib`: per-process values p1=7360; p2=7348; p3=7560; range [7348,7560].

## `attach_end_to_end_absolute_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1545102; p2=1565532; p3=1560021.5; within-process sample ranges p1=[1530107,1578687]; p2=[1556927,1582617]; p3=[1540766,1606797]; between-process median range [1545102,1565532].
- `requested_alloc_bytes`: per-process medians p1=7075620; p2=7075620; p3=7075620; within-process sample ranges p1=[7075620,7075620]; p2=[7075620,7075620]; p3=[7075620,7075620]; between-process median range [7075620,7075620].
- `peak_live_delta`: per-process medians p1=724975; p2=724975; p3=724975; within-process sample ranges p1=[724975,724975]; p2=[724975,724975]; p3=[724975,724975]; between-process median range [724975,724975].
- `rss_kib`: per-process values p1=7492; p2=7508; p3=7708; range [7492,7708].

## `strict_attach_end_to_end_two_cell_small`

- Input bytes: 3913; SHA-256: `537762ecb2a6d42ac9c7aebcdf9789d8c16f770685c5f0891cca2ad464ebcdaa`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1510462; p2=1514456; p3=1497446.5; within-process sample ranges p1=[1491587,1538137]; p2=[1498207,1534537]; p3=[1483957,1534226]; between-process median range [1497446.5,1514456].
- `requested_alloc_bytes`: per-process medians p1=6437677; p2=6437677; p3=6437677; within-process sample ranges p1=[6437677,6437677]; p2=[6437677,6437677]; p3=[6437677,6437677]; between-process median range [6437677,6437677].
- `peak_live_delta`: per-process medians p1=521591; p2=521591; p3=521591; within-process sample ranges p1=[521591,521591]; p2=[521591,521591]; p3=[521591,521591]; between-process median range [521591,521591].
- `rss_kib`: per-process values p1=7488; p2=7168; p3=7736; range [7168,7736].

## `inverse_attach_detach_two_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1767008; p2=1769127.5; p3=1768303; within-process sample ranges p1=[1739277,1794088]; p2=[1755287,1798788]; p3=[1752038,1800508]; between-process median range [1767008,1769127.5].
- `requested_alloc_bytes`: per-process medians p1=6771299; p2=6771299; p3=6771299; within-process sample ranges p1=[6771299,6771299]; p2=[6771299,6771299]; p3=[6771299,6771299]; between-process median range [6771299,6771299].
- `peak_live_delta`: per-process medians p1=519463; p2=519463; p3=519463; within-process sample ranges p1=[519463,519463]; p2=[519463,519463]; p3=[519463,519463]; between-process median range [519463,519463].
- `rss_kib`: per-process values p1=7404; p2=7576; p3=7488; range [7404,7576].

## `inverse_attach_detach_two_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1928498.5; p2=1938783.5; p3=1921719; within-process sample ranges p1=[1916778,1958389]; p2=[1927359,1965899]; p3=[1911339,1949819]; between-process median range [1921719,1938783.5].
- `requested_alloc_bytes`: per-process medians p1=7897089; p2=7897089; p3=7897089; within-process sample ranges p1=[7897089,7897089]; p2=[7897089,7897089]; p3=[7897089,7897089]; between-process median range [7897089,7897089].
- `peak_live_delta`: per-process medians p1=723289; p2=723289; p3=723289; within-process sample ranges p1=[723289,723289]; p2=[723289,723289]; p3=[723289,723289]; between-process median range [723289,723289].
- `rss_kib`: per-process values p1=7532; p2=7496; p3=7548; range [7496,7548].

## `inverse_attach_detach_one_cell_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1763783; p2=1758422.5; p3=1756037.5; within-process sample ranges p1=[1743908,1789277]; p2=[1743728,1812058]; p3=[1742897,1802978]; between-process median range [1756037.5,1763783].
- `requested_alloc_bytes`: per-process medians p1=6771297; p2=6771297; p3=6771297; within-process sample ranges p1=[6771297,6771297]; p2=[6771297,6771297]; p3=[6771297,6771297]; between-process median range [6771297,6771297].
- `peak_live_delta`: per-process medians p1=519463; p2=519463; p3=519463; within-process sample ranges p1=[519463,519463]; p2=[519463,519463]; p3=[519463,519463]; between-process median range [519463,519463].
- `rss_kib`: per-process values p1=7480; p2=7160; p3=7540; range [7160,7540].

## `inverse_attach_detach_one_cell_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1935533.5; p2=1939444; p3=1911738; within-process sample ranges p1=[1909388,1958909]; p2=[1921259,2030249]; p3=[1895379,1942149]; between-process median range [1911738,1939444].
- `requested_alloc_bytes`: per-process medians p1=7897097; p2=7897097; p3=7897097; within-process sample ranges p1=[7897097,7897097]; p2=[7897097,7897097]; p3=[7897097,7897097]; between-process median range [7897097,7897097].
- `peak_live_delta`: per-process medians p1=723289; p2=723289; p3=723289; within-process sample ranges p1=[723289,723289]; p2=[723289,723289]; p3=[723289,723289]; between-process median range [723289,723289].
- `rss_kib`: per-process values p1=7704; p2=7676; p3=7660; range [7660,7704].

## `inverse_attach_detach_absolute_small`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1771788; p2=1768903; p3=1775422.5; within-process sample ranges p1=[1755198,1811878]; p2=[1745998,1787538]; p3=[1759998,1807028]; between-process median range [1768903,1775422.5].
- `requested_alloc_bytes`: per-process medians p1=6771289; p2=6771289; p3=6771289; within-process sample ranges p1=[6771289,6771289]; p2=[6771289,6771289]; p3=[6771289,6771289]; between-process median range [6771289,6771289].
- `peak_live_delta`: per-process medians p1=519463; p2=519463; p3=519463; within-process sample ranges p1=[519463,519463]; p2=[519463,519463]; p3=[519463,519463]; between-process median range [519463,519463].
- `rss_kib`: per-process values p1=7312; p2=7556; p3=7440; range [7312,7556].

## `inverse_attach_detach_absolute_large`

- Input bytes: 4455; SHA-256: `feb3cf05909aaa5601164c7b80003389f8b78e1d49bcd44648525b1c062afa46`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1932793.5; p2=1937283.5; p3=1931224; within-process sample ranges p1=[1912078,2064739]; p2=[1918879,1967388]; p3=[1914048,2016319]; between-process median range [1931224,1937283.5].
- `requested_alloc_bytes`: per-process medians p1=7897081; p2=7897081; p3=7897081; within-process sample ranges p1=[7897081,7897081]; p2=[7897081,7897081]; p3=[7897081,7897081]; between-process median range [7897081,7897081].
- `peak_live_delta`: per-process medians p1=723289; p2=723289; p3=723289; within-process sample ranges p1=[723289,723289]; p2=[723289,723289]; p3=[723289,723289]; between-process median range [723289,723289].
- `rss_kib`: per-process values p1=7664; p2=7564; p3=7688; range [7564,7688].

## `detach_end_to_end_shared_first_two_cell`

- Input bytes: 4418; SHA-256: `887769bd5883d70c8c0f9eec0cf0406b696829cba59830ee6da89ff6bf5d41c6`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1136670; p2=1139645; p3=1130550; within-process sample ranges p1=[1122205,1180655]; p2=[1130095,1185196]; p3=[1114335,1165425]; between-process median range [1130550,1139645].
- `requested_alloc_bytes`: per-process medians p1=4740636; p2=4740636; p3=4740636; within-process sample ranges p1=[4740636,4740636]; p2=[4740636,4740636]; p3=[4740636,4740636]; between-process median range [4740636,4740636].
- `peak_live_delta`: per-process medians p1=496823; p2=496823; p3=496823; within-process sample ranges p1=[496823,496823]; p2=[496823,496823]; p3=[496823,496823]; between-process median range [496823,496823].
- `rss_kib`: per-process values p1=7596; p2=7628; p3=7404; range [7404,7628].

## `detach_end_to_end_shared_first_one_cell`

- Input bytes: 4419; SHA-256: `9330a4d7d475e6ef60091c3da081e32afed0d96ce654ef82fc90316bf99d48d1`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1139325; p2=1140850; p3=1144030; within-process sample ranges p1=[1131975,1185625]; p2=[1131615,1185495]; p3=[1131375,1191485]; between-process median range [1139325,1144030].
- `requested_alloc_bytes`: per-process medians p1=4740639; p2=4740639; p3=4740639; within-process sample ranges p1=[4740639,4740639]; p2=[4740639,4740639]; p3=[4740639,4740639]; between-process median range [4740639,4740639].
- `peak_live_delta`: per-process medians p1=496824; p2=496824; p3=496824; within-process sample ranges p1=[496824,496824]; p2=[496824,496824]; p3=[496824,496824]; between-process median range [496824,496824].
- `rss_kib`: per-process values p1=7320; p2=7696; p3=7260; range [7260,7696].

## `detach_end_to_end_shared_first_absolute`

- Input bytes: 4419; SHA-256: `4bcbc3e574c90b172a828fb112c6779dbfaab520da2f8df87e9662687bbac6a1`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1166850.5; p2=1133285; p3=1136405; within-process sample ranges p1=[1149115,1219845]; p2=[1123895,1193826]; p3=[1126145,1196235]; between-process median range [1133285,1166850.5].
- `requested_alloc_bytes`: per-process medians p1=4740639; p2=4740639; p3=4740639; within-process sample ranges p1=[4740639,4740639]; p2=[4740639,4740639]; p3=[4740639,4740639]; between-process median range [4740639,4740639].
- `peak_live_delta`: per-process medians p1=496824; p2=496824; p3=496824; within-process sample ranges p1=[496824,496824]; p2=[496824,496824]; p3=[496824,496824]; between-process median range [496824,496824].
- `rss_kib`: per-process values p1=7316; p2=7516; p3=7404; range [7316,7516].

## `detach_end_to_end_shared_final_two_cell`

- Input bytes: 4423; SHA-256: `c01ab5eb838cc1291fe99131730dd01f4629651f2b2728f72815abfcd09c934e`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1127360; p2=1127169.5; p3=1125525; within-process sample ranges p1=[1114695,1166576]; p2=[1112435,1180706]; p3=[1118235,1174226]; between-process median range [1125525,1127360].
- `requested_alloc_bytes`: per-process medians p1=4715560; p2=4715560; p3=4715560; within-process sample ranges p1=[4715560,4715560]; p2=[4715560,4715560]; p3=[4715560,4715560]; between-process median range [4715560,4715560].
- `peak_live_delta`: per-process medians p1=510332; p2=510332; p3=510332; within-process sample ranges p1=[510332,510332]; p2=[510332,510332]; p3=[510332,510332]; between-process median range [510332,510332].
- `rss_kib`: per-process values p1=7468; p2=7452; p3=7668; range [7452,7668].

## `detach_end_to_end_shared_final_one_cell`

- Input bytes: 4424; SHA-256: `c0df095f2b3a0a5fef64620c212ca1709f5ebe21ea1f90f4b23ded51b9fcdc16`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1125720; p2=1130034.5; p3=1126290; within-process sample ranges p1=[1112284,1179445]; p2=[1118415,1151535]; p3=[1116055,1156595]; between-process median range [1125720,1130034.5].
- `requested_alloc_bytes`: per-process medians p1=4715561; p2=4715561; p3=4715561; within-process sample ranges p1=[4715561,4715561]; p2=[4715561,4715561]; p3=[4715561,4715561]; between-process median range [4715561,4715561].
- `peak_live_delta`: per-process medians p1=510333; p2=510333; p3=510333; within-process sample ranges p1=[510333,510333]; p2=[510333,510333]; p3=[510333,510333]; between-process median range [510333,510333].
- `rss_kib`: per-process values p1=7344; p2=7316; p3=7556; range [7316,7556].

## `detach_end_to_end_shared_final_absolute`

- Input bytes: 4421; SHA-256: `e21d0033bdea2452a8895b3cc7af80c5cba83048b5bd71bc3d6950baab590b28`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1138020; p2=1130690; p3=1135290; within-process sample ranges p1=[1129675,1194475]; p2=[1118505,1169165]; p3=[1130095,1167595]; between-process median range [1130690,1138020].
- `requested_alloc_bytes`: per-process medians p1=4715558; p2=4715558; p3=4715558; within-process sample ranges p1=[4715558,4715558]; p2=[4715558,4715558]; p3=[4715558,4715558]; between-process median range [4715558,4715558].
- `peak_live_delta`: per-process medians p1=510330; p2=510330; p3=510330; within-process sample ranges p1=[510330,510330]; p2=[510330,510330]; p3=[510330,510330]; between-process median range [510330,510330].
- `rss_kib`: per-process values p1=7476; p2=7024; p3=7404; range [7024,7476].

## `strict_detach_end_to_end_shared_final_two_cell`

- Input bytes: 4454; SHA-256: `29184d50a733095b6e28ccb7cb563d2b71b5a63ba71e702d46c7cc41fcf772f5`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1193745; p2=1204111; p3=1208880; within-process sample ranges p1=[1182825,1237135]; p2=[1192646,1225875]; p3=[1196696,1227805]; between-process median range [1193745,1208880].
- `requested_alloc_bytes`: per-process medians p1=5251734; p2=5251734; p3=5251734; within-process sample ranges p1=[5251734,5251734]; p2=[5251734,5251734]; p3=[5251734,5251734]; between-process median range [5251734,5251734].
- `peak_live_delta`: per-process medians p1=510234; p2=510234; p3=510234; within-process sample ranges p1=[510234,510234]; p2=[510234,510234]; p3=[510234,510234]; between-process median range [510234,510234].
- `rss_kib`: per-process values p1=7388; p2=7292; p3=7464; range [7292,7464].

## `incoming_edge_shared_final_two_cell`

- Input bytes: 4819; SHA-256: `e20951490498469652936b92fed29e92e5b00d46c8ad9bb99bac1693e2096b5f`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1213415; p2=1224520.5; p3=1226245.5; within-process sample ranges p1=[1199815,1247715]; p2=[1210115,1259845]; p3=[1207725,1276135]; between-process median range [1213415,1226245.5].
- `requested_alloc_bytes`: per-process medians p1=6507529; p2=6507529; p3=6507529; within-process sample ranges p1=[6507529,6507529]; p2=[6507529,6507529]; p3=[6507529,6507529]; between-process median range [6507529,6507529].
- `peak_live_delta`: per-process medians p1=505206; p2=505206; p3=505206; within-process sample ranges p1=[505206,505206]; p2=[505206,505206]; p3=[505206,505206]; between-process median range [505206,505206].
- `rss_kib`: per-process values p1=7296; p2=7564; p3=7288; range [7288,7564].

## `detach_end_to_end_distinct_two_cell_small`

- Input bytes: 5438; SHA-256: `37721c19bd263f0bd36193f37463b1ab7a48cd5b83aae001067a0f13026041b0`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1394031.5; p2=1419641.5; p3=1401601; within-process sample ranges p1=[1381396,1420396]; p2=[1401406,1440206]; p3=[1381766,1434506]; between-process median range [1394031.5,1419641.5].
- `requested_alloc_bytes`: per-process medians p1=4991428; p2=4991428; p3=4991428; within-process sample ranges p1=[4991428,4991428]; p2=[4991428,4991428]; p3=[4991428,4991428]; between-process median range [4991428,4991428].
- `peak_live_delta`: per-process medians p1=521502; p2=521502; p3=521502; within-process sample ranges p1=[521502,521502]; p2=[521502,521502]; p3=[521502,521502]; between-process median range [521502,521502].
- `rss_kib`: per-process values p1=7732; p2=7616; p3=7428; range [7428,7732].

## `detach_end_to_end_distinct_two_cell_large`

- Input bytes: 25291; SHA-256: `9ca07160b4d43c27a963522d01ee98d1ad0cfdfb2c586cd51285a22e5f5b6632`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1841283.5; p2=1803173; p3=1798268; within-process sample ranges p1=[1825778,1870669]; p2=[1787248,1827898]; p3=[1787928,1834819]; between-process median range [1798268,1841283.5].
- `requested_alloc_bytes`: per-process medians p1=6617122; p2=6617122; p3=6617122; within-process sample ranges p1=[6617122,6617122]; p2=[6617122,6617122]; p3=[6617122,6617122]; between-process median range [6617122,6617122].
- `peak_live_delta`: per-process medians p1=948139; p2=948139; p3=948139; within-process sample ranges p1=[948139,948139]; p2=[948139,948139]; p3=[948139,948139]; between-process median range [948139,948139].
- `rss_kib`: per-process values p1=8156; p2=8172; p3=8064; range [8064,8172].

## `detach_end_to_end_distinct_one_cell_small`

- Input bytes: 5438; SHA-256: `67fb540e870961ca827c06ce1054ae8223c94bdc08c4a99354dfc7ab1ccdd67d`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1410126.5; p2=1399351; p3=1415891; within-process sample ranges p1=[1392706,1451846]; p2=[1387747,1427937]; p3=[1393986,1445146]; between-process median range [1399351,1415891].
- `requested_alloc_bytes`: per-process medians p1=4991523; p2=4991523; p3=4991523; within-process sample ranges p1=[4991523,4991523]; p2=[4991523,4991523]; p3=[4991523,4991523]; between-process median range [4991523,4991523].
- `peak_live_delta`: per-process medians p1=521502; p2=521502; p3=521502; within-process sample ranges p1=[521502,521502]; p2=[521502,521502]; p3=[521502,521502]; between-process median range [521502,521502].
- `rss_kib`: per-process values p1=7440; p2=7440; p3=7640; range [7440,7640].

## `detach_end_to_end_distinct_one_cell_large`

- Input bytes: 25297; SHA-256: `b60d704710adff1bb596f955ee7a42509085427c09908319ea7f846554d90251`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1780308; p2=1775323; p3=1777708; within-process sample ranges p1=[1769058,1814677]; p2=[1757478,1806897]; p3=[1768388,1817678]; between-process median range [1775323,1780308].
- `requested_alloc_bytes`: per-process medians p1=6617234; p2=6617234; p3=6617234; within-process sample ranges p1=[6617234,6617234]; p2=[6617234,6617234]; p3=[6617234,6617234]; between-process median range [6617234,6617234].
- `peak_live_delta`: per-process medians p1=948153; p2=948153; p3=948153; within-process sample ranges p1=[948153,948153]; p2=[948153,948153]; p3=[948153,948153]; between-process median range [948153,948153].
- `rss_kib`: per-process values p1=8388; p2=8316; p3=8364; range [8316,8388].

## `detach_end_to_end_distinct_absolute_small`

- Input bytes: 5435; SHA-256: `8eb99c8090871b633cad64f253d1d5dac470b51e6187a76c5b26f49b16ba6382`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1414276; p2=1399451.5; p3=1407591; within-process sample ranges p1=[1399907,1436896]; p2=[1386606,1426326]; p3=[1389256,1427616]; between-process median range [1399451.5,1414276].
- `requested_alloc_bytes`: per-process medians p1=4991615; p2=4991615; p3=4991615; within-process sample ranges p1=[4991615,4991615]; p2=[4991615,4991615]; p3=[4991615,4991615]; between-process median range [4991615,4991615].
- `peak_live_delta`: per-process medians p1=521499; p2=521499; p3=521499; within-process sample ranges p1=[521499,521499]; p2=[521499,521499]; p3=[521499,521499]; between-process median range [521499,521499].
- `rss_kib`: per-process values p1=7464; p2=7464; p3=7840; range [7464,7840].

## `detach_end_to_end_distinct_absolute_large`

- Input bytes: 25282; SHA-256: `1c69ad92d3e94cd42abdf02dba1e0a7f243f0ee9743c60d7d2be86690f63a6b9`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1786498; p2=1794613; p3=1795098; within-process sample ranges p1=[1778238,1817528]; p2=[1779748,1819818]; p3=[1777268,1826168]; between-process median range [1786498,1795098].
- `requested_alloc_bytes`: per-process medians p1=6617286; p2=6617286; p3=6617286; within-process sample ranges p1=[6617286,6617286]; p2=[6617286,6617286]; p3=[6617286,6617286]; between-process median range [6617286,6617286].
- `peak_live_delta`: per-process medians p1=948112; p2=948112; p3=948112; within-process sample ranges p1=[948112,948112]; p2=[948112,948112]; p3=[948112,948112]; between-process median range [948112,948112].
- `rss_kib`: per-process values p1=7704; p2=8044; p3=8172; range [7704,8172].

## `same_picture_attach_detach_two_cell`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=2786872; p2=2796532; p3=2804202.5; within-process sample ranges p1=[2758112,2807052]; p2=[2778602,2818482]; p3=[2791402,2823223]; between-process median range [2786872,2804202.5].
- `requested_alloc_bytes`: per-process medians p1=11922090; p2=11922090; p3=11922090; within-process sample ranges p1=[11922090,11922090]; p2=[11922090,11922090]; p3=[11922090,11922090]; between-process median range [11922090,11922090].
- `peak_live_delta`: per-process medians p1=549145; p2=549145; p3=549145; within-process sample ranges p1=[549145,549145]; p2=[549145,549145]; p3=[549145,549145]; between-process median range [549145,549145].
- `rss_kib`: per-process values p1=7384; p2=7492; p3=7492; range [7384,7492].

## `same_picture_attach_detach_one_cell`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=2800738; p2=2765577.5; p3=2768857; within-process sample ranges p1=[2784412,2813582]; p2=[2746462,2805052]; p3=[2748002,2781703]; between-process median range [2765577.5,2800738].
- `requested_alloc_bytes`: per-process medians p1=11922088; p2=11922088; p3=11922088; within-process sample ranges p1=[11922088,11922088]; p2=[11922088,11922088]; p3=[11922088,11922088]; between-process median range [11922088,11922088].
- `peak_live_delta`: per-process medians p1=549144; p2=549144; p3=549144; within-process sample ranges p1=[549144,549144]; p2=[549144,549144]; p3=[549144,549144]; between-process median range [549144,549144].
- `rss_kib`: per-process values p1=7440; p2=7460; p3=7284; range [7284,7460].

## `same_picture_attach_detach_absolute`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=2784262; p2=2759207; p3=2813107; within-process sample ranges p1=[2760062,2804343]; p2=[2743782,2805192]; p3=[2792853,2843413]; between-process median range [2759207,2813107].
- `requested_alloc_bytes`: per-process medians p1=11922082; p2=11922082; p3=11922082; within-process sample ranges p1=[11922082,11922082]; p2=[11922082,11922082]; p3=[11922082,11922082]; between-process median range [11922082,11922082].
- `peak_live_delta`: per-process medians p1=549141; p2=549141; p3=549141; within-process sample ranges p1=[549141,549141]; p2=[549141,549141]; p3=[549141,549141]; between-process median range [549141,549141].
- `rss_kib`: per-process values p1=7384; p2=7396; p3=7600; range [7384,7600].

## `multisheet_attach_detach`

- Input bytes: 5991; SHA-256: `b49deab44fe133821d3c04e6b95e8854f724c938ac8b02351e4873849d907c42`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=6062667; p2=5941501.5; p3=6033696.5; within-process sample ranges p1=[6034317,6104366]; p2=[5918026,5956406]; p3=[5997846,6071997]; between-process median range [5941501.5,6062667].
- `requested_alloc_bytes`: per-process medians p1=26097151; p2=26097151; p3=26097151; within-process sample ranges p1=[26097151,26097151]; p2=[26097151,26097151]; p3=[26097151,26097151]; between-process median range [26097151,26097151].
- `peak_live_delta`: per-process medians p1=619263; p2=619263; p3=619263; within-process sample ranges p1=[619263,619263]; p2=[619263,619263]; p3=[619263,619263]; between-process median range [619263,619263].
- `rss_kib`: per-process values p1=7596; p2=7604; p3=7536; range [7536,7604].

## `noop_detach_two_cell`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=612097.5; p2=604808; p3=602032.5; within-process sample ranges p1=[597212,632043]; p2=[598273,625673]; p3=[593123,611572]; between-process median range [602032.5,612097.5].
- `requested_alloc_bytes`: per-process medians p1=4483805; p2=4483805; p3=4483805; within-process sample ranges p1=[4483805,4483805]; p2=[4483805,4483805]; p3=[4483805,4483805]; between-process median range [4483805,4483805].
- `peak_live_delta`: per-process medians p1=174735; p2=174735; p3=174735; within-process sample ranges p1=[174735,174735]; p2=[174735,174735]; p3=[174735,174735]; between-process median range [174735,174735].
- `rss_kib`: per-process values p1=6396; p2=6616; p3=6692; range [6396,6692].

## `noop_detach_one_cell`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=606728; p2=607377.5; p3=613792.5; within-process sample ranges p1=[596993,616753]; p2=[597162,623603]; p3=[601553,634643]; between-process median range [606728,613792.5].
- `requested_alloc_bytes`: per-process medians p1=4483805; p2=4483805; p3=4483805; within-process sample ranges p1=[4483805,4483805]; p2=[4483805,4483805]; p3=[4483805,4483805]; between-process median range [4483805,4483805].
- `peak_live_delta`: per-process medians p1=174735; p2=174735; p3=174735; within-process sample ranges p1=[174735,174735]; p2=[174735,174735]; p3=[174735,174735]; between-process median range [174735,174735].
- `rss_kib`: per-process values p1=6576; p2=6700; p3=6520; range [6520,6700].

## `noop_detach_absolute`

- Input bytes: 3886; SHA-256: `de9b89d6a7d62eecda557ca06586313c38b3b0e4d9341e20ddf11b53dd3616ac`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=606712.5; p2=608708; p3=608772.5; within-process sample ranges p1=[592843,629413]; p2=[601063,620513]; p3=[597913,617633]; between-process median range [606712.5,608772.5].
- `requested_alloc_bytes`: per-process medians p1=4483805; p2=4483805; p3=4483805; within-process sample ranges p1=[4483805,4483805]; p2=[4483805,4483805]; p3=[4483805,4483805]; between-process median range [4483805,4483805].
- `peak_live_delta`: per-process medians p1=174735; p2=174735; p3=174735; within-process sample ranges p1=[174735,174735]; p2=[174735,174735]; p3=[174735,174735]; between-process median range [174735,174735].
- `rss_kib`: per-process values p1=6516; p2=6516; p3=6500; range [6500,6516].

## `limit_small`

- Input bytes: 33558319; SHA-256: `f2a7db899a105075f71343cf2df8f3798b9c5ba40ab694c9ce04e56f67fa8181`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=71980; p2=71280.5; p3=72035.5; within-process sample ranges p1=[69530,86121]; p2=[70091,82011]; p3=[70601,83881]; between-process median range [71280.5,72035.5].
- `requested_alloc_bytes`: per-process medians p1=559373; p2=559373; p3=559373; within-process sample ranges p1=[559373,559373]; p2=[559373,559373]; p3=[559373,559373]; between-process median range [559373,559373].
- `peak_live_delta`: per-process medians p1=114951; p2=114951; p3=114951; within-process sample ranges p1=[114951,114951]; p2=[114951,114951]; p3=[114951,114951]; between-process median range [114951,114951].
- `rss_kib`: per-process values p1=104072; p2=104024; p3=104272; range [104024,104272].

## `limit_large`

- Input bytes: 33558888; SHA-256: `591f763e0a548573ee1de424f0a719a956cc40227f0a908bbd7e6cb2e78349e6`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=77630; p2=76860.5; p3=77105.5; within-process sample ranges p1=[76881,96450]; p2=[75831,86500]; p3=[76340,86790]; between-process median range [76860.5,77630].
- `requested_alloc_bytes`: per-process medians p1=625592; p2=625592; p3=625592; within-process sample ranges p1=[625592,625592]; p2=[625592,625592]; p3=[625592,625592]; between-process median range [625592,625592].
- `peak_live_delta`: per-process medians p1=180032; p2=180032; p3=180032; within-process sample ranges p1=[180032,180032]; p2=[180032,180032]; p3=[180032,180032]; between-process median range [180032,180032].
- `rss_kib`: per-process values p1=104644; p2=104516; p3=103868; range [103868,104644].

## `mixed_caps_rejection`

- Input bytes: 4418; SHA-256: `5855cab9f8381c2be8497364dce89a9fe5567f37d99c43c2f48b96547abeb1fc`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1323666; p2=1297430.5; p3=1312226; within-process sample ranges p1=[1305666,1340526]; p2=[1283945,1316056]; p3=[1298906,1324896]; between-process median range [1297430.5,1323666].
- `requested_alloc_bytes`: per-process medians p1=3927733; p2=3927733; p3=3927733; within-process sample ranges p1=[3927733,3927733]; p2=[3927733,3927733]; p3=[3927733,3927733]; between-process median range [3927733,3927733].
- `peak_live_delta`: per-process medians p1=119467; p2=119467; p3=119467; within-process sample ranges p1=[119467,119467]; p2=[119467,119467]; p3=[119467,119467]; between-process median range [119467,119467].
- `rss_kib`: per-process values p1=7108; p2=7040; p3=7096; range [7040,7108].

## `malformed_duplicate_owner`

- Input bytes: 4330; SHA-256: `4ca82dc55c1654c9d2b9710d64ee576b2903fd89fe5263aad9534a2e96142fa2`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=290491.5; p2=292366; p3=290956; within-process sample ranges p1=[286602,326961]; p2=[288671,316952]; p3=[286471,304932]; between-process median range [290491.5,292366].
- `requested_alloc_bytes`: per-process medians p1=1719489; p2=1719489; p3=1719489; within-process sample ranges p1=[1719489,1719489]; p2=[1719489,1719489]; p3=[1719489,1719489]; between-process median range [1719489,1719489].
- `peak_live_delta`: per-process medians p1=148807; p2=148807; p3=148807; within-process sample ranges p1=[148807,148807]; p2=[148807,148807]; p3=[148807,148807]; between-process median range [148807,148807].
- `rss_kib`: per-process values p1=6448; p2=6572; p3=6644; range [6448,6644].

## `malformed_mce_owner`

- Input bytes: 4371; SHA-256: `7b59b378a17de3d9d62208e82d8c52f854d84da062a36e1801490445a96b641c`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=281891; p2=283996.5; p3=286631; within-process sample ranges p1=[277911,293992]; p2=[280131,299781]; p3=[281581,307172]; between-process median range [281891,286631].
- `requested_alloc_bytes`: per-process medians p1=1714045; p2=1714045; p3=1714045; within-process sample ranges p1=[1714045,1714045]; p2=[1714045,1714045]; p3=[1714045,1714045]; between-process median range [1714045,1714045].
- `peak_live_delta`: per-process medians p1=148979; p2=148979; p3=148979; within-process sample ranges p1=[148979,148979]; p2=[148979,148979]; p3=[148979,148979]; between-process median range [148979,148979].
- `rss_kib`: per-process values p1=6792; p2=6708; p3=6648; range [6648,6792].

## `malformed_linked_owner`

- Input bytes: 3851; SHA-256: `861c42438befa6aaf56f0d9c2a72d8beb90efcd314f6988210a9e6b59d45f435`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=269606; p2=266831; p3=267001; within-process sample ranges p1=[264512,298861]; p2=[263412,282781]; p3=[263471,286422]; between-process median range [266831,269606].
- `requested_alloc_bytes`: per-process medians p1=1703476; p2=1703476; p3=1703476; within-process sample ranges p1=[1703476,1703476]; p2=[1703476,1703476]; p3=[1703476,1703476]; between-process median range [1703476,1703476].
- `peak_live_delta`: per-process medians p1=143570; p2=143570; p3=143570; within-process sample ranges p1=[143570,143570]; p2=[143570,143570]; p3=[143570,143570]; between-process median range [143570,143570].
- `rss_kib`: per-process values p1=6680; p2=6768; p3=6664; range [6664,6768].

## `malformed_unknown_uri`

- Input bytes: 3769; SHA-256: `cdc0c7cd1ec206a1ec4ba35ac5f1ab00dbfbaed2749b3984a1b27522222fec7e`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=498392; p2=494512; p3=497227; within-process sample ranges p1=[490682,517563]; p2=[487302,507532]; p3=[488472,510562]; between-process median range [494512,498392].
- `requested_alloc_bytes`: per-process medians p1=3910985; p2=3910985; p3=3910985; within-process sample ranges p1=[3910985,3910985]; p2=[3910985,3910985]; p3=[3910985,3910985]; between-process median range [3910985,3910985].
- `peak_live_delta`: per-process medians p1=165699; p2=165699; p3=165699; within-process sample ranges p1=[165699,165699]; p2=[165699,165699]; p3=[165699,165699]; between-process median range [165699,165699].
- `rss_kib`: per-process values p1=6564; p2=6772; p3=6548; range [6548,6772].

## `multi_picture_same_drawing_16`

- Input bytes: 4126; SHA-256: `9c22d22cba85bddbfbbe26cb42065ca745f13d74c2a21d9fb78d4c4f87983703`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=13131873; p2=13121158; p3=13423924.5; within-process sample ranges p1=[13076787,13163448]; p2=[13090268,13165588]; p3=[13396419,13512870]; between-process median range [13121158,13423924.5].
- `requested_alloc_bytes`: per-process medians p1=30926141; p2=30926141; p3=30926141; within-process sample ranges p1=[30926141,30926141]; p2=[30926141,30926141]; p3=[30926141,30926141]; between-process median range [30926141,30926141].
- `peak_live_delta`: per-process medians p1=653817; p2=653817; p3=653817; within-process sample ranges p1=[653817,653817]; p2=[653817,653817]; p3=[653817,653817]; between-process median range [653817,653817].
- `rss_kib`: per-process values p1=8016; p2=7572; p3=7832; range [7572,8016].

## `multi_picture_same_drawing_64`

- Input bytes: 4876; SHA-256: `c8e389a94a20c44330e36d9ab2573be09e6e6274f7c588b03fe8c828d15f00e1`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=139445375; p2=138903262.5; p3=139107033.5; within-process sample ranges p1=[139238894,139849567]; p2=[138386031,139301414]; p3=[138668572,139731866]; between-process median range [138903262.5,139445375].
- `requested_alloc_bytes`: per-process medians p1=213800728; p2=213800728; p3=213800728; within-process sample ranges p1=[213800728,213800728]; p2=[213800728,213800728]; p3=[213800728,213800728]; between-process median range [213800728,213800728].
- `peak_live_delta`: per-process medians p1=1247097; p2=1247097; p3=1247097; within-process sample ranges p1=[1247097,1247097]; p2=[1247097,1247097]; p3=[1247097,1247097]; between-process median range [1247097,1247097].
- `rss_kib`: per-process values p1=8928; p2=8772; p3=8380; range [8380,8928].

## `multi_picture_same_drawing_256`

- Input bytes: 7692; SHA-256: `c6c0736c087dbd1b10bccf0d010316df72e1286f1956ec9e6322e472c3d47e41`.
- Shape: n=3; samples/process=20.
- `elapsed_ns`: per-process medians p1=1984573621; p2=1988579013; p3=1995196422.5; within-process sample ranges p1=[1939582976,2017354600]; p2=[1951047313,2017688600]; p3=[1955617774,2037882318]; between-process median range [1984573621,1995196422.5].
- `requested_alloc_bytes`: per-process medians p1=2463741687; p2=2463741687; p3=2463741687; within-process sample ranges p1=[2463741687,2463741687]; p2=[2463741687,2463741687]; p3=[2463741687,2463741687]; between-process median range [2463741687,2463741687].
- `peak_live_delta`: per-process medians p1=4608267; p2=4608267; p3=4608267; within-process sample ranges p1=[4608267,4608267]; p2=[4608267,4608267]; p3=[4608267,4608267]; between-process median range [4608267,4608267].
- `rss_kib`: per-process values p1=13056; p2=12784; p3=12468; range [12468,13056].
