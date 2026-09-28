# 0824 outer-whitespace trial analysis

Offline deterministic replay of retained packet evidence.

Reports: 342  Samples: 7106

## Native timing rows

| case | before p50 | after p50 | after/before | CI95 | flags |
|---|---:|---:|---:|---|---|
| synthetic/tiny/capture | 229226.0 | 228716.0 | 0.997557009 | [0.994499543, 1.0205059] | spread=True |
| synthetic/tiny/commit | 207556.0 | 206666.0 | 0.995997914 | [0.994748396, 0.999761222] | spread=True |
| synthetic/tiny/lifecycle | 1407902.0 | 1397652.0 | 0.993913395 | [0.990330314, 0.996770949] | spread=True |
| synthetic/medium/capture | 425832.0 | 431507.0 | 1.01131574 | [1.00993609, 1.016247] | spread=True |
| synthetic/medium/commit | 290891.5 | 290687.0 | 0.999294874 | [0.997483522, 1.00049794] | spread=True |
| synthetic/medium/lifecycle | 1979289.5 | 1979154.5 | 0.9987966 | [0.996135517, 1.00351351] | spread=True |
| synthetic/large/capture | 16136124.5 | 16725542.5 | 1.03636336 | [1.03455225, 1.04451261] | spread=True |
| synthetic/large/commit | 1261646.5 | 1267396.0 | 1.0051914 | [0.99913699, 1.00717266] | spread=False |
| synthetic/large/lifecycle | 25875311.5 | 26197253.5 | 1.0125054 | [1.0083384, 1.02466632] | spread=False |
| synthetic/vendor/capture | 511013.0 | 521292.5 | 1.02253869 | [1.01121189, 1.02756256] | spread=True |
| synthetic/vendor/commit | 319216.5 | 320967.0 | 1.00608641 | [1.00392659, 1.00745763] | spread=True |
| synthetic/vendor/lifecycle | 2136445.0 | 2140640.5 | 1.00196504 | [0.997187593, 1.00561071] | spread=True |
| synthetic/unicode-vendor/capture | 513552.5 | 524897.5 | 1.02120712 | [1.01634965, 1.0235576] | spread=True |
| synthetic/unicode-vendor/commit | 319826.0 | 320707.0 | 1.00395453 | [1.00115613, 1.01091161] | spread=True |
| synthetic/unicode-vendor/lifecycle | 2142321.0 | 2153520.5 | 1.00561694 | [1.00055059, 1.00874186] | spread=True |
| synthetic/valid-4attr/capture | 489402.0 | 503922.5 | 1.02913231 | [1.02427669, 1.03309368] | spread=True |
| synthetic/valid-4attr/commit | 312566.0 | 313082.0 | 0.999104535 | [0.997943985, 1.00401959] | spread=True |
| synthetic/valid-4attr/lifecycle | 2095550.5 | 2102475.0 | 1.0032161 | [1.00249686, 1.00708102] | spread=False |
| real/real/direct | 1436052.0 | 1286511.5 | 0.895934541 | [0.892216072, 0.899348985] | spread=True |

## Allocation guards

| case | calls before/after | bytes before/after | net live before/after | peak before/after |
|---|---:|---:|---:|---:|
| synthetic/tiny/capture | 1338/1338 | 117635/117635 | 43477/43477 | 62463/62463 |
| synthetic/tiny/commit | 1863/1863 | 148164/148164 | 5947/5947 | 67676/67676 |
| synthetic/tiny/lifecycle | 5670/5670 | 1023938/1023938 | 106060/106060 | 552012/552012 |
| synthetic/medium/capture | 2869/2869 | 228631/228631 | 68221/68221 | 90400/90400 |
| synthetic/medium/commit | 3061/3061 | 233162/233162 | 7666/7666 | 98477/98477 |
| synthetic/medium/lifecycle | 9568/9568 | 1331862/1331862 | 135196/135196 | 592974/592974 |
| synthetic/large/capture | 72106/72106 | 4767939/4767939 | 278201/278201 | 338955/338955 |
| synthetic/large/commit | 15498/15498 | 1111996/1111996 | 17709/17709 | 367392/367392 |
| synthetic/large/lifecycle | 104331/104331 | 8542943/8542943 | 706864/706864 | 1290863/1290863 |
| synthetic/vendor/capture | 3157/3157 | 292567/292567 | 68221/68221 | 90400/90400 |
| synthetic/vendor/commit | 3165/3165 | 269277/269277 | 7666/7666 | 98477/98477 |
| synthetic/vendor/lifecycle | 9998/9998 | 1445364/1445364 | 137151/137151 | 596387/596387 |
| synthetic/unicode-vendor/capture | 3157/3157 | 292567/292567 | 68221/68221 | 90400/90400 |
| synthetic/unicode-vendor/commit | 3165/3165 | 269277/269277 | 7666/7666 | 98477/98477 |
| synthetic/unicode-vendor/lifecycle | 9998/9998 | 1445364/1445364 | 137151/137151 | 596387/596387 |
| synthetic/valid-4attr/capture | 3049/3049 | 277207/277207 | 68221/68221 | 90400/90400 |
| synthetic/valid-4attr/commit | 3138/3138 | 265257/265257 | 7666/7666 | 98477/98477 |
| synthetic/valid-4attr/lifecycle | 9848/9848 | 1421180/1421180 | 136971/136971 | 596207/596207 |
| real/real/direct | 7682/7384 | 2719494/2697806 | 158985/158985 | 305970/305970 |

## Disposition guards

Benefit satisfied: `True`; latency vetoes: `0`; allocation violations: `0`.

RSS increases above five percent are review flags only. The analysis makes no universal, cross-format, cold-cache, tail-latency, RSS-saving, or historical-speedup claim.
