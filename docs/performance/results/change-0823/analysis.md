# 0823 scanner trial analysis

Offline deterministic replay of retained packet evidence.

Reports: 342  Samples: 7106

## Native timing rows

| case | before p50 | after p50 | after/before | CI95 | flags |
|---|---:|---:|---:|---|---|
| synthetic/tiny/capture | 229351.0 | 228186.0 | 0.994921104 | [0.987901903, 0.999082197] | spread=True |
| synthetic/tiny/commit | 207351.0 | 206421.0 | 0.995147811 | [0.992731886, 1.00096383] | spread=True |
| synthetic/tiny/lifecycle | 1406507.0 | 1392022.0 | 0.989802484 | [0.987826513, 0.992466697] | spread=False |
| synthetic/medium/capture | 427647.0 | 426522.0 | 0.997060008 | [0.995064984, 1.0034954] | spread=True |
| synthetic/medium/commit | 290811.5 | 289101.5 | 0.996750258 | [0.99052798, 0.998964567] | spread=True |
| synthetic/medium/lifecycle | 1980499.5 | 1962504.5 | 0.990261288 | [0.984262259, 0.998348028] | spread=False |
| synthetic/large/capture | 16049970.5 | 16261857.0 | 1.01282785 | [1.00600008, 1.02025999] | spread=False |
| synthetic/large/commit | 1259636.0 | 1252136.0 | 0.99441214 | [0.989442511, 0.997103303] | spread=False |
| synthetic/large/lifecycle | 25767209.5 | 25593958.5 | 0.99303801 | [0.99277988, 0.997842745] | spread=False |
| synthetic/vendor/capture | 512267.5 | 507427.5 | 0.995192476 | [0.984020767, 0.999823622] | spread=False |
| synthetic/vendor/commit | 319732.0 | 318201.5 | 0.995779154 | [0.990141773, 1.00075553] | spread=True |
| synthetic/vendor/lifecycle | 2134261.0 | 2118565.5 | 0.993314096 | [0.987765909, 0.994832527] | spread=False |
| synthetic/unicode-vendor/capture | 513233.0 | 512943.0 | 0.998738804 | [0.995664574, 1.00383464] | spread=True |
| synthetic/unicode-vendor/commit | 321572.0 | 317997.0 | 0.989177727 | [0.985459397, 1.0021389] | spread=True |
| synthetic/unicode-vendor/lifecycle | 2142266.0 | 2125221.0 | 0.992765671 | [0.988008073, 0.995214731] | spread=True |
| synthetic/valid-4attr/capture | 489312.5 | 491197.5 | 1.00026075 | [0.996653845, 1.0072624] | spread=True |
| synthetic/valid-4attr/commit | 312372.0 | 311297.0 | 0.997181857 | [0.992312737, 0.999712683] | spread=True |
| synthetic/valid-4attr/lifecycle | 2094045.5 | 2078595.5 | 0.991837705 | [0.990600275, 0.993995793] | spread=True |
| real/real/direct | 1439947.0 | 1405687.0 | 0.976061924 | [0.97479793, 0.984335705] | spread=False |

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
| real/real/direct | 7682/7682 | 2719494/2719494 | 158985/158985 | 305970/305970 |

## Disposition guards

Benefit satisfied: `False`; latency vetoes: `0`; allocation violations: `0`.

RSS increases above five percent are review flags only. The analysis makes no universal, cross-format, cold-cache, tail-latency, RSS-saving, or historical-speedup claim.
