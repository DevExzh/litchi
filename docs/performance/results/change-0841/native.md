# Current-source native diagnostic

Absolute values are medians of six process values. Paired estimates are medians of matched block ratios, not quotients of the displayed absolute medians. Whole-child RSS includes setup/readback. Fault counters are whole-child observations.

| Shape | Mode | Metric | Before | After | Paired ratio | 95% lower | 95% upper |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| tiny | write | p50_ns | 12395.0 | 12695.0 | 1.024608 | 1.015247 | 1.036072 |
| tiny | write | p95_ns | 15325.0 | 14245.0 | 0.953575 | 0.883008 | 0.975606 |
| tiny | write | p99_ns | 22895.0 | 15285.5 | 0.908269 | 0.698073 | 1.027815 |
| tiny | write | mean_ns | 13120.7 | 12931.2 | 1.003971 | 0.979916 | 1.025360 |
| tiny | write | rss_kib | 3608.0 | 3566.0 | 0.995267 | 0.966918 | 1.024141 |
| tiny | write | minor_faults | 170.0 | 173.0 | 1.017647 | 0.991412 | 1.026698 |
| tiny | write | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| tiny | lifecycle | p50_ns | 12690.0 | 12830.0 | 1.011401 | 1.000425 | 1.019419 |
| tiny | lifecycle | p95_ns | 15835.0 | 14560.0 | 0.912665 | 0.710304 | 0.978618 |
| tiny | lifecycle | p99_ns | 20070.0 | 15235.5 | 0.917894 | 0.504516 | 1.707009 |
| tiny | lifecycle | mean_ns | 13239.349999999999 | 13058.7 | 1.006087 | 0.928024 | 1.047635 |
| tiny | lifecycle | rss_kib | 3506.0 | 3550.0 | 1.005090 | 0.989819 | 1.028222 |
| tiny | lifecycle | minor_faults | 170.0 | 172.0 | 1.008858 | 0.997110 | 1.023529 |
| tiny | lifecycle | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| many | write | p50_ns | 264986.5 | 264286.0 | 0.999245 | 0.980132 | 1.012208 |
| many | write | p95_ns | 280216.0 | 279136.5 | 0.993803 | 0.976916 | 1.019734 |
| many | write | p99_ns | 282621.0 | 282491.0 | 0.996074 | 0.950664 | 1.021806 |
| many | write | mean_ns | 266673.93333333335 | 266233.75 | 0.997878 | 0.982748 | 1.014653 |
| many | write | rss_kib | 6354.0 | 6356.0 | 1.001265 | 0.994037 | 1.014267 |
| many | write | minor_faults | 966.0 | 996.0 | 1.031040 | 1.030021 | 1.034269 |
| many | write | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| many | lifecycle | p50_ns | 308221.0 | 307077.0 | 0.996078 | 0.987301 | 1.004213 |
| many | lifecycle | p95_ns | 320671.0 | 322231.5 | 1.008216 | 0.980169 | 1.019845 |
| many | lifecycle | p99_ns | 327206.5 | 327397.0 | 0.997821 | 0.966867 | 1.027794 |
| many | lifecycle | mean_ns | 309565.1166666667 | 308732.6333333333 | 0.995669 | 0.990549 | 1.005056 |
| many | lifecycle | rss_kib | 6376.0 | 6414.0 | 1.005975 | 0.996575 | 1.026073 |
| many | lifecycle | minor_faults | 965.0 | 996.5 | 1.033143 | 1.031590 | 1.034787 |
| many | lifecycle | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| payload | write | p50_ns | 319367.0 | 250061.0 | 0.777416 | 0.754436 | 0.827423 |
| payload | write | p95_ns | 358961.5 | 309726.5 | 0.872957 | 0.812667 | 0.908450 |
| payload | write | p99_ns | 598432.0 | 529092.0 | 0.878101 | 0.856369 | 0.912276 |
| payload | write | mean_ns | 329899.56666666665 | 266372.76666666666 | 0.800796 | 0.776073 | 0.836445 |
| payload | write | rss_kib | 39012.0 | 38888.0 | 0.996718 | 0.996209 | 0.998925 |
| payload | write | minor_faults | 95775.0 | 130797.0 | 1.365684 | 1.365639 | 1.365707 |
| payload | write | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| payload | lifecycle | p50_ns | 404812.0 | 330201.0 | 0.815250 | 0.782656 | 0.851678 |
| payload | lifecycle | p95_ns | 456832.0 | 407631.5 | 0.895707 | 0.789787 | 0.935891 |
| payload | lifecycle | p99_ns | 670327.5 | 619732.5 | 0.902301 | 0.869159 | 1.033701 |
| payload | lifecycle | mean_ns | 416701.25 | 348517.1 | 0.836769 | 0.817347 | 0.854773 |
| payload | lifecycle | rss_kib | 38972.0 | 38908.0 | 0.998770 | 0.996976 | 1.000514 |
| payload | lifecycle | minor_faults | 95775.0 | 130798.0 | 1.365680 | 1.365657 | 1.365689 |
| payload | lifecycle | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| unicode | write | p50_ns | 2213759.0 | 1536811.5 | 0.693009 | 0.684760 | 0.712716 |
| unicode | write | p95_ns | 2325995.0 | 1754922.5 | 0.750986 | 0.717900 | 0.796804 |
| unicode | write | p99_ns | 2577625.5 | 1781013.0 | 0.686965 | 0.646890 | 0.768284 |
| unicode | write | mean_ns | 2240127.6666666665 | 1573269.6333333333 | 0.702662 | 0.689590 | 0.729100 |
| unicode | write | rss_kib | 53252.0 | 48476.0 | 0.909821 | 0.909261 | 0.910916 |
| unicode | write | minor_faults | 158080.0 | 155636.0 | 0.984536 | 0.984527 | 0.984562 |
| unicode | write | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| unicode | lifecycle | p50_ns | 2353300.0 | 1672987.0 | 0.707957 | 0.702395 | 0.716485 |
| unicode | lifecycle | p95_ns | 2442985.5 | 1896307.5 | 0.751217 | 0.733666 | 0.797881 |
| unicode | lifecycle | p99_ns | 2722507.0 | 1957103.0 | 0.727259 | 0.681483 | 0.747035 |
| unicode | lifecycle | mean_ns | 2368795.966666667 | 1720012.5166666666 | 0.721972 | 0.710568 | 0.728607 |
| unicode | lifecycle | rss_kib | 53270.0 | 48446.0 | 0.909675 | 0.908290 | 0.910532 |
| unicode | lifecycle | minor_faults | 158080.0 | 155635.0 | 0.984533 | 0.984524 | 0.984539 |
| unicode | lifecycle | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| rich | write | p50_ns | 16475.0 | 16320.5 | 0.991153 | 0.960139 | 0.994530 |
| rich | write | p95_ns | 19560.5 | 18860.5 | 0.964602 | 0.926176 | 0.993515 |
| rich | write | p99_ns | 24030.0 | 26460.0 | 1.115737 | 0.665777 | 1.942186 |
| rich | write | mean_ns | 16961.36666666667 | 16669.75 | 0.977883 | 0.930997 | 1.020625 |
| rich | write | rss_kib | 3542.0 | 3504.0 | 0.986430 | 0.962626 | 1.011576 |
| rich | write | minor_faults | 173.0 | 171.0 | 0.988372 | 0.979819 | 0.994253 |
| rich | write | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |
| rich | lifecycle | p50_ns | 16785.5 | 16525.0 | 0.981524 | 0.978171 | 0.986866 |
| rich | lifecycle | p95_ns | 19715.0 | 18520.0 | 0.938819 | 0.907877 | 0.974225 |
| rich | lifecycle | p99_ns | 27575.5 | 19815.0 | 0.857238 | 0.700854 | 0.992428 |
| rich | lifecycle | mean_ns | 17404.933333333334 | 16624.216666666667 | 0.963013 | 0.946481 | 0.972850 |
| rich | lifecycle | rss_kib | 3524.0 | 3592.0 | 1.020732 | 0.988154 | 1.033637 |
| rich | lifecycle | minor_faults | 172.5 | 171.0 | 0.988439 | 0.979850 | 0.997093 |
| rich | lifecycle | major_faults | 0.0 | 0.0 | n/a | n/a | n/a |

## Threshold observations

0 native p50/RSS guard flags; 0 max/min block-p50 spread flags. Diagnostic only; no adoption is authorized.
