# 0455 matched transfer-chunk measurements

linear-interpolated quantiles; independent within-lane median bootstrap, 2000 resamples, seed 455+candidate lane; descriptive two-repeat ABBA, not a population-tail guarantee

| Provider | Corpus | Repeat | API p50 change | Bootstrap 95% interval | Publication p50 change | RSS change |
|---|---|---|---:|---|---:|---:|
| bytes | plain | R1 | +0.262% | [-0.079%, +0.470%] | +0.360% | +0.157% |
| bytes | media-rich | R1 | -1.318% | [-1.459%, -1.138%] | -1.830% | -0.041% |
| range | plain | R1 | -0.001% | [-0.020%, +0.010%] | -0.014% | +2.671% |
| range | media-rich | R1 | -4.775% | [-4.801%, -4.664%] | -9.227% | -0.029% |
| range | media-rich | R2 | -4.449% | [-4.541%, -4.408%] | -8.606% | -0.081% |
| range | plain | R2 | -0.005% | [-0.014%, +0.001%] | +0.021% | +1.394% |
| bytes | media-rich | R2 | -17.817% | [-17.937%, -17.675%] | -16.836% | -0.167% |
| bytes | plain | R2 | -0.011% | [-0.286%, +0.249%] | -0.059% | +0.216% |

All absolute >5% flags (including improvements) follow. Positive values are regressions.

- {"metric": "publication_ns.p99", "pair": {"corpus": "plain", "provider": "range", "repeat": "R1"}, "percent": -22.736183349687455}
- {"metric": "api_sum_ns.p99", "pair": {"corpus": "plain", "provider": "range", "repeat": "R1"}, "percent": -14.631278312041573}
- {"metric": "open_source_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R1"}, "percent": 28.241407142318707}
- {"metric": "open_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R1"}, "percent": 12.928736773620386}
- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R1"}, "percent": -9.226676922158838}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R1"}, "percent": -7.014860646925558}
- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R2"}, "percent": -8.605520688404955}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R2"}, "percent": -9.4772749927429}
- {"metric": "publication_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R2"}, "percent": -8.54149553074317}
- {"metric": "plan_ns.p99", "pair": {"corpus": "plain", "provider": "range", "repeat": "R2"}, "percent": -9.040102186582422}
- {"metric": "plan_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -19.673818082373263}
- {"metric": "plan_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -20.266923866450004}
- {"metric": "plan_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -20.800626269780686}
- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -16.83578153118671}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -16.85045052102756}
- {"metric": "publication_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -17.274227943533493}
- {"metric": "api_sum_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -17.81677273623935}
- {"metric": "api_sum_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -18.08505142556136}
- {"metric": "api_sum_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": -18.329231116125435}

Complete phase percentiles, separate allocator populations, request histograms, sink counters and resource counters are retained in `measurements.json`.
