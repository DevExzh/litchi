# 0455 matched transfer-chunk measurements

linear-interpolated quantiles; independent within-lane median bootstrap, 2000 resamples, seed 455+candidate lane; descriptive two-repeat ABBA, not a population-tail guarantee

| Provider | Corpus | Repeat | API p50 change | Bootstrap 95% interval | Publication p50 change | RSS change |
|---|---|---|---:|---|---:|---:|
| bytes | media-rich | C1 | -0.246% | [-0.465%, -0.076%] | -0.898% | -0.013% |
| bytes | media-rich | C2 | -0.742% | [-0.922%, -0.508%] | -1.301% | -0.041% |
| bytes | media-rich | C3 | -0.637% | [-0.829%, -0.505%] | -1.144% | -0.000% |
| bytes | media-rich | C4 | -27.962% | [-28.234%, -27.761%] | -36.501% | -0.164% |

All absolute >5% flags (including improvements) follow. Positive values are regressions.

- {"metric": "plan_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -17.536024139049854}
- {"metric": "plan_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -17.470985658331518}
- {"metric": "plan_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -17.40737763589234}
- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -36.50140025881985}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -36.39573372869381}
- {"metric": "publication_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -37.373813802255754}
- {"metric": "api_sum_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -27.962404031344303}
- {"metric": "api_sum_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -27.89528590899287}
- {"metric": "api_sum_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "C4"}, "percent": -28.60714061086528}

Complete phase percentiles, separate allocator populations, request histograms, sink counters and resource counters are retained in `measurements.json`.
