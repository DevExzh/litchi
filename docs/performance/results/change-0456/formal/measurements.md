# 0456 matched transfer-chunk measurements

linear-interpolated quantiles; independent within-lane median bootstrap, 2000 resamples, seed 455+candidate lane; descriptive two-repeat ABBA, not a population-tail guarantee

| Provider | Corpus | Repeat | API p50 change | Bootstrap 95% interval | Publication p50 change | RSS change |
|---|---|---|---:|---|---:|---:|
| bytes | plain | R1 | -0.512% | [-0.747%, -0.282%] | -0.716% | -4.611% |
| bytes | media-rich | R1 | +15.411% | [+15.177%, +15.989%] | +32.611% | +0.063% |
| range | plain | R1 | +0.007% | [-0.028%, +0.014%] | -0.027% | -1.470% |
| range | media-rich | R1 | -0.051% | [-0.065%, -0.007%] | -0.092% | -0.023% |
| range | media-rich | R2 | -0.385% | [-0.430%, -0.348%] | -0.690% | +0.053% |
| range | plain | R2 | +0.002% | [-0.002%, +0.004%] | -0.024% | -1.506% |
| bytes | media-rich | R2 | +16.347% | [+15.882%, +16.534%] | +32.769% | -0.041% |
| bytes | plain | R2 | -0.206% | [-0.478%, +0.233%] | -0.230% | +0.931% |

All absolute >5% flags (including improvements) follow. Positive values are regressions.

- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 32.61095979232196}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 33.43884364787228}
- {"metric": "publication_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 33.98292218166945}
- {"metric": "api_sum_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 15.411324027802987}
- {"metric": "api_sum_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 16.48759186148383}
- {"metric": "api_sum_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R1"}, "percent": 16.418960196152455}
- {"metric": "open_destination_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R1"}, "percent": 10.209651517016516}
- {"metric": "open_source_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R2"}, "percent": -12.436232830224514}
- {"metric": "open_ns.p99", "pair": {"corpus": "media-rich", "provider": "range", "repeat": "R2"}, "percent": -7.055048971160549}
- {"metric": "publication_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 32.7691857022238}
- {"metric": "publication_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 32.93922057103011}
- {"metric": "publication_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 34.25594224716253}
- {"metric": "api_sum_ns.p50", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 16.34672313875236}
- {"metric": "api_sum_ns.p95", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 16.33971220538133}
- {"metric": "api_sum_ns.p99", "pair": {"corpus": "media-rich", "provider": "bytes", "repeat": "R2"}, "percent": 16.3718339072624}

Complete phase percentiles, separate allocator populations, request histograms, sink counters and resource counters are retained in `measurements.json`.
