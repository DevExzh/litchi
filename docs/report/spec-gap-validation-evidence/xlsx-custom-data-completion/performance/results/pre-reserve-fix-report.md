# XLSX Custom Data completion performance

This report summarizes fresh-process matched observations from the authored, bounded fixtures. It does not establish native Office acceptance, tail-latency guarantees, or asymptotic complexity.

Baseline source: `1d39eea516248c758b8ea429c879a4a3245e4fbb`; samples per lane/size: `15`; warmups: `3`.
Candidate source: `1d39eea516248c758b8ea429c879a4a3245e4fbb`; both captures use the same harness profile and input matrix.
Diagnostic appendix only: this candidate capture predates the actual-edge reserve fix and is excluded from the accepted baseline/candidate comparison. Use `matched-report.md` and `matched-statistics.md` for the corrected candidate.

| size | lane | storages | bindings | payload | median ms | p90 ms | median requested bytes | median peak live bytes | median RSS KiB | candidate elapsed | candidate requested | candidate peak | candidate RSS |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| small | read | 2 | 16 | 1024 B | 0.803644 | 0.813454 | 949575 | 297902 | 4316 | +18.5% | +725.3% | +2203.8% | +4.9% |
| small | noop-commit-save | 2 | 16 | 1024 B | 1.587338 | 1.609058 | 1918354 | 316806 | 4504 | +20.4% | +718.1% | +4223.7% | +0.4% |
| small | payload-replacement | 2 | 16 | 1024 B | 3.078105 | 3.106225 | 4196572 | 513617 | 4800 | +29.3% | +659.0% | +3896.2% | +10.8% |
| small | rename-binding-rewrite | 2 | 16 | 1024 B | 3.174015 | 3.205865 | 4250631 | 525472 | 5032 | +26.4% | +650.8% | +3806.4% | +5.5% |
| small | remove-inverse | 2 | 16 | 1024 B | 4.692472 | 4.731193 | 5510520 | 332849 | 4552 | +27.2% | +752.7% | +8113.6% | +5.7% |
| medium | read | 16 | 128 | 4096 B | 5.626517 | 5.672247 | 7153519 | 2292248 | 6860 | +14.8% | +100.7% | +213.1% | -0.2% |
| medium | noop-commit-save | 16 | 128 | 4096 B | 11.201294 | 11.274784 | 14366009 | 2459575 | 7120 | +16.0% | +100.3% | +475.7% | -0.4% |
| medium | payload-replacement | 16 | 128 | 4096 B | 22.138217 | 22.203207 | 28438013 | 2508860 | 7356 | +19.8% | +103.0% | +738.1% | +3.4% |
| medium | rename-binding-rewrite | 16 | 128 | 4096 B | 22.648829 | 22.730979 | 28866398 | 2591391 | 7632 | +19.3% | +101.5% | +714.3% | +0.1% |
| medium | remove-inverse | 16 | 128 | 4096 B | 34.325415 | 34.435746 | 42078167 | 2548532 | 7372 | +19.2% | +104.6% | +991.8% | +2.9% |
| large | read | 64 | 512 | 16384 B | 22.926921 | 23.122951 | 30439339 | 10676156 | 17840 | +13.5% | +27.3% | +0.0% | -0.1% |
| large | noop-commit-save | 64 | 512 | 16384 B | 46.480054 | 47.027537 | 61078721 | 12121003 | 19640 | +14.5% | +27.2% | +56.2% | +5.0% |
| large | payload-replacement | 64 | 512 | 16384 B | 91.21766 | 91.944184 | 118819322 | 13115084 | 21516 | +17.2% | +29.4% | +103.9% | +1.1% |
| large | rename-binding-rewrite | 64 | 512 | 16384 B | 92.902609 | 94.059754 | 118475216 | 12477471 | 20676 | +17.0% | +29.5% | +109.3% | +1.1% |
| large | remove-inverse | 64 | 512 | 16384 B | 141.946988 | 142.452489 | 175944228 | 13275572 | 20656 | +16.6% | +30.0% | +140.4% | +5.2% |

The timed boundary includes package ingress, the public Custom Data operation, XLSX serialization, and result reopening. Allocator values include the observer's atomic accounting overhead. `requested bytes` is direct allocation plus new realloc bytes; `peak live bytes` is aggregate allocator accounting; RSS comes from `/usr/bin/time -v`.

The three sizes vary storage count, connection count, and payload bytes together. Their rows characterize the authored workloads and should not be read as a proof of an asymptotic slope. Percent deltas compare medians only; p90 columns are descriptive sample percentiles.
