# XLSX Custom Data completion performance

This report summarizes fresh-process matched observations from the authored, bounded fixtures. It does not establish native Office acceptance, tail-latency guarantees, or asymptotic complexity.

Baseline source: `1d39eea516248c758b8ea429c879a4a3245e4fbb`; samples per lane/size: `15`; warmups: `3`.
Candidate capture: pending the root freeze receipt.

| size | lane | storages | bindings | payload | median ms | p90 ms | median requested bytes | median peak live bytes | median RSS KiB | candidate elapsed | candidate requested | candidate peak | candidate RSS |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| small | read | 2 | 16 | 1024 B | 0.803644 | 0.813454 | 949575 | 297902 | 4316 | pending | pending | pending | pending |
| small | noop-commit-save | 2 | 16 | 1024 B | 1.587338 | 1.609058 | 1918354 | 316806 | 4504 | pending | pending | pending | pending |
| small | payload-replacement | 2 | 16 | 1024 B | 3.078105 | 3.106225 | 4196572 | 513617 | 4800 | pending | pending | pending | pending |
| small | rename-binding-rewrite | 2 | 16 | 1024 B | 3.174015 | 3.205865 | 4250631 | 525472 | 5032 | pending | pending | pending | pending |
| small | remove-inverse | 2 | 16 | 1024 B | 4.692472 | 4.731193 | 5510520 | 332849 | 4552 | pending | pending | pending | pending |
| medium | read | 16 | 128 | 4096 B | 5.626517 | 5.672247 | 7153519 | 2292248 | 6860 | pending | pending | pending | pending |
| medium | noop-commit-save | 16 | 128 | 4096 B | 11.201294 | 11.274784 | 14366009 | 2459575 | 7120 | pending | pending | pending | pending |
| medium | payload-replacement | 16 | 128 | 4096 B | 22.138217 | 22.203207 | 28438013 | 2508860 | 7356 | pending | pending | pending | pending |
| medium | rename-binding-rewrite | 16 | 128 | 4096 B | 22.648829 | 22.730979 | 28866398 | 2591391 | 7632 | pending | pending | pending | pending |
| medium | remove-inverse | 16 | 128 | 4096 B | 34.325415 | 34.435746 | 42078167 | 2548532 | 7372 | pending | pending | pending | pending |
| large | read | 64 | 512 | 16384 B | 22.926921 | 23.122951 | 30439339 | 10676156 | 17840 | pending | pending | pending | pending |
| large | noop-commit-save | 64 | 512 | 16384 B | 46.480054 | 47.027537 | 61078721 | 12121003 | 19640 | pending | pending | pending | pending |
| large | payload-replacement | 64 | 512 | 16384 B | 91.21766 | 91.944184 | 118819322 | 13115084 | 21516 | pending | pending | pending | pending |
| large | rename-binding-rewrite | 64 | 512 | 16384 B | 92.902609 | 94.059754 | 118475216 | 12477471 | 20676 | pending | pending | pending | pending |
| large | remove-inverse | 64 | 512 | 16384 B | 141.946988 | 142.452489 | 175944228 | 13275572 | 20656 | pending | pending | pending | pending |

The timed boundary includes package ingress, the public Custom Data operation, XLSX serialization, and result reopening. Allocator values include the observer's atomic accounting overhead. `requested bytes` is direct allocation plus new realloc bytes; `peak live bytes` is aggregate allocator accounting; RSS comes from `/usr/bin/time -v`.

The three sizes vary storage count, connection count, and payload bytes together. Their rows characterize the authored workloads and should not be read as a proof of an asymptotic slope. Percent deltas compare medians only; p90 columns are descriptive sample percentiles.
