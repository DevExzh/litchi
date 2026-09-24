# XLSX Custom Data completion performance

This report summarizes fresh-process matched observations from the authored, bounded fixtures. It does not establish native Office acceptance, tail-latency guarantees, or asymptotic complexity.

Baseline source: `1d39eea516248c758b8ea429c879a4a3245e4fbb`; samples per lane/size: `15`; warmups: `3`.
Candidate source: `1d39eea516248c758b8ea429c879a4a3245e4fbb`; both captures use the same harness profile and input matrix.
Candidate identity is the selected-file freeze in `gates/freeze.json` (receipt SHA-256 `f19c50c75401cc9193d22d5605a0376720630e4185d15717db93338cdb8f39ec`; corrected `embedded_data.rs` SHA-256 `0d0a8407724d4ba4c53e2613b547ff227268ac88f8b3e3003bd5d6038663a815`); its detached base commit string is shared with the baseline because the frozen candidate source contains selected working-tree changes. The earlier reserve-ceiling capture is diagnostic only; the accepted candidate uses the corrected fallible one-edge reserve behavior and is the source of all candidate deltas below. See `matched-statistics.md` for allocation-call ratios and `candidate-pre-reserve-fix` for the excluded diagnostic evidence.

| size | lane | storages | bindings | payload | baseline median ms | baseline p90 ms | baseline median requested bytes | baseline median peak live bytes | baseline median RSS KiB | candidate elapsed | candidate requested | candidate peak | candidate RSS |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| small | read | 2 | 16 | 1024 B | 0.803644 | 0.813454 | 949575 | 297902 | 4316 | +17.8% | +7.6% | +0.1% | +5.6% |
| small | noop-commit-save | 2 | 16 | 1024 B | 1.587338 | 1.609058 | 1918354 | 316806 | 4504 | +18.1% | +7.5% | +0.3% | +0.6% |
| small | payload-replacement | 2 | 16 | 1024 B | 3.078105 | 3.106225 | 4196572 | 513617 | 4800 | +25.1% | +9.3% | +0.2% | +5.4% |
| small | rename-binding-rewrite | 2 | 16 | 1024 B | 3.174015 | 3.205865 | 4250631 | 525472 | 5032 | +24.5% | +9.5% | +0.1% | +0.9% |
| small | remove-inverse | 2 | 16 | 1024 B | 4.692472 | 4.731193 | 5510520 | 332849 | 4552 | +24.4% | +10.6% | +0.4% | +0.4% |
| medium | read | 16 | 128 | 4096 B | 5.626517 | 5.672247 | 7153519 | 2292248 | 6860 | +14.5% | +5.5% | +0.0% | -0.2% |
| medium | noop-commit-save | 16 | 128 | 4096 B | 11.201294 | 11.274784 | 14366009 | 2459575 | 7120 | +15.0% | +5.4% | +0.0% | -0.1% |
| medium | payload-replacement | 16 | 128 | 4096 B | 22.138217 | 22.203207 | 28438013 | 2508860 | 7356 | +18.3% | +7.1% | +0.0% | +0.1% |
| medium | rename-binding-rewrite | 16 | 128 | 4096 B | 22.648829 | 22.730979 | 28866398 | 2591391 | 7632 | +17.9% | +7.0% | +0.0% | -0.2% |
| medium | remove-inverse | 16 | 128 | 4096 B | 34.325415 | 34.435746 | 42078167 | 2548532 | 7372 | +17.3% | +7.4% | +0.0% | -0.2% |
| large | read | 64 | 512 | 16384 B | 22.926921 | 23.122951 | 30439339 | 10676156 | 17840 | +13.6% | +4.9% | +0.0% | -0.1% |
| large | noop-commit-save | 64 | 512 | 16384 B | 46.480054 | 47.027537 | 61078721 | 12121003 | 19640 | +13.6% | +4.9% | +0.0% | +3.9% |
| large | payload-replacement | 64 | 512 | 16384 B | 91.21766 | 91.944184 | 118819322 | 13115084 | 21516 | +16.3% | +6.5% | +0.0% | +1.1% |
| large | rename-binding-rewrite | 64 | 512 | 16384 B | 92.902609 | 94.059754 | 118475216 | 12477471 | 20676 | +16.1% | +6.5% | +0.0% | +1.1% |
| large | remove-inverse | 64 | 512 | 16384 B | 141.946988 | 142.452489 | 175944228 | 13275572 | 20656 | +16.1% | +6.8% | +0.0% | +1.3% |

The timed boundary includes package ingress, the public Custom Data operation, XLSX serialization, and result reopening. Allocator values include the observer's atomic accounting overhead. `requested bytes` is direct allocation plus new realloc bytes; `peak live bytes` is aggregate allocator accounting; RSS comes from `/usr/bin/time -v`.

The three sizes vary storage count, connection count, and payload bytes together. Their rows characterize the authored workloads and should not be read as a proof of an asymptotic slope. Percent deltas compare medians only; p90 columns are descriptive sample percentiles.
