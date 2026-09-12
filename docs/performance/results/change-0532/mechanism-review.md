# CFB claim mechanism and baseline summary

All measured constructors preserve their existing validation. No optimization is adopted. The same 15-instruction successful claim path accounts for a substantial share of large-sector cases.

| Workload | Repeat | Physical sectors / claims per constructor | Claim exclusive share | Physical reconciliation share | Stream validation exclusive share |
|---|---:|---:|---:|---:|---:|
| xls_owned_source_open_one_cell / 256-comments-opaque-heavy | 1 | 33,194 | 17.8164% | 14.2536% | 14.1656% |
| cfb_open / tiny | 1 | 6 | 0.1836% | 0.1754% | 1.7480% |
| cfb_open / many-small | 1 | 614 | 0.3282% | 0.2631% | 3.8291% |
| cfb_open / few-large | 1 | 33,031 | 18.9505% | 15.1609% | 15.0530% |
| xls_owned_source_open_one_cell / 256-comments-opaque-heavy | 2 | 33,194 | 17.8209% | 14.2572% | 14.1692% |
| cfb_open / tiny | 2 | 6 | 0.1836% | 0.1754% | 1.7480% |
| cfb_open / many-small | 2 | 614 | 0.3282% | 0.2631% | 3.8291% |
| cfb_open / few-large | 2 | 33,031 | 18.9505% | 15.1609% | 15.0530% |

The percentages use disjoint exclusive function costs divided by the matching parent-constructor instructions. Inclusive child and parent values overlap and are not added. Physical-sector counts follow the fixed generators’ 512-byte `OleWriter::new()` geometry, corroborated by every timed dump. This is not a statement about arbitrary CFB inputs.

The raw Callgrind claim count and exclusive instructions satisfy `Ir = 15 × calls` in all 40 timed dumps. All six generated helper variants have a 363-byte body, a 112-byte stack reservation, checked bounds and ownership branches, and the same 15-instruction success path. Caller assembly retains explicit calls. Static code size is not execution time, and no speedup is extrapolated from these counts.

Native and allocation evidence remains separate: 24,000 native samples, 720 allocation samples, no greater-than-five-percent native/RSS repeat flag, and identical paired allocation vectors for calls, reallocations, allocated bytes and incremental region peak in all 12 scenarios. Two repeats do not establish a universal noise bound. Instrumented timing and whole-process memory are not inferred from allocation regions.

The next candidate is private cold error helpers plus ordinary inlining, preserving the exact checked claim body and collect-then-claim order. It requires a fresh frozen ABBA comparison and all correctness/guard gates before adoption. See `next-candidate.md`.
