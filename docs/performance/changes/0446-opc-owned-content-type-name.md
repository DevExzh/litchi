# 0446: Transfer owned content-type Part names

The content-type parser passed an owned Part-name String by reference into
`PackURI::new`, whose `Into<String>` conversion cloned it. Passing ownership
removes that allocation. The same bytes, validation order, resource limits,
case folding, freshness checks and required manifest parses remain.

The frozen plain-source OPC one-Part/root-relationship addition matrix retains
24 reports/720 samples in before/after ABBA order at 64/1024/4096 Parts, with
normal and allocator builds, CPU 2, one worker, 30 samples and three warmups.

| Existing Parts | Before calls | After calls | Reduction | Requested-byte reduction |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 3,526 | 3,333 | 5.474% | 0.219% |
| 1,024 | 51,538 | 48,465 | 5.963% | 1.070% |
| 4,096 | 205,145 | 192,856 | 5.990% | 1.329% |

Both repeats save exactly `3*N+1` allocation calls and `78*N+25` requested
bytes in every sample, matching the three required content-type parses. The
frozen 5% medium/large allocation gate passes. Peak allocation above entry stays
705,777 / 2,716,017 / 9,181,553 bytes; endpoint live delta stays zero.

Normal medium p50 falls 2.096%/2.846%; large falls 2.367%/2.204%. These do not
meet the separate 3% gate, so **no latency improvement is claimed**. No paired
latency/RSS regression or repeat drift exceeds 5%. Six absolute paired triggers
are allocation-call improvements. Every result, bootstrap interval, tail, RSS
value and comparison is in [measurements](../results/change-0446/measurements.md)
and [summary.json](../results/change-0446/summary.json).

Four profiles retain whole-process counters and weighted stacks. Content-type
parsing remains 29.810% inclusive within the candidate run-frame subset, which
includes untimed setup/probes/warmups and is not the exact timed region. Inclusive
rows overlap. Both stack exports have no unparsed lines; the guest's zero L1
miss event supports no cache claim. Continued parser/map and source-catalog
allocation work merits measurement before broad caching or SIMD changes.

All 458 OPC and 373 harness tests pass. Owner strict lint is clean; harness
strict lint has zero new diagnostics relative to retained debt. Warning-denied
docs, non-iWork workspace/feature checks, formatting and boundaries pass. The
independent fixture/report oracle and 29 corruption probes pass. The initial
missing-data-path/protocol preflight failures are retained; no workload or source
edit occurred before the actual freeze. See [validation](../results/change-0446/validation-notes.md).

Retain this one-line production optimization for its allocation reduction.
No public API, dependency, unsafe code or concurrency behavior changes. The
coverage index and opt-in registry are unchanged. This is synthetic package
addition, not semantic Office owner/native/cold/range/scaling evidence. The full
non-iWork goal remains active and incomplete.

Reproduce portable evidence with `python3 -B docs/performance/results/change-0446/verify.py --sealed --cleanup`.
Capture commands, frozen protocol, candidate sources, binary/source hashes,
original failures and cleanup receipts are retained in the result bundle.
