# 0494: opened DOCX edit/save provider baseline

This extends the existing one-paragraph edit and sequential-save benchmark across
six explicit source providers and a separately verified cold-filesystem path.
It adds benchmark evidence and checks; production DOCX/OPC code is unchanged.

## Warm-provider observations

The synthetic DOCX has 200 paragraphs and eight 2 MiB media members. Each cell
retains 30 samples after three warmups. Two repeats reverse role and provider
order. Values below are descriptive normal-executable medians with bootstrap
95% intervals; p99 is the maximum of 30 observations, not a population-tail estimate.

| Provider | Repeat | p50 ms | Median 95% interval ms | p95 ms | p99 ms |
| --- | ---: | ---: | --- | ---: | ---: |
| owned | 1 | 4.539 | 4.502–4.675 | 5.056 | 6.219 |
| owned | 2 | 4.430 | 4.385–4.501 | 4.673 | 4.711 |
| instrumented | 1 | 2.298 | 2.283–2.313 | 2.465 | 2.473 |
| instrumented | 2 | 4.521 | 4.440–4.579 | 4.747 | 5.357 |
| file-warm | 1 | 4.721 | 4.687–4.749 | 4.851 | 5.000 |
| file-warm | 2 | 2.318 | 2.298–2.328 | 2.375 | 2.387 |
| short | 1 | 5.530 | 3.209–5.595 | 5.713 | 5.716 |
| short | 2 | 5.503 | 5.467–5.534 | 6.021 | 8.047 |
| delayed | 1 | 569.581 | 566.131–574.345 | 594.229 | 602.533 |
| delayed | 2 | 563.998 | 562.823–565.328 | 600.443 | 607.285 |
| range-zero | 1 | 185.923 | 185.295–188.298 | 197.178 | 201.202 |
| range-zero | 2 | 184.997 | 184.853–185.251 | 202.717 | 232.098 |

The two range controls use the same 64 KiB maximum and 100 MiB/s transfer
model. `range-zero` has zero fixed request latency, while `delayed` adds 1 ms
per request. Neither is a real-network measurement. Their nominal transfer
service totals 160.212 ms per operation; the delayed arm adds 377 ms of fixed
service. Observed latency includes operation work and scheduling overhead.

The instrumented/file/delayed paths make 377 nonempty logical reads and return
16,799,430 bytes. The 4 KiB short-read arm makes 4,217 calls and returns the same
bytes. Its 142,628,550 requested logical bytes count progressively shortened
read buffers; underlying reads still request and return only 16,799,430 bytes.
Every output is exactly 16,793,048 bytes in 339 writes, with a 65,536-byte largest
write. Every row materializes one main Part and commits one changed operation.

Normal instrumented medians nearly double between repeats, while the warm-file
median roughly halves. Several allocator repeat-two local cells have much wider
tails. These shared-host observations do not establish a stable local-provider
ranking or a local speedup. Full allocator latency and per-repeat variation
flags remain in the machine-readable analysis; no geometric mean
or before/after production improvement is claimed.

## Memory scope

Every warm allocator operation records 22,859 allocation calls, including 185
reallocations, 22,674 standalone deallocations, and 5,721,361 allocated/deallocated
bytes. Peak live bytes rise 606,986 bytes above the operation-entry baseline.
Source ownership, expected output, trace reservation, and the bounded output
sink are already allocated outside that region: this is not total end-to-end
memory. Whole-child GNU-time RSS ranges from 131,564 to 155,260 KiB across the
24 processes and includes setup, report storage, and post-clock correctness work.

## Verified-cold observations

All 120 formal children passed zero-residency, clean-page and positive storage-I/O
admission, with one materialization, one edit, exact output and stable source.
Each operation returned 16,865,530 logical source bytes in 379 nonempty reads.
Storage `read_bytes` increased by 16,793,600 bytes in every child. The aligned
output is 16,793,612 bytes; its ZIP-tail padding differs from the warm corpus.

| Executable | Repeat | p50 ms | Median 95% interval ms | p95 ms | p99 ms |
| --- | ---: | ---: | --- | ---: | ---: |
| normal | 1 | 148.065 | 95.287–189.559 | 254.398 | 284.198 |
| allocator | 1 | 98.139 | 94.311–105.043 | 244.640 | 246.965 |
| allocator | 2 | 93.802 | 93.105–98.242 | 173.426 | 177.396 |
| normal | 2 | 238.430 | 214.375–528.713 | 676.200 | 712.691 |

Normal repeat-two latency is much higher and the confidence intervals are wide.
This is evidence of substantial variability in these cold observations, not a
stable hardware-latency estimate. The cold clock includes FileSource construction
and version fences and uses a padded source; it must not be compared with warm
file timing as if cache residency were the only difference. Child RSS/VmHWM
and parent/waited-child GNU-time RSS are retained separately in the
[cold analysis](../results/change-0494/analysis/cold-formal1.json).

All 60 cold allocator operations record 22,864 allocation calls, including 185
reallocations, 22,679 standalone deallocations, and 5,722,133 allocated/deallocated
bytes. The operation peak increment is 607,702 bytes. A separate bundle audit
checks live-byte conservation and peak bounds for all 60 formal and three pilot
allocator rows without changing the frozen capture protocol.

## Whole-process profiles

Separate owned, warm-file, and short-read observer processes each run five
measured operations after two warmups. Their whole-process IPC values are
2.147, 2.116, and 2.372 respectively. Corrected syscall summaries account for
13,443, 127,427, and 2,332,877 calls. These counters include corpus setup,
executable hashing, untimed correctness work, and JSON trace serialization;
the large write counts do not measure document sink calls. LLC counters are
unsupported. Recorded L1 zeroes do not establish zero cache misses.

The retained owned `perf record` contains 1,382 samples over 100 operations
and three warmups. Its original stack export requested an unavailable CPU
field; recovery reprocesses the same recording without that field. Corrected
classification finds 23 samples with both `run_sample` and publication markers,
about 1.99% of total sampled period. Most marker-visible period belongs to
untimed output oracles. Bare inline symbols and partial callchains limit
attribution, so this small publication subset does not establish a reliable
ranking of production hot loops or exact phase durations.

Original failed exports and incorrect parsed summaries remain historical;
[separate corrected receipts](../results/change-0494/profiling-counter-recovery-r5/observer-correction.json)
bind the raw inputs and recompute their totals.
The formal timing and allocation captures are independent of these observers.

## Verification and scope

All 720 warm formal samples passed independent recomputation against their raw
reports. The final Rust harness suite passed 468 tests with one existing opt-in
real-producer security test ignored. Warning-denied Clippy/rustdoc, formatting,
the crate-boundary check, and 72 helper tests passed. The boundary check retains
11 explicit preexisting iWork migration debts. Cleanup removed 4,004,667,392
allocated bytes of owned build/scratch files and retained the two authenticated
executables. The bundle verifier checks the raw evidence, final helper custody,
and cleanup inventory.

See [methods](../results/change-0494/methods.md), [warm analysis](../results/change-0494/analysis/warm-formal1.json),
and [the evidence bundle](../results/change-0494/README.md) for source/build binding,
raw ranges, all allocator rows, correctness checks, and interpretation limits.

The full goal remains incomplete. Managed ordinary edits, genuinely borrowed
sources, atomic filesystem save, independent producers, concurrent scaling, and
the broader CRUD matrix remain required.
