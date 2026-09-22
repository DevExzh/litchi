# 0731 — public PPT instruction attribution and dispatch limitation

The current public PPT slide-removal lifecycle is reproducibly qualified, but
Callgrind cannot identify its native latency bottleneck on this host. The profile
attributes 67.777% of collected instructions to artifact hashing through software
SHA-256. A separate feature witness shows `sha=true` natively and `sha=false`
under Valgrind. The dependency's runtime dispatch explains that distinction.
Do not transfer the hashing fraction, or an Amdahl ceiling derived from it, to
native execution. Production remains unchanged at `4cb097a4f6`;
`performance_claim: none`.

| Metric | Three-process result |
| --- | ---: |
| Native public lifecycle p50 | 1058.43–1070.69 μs |
| Native mean | 1054.81–1058.57 μs |
| Native p95 | 1168.56–1179.72 μs |
| Native p99 / maximum | 1181.69–1219.42 μs |
| Allocated bytes / calls | 11,776,674 / 5,663 |
| Peak live / end-of-region retained bytes | 2,686,521 / 390,144 |
| Collected Callgrind instructions | 59,578,131–59,578,178 |

Native timing uses three independent processes, each with three warmups and
50 measured lifecycles. Relative to process zero, the other processes' median
changes are −0.324% and −1.145%, and mean changes are +0.047% and −0.308%; none
crosses the prospective absolute 5% central-control threshold. These ranges
are descriptive, not confidence intervals. Allocation uses three separate
one-sample/no-warmup processes and agrees exactly across repeats. It also
matches the established 0728 allocation result, without establishing a latency
comparison against that older build.

The edit removes the second live slide, slide ID 472 / persist ID 10, from
`45543.ppt`, taking eleven slides to ten. Every expected and measured output,
stream inventory, normalized directory metadata, logical-length proof, semantic
survivor witness and negative oracle control matches the sealed 0728 contract.
This retains the original limitation: dependencies within the changed PowerPoint
Document stream are not comprehensively verified. It does not promote broader
producer or CRUD support.

## Collected ownership and call edges

A single non-inlined probe wrapper surrounds the unchanged public
open/edit/remove/commit/output-copy function, including destruction of its local
owners. Callgrind starts with collection disabled and toggles only in that
wrapper. Three profiling processes each execute exactly one measured lifecycle
with no warmup. Expected-output construction, negative controls, untimed output
verification and destruction of the returned output are outside collection.
The wrapper's inclusive count equals the entire collected instruction total;
raw self costs sum to that same total and agree with `callgrind_annotate`.

| Collected edge | Instructions, first process | Fraction of collected owner |
| --- | ---: | ---: |
| Public workflow → transaction commit | 52,560,485 | 88.22% |
| Public workflow → snapshot open | 2,582,792 | 4.34% |
| Public workflow → remove slide | 1,813,821 | 3.04% |
| Public workflow → edit construction | 1,762,979 | 2.96% |
| Commit → artifact hash | 40,380,397 | 67.78% |
| Commit → persisted-slide readback | 3,248,792 | 5.45% |
| Commit → embedded writer finish | 3,087,582 | 5.18% |
| Commit → final snapshot reopen | 2,663,275 | 4.47% |

The final four rows are nested inside commit and must not be added to its row.
All three full self-cost and edge tables are retained in `analysis.json`.
Aggregate function costs can cover repeated calls; they are not independent
phases. In particular, readback occurs before and after publication. The
software SHA compression self cost is 40,370,752 instructions in every repeat.
Memory-copy and memory-set self costs are 14.94% and 9.82% of the collected
owner, but instruction counts are not copied-byte counts or memory bandwidth.

## Dispatch control and next step

The post-profile interpretation control is separate from the frozen nine-run
matrix. The same small feature-detection executable reports:

| Execution | SHA | AVX2 | AVX512F |
| --- | --- | --- | --- |
| Native, CPU 12 | true | true | true |
| Valgrind, CPU 12 | false | true | false |

The raw profile proves software SHA execution under Callgrind. The local
`sha2` 0.11.0 source selects its x86 SHA implementation when the SHA, SSE2,
SSSE3 and SSE4.1 feature checks pass. Native accelerated dispatch is an inference
from that source and host features, not a native sampled stack or phase timing.
The packet preserves the dependency-source digest, relevant dispatch excerpt,
feature executable identity, commands and outputs. No profile was discarded or
rerun to obtain a preferred fraction.

The next bounded experiment should attribute **native public PPT phases**, with
ordinary/observed controls and the same owner lifetimes. Measure commit's
artifact hashing, before/after persisted-slide capture, embedded writer finish,
unrelated-stream preservation and final reopen separately, without removing any
of those checks. The two artifact hashes populate durable patch preconditions;
this evidence does not authorize replacing the digest or skipping its work.
The DOC retained-render result is not an equivalent PPT mechanism.

## Verification and limits

The copied probe changes only its local crate identity and the collection
wrapper. Four existing oracle unit tests, formatting, warning-denied Clippy,
warning-denied rustdoc and release build pass. Source, accepted constraints,
fixture, exact oracle, probe, commands and binary receipts are bound before the
main capture. The analyzer checks all nine processes and native statistics;
ten isolated actual-analyzer controls exercise semantic/output corruption,
missing controls/processes, command changes, zero time, profile digests and
instruction arithmetic. Independent review checks the call edges and dispatch
interpretation. All raw samples, profiles, annotations and controls remain.

No production code changed, so these probe gates are not claimed as a new
full-workspace correctness run. Measurements use CPU 12 on AMD EPYC 9R45,
Rust 1.95.0, release/default-feature builds and warm OS caches. Callgrind Ir is
simulated instruction attribution, not a hardware counter, native latency
fraction, cache result or predicted speedup. No RSS, cold-storage, concurrency,
Office or cross-platform claim is made. The non-iWork goal remains active.

[Evidence, review and replay](results/change-0731/README.md).
