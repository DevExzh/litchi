# Change 0476: reuse owned ZIP Deflate state

`performance_claim: none; descriptive paired allocation-work experiment`

`claim_authorized: false`

The owned ZIP writer previously allocated a fresh raw Deflate backend and a
32 KiB encoder output buffer for every member. The 0475 exact-stack profiles
identified those two allocations as 99.543891% of requested bytes under the
large PPTX benchmark's run ancestry. This change retains one backend and one
fixed output buffer between successfully finalized members.

Pending output follows the previous flate2 adapter's write and sync-flush
ordering. Input accounting, CRC, compressed limits, short writes and failure
timing remain covered by fresh-encoder differential tests. The codec resets
and returns to the archive only after final Deflate output, descriptor and
central-record publication succeed. Incomplete or failed entries discard it.
Store members preserve the available cache; the owned API still uses default
compression. No public API, dependency, runtime or unsafe code is added.

The [evidence bundle](../results/change-0476/README.md) compares authenticated
0474 normal/allocator executables with a clean candidate build at
`bbbdedaa4cd2229403cc6c6f39d66548d3887816`. The only production source
difference is the ZIP writer; one integration test file is added. Frozen
drivers, full source manifests, fixture identities, commands and binary hashes
bind each capture. Four pilots are excluded from the formal comparison.

All 24 main lanes pass, retaining 720 samples. Each allocator row below is
identical in all sixty samples per arm and shape. Requested bytes measure
cumulative operation allocation work. Peak heap is the region high-water
counter minus that sample's live-byte baseline.

| Slides | Requested bytes, control → candidate | Reduction | Peak heap bytes, control → candidate |
| ---: | ---: | ---: | ---: |
| 8 | 21,964,902 → 499,462 | 97.726090% | 435,541 → 435,701 |
| 256 | 227,629,482 → 1,415,242 | 99.378269% | 681,659 → 681,819 |
| 8,192 | 6,809,604,013 → 31,428,173 | 99.538473% | 8,875,092 → 8,875,252 |

Allocation calls decrease from 954/10,145/310,993 to 850/9,049/278,153.
Every allocator sample has zero failed allocation calls and zero live-byte
change at exit. The large case exceeds the frozen 95% requested-byte reduction
criterion. Peak heap rises by 160 bytes at each size, retaining its structural
growth. This is an allocation-work reduction, not a constant-total-memory result.

Normal p50 observations in milliseconds remain separate from allocator timing:

| Slides | R1 control | R1 candidate | R2 control | R2 candidate |
| ---: | ---: | ---: | ---: | ---: |
| 8 | 1.033794 | 0.906510 | 1.036320 | 0.906469 |
| 256 | 8.507085 | 6.544581 | 8.452471 | 6.519916 |
| 8,192 | 254.850610 | 191.026502 | 255.634918 | 185.756656 |

All normal mean/p50/p95/p99 repeat changes are below five percent in magnitude.
The largest is candidate-large p99 at -4.708%. These are descriptive paired
observations; no new registered latency claim is added. Output archive lengths,
hashes and producer oracles remain identical. The large case still writes
16,421 members and exactly 7,940,406 bytes with archive SHA-256
`c7b08da644e651046d368b1baaff9a12c6d7218c4f96914dacb1033722e4b527`.

Four process-counter lanes retain 120 samples; four ten-row shared-format
guards retain another 1,200. All full output/source/sink projections match.
The [guard review](../results/change-0476/guard-review.md) retains one tiny
XLSX p99 penalty in the first pair, not reproduced in the second pair, and
two candidate p99 drift flags. No capture is discarded or replaced. Hardware
counter scheduling is approximately 83%, LLC misses are unsupported, and
whole-process counts include preflight and observers. Main and guard RSS
comparisons stay within five percent; RSS is not an operation-heap measure.

Seventeen required Rust/repository gates pass: 491 ZIP tests, 475 OPC tests,
180 selected streaming unit tests, the named Office streaming integrations,
363 benchmark tests and five allocator tests, scoped formatting, strict Clippy,
rustdoc, crate boundaries and the strict ten-claim registry. Existing ignored
tests are retained in the logs. The initial test-only compile failure, corrected
atomic-limit reference mismatch and cancelled broad integration compilation
remain recorded. The whole-workspace format check finds one unchanged Keynote
formatting difference; iWork is explicitly excluded, and its source is untouched.

Eleven Python evidence tests, exact summary replay, final live source/binary
verification, and a sealed fresh-copy replay pass. Five resealed mutations
are rejected after owned runtime cleanup. The initial verifier pilot-count,
ledger and helper-binding mismatches and an outdated Python test fixture remain
recorded with the corrected results. Shared Cargo caches and user-owned files
are preserved.

ZIP/OPC name indexes and the central directory remain allocated until
finalization. An explicit total-memory design must account for those owners.
Fresh creation remains separate from logical append, Part addition and arbitrary
editing/repackaging. Native-producer breadth, source variants, feature breadth,
worker scaling and the full non-iWork goal remain open.
