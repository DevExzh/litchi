# 0453: share managed decoded payloads in PPTX copy plans

Retained compressed image/chart captures previously sat beside a second decoded
copy in the PPTX plan. The plan now shares OPC's managed `PartData` allocation
when initial capture succeeds. Exact metadata/byte checks authorize reuse during
planner rerun. Ordinary-read fallback retains its original owned-copy behavior;
a late writer Memory refusal performs a checked decoded copy under the existing
destination staging reservation. The larger inline handle is additionally charged.

## Measured result

The final baseline uses production from `456e19246506ffba230052df2b25b072ce4a8749`.
Both builds use identical final standalone harness sources; exact manifests and
binary hashes distinguish working-tree builds carrying the same revision.
The primary matrix has 480 ordinary timing samples and 240 separate allocator
samples, each process with 3 warmups/30 samples, forward/reverse order, CPU 2,
one worker. [Protocol](../results/change-0453/protocol.md) and
[machine](../results/change-0453/machine.json) record the setup.

| Media-rich metric | Before | After | Change |
|---|---:|---:|---:|
| Plan allocated bytes, both repeats | 51,760,790 | 34,983,382 | −32.413% |
| Plan retained live growth, both repeats | 50,425,902 | 33,648,494 | −33.271% |
| Plan allocation calls | 3,404 | 3,388 | −16 calls |
| Publication absolute region peak | 238,097,429 | 221,320,019 | −7.046% |
| Bytes API p50 R1, ms | 26.122 | 25.340 | −2.995% |
| Bytes API p50 R2, ms | 26.123 | 25.282 | −3.217% |

Planning removes 16,777,408 allocated and retained bytes in both repeats.
Publication allocation and peak growth remain unchanged; its absolute live peak
is lower because less memory is already live at entry. Region peaks include
preexisting allocations and are not RSS. Process RSS includes untimed fixture
construction. Simulated range/media API changes −0.054%/−0.066%; provider pacing
uses 64 KiB returns, 200 microseconds/call, and 25 MiB/s separate sleeps.
Source/destination read calls and bytes, source work and exact output are unchanged.
Publication continues to make zero source data reads.

Source planning reservations remain 33,622,602 bytes. Destination planning
reservations increase 16,826,487→16,826,615 bytes (+128); full decoded fallback
admission remains conservative. This is an actual allocation/copy reduction,
not reduced admission, bounded streaming, or a universal zero-copy result.

## Regression review

Plain bytes API medians change +0.053%/+0.296%. Two primary plain p99 triggers
(+10.883%/+10.174%) each contain one roughly 0.33 ms spike, in different phases;
the cause is not established. A separately frozen 240-sample two-ABBA investigation
does not reproduce them: block p99 changes +1.508%/+0.001%, with pooled descriptive
p50/p95/p99 +0.132%/+0.122%/+1.082%. Original results remain. No stable population
p99, physical-network, cold-I/O or scaling claim follows.

All twelve primary flags are reviewed: two positive tails and ten negative
allocator metrics. No repeat trigger exceeds 5%. See [measurements](../results/change-0453/measurements.md),
[tail investigation](../results/change-0453/confirmation-summary.md) and
[full review](../results/change-0453/regression-review.md).

## Correctness and architecture

A one-byte-cache test verifies shared managed ownership, exact rerun identity
across different reread allocations, mismatch/cancellation handling, no bare-Arc
escape, and release after package/capture drop. Public image/chart tests exercise
independent finite source/destination budgets, writer-authorization and target-name
Memory refusals, actual Store→Deflate fallback, unchanged decoded members and
source-media reads, eager/source reopen, and complete budget release.

Final checks pass 850 PPTX, 471 OPC and 381 harness tests (1,702 total), strict
all-target/all-feature lint, warning-denied docs, non-iWork workspace feature
check, explicit-file formatting, and crate boundaries. The unchanged OPC fuzz
target passes 1,000 ASan/sancov runs; this is substrate evidence, not new PPTX fuzz
coverage. The required broad harness gate exposed 50 preexisting mechanical lint
issues, repaired without suppression in the same harness used by both builds.
Failed preliminary receipts remain visible in [validation notes](../results/change-0453/validation-notes.md).

No public API, dependency, unsafe implementation, ambient provider or runtime pool
is added. Accepted ADRs remain unchanged; managed reservations, exact semantic
checks, source identity and fail-closed publication stay in their existing owners.
See [source review](../results/change-0453/source-review.md).

## Replay and remaining work

```sh
python3 -B docs/performance/results/change-0453/verify.py --sealed --cleanup
```

The exported bundle verifies without the repository or deleted build products.
It binds source epochs, binaries, real commands, corpus/output identity, allocator
oracles, derived primary/confirmation values, review flags, fuzz inventory and
owned cleanup. Mutation probes test portable replay before and after cleanup. Owned cleanup
removes 1,646 inventoried files totaling 1,465,309,287 bytes; shared build caches
and the user-owned `docs/GOAL.md` remain untouched.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Native application breadth,
cold I/O, bounded existing append, repackaging and scaling remain required.
The full non-iWork goal is active and uncompleted.
