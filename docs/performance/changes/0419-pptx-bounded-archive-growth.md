# 0419: bound and amortize PPTX candidate archive growth

The owned PPTX cross-copy planner reserved exactly each incoming ZIP chunk.
On the media-rich lifecycle this produced about 36.8 billion cumulative
requested allocation bytes per operation. The allocator counts the entire
new realloc size, even when growth happens in place. The baseline whole-command
heap trace attributed 98.50% of requested bytes directly to this writer.

Commit `0320a6a88` grows the private buffer geometrically, capped at the existing
archive limit. Before owned OPC ingress retains the `Vec`, it makes one
fallible exact-size reservation and copy if spare capacity exists. The code
keeps checked length arithmetic, typed allocation refusal, detached candidate
validation and source authorization. It changes no public API or dependency.

Compared with clean control `7c7561160`, the matched resource diagnostic found:

| Metric | Media-rich lifecycle | Plain lifecycle |
|---|---:|---:|
| Requested allocation volume | 36,811.875 MB → about 369.980 MB (−98.995%) | 24.559 MB → 15.664 MB (−36.217%) |
| Allocation calls | about 61,745 → 59,715 (−3.288%) | 49,681 → 49,335 (−0.696%) |
| Reallocation calls | 9,337 → 7,305 (−21.763%) | 6,433 → 6,085 (−5.410%) |
| End-of-region live bytes | unchanged at 287,885,879 | unchanged at 884,426 |
| Normal median timing, paired directions | −1.60% / −1.67% | **+2.07% / +0.38%** |

Requested-volume reductions do not establish reduced physical copying or an
equivalent reduction in memory use. Whole-process RSS was slightly lower in
every pair (less than 0.16%); treat it as unchanged at this measurement scope.
Process high-water allocator snapshots were unchanged except a 1,568-byte
candidate-leg variation. They are not operation-local peak measurements.

The optimization is retained for the measured reduction in allocation requests
and calls, with the small adverse plain median explicitly accepted. Its final
copy can add work and transient memory; it is not a universal speedup. Each
normal leg has 100 samples/10 warmups, so these timing numbers are diagnostics
and no release latency claim is registered. Allocator legs have 30 samples/
three warmups, measured separately. All same-revision drift checks are below
5%; all paired adverse timing/resource observations are below the declared
5% review threshold. Sources, output bytes and semantic/refusal gates match.

The whole-command heap traces fell from 154.775 GB to 3.160 GB requested
volume. Direct writer reservations fell from 8,478 events / 152.447 GB to
119 events / 0.563 GB; the candidate also incurs final compaction allocations,
included in the whole-command total. Setup and correctness gates are included
in these traces. The interpreted trace replay reproduces the unfiltered
Heaptrack histogram's count and requested-byte total before classifying stacks.

Validation: 843 PPTX all-feature tests passed, two ignored; 552 focused tests
also passed. New tests cover exact output limits, one-byte-short refusal
without mutation, refusal before reservation, zero/empty writes and buffer
handoff. Clippy passed with the same three previously documented command-local
exemptions (`chunks_exact_to_as_chunks`, `clone_on_copy`, `needless_lifetimes`).
The [bundle](../results/change-0419/README.md) retains raw vectors, source/binary
identities, protocols, heap traces, replay tools and the corrected capture
wrapper's failed first attempt. [Resource review](../results/change-0419/resource-review.md)
records the retention limits; [the table](../results/change-0419/result-table.md)
keeps every leg separate.

The output-size bound is not an aggregate live-memory budget. Large workloads
near the default 128 MiB archive bound still need separate memory evidence.
Decoded payload sharing, matched source-backed lifecycles, native producers,
physical cold/range I/O and scaling remain open. The full non-iWork goal is
not complete.
