# Change 0438: reject ODP fixed-markup batching

The measured candidate improved median fresh ODP creation latency by only
0.703–1.982%, below the predeclared 5% retention threshold. The production
change and its private tests were reverted. The tested source, applicable
patch, checks, raw samples and profiles remain in the
[evidence bundle](../results/change-0438/README.md).

The hypothesis was that grouping adjacent fixed XML pieces would remove
enough repeated Work-budget accounting to improve end-to-end creation.
Four markup groups used a 256-byte stack buffer, with original piecewise
fallback on byte/Work limits. No user-text batching, grammar, public API,
dependency, parallelism or common-owner change was involved.

Both binaries ran the same `odp_streaming_create` selector on CPU 2 with one
worker, 30 samples and three warmups. ABBA order produced 24 formal reports,
720 samples and four fresh profiles. Source manifests bind the exact two-file
candidate delta; both builds used HEAD `3841cc99752baad0816a44fb561e6c610e8e1482`,
with the candidate uncommitted. Binary hashes distinguish the implementations.

| Slides | R1 p50 change | R2 p50 change |
|---:|---:|---:|
| 64 | −0.703% | −1.063% |
| 4,096 | −1.136% | −0.972% |
| 8,192 | −1.224% | −1.982% |

All four required medium/large mean 95% confidence intervals were favorable
and nonoverlapping. All four median gates failed. This is a small measured
effect, insufficient to retain the extra implementation complexity under the
rule recorded before implementation. There were no adverse comparison or
repeat-drift flags above 5%; all 66 matched comparisons remain visible.

Exact archive, content, styles, metadata, semantics and caller-sink identities
held. Chronological operation allocation vectors matched, including a
420,352-byte peak above entry in all 360 formal allocator samples. Whole-process
RSS ranged from 84,537,344 to 84,725,760 bytes. Raw absolute allocator totals
remain retained; their path-dependent process offsets are not operation costs.

Fresh whole-executable consume self samples were 9.03% before and 8.55% after.
Instructions were 29,925,010,026 and 29,558,424,257. These profiles include setup,
corpus generation, warmups, samples and hashing; the later Python oracle and
symbolization are outside the capture. They do not establish an operation-only
causal fraction. Symbolization warnings are retained, and the reported zero
L1-miss counter is not used to claim zero misses.

The candidate passed 356 ODP release tests, including seven new differential
tests, scoped warning-denied Clippy with the existing common enum allowance,
warning-denied rustdoc, minimal-feature checking, formatting and crate
boundaries. Baseline ODP tests passed before implementation and all 349 passed
again after restoration. Portable replay rejects eight independent evidence
mutations. Cleanup removed five task directories totaling 1,827,055,009 bytes
while preserving both Cargo targets and the untracked user goal. There is no new
native, cold/range, scaling or existing-document append evidence in this batch.

The rejected hypothesis is now closed for this corpus and implementation.
Further work should profile a different source of end-to-end cost, especially
required XML/publication validation or compression, before changing it.
Broader non-iWork CRUD coverage and the original performance goal remain open.
