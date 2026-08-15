# Change 0140: ODP repeated-text cache release evidence

## Decision

Change 0140 accepts a narrow release result for the matched selectors added by
[change 0139](0139-odp-repeated-text-cache-evidence.md):

- `odp_source_backed_repeated_text_uncached`
- `odp_source_backed_repeated_text_cached`

On the fixed media-rich ODP corpus, four repeated full-text projections through
`SourceBackedPresentation` are materially faster with the threshold-two cache.
Both paired release directions agree at p50, p95, p99, and mean. Whole-process
Heaptrack allocation-call counts also improve identically in both A/B
profiles. Peak heap and `/usr/bin/time -v` process VmHWM are effectively
unchanged, so no peak-memory reduction is accepted.

This is a selector-pair result from one clean revision, not a claim that ODP
open, single-call text projection, physical I/O, decompression, cold-cache, or
generic ODF behavior improved.

## Frozen build and corpus

The release binary was built from a clean detached worktree at exact revision:

```text
revision: a445bd4cfb3a4e20964473f825fd6cdb0639f83a
binary SHA-256: 9856aec89f9dcb19a27310eaa345cd73ea83061cd8110e74610e24a4597a2a91
binary bytes: 39,371,256
rustc: 1.95.0 (59807616e 2026-04-14)
allocator: Rust system allocator
CPU: AMD EPYC 9575F 64-Core Processor
affinity: CPU 2
kernel: Linux 6.8.0-101-generic
```

Every raw JSON record reports that revision, `git_worktree_dirty: false`, and
affinity `2`. The deterministic corpus remains the exact change-0139 fixture:

```text
slides: 12
archive members: 13
Pictures members: 8
uncompressed Pictures payload: 16,777,216 B
source archive: 16,786,129 B
source archive SHA-256: c5e98dac88846d7b8264f0af4e893d80e21672222c35c3b8890f78cff53242d3
canonical full-text SHA-256: 460bfe509d9c35eb05728c4ff847e0a080aec9bf7a2684ee80b2f9e46b37e3c7
uncompressed Pictures SHA-256: bac87991b97be1a282eabbe32c245dc504bd4344aa01c6d0619b00d41f63983c
```

## Latency method and result

Four fresh processes ran in strict `A1, B1, B2, A2` order on CPU 2. `A` is
the public uncached control and `B` is the cached candidate. Each process used
20 warmups and 200 measured samples. Per sample, owner construction and four
output slots remain outside timing; the timer contains exactly four full-text
projections.

| Leg | Selector | p50 ns | p95 ns | p99 ns | mean ns |
|---|---|---:|---:|---:|---:|
| A1 | uncached | 562,063 | 610,509 | 639,480 | 566,597.395 |
| B1 | cached | 304,629 | 334,229 | 384,260 | 307,443.045 |
| B2 | cached | 301,845 | 332,177 | 371,721 | 305,538.500 |
| A2 | uncached | 562,328 | 613,203 | 680,957 | 569,295.905 |

Paired candidate reductions are:

| Pair | p50 | p95 | p99 | mean |
|---|---:|---:|---:|---:|
| A1/B1 | 45.80% | 45.25% | 39.91% | 45.74% |
| A2/B2 | 46.32% | 45.83% | 45.41% | 46.33% |

Both directions therefore accept a latency improvement for this four-call
source-backed projection shape. The second pair's larger A2 p99 maximum does
not reverse p99 or mean, and no single-leg statistic is used as the decision.

## Allocation and process-memory evidence

Heaptrack 1.5.0 captured four fresh CPU-pinned profiles in the same
`A1, B1, B2, A2` order, using three warmups and 30 samples per process. These
are whole-process counts: they include corpus setup, owner preparation, timed
work, and the untimed correctness replay.

| Leg | allocation calls | temporary allocations | peak heap |
|---|---:|---:|---:|
| A1 | 985,072 | 363,854 | 89.22M |
| B1 | 844,141 | 301,106 | 89.22M |
| B2 | 844,141 | 301,106 | 89.22M |
| A2 | 985,072 | 363,854 | 89.22M |

The cached selector records 14.31% fewer allocation calls and 17.25% fewer
temporary allocations in both directions. Peak heap is unchanged, and the
profiles do not expose operation-local allocated-byte totals; no such claims
are made.

Matched `/usr/bin/time -v` fresh processes used the same three-warmup,
30-sample shape. Maximum resident set sizes were 97,244 KiB (A1), 97,244 KiB
(B1), 97,084 KiB (B2), and 97,240 KiB (A2). One pair is identical and the
other differs by only 0.16%, so process VmHWM is classified as near-neutral
and no RSS reduction is accepted.

## Correctness and source-work gates

Every latency and resource record retains the change-0139 gates:

- exact archive, text, and Pictures hashes;
- 12 slides, 13 archive members, eight Pictures, and 16 MiB media identity;
- four output projections equal to the eager oracle;
- zero post-preparation `ReadAt` calls, bytes, compressed-range overlap, and
  picture-payload reads;
- exact per-call source observations `[3,3,3,3]` for the control and
  `[3,5,2,2]` for the candidate.

The observed improvement is therefore parse/projection/cache work over already
prepared source-backed content. It is not evidence of reduced physical I/O or
decompression.

## Artifacts and verification

The compact result is
[`odp-repeated-text-cache-0140-summary.json`](../results/odp-repeated-text-cache-0140-summary.json).
The four raw latency JSON records, four raw RSS JSON records, four
`/usr/bin/time -v` logs, and four Heaptrack profiles share the
`odp-repeated-text-cache-0140-` prefix in
[`docs/performance/results`](../results/). Their digests are recorded in
[`odp-repeated-text-cache-0140.sha256`](../results/odp-repeated-text-cache-0140.sha256).

The exact raw-record gate verifies revision/cleanliness/affinity, sample
cardinality, case order, corpus identity, output hashes, zero replay reads, and
freshness vectors. `heaptrack_print 1.5.0` parses all four profiles and reports
the counts above. All four time logs report exit status zero.

## Applicability and limitations

Accepted claims apply only to the named selector pair, exact corpus, prepared
`SourceBackedPresentation`, and four-call full-text projection shape. They do
not cover single-call queries, ODP open, slide-object projection, edit/save,
physical or cold-cache I/O, decompression, operation-local allocated bytes,
peak heap, RSS, other ODP producers, other ODF formats, or iWork.
