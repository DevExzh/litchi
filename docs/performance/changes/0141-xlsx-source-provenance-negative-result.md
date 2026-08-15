# Change 0141: rejected XLSX source-provenance reuse

## Decision

Change 0141 tested a private source-provenance fast path for seven typed XLSX
source-backed publishers. The candidate retained `SourceLineage` and
`SourceVersion` in each semantic snapshot, then skipped the publication-time
semantic reload when both matched the publishing package.

The candidate is **rejected and fully reverted** by `a12387478`. On the fixed
media-rich XLSX corpus it was 1.04% slower on the pooled seven-case p50
geometric mean. Individual paired directions were inconsistent, the largest
pooled regression was 3.84% p50 for calculation metadata, whole-process
allocation calls improved by only 2.84%, and neither peak heap nor process
VmHWM changed materially. This is not enough benefit to retain seven new
provenance-bearing snapshot paths and their conflict surface.

The negative result applies only to this end-to-end publication shape. It does
not disprove provenance reuse in a path where semantic reload is a larger
fraction of work or where the identity check replaces physical I/O.

## Frozen revisions and environment

```text
control revision: b5ace54a792a0b82d819939d6a25c6e8f8ad0725
control binary SHA-256: c632b2733656a9d932b1299faad0cdc0753fd08364f423da0c8f864ba4b7c4f8
control binary bytes: 39,381,960

candidate revision: eccd8de780bbf0ba0347f62bd156d4c19b7155d5
candidate binary SHA-256: 86842e4c42999ac1a1cf974b4a6227abeb40ab196148d2c0dbe47f9818e97ebf
candidate binary bytes: 39,400,432
revert revision: a12387478

rustc: 1.95.0 (59807616e 2026-04-14)
allocator: Rust system allocator
CPU: AMD EPYC 9575F 64-Core Processor
affinity: CPU 2
kernel: Linux 6.8.0-101-generic x86_64
```

Both binaries were built from clean detached worktrees. Every raw latency
record reports the exact expected revision, release profile,
`git_worktree_dirty: false`, and affinity `2`.

## Scenario and corpus

The run covers the seven affected source-backed publishers:

- calculation metadata;
- defined names;
- page breaks;
- page margins;
- page setup;
- print options;
- sheet protection.

All use the deterministic 16 MiB-media XLSX corpus with SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`:
12 ordinary Parts, 17 ZIP members, eight incompressible 2 MiB PNG Parts,
16,782,412 logical Part bytes, and a 16,786,830-byte archive. The harness
constructs and verifies the expected edit outside timing, reopens the output,
checks exact semantics and topology, and publishes through a bounded sequential
sink.

## Release ABBA result

Four fresh processes ran in strict `A1, B1, B2, A2` order with 20 warmups and
200 measured samples per case. The timer includes source-backed open, edit and
commit planning, and sequential publication.

| Case | A1 p50 ms | B1 p50 ms | B2 p50 ms | A2 p50 ms | pooled p50 change |
|---|---:|---:|---:|---:|---:|
| Calculation metadata | 4.505 | 4.834 | 4.648 | 4.652 | +3.84% |
| Defined names | 4.701 | 4.843 | 4.902 | 4.801 | +2.52% |
| Page breaks | 4.684 | 4.762 | 4.783 | 4.788 | +0.79% |
| Page margins | 4.651 | 4.709 | 4.791 | 4.738 | +1.34% |
| Page setup | 4.715 | 4.751 | 4.765 | 4.811 | -0.04% |
| Print options | 4.802 | 4.640 | 4.896 | 4.698 | +0.42% |
| Sheet protection | 4.769 | 4.840 | 4.737 | 4.899 | -1.52% |

Positive values are regressions. The pooled tail and mean changes are:

| Case | p95 | p99 | mean |
|---|---:|---:|---:|
| Calculation metadata | +5.26% | +7.49% | +3.94% |
| Defined names | +1.34% | +0.21% | +2.64% |
| Page breaks | -0.17% | +1.56% | +0.74% |
| Page margins | -0.22% | -1.68% | +0.34% |
| Page setup | -1.52% | +2.22% | +0.22% |
| Print options | +0.24% | -13.23% | +0.28% |
| Sheet protection | -4.03% | -4.57% | -1.40% |

The p50 geometric mean of the seven candidate/control ratios is `1.010373`,
or a 1.04% regression. The mixed paired directions and small magnitudes do not
support a latency improvement.

## Allocation and process-memory evidence

Heaptrack 1.5.0 ran all seven cases with zero warmups and 20 samples per case.
These are whole-process counts, including setup and untimed verification.

| Revision | allocation calls | temporary allocations | peak heap | peak RSS with Heaptrack |
|---|---:|---:|---:|---:|
| Control | 675,330 | 83,519 | 152.90M | 157.38M |
| Candidate | 656,136 | 81,745 | 152.90M | 157.44M |

Allocation calls fall 2.84% and temporary allocations 2.12%, but peak heap is
unchanged and Heaptrack RSS is near-identical. A separate
`/usr/bin/time -v` three-warmup/30-sample run observed 147,916 KiB control and
146,900 KiB candidate VmHWM (-0.69%). It is one whole-process direction and is
classified as neutral rather than an accepted RSS improvement.

## Correctness and work counters

Across all four latency legs and all 5,600 measured observations:

- corpus descriptors and source archive hashes are identical;
- case-specific output hashes are identical;
- accepted bytes, write calls, largest write, and write-size buckets are
  identical;
- logical source read calls/bytes, ordinary-payload reads, maximum in-flight
  reads, and materialization vectors are identical;
- workbook-owned cases materialize one ordinary payload and worksheet-owned
  cases materialize two.

The old semantic reload generally hit the retained source cache, so removing it
did not reduce logical `ReadAt` work or materialization counts. These counters
do not measure physical disk I/O, decompression, recompression, or memory-copy
bytes, and no claim is made for those dimensions.

## Artifacts

The compact machine-readable result is
[`xlsx-source-provenance-0141-summary.json`](../results/xlsx-source-provenance-0141-summary.json).
The four raw ABBA records, Heaptrack captures and print summaries, and
`/usr/bin/time -v` records share the `xlsx-source-provenance-0141-` prefix in
[`docs/performance/results`](../results/). Their digests are recorded in
[`xlsx-source-provenance-0141.sha256`](../results/xlsx-source-provenance-0141.sha256).

The raw gate verifies clean revision provenance, affinity, sample cardinality,
corpus identity, output identity, sink equality, and exact source-counter
vectors. Both Heaptrack captures parse with Heaptrack 1.5.0 and both time logs
report exit status zero.

## Follow-up boundary

Do not reintroduce generic lineage/version fields into these seven snapshots as
a performance optimization without a different measured mechanism. Higher-ROI
XLSX work should remove physical publication work, broad graph validation,
whole-Part reconstruction, or a larger semantic parse from the timed closure.
The existing cell-value provenance path remains separate because it carries a
larger validated worksheet closure and has its own accepted evidence.
