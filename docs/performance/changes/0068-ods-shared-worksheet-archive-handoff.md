# Change 0068: ODS shared worksheet archive handoff

Date: 2026-08-12

Production base: `3341b965dba2039638b28afd370d4518bb7b0b3e`

Status: accepted

## Decision and scope

Retain one exact `Arc<Vec<u8>>` archive owner across the private unified ODS
worksheet handoff, worksheet snapshot/patch lineage, and ODF package reader.
The public API remains `Vec<u8>` on ingress and `&[u8]` on borrow. Durable
unified patches remain `Arc<[u8]>` plus the established `BlobBundle`; other ODS
semantic domains and ODF family crates are unchanged.

The change is private to `litchi-ods` and adds no dependency, cache, runtime,
lock, global state, unsafe code, persisted data, or archive abstraction.

## Problem and implementation

The accepted row-local and raw-member publication paths made the media-rich
ODS transaction fast enough that repeated archive ownership conversions became
material. A unified worksheet edit cloned its current package into the nested
worksheet snapshot. That snapshot converted the owned `Vec` into
`Arc<[u8]>`, copied it back into an ODF package at parse and commit, converted
the target package again, copied the result back to the unified edit, then
cloned it once more for final package validation.

The retained path now:

- stores worksheet snapshot and patch source/target owners as
  `Arc<Vec<u8>>` while keeping all public byte borrows unchanged;
- opens the private ODS package through ODF-common's established
  `family::Package::from_shared_bytes` boundary;
- temporarily moves the unified edit's `Vec` into that shared owner, restores
  the exact allocation after a no-op or error, and moves a changed target back
  out after consuming the nested commit;
- validates staged candidates through the same shared package owner and moves
  the allocation back only after validation succeeds; and
- retains the prior full worksheet parse, row-splice provenance, compactness,
  package publication, unified effect validation, durable patch construction,
  final snapshot reopen, and typed readback order.

`Arc::try_unwrap` is only an ownership fast path. An unexpected outstanding
owner falls back to the former exact clone, so correctness does not depend on
uniqueness. A regression test proves that a closure which stages a worksheet
change and then fails restores both the original bytes and the original
allocation pointer. A second test proves public `Vec` adoption, exact
snapshot/package `Arc` identity, and successful pointer-provenance row-splice
publication.

## Matched corpus and protocol

The unchanged `ods_media_one_edit_save` control contains 2,048 cells and eight
deterministic incompressible 2 MiB media members. The archive is 16,790,689
bytes with 16,887,808 uncompressed payload bytes and SHA-256
`46b7f61cb74639115f6d120dc6498b97d6b310d51c78c4fb85ac60d6fc758b14`.
Each timed operation opens the public unified snapshot, edits one middle cell,
commits the complete durable patch and materializes the final bytes. Outside
timing, every result is reopened and checked for the target cell, package and
manifest inventory, all eight exact media payloads, and deterministic output.

The frozen control binary is
`397121615aba9308a2f5fb3f9126ab5436a00bb3932945aaab276f7523c4b7b1`;
the final candidate is
`5115bb29f7f18324a8ab262757b7db01baf87d6519a0e42cb34de93161bfba82`.
The harness source is identical. On CPU 2, the retained sequence was before A,
after A, after B, before B, then after C, before C, before D, after D. Every leg
used 50 warmups and 500 samples, yielding 2,000 samples per state. Exact legs,
pooled distributions, guards, profiles, counters and memory are in the
[`measurement summary`](../results/ods-worksheet-shared-ownership-summary.json).

## Results

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| pooled samples | 2,000 | 2,000 | — |
| p50 | 76.440421 ms | 60.140386 ms | **-21.32%** |
| mean | 76.703288 ms | 60.364275 ms | **-21.30%** |
| mean 95% interval | 76.626459-76.780116 ms | 60.304795-60.423755 ms | disjoint |
| p95 | 79.504163 ms | 62.686992 ms | **-21.15%** |
| p99 | 83.001595 ms | 64.755446 ms | **-21.98%** |

Every ordered leg improves independently. Control p50 spans
75.958-77.005 ms and candidate p50 spans 59.769-60.499 ms, both within the 5%
drift policy.

The worksheet snapshot/patch-only intermediate binary
`e1ebd46b29ef736654bd7dbb64a8b22169af55dd93879ded7e0b32b52ebf8137`
improved pooled p50 only 4.01% across the same 2,000 samples per state. It was
not accepted. Extending the same ownership proof across the adjacent unified
handoff and candidate validation removed the remaining package transfers and
produced the retained result.

## Ordinary ODS guards

Two balanced 50-sample legs per state cover all large ordinary ODS open/read
and edit cases. The affected full transactions remain inside the regression
gate:

| Case | p50 | Mean | p95 |
|---|---:|---:|---:|
| open | -0.39% | +0.20% | +1.10% |
| exact no-op edit/save | +0.93% | +0.43% | +1.52% |
| one-cell edit/save | +0.60% | +0.36% | +0.10% |

The list/one-cell accessors are too short for a layout claim. Cell sweep and
full-cell-text happened to improve 11.97% and 7.99% p50, respectively, but are
treated only as clean guards because this ownership path is not used by their
timed read loops.

## Allocation, memory, and CPU evidence

One-sample Heaptrack attribution reports allocation calls 382,532 -> 382,522,
temporary allocations flat at 85,601, and peak heap 140.05 -> 109.20 MiB
(-22.03%). Heaptrack-inclusive RSS falls 164.66 -> 132.46 MiB. Both states
retain the same 1.78 KiB profiler/runtime leak. Four uninstrumented GNU Time
processes per state report mean maximum RSS 156,855 -> 124,590 KiB (-20.57%).

Two matched process-wide `perf stat` repeats per state used two warmups and ten
samples:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 1,827.08 ms | 1,613.57 ms | -11.69% |
| cycles | 8.910 billion | 7.956 billion | -10.71% |
| instructions | 18.032 billion | 17.388 billion | -3.57% |
| branches | 2.757 billion | 2.652 billion | -3.81% |
| branch misses | 16.126 million | 15.733 million | -2.43% |
| cache references | 677.383 million | 550.669 million | -18.71% |
| cache misses | 44.281 million | 33.875 million | -23.50% |
| page faults | 351,948 | 255,826 | -27.31% |

Each process reports one initial CPU migration while establishing the pinned
execution. Process-wide profiles include deterministic corpus creation and
untimed complete verification, so their counter reductions understate the
scoped transaction latency delta.

## Validation and limitations

Passed on the final source:

- the complete all-feature ODS suite: 245 tests across the library and every
  integration target;
- focused allocation-identity, row-local raw preservation, typed readback,
  patch/inverse, stale-source, signed/encrypted policy, limits, exact no-op and
  failure-atomic handoff coverage;
- warning-denied ODS library and changed integration-test Clippy plus
  warning-denied ODS rustdoc;
- all 36 performance-harness tests and warning-denied all-target Clippy;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation cleanup from `1194fbc7f`; and
- formatter, JSON, whitespace, and final-diff checks.

The repository-wide ODS all-target Clippy command retains only the unrelated
pre-existing test-only findings recorded by earlier ODS changes. The changed
targets are warning-clean. There is no dedicated ODS fuzz manifest.

This result is a generated warm-memory, same-topology, compact-row transaction.
It does not add source-backed positional reads, resource/structural editing,
new signature or encryption policy, a durable-patch ownership change, or
real-producer coverage. OLE2, OOXML, RTF, other ODF family crates, and every
iWork/IWA crate are unchanged by this batch.
