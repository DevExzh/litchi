# Change 0062: native PPT root adopts validated text publication

Date: 2026-08-12

Production base: `4abe2e0197df0fb8075d973b84b9463863938a52`

Status: accepted

## Hypothesis and implementation

`slide_order::Transaction::set_shape_text` already delegated one isolated shape
replacement to the complete native text owner. That owner source-checks the live
slide record, replaces exactly that persisted slide record, finishes the CFB
editor, reopens the generic presentation, and independently reads back the
selected shape. The root transaction then copied the resulting artifact and
immediately repeated the stronger root snapshot open even though the live
document record, slide order, review-history state, and root policy had not
changed.

The text commit now carries a crate-private publication certificate containing
its exact source artifact, exact output allocation, and sole replaced slide
persist ID. The root accepts that certificate only when its source bytes equal
the current working snapshot, the persist ID equals the selected existing
slide, and neither expected nor published identity is the live document record.
It then adopts the already validated output allocation and carries forward the
unchanged document structure, document persist ID, review-history flag, and
default limits.

Nondefault `RecordLimits` retain the original complete
`Snapshot::from_bytes_with_limits` path. Structural, visibility, transition,
anchor, media, insert, remove, reorder, durable-patch, signed/protected-source,
and final root-publication paths are unchanged. There is no public API,
dependency, runtime, lock, cache, unsafe code, or global state change.

The root now also constructs the private text snapshot from its existing
`Arc<[u8]>`; validation opens that shared immutable source directly. The public
standalone text snapshot and direct text-edit transaction retain their existing
owned-ingress behavior.

## Matched corpus and protocol

The existing public `ppt_semantic_one_edit_save` case is the acceptance
scenario. Snapshot capture remains outside timing. The timed region creates a
root edit, performs one same-length middle-shape replacement, commits it, and
materializes the output bytes. Every iteration compares exact deterministic
bytes; complete patch replay, inverse restoration, selected-shape readback, and
all 144 shapes through the generic presentation facade stay outside timing.

The deterministic large artifact contains 144 text boxes, four CFB streams,
9,072 logical text bytes, and 40,960 package bytes. Its SHA-256 is
`229052cd918c0e5b7ef44070bafe20833531eee119b5943b18499503e225ff52`;
the 37,385-byte `PowerPoint Document` stream SHA-256 is
`bef446ada643821b87531c06be7564b7ff8ca5539bb6a39766fbd28c11f65523`.

The frozen control executable SHA-256 is
`6acb71f5e4ec5366aaf31233b3ab9877d7b39f5a339bd18f504e88ac336075ea`.
The candidate SHA-256 is
`718297e1e86e716284d82f233ab09fd143d07e4bc6ad04c477f6d62b4f07344a`.
Both are release builds from Rust 1.95.0 on Linux 6.8.0-101-generic, AMD EPYC
9575F, system allocator. Latency runs were pinned to CPU 2.

The balanced sequence was before A, after A, after B, before B. Each primary
leg used 50 warmups and 1,000 samples. P50 drift was 0.62% within the control
and 2.29% within the candidate; mean drift was 0.15% and 3.48%. Exact pooled
values are retained in the
[`measurement summary`](../results/ppt-root-text-adoption-summary.json).

## Results

| Large root one-shape edit/save | Before | After | Delta |
|---|---:|---:|---:|
| pooled samples | 2,000 | 2,000 | — |
| p50 | 352.306 us | 286.805 us | **-18.59%** |
| mean | 355.365 us | 292.012 us | **-17.83%** |
| mean 95% interval | 354.596-356.135 us | 291.063-292.960 us | disjoint |
| p95 | 386.252 us | 322.213 us | **-16.58%** |
| p99 | 409.916 us | 374.435 us | **-8.66%** |

The improvement is larger than the prior isolated 34.227 us root-open p50
because the handoff also removes an artifact copy and lets the private root
text snapshot borrow its existing allocation.

## Unaffected and size guards

The large direct text editor does not execute the root adoption path. Ordinary
open, exact root no-op, and root snapshot open are likewise unaffected. Each
guard pooled 2,000 samples per state, except the repeated root-open guard at
8,000. The first 2,000-sample root-open observation had a +10.53% p99 trigger
despite neutral p50/mean; it was retained and repeated at higher power. The
repeat is neutral through p99.

| Workload | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| Direct text edit/save | +0.95% | +0.63% | +1.48% | -5.42% |
| Exact root no-op edit/save | -0.42% | -1.04% | -1.87% | -10.98% |
| Ordinary semantic open | +0.77% | +0.24% | -2.36% | -9.24% |
| Root snapshot open, repeated | +0.18% | +0.23% | +0.65% | -1.93% |
| Tiny root one-shape edit/save | -16.58% | -17.31% | -19.18% | -17.96% |

## Allocations, counters, and memory

Matched Heaptrack processes used 1,000 primary samples and one final complete
verifier:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 12,283,120 | 11,479,317 | **-6.54%** |
| temporary allocations | 1,356,122 | 1,301,067 | -4.06% |
| peak heap | 941.22 KiB | 938.00 KiB | -0.34% |
| Heaptrack RSS | 12.55 MiB | 12.56 MiB | +0.08% |
| leaked bytes | 544 B | 544 B | unchanged |

Uninstrumented GNU Time ABBA processes used 100 warmups and 10,000 samples per
leg. Maximum RSS was 30,976/30,848 KiB before and 30,848/30,976 KiB after: the
30,912 KiB state means are exact.

Matched process-wide `perf stat` runs at the same sample count give:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 15,426.4 ms | 14,075.6 ms | -8.76% |
| cycles | 75.977 billion | 69.129 billion | -9.01% |
| instructions | 232.190 billion | 216.510 billion | -6.75% |
| branches | 41.913 billion | 38.888 billion | -7.22% |
| branch misses | 420.083 million | 367.249 million | -12.58% |
| cache references | 6.731 billion | 6.231 billion | -7.43% |
| cache misses | 634.705 million | 578.468 million | -8.86% |
| page faults | 648,986 | 112,097 | -82.73% |
| context switches | 257 | 187 | -27.24% |
| CPU migrations | 2 | 2 | unchanged |

Whole-process counter deltas are smaller than scoped latency because corpus
construction and complete untimed verification remain in the profiled process.
Sampled call graphs retain the expected CFB and document-structure frames on
the control while the candidate no longer has the second root-open owner.

## Correctness and validation

Focused tests prove that adopted default-limit snapshots equal a complete root
reopen in exact bytes, document structure, document persist identity,
review-history flag, and limits; unrelated CFB streams remain exact. Separate
tests reject foreign source bytes, a wrong slide persist ID, and the live
document persist ID. Consecutive text edits and text-plus-anchor-plus-reorder
composition reopen exactly. A custom-limit test proves that the conservative
full-revalidation fallback remains active.

The final source passes:

- the complete `litchi-ppt --all-features` suite: 1,064 unit tests, every
  integration target, and all doc tests;
- warning-denied all-target/all-feature PPT Clippy and warning-denied rustdoc;
- all 34 performance-harness tests and warning-denied all-target Clippy;
- warning-denied all-target/all-feature `litchi-odf-common` Clippy and rustdoc,
  revalidating the deprecation cleanup committed in `1194fbc7f`;
- formatter, JSON parsing, whitespace, and staged-scope checks.

This tranche changes only native PPT/OLE2 production code. OOXML, RTF, ODF,
and every iWork/IWA crate are unchanged.
