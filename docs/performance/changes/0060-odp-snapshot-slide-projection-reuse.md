# Change 0060: ODP snapshot slide-projection reuse

Date: 2026-08-12

Production base: `442d46c5d200c3bf1af58dc46fa4cf4c3e3b7240`

Status: accepted

## Hypothesis and attribution

An ODP editing `Snapshot` already retained the complete validated slide
projection in `Arc<[Slide]>`. Starting a transaction nevertheless reopened the
same immutable package and called `Presentation::slides()` again while building
the detached mutable draft. That second call repeated transition-style
resolution and a complete namespace-aware `content.xml` traversal before any
edit could be staged.

The existing `odp_semantic_noop_edit_save` case constructs the presentation and
editing snapshot before timing, then times transaction creation, commit, and
owned output materialization. For an exact no-op there is no changed candidate
parse, making the duplicate staging traversal directly observable. The
hypothesis was that reusing the source-bound snapshot projection would remove
that work while retaining the independent package, security, preservation, and
publication checks.

## Implementation and proof boundary

`Snapshot::transaction` still reopens the exact shared package bytes as a
`Presentation` and applies the existing signature/encryption policy before
staging. It now passes its already validated slide slice to the private
`MutablePresentation` constructor instead of asking that presentation to parse
the same slides again.

The private constructor still:

- parses settings, declarations, page metadata, MIME type, styles, and the
  lossless `ContentSource` independently;
- requires raw source-page coverage to equal the supplied slide count and
  preserves the existing typed refusal on incomplete coverage;
- clones the supplied values into both detached mutable slides and immutable
  source-slide comparison state, so edits cannot mutate the source projection;
- retains exact source-package bytes and raw page fragments for untouched-page
  preservation; and
- leaves candidate generation, compact XML audit, package publication, final
  snapshot construction, complete semantic reopen/readback, reversible patch,
  inverse, limits, signatures, encryption, media, and manifest checks unchanged.

The handoff is private and has one caller. A snapshot's bytes and slide
projection are created together from the same immutable `Arc<Vec<u8>>`; no
public caller can combine a projection with another package. Focused tests prove
draft/source clone isolation and that incomplete page coverage still refuses.
No public API, dependency, cache, runtime, lock, executor, unsafe code, archive
type, output format, or validation policy changed.

## Measurement method

Environment: release Rust 1.95.0 / LLVM 22.1.2, Linux
6.8.0-101-generic, x86-64 AMD EPYC 9575F VM, Rust system allocator, CPU 11
pinned with `taskset`. The fixed large corpus has 100 slides, 8,700 logical
payload bytes, five ZIP members, a 3,424-byte archive, and SHA-256
`afb69ac66dffbc9f3ef19db360161af636abb818bfce689c6f4964fc521778c6`.
Every no-op sample verifies exact source bytes, unchanged commit state, complete
ODP reopen, and the full slide projection outside the timed interval.

The control executable SHA-256 is
`955dad8ec0d12103b63dc677a2eeab8c9c61351f285225a7fcd6fa4289d9cee1`;
the candidate is
`22175370a69e22986e3ac2c1e813ae54524da27b1e51b2be4baa080dba33b911`.
The exact identities and `.text` hashes are recorded in
[`odp-slide-projection-sha256.txt`](../results/odp-slide-projection-sha256.txt).

The primary used one ABBA and one reverse BAAB sequence. Each leg had 50
warmups and 1,000 measured samples, yielding 4,000 observations per state.
Pooled samples and confidence intervals are in the
[`primary summary`](../results/odp-slide-projection-primary-summary.json).

## Latency result

| Large exact no-op edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 1.728214 ms | 0.692022 ms | **-59.96% (2.50x)** |
| mean | 1.748348 ms | 0.700785 ms | **-59.92%** |
| p95 | 1.864988 ms | 0.756270 ms | **-59.45%** |
| p99 | 2.066914 ms | 0.867494 ms | **-58.03%** |

The before mean 95% interval is 1.745502-1.751194 ms; the after interval is
0.699393-0.702177 ms. Every individual primary leg improves by more than 58%
at p50 and mean.

The same handoff also helps real changes because draft construction precedes
every edit. The large one-slide append/save guard uses 1,000 samples per state:

| Large one-edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 4.499871 ms | 3.564817 ms | **-20.78%** |
| mean | 4.542664 ms | 3.597516 ms | **-20.81%** |
| p95 | 4.764500 ms | 3.801051 ms | **-20.22%** |
| p99 | 5.055383 ms | 4.180198 ms | **-17.31%** |

## Scaling and guardrails

The smaller exact-no-op shapes improve as the redundant traversal grows:

| Shape | Slides | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|---:|
| tiny | 3 | 122.274 us | 85.243 us | **-30.29%** | -31.76% | -34.68% |
| medium | 12 | 269.029 us | 144.470 us | **-46.30%** | -46.54% | -45.95% |

Large read-only controls do not call `Snapshot::transaction` and remain neutral
at their central metrics:

| Unchanged case | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| open | -0.85% | -2.39% | -6.46% | -11.17% |
| list slides | -0.92% | -0.11% | +2.29% | +35.33% |
| one slide | +1.10% | +0.85% | -0.64% | -2.23% |
| full text | +0.51% | +1.01% | +3.53% | +9.55% |

The list/full-text p99 tails triggered review. The movement is concentrated in
one candidate leg: the other candidate list leg is faster than both control
legs at p95 and p99, while p50/mean and the pooled p95 remain within 3.6% for
both cases. Neither read path reaches the changed constructor, and the
`Presentation::{slides,slide,text}` symbol sizes remain exactly 45, 39, and 816
bytes in both executables. The tail is retained as linked-executable/scheduler
variance; no unrelated read-path workaround was added.

The fixed 16 MiB media-rich text-box edit/save control is neutral at p50
(-0.39%) and improves mean/p95 by 1.76%/3.29%. Its package-level source-backed
publication cost dominates the saved small slide traversal. Complete guard
distributions are in the
[`guard summary`](../results/odp-slide-projection-guard-summary.json).

## Allocation, counters, and memory

Heaptrack used two warmups and 20 measured large no-op commits per process. It
profiles deterministic corpus construction and complete post-timing
verification as well as the timed operation:

| Heaptrack process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 915,236 | 731,008 | **-20.13%** |
| temporary allocations | 393,448 | 305,052 | **-22.47%** |
| peak heap | 696.20 KiB | 696.20 KiB | flat |
| profiler peak RSS | 13.06 MiB | 12.61 MiB | -3.45% |

The identical 1.78 KiB profiler/runtime leak remains. Four uninstrumented GNU
Time processes per state report mean maximum RSS of 30,912 KiB before and
30,880 KiB after (-0.10%, flat).

Matched process-wide `perf stat` ABBA legs used 20 warmups and 500 samples:

| Counter, mean/process | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 2,704.810 ms | 2,105.925 ms | **-22.14%** |
| cycles | 13.269 billion | 10.395 billion | **-21.65%** |
| instructions | 61.347 billion | 48.687 billion | **-20.64%** |
| branches | 14.670 billion | 11.589 billion | **-21.00%** |
| branch misses | 33.892 million | 26.354 million | **-22.24%** |
| cache references | 252.867 million | 206.378 million | -18.39% |
| cache misses | 9.257 million | 5.706 million | **-38.36%** |
| page faults | 8,646 | 8,648 | +0.02% |

CPU migrations are zero in all four counter legs. The process-wide reductions
are smaller than the scoped latency win because corpus creation, initial
snapshot parsing, and complete semantic verification remain outside the timer
but inside the profiled process. Matched profiles retain zero lost samples.

## Validation

Passed on the final source:

- focused validated-projection isolation and incomplete-coverage tests;
- complete `litchi-odp --all-features` suite: 125 unit tests, all integration
  targets, and 21 doc tests;
- warning-denied all-target/all-feature ODP Clippy and warning-denied rustdoc;
- all 33 release performance-harness tests and warning-denied all-target
  Clippy;
- warning-denied ODF-common all-target Clippy and rustdoc, revalidating the
  deprecation cleanup from `1194fbc7f`;
- formatter, JSON parsing, and diff/whitespace checks.

There is no dedicated ODP fuzz manifest in the repository. The workspace-wide
iWork/IWA gate was not run because those crates remain explicitly excluded
while other agents modify them.

## Limitations and next work

This change removes only the duplicate complete slide parse used to seed a
transaction draft. It does not cache parsed settings/declarations/page
metadata, skip package/security reopening, adopt a changed candidate, weaken
complete final reopen/readback, or add positional ODF source I/O. Repeated
independent read queries, structural/resource publication, real-producer media,
and broader malformed/security performance matrices remain open. OLE2, OOXML,
RTF, and every iWork/IWA crate are unchanged by this batch.
