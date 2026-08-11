# Change 0065: ODP final slide-snapshot handoff

Date: 2026-08-12

Production base: `3cd4ff394ae05d1aa5cfca46e77b535c1e6f3412`

Status: accepted

## Hypothesis and attribution

A changed ODP slide transaction already constructed a complete `Snapshot` from
the serialized candidate so it could compare the parsed slides with the staged
draft. After all publication checks, an exact slide-only commit parsed the same
immutable bytes into another `Snapshot`. The second projection repeated the
complete styled slide traversal, slide allocation and resource-byte count.

The existing `odp_semantic_one_edit_save` case is the exact end-to-end target:
it starts from an already-open editing snapshot, appends one slide, commits,
and materializes the owned output inside the timed interval. Complete reopened
slide verification remains outside the timer.

## Implementation and proof boundary

`Transaction::commit` retains the already mandatory slide candidate through
the rest of publication. It adopts that candidate only when the transaction is
exactly slide-only: the slide draft changed and the RDF, chart, design,
annotation and rich-content operation sets are all empty. An internal pointer
assertion documents that the retained candidate and final package refer to the
same immutable byte allocation.

The handoff happens only at final snapshot construction. The transaction still:

- serializes the complete staged slide model and parses it back before any
  later domain operation;
- compares the parsed slides with the staged draft;
- opens the final package independently;
- validates raw referenced XML preservation and compact generated XML;
- reopens every staged embedded-media Part and checks its manifest entry;
- performs every applicable RDF, chart, design and annotation readback;
- builds the same source-checked patch and inverse; and
- leaves exact no-op and every compound-domain commit on their existing paths.

The final `OwnedPackage` reopen is deliberately retained because the compact,
media and auxiliary-domain checks consume it. Compound commits retain both the
early slide readback and the ordinary final snapshot parse, so error order and
the unified publication boundary do not change. No public API, dependency,
runtime, cache, lock, executor, unsafe code, archive type or format changed.

## Measurement method

The fixed large corpus has 100 slides, five ZIP members, 8,700 logical text
bytes and a 3,424-byte archive. Its SHA-256 is
`afb69ac66dffbc9f3ef19db360161af636abb818bfce689c6f4964fc521778c6`.
Runs used release Rust 1.95.0 / LLVM 22.1.2, Linux 6.8.0-101-generic,
the Rust system allocator, an AMD EPYC 9575F VM and CPU 11 pinned by
`taskset`.

The control executable SHA-256 is
`748ae86dc8d8981c04d95c9228f80b798e7d8677b36bc2fdbc76d146272926db`;
the candidate is
`48ec19374e8ffd47fa6001cc25fdc451ad68cf93b4e48542474d030436d2bbba`.
Their `.text` SHA-256 values are recorded in the
[`summary`](../results/odp-final-snapshot/summary.json).

Two initial balanced cycles were retained as discarded evidence because one
or both states exceeded the predeclared 5% within-state central-metric drift
gate. The accepted warmed ABBA cycle used 100 warmups and 1,000 samples per leg
in `before-e`, `after-e`, `after-f`, `before-f` order. Its two legs per state
pool 2,000 observations. The control legs differ 2.12% at p50 and 3.59% at
mean; candidate legs differ 0.54% and 0.38%.

## Latency result

| Large one-slide edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 3.572775 ms | 2.417132 ms | **-32.35%** |
| mean | 3.637947 ms | 2.440153 ms | **-32.92%** |
| p95 | 3.989373 ms | 2.555322 ms | **-35.95%** |
| p99 | 4.358681 ms | 2.786120 ms | **-36.08%** |

The before mean 95% interval is 3.628132-3.647762 ms; the after interval is
2.435363-2.444944 ms. Both individual candidate legs are more than 32% faster
than the corresponding control center.

The same removed final projection scales materially on smaller decks:

| Shape | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|
| tiny | **-21.14%** | -22.79% | -32.76% |
| medium | **-25.68%** | -26.60% | -29.13% |

## Guardrails

The exact-no-op commit is ineligible and improves 1.54% p50/mean. The fixed
16 MiB media-rich content-domain text-box edit is also ineligible and improves
1.81% p50, 1.59% mean and 0.68% p95.

The first read-only pool had neutral central metrics but noisy p99 tails, so a
separate reverse-order 2,000-sample/state repeat was retained. Its results are:

| Unchanged read case | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|
| open | -1.62% | -0.72% | +2.99% | +13.55% |
| list slides | -0.43% | +0.12% | -0.47% | +18.53% |
| one slide | -0.69% | -1.28% | -3.97% | -11.71% |
| full text | +0.09% | +0.25% | -1.17% | +0.41% |

The open/list p99 movements reproduce only in the extreme tail: their p50,
mean and p95 remain within 3%, and no read-only path reaches the changed commit
handoff. The candidate and control `Presentation::{slide,slides,text}` symbols
have identical sizes (39, 45 and 816 bytes). As in change 0060, the tails are
retained as linked-executable/host variance; no unrelated read-path change was
added.

## Allocation, memory and counters

Heaptrack covered two warmups and 20 complete large changed commits per
process, including deterministic construction and untimed semantic
verification:

| Process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 1,210,168 | 1,007,922 | **-16.71%** |
| temporary allocations | 510,026 | 419,298 | **-17.79%** |
| peak heap | 947.20 KiB | 947.20 KiB | flat |
| profiler peak RSS | 13.58 MiB | 13.46 MiB | -0.88% |
| leaked memory | 1.78 KiB | 1.78 KiB | flat |

Four uninstrumented GNU Time processes per state report mean maximum RSS of
30,880 KiB before and 30,816 KiB after (-0.21%); every process has zero major
faults.

Matched `perf stat` ABBA processes used 20 warmups and 500 samples. Process-wide
task clock falls 15.88%, cycles 15.49%, instructions 16.25% and branches
16.54%; all four legs have zero CPU migrations. These deltas are smaller than
the scoped latency result because corpus construction and complete verification
remain inside the profiled process.

## Validation

Passed on the final source:

- focused slide-only publication, compact/raw XML, staged-media and unified
  slide-plus-auxiliary-domain transaction tests;
- complete debug and release `litchi-odp --all-features` suites;
- warning-denied all-target/all-feature ODP Clippy and warning-denied rustdoc;
- complete performance-harness tests and warning-denied harness Clippy;
- warning-denied ODF-common all-target Clippy and rustdoc, revalidating the
  deprecation cleanup from `1194fbc7f`;
- formatter, JSON parsing and diff/whitespace checks.

There is no dedicated ODP fuzz manifest in the repository. The workspace-wide
iWork/IWA gate was not run because those crates remain explicitly excluded
while other agents modify them.

## Limitations and next work

This change removes only the redundant final slide projection for exact
slide-only commits. It does not skip the final package reopen, accelerate
compound-domain commits, add parsed-package retention, change raw ZIP I/O, or
broaden CRUD. Repeated independent semantic scans, structural/resource edits,
real-producer media, malformed/security matrices and positional ODF source I/O
remain open. OLE2, OOXML, RTF and every iWork/IWA crate are unchanged by this
batch.
