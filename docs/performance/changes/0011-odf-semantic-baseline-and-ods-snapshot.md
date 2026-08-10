# Change 0011: ODF semantic baseline and ODS snapshot parse reuse

Date: 2026-08-10

## Decision

Accept a narrow ODS snapshot optimization and add an opt-in public-API ODF
semantic matrix. `document::Snapshot::from_bytes` now constructs one validated
`Package`, checks its resource count, and moves that same package into the
`Spreadsheet` facade for complete semantic readback. The previous path cloned
the package bytes, parsed the archive/package twice, and discarded the first
parsed package.

The exact `Arc<[u8]>` snapshot source, size/resource bounds, full facade
readback, immutable edit owner, exact no-op publication, and changed-output
reopen contracts are unchanged. The only newly visible method is
`Spreadsheet::from_package`, and it is crate-private; no archive or package
type crosses the public facade.

## Benchmark coverage

The harness adds 21 opt-in cases: seven each for ODT, ODS, and ODP. Each family
covers owned open, semantic listing, one semantic object, complete text, small
creation, exact no-op edit/save, and one supported semantic edit/save. Creation
runs only for the tiny shape. All edits publish through the format owner's
public transaction and reopen the resulting bytes for complete semantic
validation.

The deterministic shapes are:

| Shape | ODT paragraphs | ODS sheets x rows x columns | ODP slides |
|---|---:|---:|---:|
| tiny | 24 | 1 x 8 x 8 | 3 |
| medium | 200 | 2 x 32 x 32 | 12 |
| large | 10,000 | 2 x 128 x 128 | 100 |

ODP opened slides are preservation-only under rewrite, so its supported
one-edit case appends one authored slide and verifies every retained slide plus
the addition. The ODP builder records wall-clock metadata; the corpus retains
its public authored content/styles but republishes fixed `meta.xml` so hashes
are stable. The default matrix remains 36 cases / 198 records; the harness now
has 81 selectable cases.

## Matched latency result

The release executables were frozen on production base `01e93c778` with the
same completed harness:

- before SHA-256:
  `8bcc78531cb27ba2d1995393bf01f61322340295cb8cfc3c65e20816d49d3beb`
- after SHA-256:
  `e262b4417847d7b855029f40080dc8a01a19431e668af13f196874447a37de28`

Both states ran pinned to CPU 2 in before-A, after-A, after-B, before-B order,
with three warmups and 15 measured samples per leg. The table pools both legs
(30 samples per state). Times are milliseconds; mean intervals are two-sided
Student's-t 95% intervals.

The raw reports mark the worktree dirty because the completed harness and
performance documents were not yet committed and an unrelated pre-existing
documentation edit was present. The before executable was frozen before either
`litchi-ods` production source file changed.

| Case | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| ODS no-op edit/save, medium | 4.769 / 5.359 / 5.418 | 4.414 / 4.658 / 5.341 | **-7.45%** | 4.830 (4.753-4.907) | 4.413 (4.327-4.498) | **-8.64%** |
| ODS no-op edit/save, large | 76.894 / 84.315 / 88.138 | 67.838 / 70.810 / 73.001 | **-11.78%** | 77.246 (76.069-78.423) | 67.915 (67.321-68.509) | **-12.08%** |
| ODS one-cell edit/save, medium | 24.015 / 25.814 / 26.800 | 23.159 / 24.228 / 24.685 | **-3.57%** | 24.122 (23.825-24.419) | 23.247 (23.033-23.461) | **-3.63%** |
| ODS one-cell edit/save, large | 384.150 / 403.866 / 412.984 | 376.237 / 392.146 / 395.864 | **-2.06%** | 385.206 (381.244-389.167) | 376.769 (373.494-380.044) | **-2.19%** |

The no-op cell isolates snapshot construction most clearly. Changed
publication still dominates the one-cell cases because it rewrites and
revalidates the edited spreadsheet package.

Raw samples:
[`before A`](../results/abba-odf-before-a.json),
[`after A`](../results/abba-odf-after-a.json),
[`after B`](../results/abba-odf-after-b.json), and
[`before B`](../results/abba-odf-before-b.json).

## Allocation and memory result

Whole-process Valgrind Memcheck on one medium no-op sample reports allocated
bytes falling from 23,097,897 to 22,760,212 (**-1.46%**). Allocation calls rose
from 211,963 to 215,793 (+1.81%), so this change is not claimed as an
allocation-count win. The source-level package-sized clone and duplicate parse
are removed, but ownership/lifetime differences change the allocator call
shape.

Heaptrack over five large no-op samples reports flat 41.70 MB peak heap and
46.87 MB versus 46.81 MB profiler RSS (-0.13%). A reverse-order GNU Time
sample reports 42,548 KiB versus 42,416 KiB maximum RSS (-0.31%). These are
whole-process figures dominated by corpus construction and complete semantic
verification; they establish no peak-memory regression, not a per-snapshot
peak reduction.

## Correctness and contract gates

- The complete all-feature `litchi-ods` suite passes: 126 unit tests and every
  integration suite, including exact no-op, source-checked patch, inverse,
  signed/protected refusal, resource limits, malformed input, and real Calc
  reopen coverage.
- The performance harness passes 20 tests, formatting, warning-denied Clippy,
  and a release 21-record tiny ODF smoke. The ODS all-feature library
  warning-denied Clippy gate passes. Its all-target gate remains blocked by
  pre-existing test-only lints (`items_after_test_module`, `useless_format`,
  and `cloned_ref_to_slice_refs`) and is not reported as passing. The harness
  default count test remains 36 cases / 198 records.
- CI watches all four ODF crates, runs the 21-record tiny smoke on pushes and
  pull requests, and publishes a 39-record tiny/large ODF matrix on scheduled
  or manual runs.
- No public archive type, dependency inversion, unsafe code, ambient input,
  hidden runtime, global lock, or iWork/IWA change is introduced.

## Remaining limitations

ODT and ODP are benchmark coverage only in this tranche. Ordinary ODF open is
still eager; repeated ODT/ODP semantic queries can rescan complete XML; and
changed ODF publication still recompresses unchanged package members. ODS
one-cell publication remains dominated by complete rewrite/readback.

The generated text/grid/deck corpora do not replace real-producer, media,
unknown-extension, malformed, signed, encrypted, external-link, cold-source,
conversion/export, or broad edit/patch matrices. iWork is deliberately
deferred while its `iwa-*` crates are being modified independently.
