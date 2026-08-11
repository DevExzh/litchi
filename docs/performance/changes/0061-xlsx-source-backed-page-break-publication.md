# Change 0061: source-backed XLSX page-break publication

Date: 2026-08-12

Production base: `eefad744a1e626a9abbf3ac12d5d11b61fb2e46f`

Status: accepted

## Hypothesis and implementation

Worksheet page breaks are a safe one-Part XLSX mutation: the existing typed
owner changes only the selected worksheet XML and does not alter formulas,
calculation policy, shared strings, styles, relationships, or package topology.
The hypothesis was that publishing this edit through the established OPC
one-Part overlay would avoid materializing and recompressing unrelated
drawings and media.

`page_breaks::SourceBackedEditor` is an additive, non-cloneable editor over a
caller-provided `Arc<dyn ReadAt>`. A source edit loads the workbook catalog and
one selected normal worksheet, stages the existing `PageBreaks` model, uses the
existing MCE-aware `replace` rewriter, and reparses the complete changed
worksheet before producing the ordinary exact reversible `Commit` and `Patch`.
Publication recaptures the exact closure, then consumes
`SourceBackedPackage::write_part_overlay_to_stream`.

The retained source identity is stronger for both owning and source-backed
patches: package `officeDocument` relationship, workbook URI/content type/XML,
selected workbook relationship ID/type/target/mode, and worksheet
URI/content type/XML. A retargeted relationship can no longer apply a patch to
an orphaned worksheet. Unrelated relationships remain outside the exact
closure.

Exact no-ops reproduce the complete source archive, including signed input.
Changed signed input, stale or foreign closures, retargeted relationships,
non-worksheet selections, projected MCE state, read limits, unsupported ZIP
layouts, and partial sink failure retain typed refusals. No dependency,
runtime, unsafe code, archive abstraction, or ordinary facade capability was
added.

## Matched corpus and protocol

Both controls use the existing deterministic media-rich archive, exposed with
a page-break-specific manifest. It contains 12 ordinary Parts and 17 ZIP
members: one workbook, one worksheet, one drawing, eight referenced
incompressible 2 MiB PNG Parts, and minimal supporting Parts. The archive is
16,786,830 bytes with 16,782,412 logical Part bytes and SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths add one manual horizontal break at row 100 on `Sheet1`. The eager
control performs positional open, `into_opc_package`, the existing owning
page-break transaction, and the ordinary sequential writer. The candidate
performs positional open, the new source edit, and one-Part overlay
publication. Both emit the same 16,786,878-byte artifact with SHA-256
`1e3b7a9f763feaed4ad4888aa8aa0cd3773cdb9fd9f12e16f3c05b7fd0cd95b3`.

Corpus creation, expected-output construction, complete XLSX reopen,
page-break and calculation-metadata readback, Part/content-type/relationship
comparison, exact media checks, source/sink inspection, and hashing stay
outside timing. The retained balanced sequence was eager A, source A, source
B, eager B on CPU 2. Each leg used 20 warmups and 100 samples. P50 and mean
drift remain below 1.2% within both states. Exact values are retained in the
[`measurement summary`](../results/xlsx-page-break-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | — |
| p50 | 216.789359 ms | 4.647479 ms | **-97.86% (46.65x)** |
| mean | 217.515955 ms | 4.649883 ms | **-97.86% (46.78x)** |
| mean 95% interval | 217.103377-217.928533 ms | 4.621163-4.678604 ms | disjoint |
| p95 | 222.424512 ms | 4.937780 ms | **-97.78% (45.05x)** |
| p99 | 227.706846 ms | 5.100022 ms | **-97.76% (44.65x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,786,878 | 16,786,878 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

The two required materializations are `xl/workbook.xml`, used to resolve and
bind `Sheet1`, and `xl/worksheets/sheet1.xml`, used for the semantic edit. The
other ten ordinary Parts are raw-copied. Physical source reads remain
archive-sized because untouched compressed members still have to reach the
sequential sink.

## Allocation, counters, and memory

One-sample Heaptrack attribution reports allocation calls 15,054 -> 12,653
(-15.95%) and temporary allocations 1,700 -> 1,530 (-10.00%). Peak heap is
flat at 152.83 -> 152.80 MiB. Heaptrack peak RSS rises 2.62% under
instrumentation, while four uninstrumented GNU Time processes per state report
mean maximum RSS 142,435 -> 142,587 KiB (+0.11%, flat). Both sides retain the
same 1.78 KiB profiler/runtime leak.

Matched process-wide `perf stat` runs used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 3,958.54 ms | 1,089.88 ms | -72.47% |
| cycles | 19.442 billion | 5.416 billion | -72.14% |
| instructions | 48.478 billion | 10.469 billion | -78.40% |
| branches | 8.222 billion | 1.354 billion | -83.53% |
| branch misses | 176.209 million | 14.276 million | -91.90% |
| cache references | 1.379 billion | 410.339 million | -70.24% |
| cache misses | 26.744 million | 21.577 million | -19.32% |
| page faults | 306,161 | 207,834 | -32.12% |

CPU migrations were zero. These process-wide reductions are smaller than the
scoped latency win because deterministic corpus creation and complete untimed
verification remain inside the profiled process.

## Validation and limitations

Passed on the final source:

- complete `litchi-xlsx --all-features` suite: 732 unit tests, every integration
  target including the four source-backed page-break cases, and two doc tests;
- warning-denied XLSX library Clippy and focused source-backed integration
  Clippy;
- all 34 performance-harness tests and warning-denied all-target Clippy;
- deterministic debug and release smoke output, fixed CI source/output hashes,
  materialization counts, source reads, and sink bounds;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation cleanup from `1194fbc7f`;
- formatter, JSON/YAML parsing, and whitespace checks.

The broad XLSX all-target Clippy and public rustdoc commands remain blocked by
the same pre-existing unrelated module-inception, needless-question-mark, and
private/broken intra-doc-link findings recorded by change 0046. The focused
changed targets are warning-clean.

This capability does not claim general source-backed XLSX editing. Cell values,
formulas, styles, shared strings, calculation-chain changes, sheet topology,
relationships, and content types remain outside its one-Part closure. OLE2,
RTF, ODF production code, and every iWork/IWA crate are unchanged by this
batch.
