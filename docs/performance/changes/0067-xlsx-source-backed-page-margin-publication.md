# Change 0067: source-backed XLSX page-margin publication

Date: 2026-08-12

Production base: `59d1a17d85df086e90d7b0fd8cfb18267db106ed`

Status: accepted

## Hypothesis and implementation

Direct worksheet page margins are a safe one-Part XLSX mutation. The existing
typed owner validates all six physical values, changes only the selected
worksheet XML, and does not alter cells, formulas, calculation policy, styles,
shared strings, relationships, or package topology. The hypothesis was that
publishing that exact edit through the established OPC one-Part overlay would
avoid materializing and recompressing unrelated drawings and media.

`page_margins::SourceBackedEditor` is an additive, non-cloneable editor over a
caller-provided `Arc<dyn ReadAt>`. A source edit loads the workbook catalog and
one selected normal worksheet, stages `Option<Margins>`, exposes only `set`
and `remove`, uses the existing MCE-aware byte-minimal worksheet rewriter, and
reparses the complete changed worksheet before producing an exact reversible
`Commit` and `Patch`. Publication recaptures the exact closure, then consumes
`SourceBackedPackage::write_part_overlay_to_stream`.

The retained identity includes the package `officeDocument` relationship,
workbook URI/content type/XML, selected workbook relationship
ID/type/target/mode, worksheet URI/content type/XML, checked sheet name and
position, and typed margins. A retargeted relationship cannot apply a patch to
an orphaned worksheet. Unrelated relationships remain outside the closure.

Exact semantic no-ops reproduce the complete source archive, including signed
input and source lexical `-0`. A changed margin canonicalizes normalized zero.
Changed signed input, stale or foreign closures, retargeted relationships,
chartsheets, projected MCE state, read limits, unsupported ZIP layouts, and
partial sink failure retain typed refusals. The MCE reader now correctly
ignores an inherited default namespace declaration on a projected
`pageMargins` element, while the rewriter still refuses to mutate projected
state. No dependency, runtime, unsafe code, archive abstraction, ordinary
workbook capability, or copy/move operation was added.

## Matched corpus and protocol

Both controls use the deterministic media-rich XLSX archive under a
page-margin-specific manifest. It contains 12 ordinary Parts and 17 ZIP
members: one workbook, one worksheet, one drawing, eight referenced
incompressible 2 MiB PNG Parts, and minimal supporting Parts. The archive is
16,786,830 bytes with 16,782,412 logical Part bytes and SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths set `Sheet1` margins to left/right/top/bottom/header/footer values
0.7/0.8/1.0/1.1/0.3/0.4 inches. The eager control performs positional open,
`into_opc_package`, the existing ordinary workbook transaction, and the
ordinary sequential writer. The candidate performs positional open, the new
source edit, and one-Part overlay publication. Both emit the same
16,786,883-byte artifact with SHA-256
`2b50a470f5066dd078c0566ca9e203f26843026b58b6a7a6eca14ac8435429a7`.

The two public paths are measured from the same frozen release binary, SHA-256
`397121615aba9308a2f5fb3f9126ab5436a00bb3932945aaab276f7523c4b7b1`;
the unchanged eager path is the matched before-state control. Corpus creation,
expected-output construction, complete XLSX reopen, page-margin and
calculation-metadata readback, Part/content-type/relationship comparison,
exact media checks, source/sink inspection, and hashing stay outside timing.

The retained sequence on CPU 2 was eager A, source A, source B, eager B,
followed by reverse source C, eager C, eager D, source D. Each leg used 50
warmups and 500 samples, yielding 2,000 samples per state. Exact values are in
the [`measurement summary`](../results/xlsx-page-margin-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 2,000 | 2,000 | — |
| p50 | 216.798932 ms | 4.492344 ms | **-97.93% (48.26x)** |
| mean | 217.042560 ms | 4.495567 ms | **-97.93% (48.28x)** |
| mean 95% interval | 216.917354-217.167765 ms | 4.483745-4.507390 ms | disjoint |
| p95 | 221.698139 ms | 4.942350 ms | **-97.77% (44.86x)** |
| p99 | 225.958358 ms | 5.237890 ms | **-97.68% (43.14x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,786,883 | 16,786,883 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

The two required materializations are `xl/workbook.xml`, used to resolve and
bind `Sheet1`, and `xl/worksheets/sheet1.xml`, used for the semantic edit. The
other ten ordinary Parts are raw-copied. Physical source reads remain
archive-sized because untouched compressed members still have to reach the
sequential sink.

## Allocation, counters, and memory

One-sample Heaptrack attribution reports allocation calls 14,277 -> 12,550
(-12.10%) and temporary allocations 1,715 -> 1,515 (-11.66%). Peak heap is
flat at 152.84 -> 152.81 MiB. Heaptrack peak RSS rises 1.38% under
instrumentation, while four uninstrumented GNU Time processes per state report
mean maximum RSS 142,219 -> 142,157 KiB (-0.04%, flat). Both sides retain the
same 1.78 KiB profiler/runtime leak.

Two matched process-wide `perf stat` repeats per state used two warmups and ten
samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 3,948.36 ms | 1,041.61 ms | -73.62% |
| cycles | 19.376 billion | 5.130 billion | -73.52% |
| instructions | 48.538 billion | 10.378 billion | -78.62% |
| branches | 8.232 billion | 1.338 billion | -83.74% |
| branch misses | 175.975 million | 14.161 million | -91.95% |
| cache references | 1.382 billion | 403.662 million | -70.78% |
| cache misses | 26.802 million | 19.397 million | -27.63% |
| page faults | 316,160 | 193,037 | -38.94% |

CPU migrations were zero. These process-wide reductions are smaller than the
scoped latency win because deterministic corpus creation and complete untimed
verification remain inside the profiled process.

## Validation and limitations

Passed on the final source:

- complete `litchi-xlsx --all-features` suite: 732 unit tests plus every
  integration target and doctest;
- six focused source-backed page-margin integration cases covering add,
  replace, remove, exact no-op, signed zero, changed signed refusal,
  patch/inverse, stale/foreign/source-version conflicts, relationship
  retargeting, chartsheet refusal, MCE projection, limits, unselected
  Part/media identity, and partial sinks;
- warning-denied XLSX library and focused-integration Clippy;
- complete performance-harness tests and warning-denied all-target Clippy;
- deterministic debug/release smoke output plus fixed CI archive/output hashes,
  materialization counts, source reads, and sink bounds;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation cleanup from `1194fbc7f`;
- formatter, JSON/YAML parsing, and whitespace checks.

Broad XLSX all-target Clippy and public rustdoc retain the pre-existing
unrelated module-inception, needless-question-mark, and private/broken
intra-doc-link findings recorded by change 0046. The focused changed targets
are warning-clean.

This capability does not claim general source-backed XLSX editing. Cell
values, formulas, styles, shared strings, page-break collections,
calculation-chain changes, sheet topology, relationships, and content types
remain outside its one-Part closure. OLE2, RTF, ODF production code, and every
iWork/IWA crate are unchanged by this batch.
