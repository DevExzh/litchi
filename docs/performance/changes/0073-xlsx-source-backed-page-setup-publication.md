# Change 0073: source-backed XLSX page-setup publication

Date: 2026-08-12

Production base: `0b5def79c63360dfd0a034b76064265bef351c94`

Status: accepted

## Hypothesis and implementation

Relationship-free worksheet `pageSetup` settings are a safe one-Part XLSX
mutation. The existing bounded typed codec changes only the selected worksheet
XML and does not alter cells, formulas, calculation policy, styles, shared
strings, relationships, printer resources, or package topology. The hypothesis
was that publishing this edit through the accepted OPC one-Part overlay would
avoid materializing and recompressing unrelated drawings and media.

`page_setup::SourceBackedEditor` is an additive, non-cloneable editor over a
caller-provided `Arc<dyn ReadAt>`. It loads the workbook catalog and one normal
worksheet, stages `Option<Setup>`, uses the existing direct-property rewriter,
and reparses the complete changed worksheet before creating an exact reversible
`Commit` and `Patch`. Publication recaptures the retained source closure and
consumes `SourceBackedPackage::write_part_overlay_to_stream`.

The retained identity includes the package `officeDocument` relationship,
workbook URI/content type/XML, selected workbook relationship
ID/type/target/mode, worksheet URI/content type/XML, checked sheet name and
position, and the complete selected-worksheet outbound relationship set. A
retargeted drawing or other worksheet relationship therefore conflicts rather
than authorizing a changed XML payload against a different closure.

The facade refuses any `pageSetup r:id` and any Transitional or Strict
printer-settings relationship. Supporting printer settings safely would
require a wider closure that validates and retains the relationship plus the
inert DEVMODE Part. Exact semantic no-ops reproduce the complete source
archive, including signed input. Changed signed input, stale sources,
retargeted relationships, chartsheets, projected MCE state, read limits,
unsupported ZIP layouts, and partial sinks retain typed refusals. The page
setup MCE preprocessing limits are now explicitly aligned with the codec's
32 MiB input/output and depth-128 policy.

No dependency, runtime, unsafe code, global cache, archive abstraction,
ordinary workbook capability, relationship edit, or topology operation was
added.

## Matched corpus and protocol

Both controls use the deterministic media-rich XLSX archive under a
page-setup-specific manifest. It contains 12 ordinary Parts and 17 ZIP
members: one workbook, one worksheet, one drawing, eight referenced
incompressible 2 MiB PNG Parts, and minimal supporting Parts. The archive is
16,786,830 bytes with 16,782,412 logical Part bytes and SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths insert A4 paper, landscape orientation, 85% scale,
over-then-down order, and an explicit false printer-default policy on
`Sheet1`. The eager control performs positional open, eager OPC ownership, the
ordinary workbook transaction, and ordinary sequential publication. The
candidate performs positional open, the new source edit, and one-Part overlay
publication. Both emit the same 16,786,902-byte artifact with SHA-256
`fd866304c7aab42bcd5195e38bef5dd76f6192b05bcb551786a2096ad898fcb1`.

The two paths were measured from frozen release binary SHA-256
`056208e98dbcb0d18d9be891f4bc736a80a8b3adbf13dbd0210e369b63d0dbce`.
Corpus construction, expected-output construction, complete XLSX reopen,
page-setup and calculation-metadata readback, Part/content-type/relationship
comparison, printer-reference absence, exact media checks, source/sink
inspection, and hashing remained outside timing.

The retained serial sequence on CPU 2 was eager A, source A, source B, eager B.
Each leg used 50 warmups and 500 samples, yielding 1,000 samples per state.
Exact aggregate evidence is in the
[`measurement summary`](../results/xlsx-page-setup-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 1,000 | 1,000 | — |
| p50 | 218.626 ms | 4.847 ms | **-97.78% (45.10x)** |
| mean | 219.188 ms | 4.846 ms | **-97.79% (45.23x)** |
| mean 95% interval | 219.000-219.376 ms | 4.833-4.859 ms | disjoint |
| p95 | 224.703 ms | 5.186 ms | **-97.69% (43.33x)** |
| p99 | 229.259 ms | 5.409 ms | **-97.64% (42.38x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,786,902 | 16,786,902 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

The two semantic materializations are `xl/workbook.xml`, used to resolve and
bind `Sheet1`, and `xl/worksheets/sheet1.xml`, used for the typed edit. The
other ten ordinary Parts are raw-copied. Physical reads remain archive-sized
because unchanged compressed members still have to reach the sequential sink.
Between-leg p50/mean drift is 0.16%/0.16% for eager and 1.42%/0.62% for
source-backed, within the 5% policy.

## Allocation, counters, and memory

One-sample Heaptrack attribution reports allocation calls 14,336 to 12,831
(-10.50%) and temporary allocations 1,800 to 1,652 (-8.22%). Peak heap is flat
at 152.84 to 152.81 MiB. Heaptrack-inclusive RSS rises 2.40%, while two
uninstrumented GNU Time processes per state report mean maximum RSS 141,098 to
141,648 KiB (+0.39%, flat). Both retain the same 1.78 KiB of profiler/runtime
leakage.

Two process-wide `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 3,956.25 ms | 1,107.48 ms | -72.01% |
| cycles | 19.450 billion | 5.493 billion | -71.76% |
| instructions | 48.696 billion | 10.783 billion | -77.86% |
| branches | 8.260 billion | 1.409 billion | -82.94% |
| branch misses | 175.711 million | 14.469 million | -91.77% |
| cache references | 1.386 billion | 414.552 million | -70.10% |
| cache misses | 27.924 million | 20.043 million | -28.22% |
| page faults | 342,763 | 258,110 | -24.70% |

CPU migrations were zero. Process-wide reductions are smaller than scoped
latency because deterministic corpus construction and complete untimed
verification remain inside the profiled process.

## Validation and limitations

Passed on the final source:

- focused relationship-free page-setup tests for add, replace, remove, exact
  no-op, signed no-op identity, changed signed refusal, patch/inverse,
  stale/foreign/source-version conflicts, complete worksheet-relationship
  mutation, printer-reference refusal, chartsheets, MCE, limits, partial sinks,
  complete reopen, eager byte equivalence, and unselected media/Part identity;
- complete XLSX and performance-harness suites, deterministic smoke and CI
  pinned corpus/output/materialization assertions;
- warning-denied XLSX production-library and performance-harness Clippy (the
  broader all-target XLSX command still reports three pre-existing test-style
  lints outside this batch);
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation fix from `1194fbc7f`; and
- formatter, workflow parsing, whitespace and final-diff checks.

This capability does not claim general source-backed XLSX editing or exact
whole-artifact authorization beyond the retained semantic closure. Cells,
formulas, styles, shared strings, printer settings, protection, relationships,
sheet topology, content types, and changed signed sources remain outside the
one-Part capability. OLE2, RTF, ODF production code and every iWork/IWA crate
are unchanged by this batch.
