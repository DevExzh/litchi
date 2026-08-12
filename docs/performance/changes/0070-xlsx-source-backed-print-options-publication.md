# Change 0070: source-backed XLSX print-options publication

Date: 2026-08-12

Production base: `9670259f28c394ced897e9549a0c79c9e70c438d`

Status: accepted

## Hypothesis and implementation

Direct worksheet `printOptions` are a safe one-Part XLSX mutation. The existing
typed owner validates all five boolean wire flags, changes only the selected
worksheet XML, and does not alter cells, formulas, calculation policy, styles,
shared strings, relationships, printer settings, or package topology. The
hypothesis was that publishing the exact edit through the accepted OPC one-Part
overlay would avoid materializing and recompressing unrelated drawings and
media.

`print_options::SourceBackedEditor` is an additive, non-cloneable editor over a
caller-provided `Arc<dyn ReadAt>`. A source edit loads the workbook catalog and
one selected normal worksheet, stages `Option<PrintOptions>`, exposes only
`set` and `remove`, invokes the existing MCE-aware direct-property rewriter,
and reparses the complete changed worksheet before producing an exact
reversible `Commit` and `Patch`. Publication recaptures the source-bound
closure and then consumes `SourceBackedPackage::write_part_overlay_to_stream`.

The retained identity includes the package `officeDocument` relationship,
workbook URI/content type/XML, selected workbook relationship
ID/type/target/mode, worksheet URI/content type/XML, checked sheet name and
position, and typed print options. A retargeted relationship cannot apply a
patch to an orphaned worksheet. Exact semantic no-ops reproduce the complete
source archive, including signed input. Changed signed input, stale or foreign
closures, retargeted relationships, chartsheets, projected MCE state, read
limits, unsupported ZIP layouts, and partial sink failure retain typed
refusals.

No dependency, runtime, unsafe code, global cache, archive abstraction,
ordinary workbook capability, relationship edit, or topology operation was
added.

## Matched corpus and protocol

Both controls use the deterministic media-rich XLSX archive under a
print-options-specific manifest. It contains 12 ordinary Parts and 17 ZIP
members: one workbook, one worksheet, one drawing, eight referenced
incompressible 2 MiB PNG Parts, and minimal supporting Parts. The archive is
16,786,830 bytes with 16,782,412 logical Part bytes and SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths enable horizontal centering, printed headings, and printed gridlines
on `Sheet1`. The eager control performs positional open,
`into_opc_package`, the existing ordinary workbook transaction, and the
ordinary sequential writer. The candidate performs positional open, the new
source edit, and one-Part overlay publication. Both emit the same
16,786,886-byte artifact with SHA-256
`6eee37e1c0e4e9cdf1f364fdf1cbc90d58a6c25acc6fce222670f937abd5a74c`.

The two paths are measured from the same frozen release binary, SHA-256
`d24f350ceceae47c48bc465a4089910cc2936de6c2c497f38d9506ceec6140aa`.
Corpus creation, expected-output construction, complete XLSX reopen,
print-options and calculation-metadata readback, Part/content-type/relationship
comparison, exact media checks, source/sink inspection, and hashing stay
outside timing.

The retained sequence on CPU 2 was eager A, source A, source B, eager B. Each
leg used 50 warmups and 500 samples, yielding 1,000 samples per state. Exact
raw samples and attribution are indexed in the
[`measurement summary`](../results/xlsx-print-options-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 1,000 | 1,000 | — |
| p50 | 219.294 ms | 4.668 ms | **-97.87% (46.98x)** |
| mean | 219.958 ms | 4.671 ms | **-97.88% (47.09x)** |
| mean 95% interval | 219.742-220.174 ms | 4.653-4.688 ms | disjoint |
| p95 | 225.747 ms | 5.170 ms | **-97.71% (43.66x)** |
| p99 | 232.393 ms | 5.435 ms | **-97.66% (42.76x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,786,886 | 16,786,886 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

The two required semantic materializations are `xl/workbook.xml`, used to
resolve and bind `Sheet1`, and `xl/worksheets/sheet1.xml`, used for the typed
edit. The other ten ordinary Parts are raw-copied. Physical source reads remain
archive-sized because untouched compressed members still have to reach the
sequential sink. Between-leg p50/mean drift is 0.19%/0.25% for the eager state
and 2.61%/1.99% for the source-backed state, within the 5% policy.

## Allocation, counters, and memory

One-sample Heaptrack attribution reports allocation calls 14,290 to 12,561
(-12.10%) and temporary allocations 1,718 to 1,523 (-11.35%). Peak heap is
flat at 152.84 to 152.81 MiB. Heaptrack-inclusive RSS rises 1.49%, while two
uninstrumented GNU Time processes per state report mean maximum RSS 141,462 to
141,672 KiB (+0.15%, flat). Both sides retain the same 1.78 KiB of
profiler/runtime leakage.

Two matched process-wide `perf stat` repeats per state used two warmups and ten
samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 3,895.50 ms | 1,064.42 ms | -72.68% |
| cycles | 19.217 billion | 5.180 billion | -73.05% |
| instructions | 48.188 billion | 10.514 billion | -78.18% |
| branches | 8.171 billion | 1.362 billion | -83.33% |
| branch misses | 175.697 million | 14.358 million | -91.83% |
| cache references | 1.363 billion | 397.998 million | -70.80% |
| cache misses | 25.377 million | 16.943 million | -33.23% |
| page faults | 257,968 | 215,238 | -16.56% |

CPU migrations were zero. These whole-process reductions are smaller than the
scoped latency win because deterministic corpus creation and complete untimed
verification remain inside the profiled process.

## Validation and limitations

Passed on the final source:

- complete `litchi-xlsx --all-features` unit, integration and doc-test suite;
- focused source-backed print-options integration coverage for add, replace,
  remove, exact no-op, signed no-op identity, changed signed refusal,
  patch/inverse, stale/foreign/source-version conflicts, relationship
  retargeting, chartsheet refusal, MCE projection, limits, unselected
  Part/media identity, and partial sinks;
- complete performance-harness tests and deterministic debug/release smoke;
- warning-denied focused XLSX and performance-harness Clippy;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the deprecation cleanup from `1194fbc7f`; and
- formatter, JSON/YAML parsing, whitespace and final-diff checks.

This capability does not claim general source-backed XLSX editing. Cells,
formulas, styles, shared strings, page setup/printer relationships, protection,
sheet topology, relationships, and content types remain outside its one-Part
closure. OLE2, RTF, ODF production code and every iWork/IWA crate are unchanged
by this batch.
