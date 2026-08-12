# Change 0076: source-backed XLSX defined-name publication

Date: 2026-08-12

Production base: `2fead7927f6111acccf49aaaf543d285d88d0f90`

Status: accepted

## Hypothesis and implementation

Workbook defined names are owned wholly by the direct `definedNames` child of
the resolved workbook Part. Replacing that catalog without changing sheet
topology, relationships, cells, formulas, calculation chains, styles, shared
strings, or content types is therefore a safe one-Part mutation. Publishing it
through the accepted OPC overlay should avoid materializing and recompressing
the eleven unrelated Parts in the media-rich control.

`defined_names::SourceBackedEditor` is an additive, non-cloneable editor over a
caller-provided `Arc<dyn ReadAt>`. It snapshots the unique package
`officeDocument` owner, exact workbook URI/content type/XML and ordered sheet
catalog, stages a complete typed `Vec<DefinedName>`, and uses the existing
checked catalog rewriter. The complete changed workbook is reparsed before an
exact reversible `Commit` and `Patch` can be published. Publication recaptures
the retained source closure and consumes
`SourceBackedPackage::write_part_overlay_to_stream`.

The rewriter refuses structure-protected workbooks and any MCE or unknown child
inside `definedNames`; it validates duplicate names, limits and local sheet
scope. Exact semantic no-ops reproduce the source archive byte-for-byte,
including signed input. Changed signed input, stale or foreign commits, source
version changes, invalid local scope, limits and partial sinks retain typed
failures before successful output. No dependency, runtime, unsafe code, global
cache, relationship edit, topology operation or iWork/IWA code was added.

## Matched corpus and protocol

Both controls use a defined-name-specific view of the deterministic media-rich
XLSX archive: 12 ordinary Parts and 17 ZIP members, including one workbook, one
worksheet, one drawing and eight referenced incompressible 2 MiB PNG Parts.
The archive is 16,786,830 bytes with 16,782,412 logical Part bytes and SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths replace the workbook catalog with one global range and one hidden
sheet-local cell name. The eager control performs positional open, eager OPC
ownership, ordinary workbook edit and ordinary sequential publication. The
candidate performs positional open, the source-backed edit and one-Part
overlay publication. Both emit the same 16,786,946-byte artifact with SHA-256
`f59c85a4003018c58732db832aff9ac3577cff7e7af37d9372ddc4ca5679a615`.

The frozen control binary SHA-256 is
`213106275522d57f9ebf7eb38fd589284dd5c3de155bf61ba24de4c959f9fa81`;
the candidate binary SHA-256 is
`1ba55cb94c15bd60db0c899e959256e2f3157cc08743a067c6e2e705f76ae43f`.
Corpus construction, complete XLSX reopen, defined-name and calculation-policy
readback, topology/relationship/content-type checks, exact unselected media,
source/sink inspection and hashing remained outside timing.

The retained serial sequence on CPU 3 was eager A, source A, source B, eager B.
Each leg used 30 warmups and 200 samples, yielding 400 samples per state. The
first exploratory sequence was discarded before inference because candidate
between-leg p50 drift exceeded the 5% policy. Exact retained evidence is in the
four `abba-xlsx-defined-names-stable-*` reports and the
[`measurement summary`](../results/xlsx-defined-names-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 400 | 400 | — |
| p50 | 220.101 ms | 4.752 ms | **-97.84% (46.32x)** |
| mean | 220.626 ms | 4.829 ms | **-97.81% (45.69x)** |
| mean 95% interval | 220.300-220.951 ms | 4.785-4.873 ms | disjoint |
| p95 | 225.785 ms | 5.380 ms | **-97.62% (41.97x)** |
| p99 | 233.721 ms | 5.662 ms | **-97.58% (41.28x)** |
| semantic Part materializations | 12 | 1 | -91.67% |
| output bytes | 16,786,946 | 16,786,946 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

Only `xl/workbook.xml` is materialized by the source-backed path. The other 11
ordinary Parts are copied from their physical source spans. Payload reads stay
archive-sized because unchanged compressed members still have to reach the
sequential sink. Between-leg p50/mean drift is 0.02%/-0.03% for eager and
-3.35%/-3.71% for source-backed, within the 5% policy.

## Allocation, counters and memory

One-sample Heaptrack attribution reports allocation calls 14,887 to 12,986
(-12.77%) and temporary allocations 2,120 to 1,828 (-13.77%). Peak heap is
flat at 152.83 to 152.80 MiB, and Heaptrack-inclusive RSS falls 1.28%. Two
uninstrumented GNU Time processes per state report mean maximum RSS 143,498 to
142,612 KiB (-0.62%, flat). Both retain 1.78 KiB of profiler/runtime leakage.

Three `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 3,993.91 ms | 1,076.10 ms | -73.06% |
| cycles | 19.482 billion | 5.322 billion | -72.68% |
| instructions | 48.262 billion | 10.399 billion | -78.45% |
| branches | 8.186 billion | 1.344 billion | -83.59% |
| branch misses | 175.858 million | 14.274 million | -91.88% |
| cache references | 1.383 billion | 404.439 million | -70.75% |
| cache misses | 30.651 million | 21.009 million | -31.46% |
| page faults | 271,380 | 197,376 | -27.27% |

CPU migrations were zero. The complete profiles lost no samples; the
cycle-weighted `deflate_medium` and `longest_match` attribution falls 92.70%
and 92.78%, respectively. Process-wide reductions are smaller than scoped
latency because corpus construction and complete untimed verification remain
inside the profiled process.

## Validation and limitations

Passed on the final source:

- focused source-backed defined-name tests for replace, clear, exact no-op,
  patch/inverse, complete reopen, eager byte equivalence, exact unselected Part
  and media identity, signed/stale/foreign/version conflicts, protection, MCE,
  invalid local scope, limits and partial sinks;
- complete XLSX and performance-harness suites plus CI-pinned corpus, output,
  sink and materialization assertions;
- warning-denied XLSX production-library and performance-harness Clippy;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation fix from `1194fbc7f`; and
- formatter, workflow parsing, JSON parsing, whitespace and final-diff checks.

This does not claim general source-backed XLSX editing. Cells, formulas, cached
results, calculation chains, styles, shared strings, relationships, sheet
topology, content types, MCE-projected catalogs, protected workbooks and
changed signed sources remain outside the one-Part capability. OLE2, RTF and
ODF production code are unchanged by this batch; iWork/IWA remains deferred.
