# Change 0080: source-backed XLSX auto-filter publication

Date: 2026-08-12

Production base: `6df5d4a1fbe53a8216e63f24cc1392be60b714a8`

Status: accepted

## Hypothesis and implementation

Worksheet auto filters already had a typed reader and writer, but publishing a
small filter or sort-state change still required eager OPC conversion, which
inflated and recompressed every media Part.

`litchi_xlsx::auto_filter::SourceBackedEditor` now owns one immutable
positional source. Its snapshot binds the exact workbook and selected normal
worksheet owners, the complete outbound worksheet relationship set, worksheet
protection state, and the styles relationship, styles XML, and differential
format count when present. The isolated edit stages one complete optional
`Definition`, supporting distinct add/replace, clear, and exact no-op results.
Commit applies the bounded byte-preserving worksheet rewriter, performs typed
post-write readback, and returns an exact reversible source-specific patch.
Publication consumes the accepted one-Part OPC overlay.

The direct worksheet `autoFilter` subtree and nested `sortState` are in scope.
Table-owned filters, cells, formulas, tables, relationships, Parts, signatures,
and topology are not. MCE-selected filters, filter-locked or sort-locked
worksheets, invalid DXF references, stale/foreign sources, changed relationship
or styles closures, changed signed sources, malformed or over-limit XML,
unsupported ZIP layouts, and partial sinks retain typed refusals. Exact no-ops
reproduce the complete source artifact. No dependency, unsafe code, global
cache, or iWork/IWA code was added.

The byte rewriter caps selected worksheet XML at 32 MiB, XML nesting at 128,
and events at one million. It refuses DTD/processing-instruction inputs and
ambiguous MCE projection, retains Strict versus Transitional spreadsheet
namespace dialect, inserts in schema order, and preserves unrelated bytes.

## Matched corpus and protocol

Both controls use one workbook, one normal worksheet, one styles Part, one
drawing, and eight referenced deterministic incompressible 2 MiB PNG Parts.
The worksheet starts with one typed value filter and descending sort. The
corpus has 12 ordinary Parts, 17 ZIP members, 16,782,720 logical Part bytes,
and a 16,786,945-byte archive with SHA-256
`57678991c8cabfceda63b278f4e50fee87fd7f540f0c2f0eff8cb048f457d421`.

Both paths replace the filter range, selected values, sort range, direction,
and case-sensitivity with the same typed state. The eager control first
materializes the complete OPC package, then loads the same semantic closure
and invokes the same rewriter. The candidate performs one guarded source
transaction and publishes one selected worksheet overlay. Both produce the
same 16,786,968-byte artifact with SHA-256
`34893a3eaf685f20569a6ea383de022825b0fcb434ff49039eaa27bb262b8561`.
Typed reopen, filter/sort equality, calculation metadata, package topology,
relationships, content types, untouched Part/media bytes, hashing, source
counters, and sink bounds remain outside timing.

Both cases share frozen release binary SHA-256
`871856f3f7993428a9a7e667a20e9033d71292101bd275bee835232dcdc677a8`.
The retained CPU-2 ABBA order was eager A, source-backed A, source-backed B,
eager B, with ten warmups and 100 samples per leg (200 per state). Raw reports
are [`before A`](../results/abba-xlsx-auto-filter-before-a.json),
[`after A`](../results/abba-xlsx-auto-filter-after-a.json),
[`after B`](../results/abba-xlsx-auto-filter-after-b.json), and
[`before B`](../results/abba-xlsx-auto-filter-before-b.json). Aggregated
evidence is in the
[`measurement summary`](../results/xlsx-auto-filter-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | — |
| p50 | 218.438 ms | 4.755 ms | **-97.82% (45.94x)** |
| mean | 219.123 ms | 4.738 ms | **-97.84% (46.25x)** |
| p95 | 224.812 ms | 5.078 ms | **-97.74% (44.27x)** |
| p99 | 234.040 ms | 5.141 ms | **-97.80% (45.52x)** |
| semantic Part materializations | 12 | 3 | -75.00% |
| output bytes | 16,786,968 | 16,786,968 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | bounded |

Both same-state legs remain within the five-percent drift policy. The candidate
materializes the workbook, selected worksheet, and styles Part; the other nine
Parts remain compressed and are raw-copied into the sequential output.

## Allocation, counters and memory

One-sample Heaptrack attribution covers the whole process, including corpus
construction and untimed verification. Allocation calls are 16,822 eager
versus 16,495 source-backed (-1.94%); temporary allocations are 2,181 versus
2,171. Peak heap is 152.84 versus 152.81 MiB (flat). Uninstrumented maximum
RSS is 143,964 versus 140,760 KiB (-2.23%). Heaptrack's profiler-inclusive RSS
is recorded but is not used for acceptance.

Three `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| cycles | 20.176 billion | 6.386 billion | -68.35% |
| instructions | 50.870 billion | 13.489 billion | -73.48% |
| branches | 8.651 billion | 1.893 billion | -78.11% |
| branch misses | 188.949 million | 26.952 million | -85.74% |
| cache references | 1.435 billion | 472.267 million | -67.10% |
| cache misses | 29.136 million | 18.682 million | -35.88% |

Latency, materialization, instruction, allocation, peak-heap, and RSS gates all
clear the acceptance thresholds.

## Correctness and regression closure

Focused integration tests cover changed publication, add/replace/clear/no-op,
ordinary complete reopen, exact unselected payload/content-type/relationship
preservation, style-DXF acceptance and refusal, patch replay and inverse,
signed no-op identity, changed signed refusal, source-version and foreign-source
conflicts, protection and MCE refusal, and partial-sink failure. Unit tests
cover Strict output, sort-state round trips, schema-order insertion, unrelated
byte preservation, DTD/MCE refusal, and the 32 MiB hostile-input bound.

The complete XLSX all-feature suite, focused harness equivalence test,
library-only XLSX denied-warning Clippy, harness denied-warning Clippy,
formatting, CI deterministic hash/materialization assertions, workflow parse,
and ADR aggregate pass. ODF-common tests, denied-warning Clippy, and rustdoc
also pass. The full harness has one unrelated existing PPTX batch
materialization failure; XLSX all-target Clippy and rustdoc remain blocked by
unrelated existing test lints/private documentation links. The boundary audit
reports only the repository's existing unclassified edges; this change adds no
dependency edge.

This batch also removes the remaining deprecated `GenericArray::clone_from_slice`
construction from the ODF-common Blowfish test fixture. It uses a default
block followed by `copy_from_slice`, matching the production fix and retaining
the same ciphertext fixture; ODF-common denied-warning Clippy and rustdoc are
separate gates.

## Alternatives retained

The next native DOC/OLE2 owner opportunity remains benchmark-first because its
memory risk is not yet attributed. RTF's forward-cursor batch and ODT's lazy
paragraph-count scan remain bounded future candidates. A previously measured
shared CFB payload prototype regressed end-to-end publication by 32% and remains
rejected.
