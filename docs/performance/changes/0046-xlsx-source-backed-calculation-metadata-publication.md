# Change 0046: source-backed XLSX calculation-metadata publication

Date: 2026-08-11
Production base: `c60a72b1681c2dbbe575f74179de092add8fb8f0`
Status: accepted

## Hypothesis

The accepted OPC one-Part publisher already proved that a same-topology edit
can regenerate one selected Part and raw-copy every other ZIP member. DOCX and
PPTX use that mechanism, but XLSX still had no semantic publication boundary.
General cell editing is not a safe one-Part mutation because it can rewrite the
workbook calculation policy and remove the calculation chain. Workbook
calculation metadata is different: the existing typed transaction owns only
`xl/workbook.xml` and already refuses unmodeled MCE state.

The hypothesis was that applying this exact one-Part semantic closure to a
media-rich workbook would eliminate eager inflation, ownership, validation,
and recompression of untouched worksheets, drawings, and images while
preserving the existing transaction and package contracts.

## Implementation

`litchi_xlsx::calculation_properties::SourceBackedEditor` is an additive,
non-cloneable source editor over `Arc<dyn ReadAt>`. Opening validates the OPC
catalog and unique workbook owner, loads only the workbook XML, and uses the
existing bounded calculation-properties inspection. `SourceEdit` exposes the
same `calcPr` and calculation-feature staging verbs as the owning transaction.

A changed commit runs the established lossless rewriter, reparses the complete
candidate workbook XML, and checks the exact staged semantics. Publication
recaptures the workbook URI, content type, owner relationship, exact source XML,
and source version before consuming the package into
`SourceBackedPackage::write_part_overlay_to_stream`. Exact no-ops copy the
complete source artifact. Changed signed sources, stale sources, foreign
workbook closures, projected MCE state, over-limit payloads, unsupported ZIP
layouts, and output failures retain typed refusals.

The API does not claim general source-backed XLSX editing. Cells, formulas,
cached results, styles, shared strings, calculation-chain topology,
relationships, and other Parts remain outside its mutation closure.

## Matched corpus and protocol

The deterministic `litchi-xlsx-calculation-metadata-source-edit-media-v1`
corpus contains one workbook, one worksheet, one DrawingML drawing, eight
referenced incompressible 2 MiB PNG Parts, and the ordinary minimal supporting
Parts. It has 12 ordinary Parts, 17 ZIP members, 16,782,412 logical Part bytes,
and a 16,786,830-byte archive.

- input SHA-256:
  `c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`
- target: `xl/workbook.xml`, `calcId=7` to `calcId=91`
- byte-identical output SHA-256 from both paths:
  `96f60a09a8e87204a533e178bdda115b0202e8aa8ac4fe5673a881f981d3e98d`

The matched eager control performs positional OPC open, complete
`into_opc_package` materialization, the existing owning XLSX transaction, and
the ordinary sequential writer. The candidate performs positional open, the
new source edit, and one-Part overlay publication. Corpus construction,
expected-output construction, complete reopen, calculation-metadata readback,
Part/content-type/relationship/topology comparison, exact media checks, source
and sink inspection, and output hashing stay outside timing.

The retained balanced sequence was control A, candidate A, candidate B,
control B. Every leg used ten warmups and 100 measured samples, pinned to CPU 2
with the allocator thresholds recorded in the raw JSON environment. A first
exploratory cycle was discarded before reporting because the two candidate
legs exceeded the 5% drift gate. The retained legs stay within 5%.

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | — |
| p50 | 215.456869 ms | 1.611870 ms | **-99.2519% (133.67x)** |
| mean | 216.098239 ms | 1.619241 ms | **-99.2507% (133.46x)** |
| mean 95% t interval | 215.555146-216.641332 ms | 1.603506-1.634976 ms | disjoint |
| p95 | 222.657959 ms | 1.803250 ms | **-99.1901% (123.48x)** |
| p99 | 227.428911 ms | 1.913619 ms | **-99.1586% (118.85x)** |
| semantic Part materializations | 12 | 1 | -91.67% |
| output bytes | 16,786,830 | 16,786,830 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | exact |

Independent p50/mean legs were 213.407/214.280 ms and
217.277/217.917 ms for the control, then 1.574/1.582 ms and
1.640/1.656 ms for the candidate.

The candidate avoids semantic materialization and recompression of 11 Parts,
covering 16,782,109 unselected logical bytes. Physical source reads remain
archive-sized because those members must still be copied to the output. This
is raw compressed-member passthrough, not zero-I/O or zero-copy output.

## Counter and memory attribution

Three `perf stat` process repeats per state (two warmups and ten samples) show:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task-clock | 3,896.15 ms | 1,113.51 ms | -71.42% |
| cycles | 19,226,776,745 | 5,457,630,415 | -71.61% |
| instructions | 48,123,852,095 | 10,691,537,042 | -77.78% |
| branches | 8,161,172,031 | 1,394,202,313 | -82.92% |
| branch misses | 176,646,559 | 14,474,400 | -91.81% |
| cache references | 1,373,621,717 | 404,548,650 | -70.55% |
| cache misses | 27,636,957 | 18,863,350 | -31.75% |

The whole-process counter scope includes deterministic corpus construction and
untimed verification, so it understates the timed publication improvement.
The control profile retains 69.94% exclusive cycles in the two leading zlib-rs
Deflate frames; those frames total 16.85% after the untouched entries leave the
recompression path. Both profiles lost zero samples.

One-sample Heaptrack attribution reports allocation calls 12,439 -> 11,094
(-10.81%) and temporary allocations 1,484 -> 1,366 (-7.95%). Peak heap is flat
at 152.83 -> 152.80 MiB. Heaptrack RSS rises 1.30%, but the uninstrumented GNU
Time authority falls 143,148 -> 141,436 KiB (-1.20%); no RSS regression is
accepted or hidden. Both sides retain the same 1.78 KiB tool/runtime leak.

## Correctness and ADR compliance

- ADR 0001/0002/0011: the format owner exposes typed calculation semantics;
  physical ZIP ownership stays in `litchi-opc`, with no dependency change.
- ADR 0003: edits are isolated; commits and inverse patches remain exact-source
  checked; the source snapshot is immutable and publication is atomic before
  writing begins.
- ADR 0005: ingress is caller-provided `ReadAt`, output is sequential, the
  source version remains monitored, and the optimization is supported by
  matched latency, counters, allocation, heap, and RSS evidence.
- ADR 0006: all untouched Parts and metadata remain exact in the logical
  package comparison; the OPC publisher raw-copies their physical records.
  Signed changes refuse, exact signed no-ops remain byte-identical, MCE
  projection refuses, and validation never repairs.
- ADR 0018: calculation-chain ownership is unchanged. This capability cannot
  add, remove, or rewrite the chain and does not infer formula dependencies.

Focused tests cover changed and no-op publication, exact unchanged Part/media
payloads, owner and Part metadata, semantic reopen, forward patch replay and
exact inverse restoration, signed
sources, foreign workbook XML, source-version changes, MCE, OPC limits, and
partial sink failure. The existing calculation-metadata schema, formula,
security, package, and owning-transaction suites remain the guardrail.

The full all-feature XLSX test command passes 732 unit tests, all integration
suites (including the four new source-backed cases), and two doctests. The
32-test benchmark harness, library warning-denied Clippy, focused integration
warning-denied Clippy, format checks, YAML/JSON parsing, evidence hashes and
exact CI smoke assertions pass. Three unrelated pre-existing all-target
Clippy findings remain in `xml_maps_api_compat`, `package::xldm::tests` and
`timelines::tests`; warning-denied public rustdoc likewise remains blocked by
pre-existing private/broken intra-doc links outside this module. The broad
crate-boundary checker reports existing unclassified workspace edges; this
batch changes no manifest or dependency edge.

## Evidence

- four raw ABBA JSON files under `docs/performance/results/`
- `xlsx-calculation-metadata-edit-perf-stat-{before,after}.csv`
- `xlsx-calculation-metadata-edit-time-{before,after}.txt`
- `xlsx-calculation-metadata-edit-profile.txt`
- `xlsx-calculation-metadata-edit-sha256.txt`

The stable harness exposes both matched cases and CI runs the exact corpus,
source/output hashes, Part counts, materialization counts, and sink bounds.
