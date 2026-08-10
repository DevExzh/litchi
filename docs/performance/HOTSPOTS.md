# Performance hotspot inventory

Status: source-audited; initial ZIP/OPC and CFB substrate measurements captured
Branch: `feat/office-format-completeness`
Evidence through: [`change 0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md)

This document records facts established by source inspection. It is not a
performance-results report. A path is called a bottleneck only after the
process benchmark and profiler evidence in `BASELINE.md` confirms its effect
on a named corpus and scenario.

## Shared OOXML data path

```text
path / Read / Vec / &[u8]
  -> litchi-opc physical reader
     -> complete source Vec for path and generic Read ingress
     -> soapberry-zip central-directory index
     -> content types and package relationships
     -> relationship-graph validation
     -> classify every physical member
     -> decompress every admitted Part
  -> OpcPackage
     -> HashMap<PackURI, Box<dyn Part>>
     -> second source-XML index
  -> DOCX / PPTX / XLSX mandatory catalog
  -> lazy format-owned semantic parse of a selected Part
  -> Edit plan and dependency validation
  -> candidate Part reconstruction and readback
  -> PackageWriter
     -> exact owned-source copy when no mutable API was entered
     -> otherwise regenerate manifests and relationship Parts
     -> build, audit, and retain one deterministic publication plan
     -> Deflate every Part into a sequential sink
```

Current work shape:

- Legacy path and generic-reader OPC ingress still has a contiguous-buffer
  path, while source-backed ingress uses an immutable positional source with
  source versions and a validated ZIP index.
- `PackageReader::load_parts_lazy` is not physically lazy: it classifies and
  decompresses every admitted Part, including unreferenced Parts that must be
  preserved.
- Ordinary bulk opens are serial. Explicit eager opens opt into local bounded
  ZIP sessions through `litchi-core::ExecutionContext` and OPC `OpenSession`;
  there is no hidden global Rayon pool.
- `OpcPackage` retains every inflated Part. XML Parts also participate in a
  second source-XML map. Part lookup has an exact hash-map fast path and a
  linear ASCII-case-insensitive fallback.
- Exact unchanged owned OPC output reuses the complete source archive. Owned,
  same-topology mutation now retains private provenance and raw-copies every
  semantically unchanged ZIP member; topology changes, borrowed ingress, and
  unsupported ZIP layouts still use the complete rewrite.
- `PackageWriter` previously reconstructed generated XML and Part order during
  emission. The measured `PublicationPlan` change now constructs, audits, and
  reuses that state once. It reduced allocation calls by 37.0% in the profiled
  256-Part save and mean latency by 5.49% in the 2,048-Part compressible save;
  full-Part recompression remains unchanged on the fallback path. Targeted raw
  publication separately improves p50 by 58-96% across the synthetic cells,
  while retained-source peak heap rises 37%; see change 0008.

## XLSX selective read and edit path

```text
whole-package OPC materialization
  -> workbook catalog and relationship parse
  -> worksheet handles with OnceLock<Store>
  -> first cell/range query
     -> parse the complete selected worksheet XML
     -> materialize and sort the complete sparse Store
  -> targeted edit commit
     -> compare against complete Store
     -> scan complete lossless worksheet layout
     -> allocate/copy complete replacement worksheet XML
     -> compact and reparse complete replacement for publication proof
     -> clone shared OPC graph and replace changed Part
  -> save
     -> recompress complete package
```

Confirmed source facts:

- The legacy eager path still materializes all admitted Parts. The additive
  source-backed XLSX facade avoids timed source reads while listing after open;
  cache bytes are finite but not yet charged to a hierarchical `Budget`.
- One first cell access parses the entire selected worksheet. The non-evicting
  `OnceLock` retains it for the snapshot lifetime.
- The sparse cell store is row-major and supports binary-search point lookup.
  A compact immutable row-start index now skips preceding rows for narrow
  ranges. The measured range query improves about 80%; full scan and first-cell
  guardrails remain near neutral.
- A targeted cell edit performs a semantic parse, an independent lossless
  layout scan, full replacement-byte construction, and a full changed-sheet
  semantic readback before publication.
- Bulk cell actions are held in address order, then regrouped into nested
  row/cell `BTreeMap`s during worksheet emission.
- An empty edit returns the original immutable workbook allocation. When the
  workbook came from owned ingress, saving that no-op snapshot now preserves
  the exact validated OPC source; borrowed ingress still performs a rewrite.

## DOCX and PPTX paths

DOCX format views are borrowed after eager OPC materialization. Repeated
`paragraphs`, `tables`, and `blocks` queries rescan and allocate result vectors.
Single-index paragraph lookup now scans the complete bounded XML but constructs
only the selected shared range; the 10,000-paragraph cell improves 4.72% p50
and removes ten collection-growth allocations per call. Table lookup still
builds the complete collection. Canonical direct-body paragraph batches now
plan every replacement against one snapshot, emit the disjoint ranges in one
forward pass, parse one candidate, and read back every selected paragraph. The
10,000-paragraph / 100-edit save improves 94.99% p50 (19.97x) and allocation
calls fall 94.11%. Scalar edits, unordered/nested selections, structural edits,
and complete transaction-capture costs are unchanged.

PPTX ordinary reads defer slide payload parsing, but repeatedly parse the
presentation slide-reference list. Exact-name slide lookup resolves and parses
all candidate slide names. The opened-transaction snapshot is deliberately
stronger and more expensive: it resolves every slide, notes graph, Part and
relationship fingerprint, and retains a cloned shared OPC graph. Commit
recaptures and re-fingerprints the candidate after readback. Shape-text edits
now reuse the selected scene when mapping its raw span, removing one redundant
scene parse per change. The 100-edit cell improves 9.37% p50/mean and allocation
calls fall 11.67%; the single-edit end-to-end guardrail remains neutral because
complete capture/commit work dominates.

These paths have strong preservation and atomicity tests plus generated-text
timing/allocation evidence. Real-producer, media/dependency, malformed,
security, copied-byte and cold-source matrices remain missing.

## ODF paths

ODT, ODS and ODP ordinary opens eagerly read and parse their ZIP packages.
The opt-in public semantic matrix now measures owned open, listing, one object,
full text, small creation, exact no-op and one supported edit/save across all
three owners. ODT/ODP repeated semantic queries still rescan complete XML.
Changed ODF publication still regenerates and compresses package members.

ODS unified snapshot construction previously cloned package bytes and parsed
the same ODS package twice: once for package/resource validation and again for
complete `Spreadsheet` readback. It now moves the one validated package into a
crate-private facade constructor. Large no-op edit/save p50 falls 11.78%; the
large changed case improves 2.06% because full spreadsheet rewrite/readback
dominates. Exact source bytes, resource bounds and facade readback remain.

## Legacy CFB data path

```text
Read + Seek
  -> header and complete FAT
  -> complete directory bytes
     -> structural validation pass
     -> public entry decoding pass
  -> complete MiniFAT metadata
  -> validate every stream allocation chain
  -> semantic DOC / XLS / PPT owner
     -> lookup a child by cached validated sibling-tree keys
     -> materialize selected stream Vecs
  -> edit/rebuild
     -> retain all output stream Vecs
     -> copy borrowed stream slices into OleWriter
     -> assemble MiniFAT/FAT/directory sector buffers
     -> Write + Seek output
```

Confirmed source facts:

- `SharedOleFile` provides positional CFB access and explicit bounded bulk
  operations. Four 4 MiB streams reach 5.93x p50 at 12 visible CPUs, while 256
  1 KiB streams regress at high worker counts; thresholds remain essential.
- Open eagerly materializes FAT, directory, MiniFAT, and allocation topology,
  while ordinary large stream payloads remain lazy.
- MiniFAT now parses directly into its final `Vec<u32>`; FAT/DIFAT/MiniFAT use
  one bounded sector buffer and directory sectors batch into the final buffer.
- Child lookup now descends the validated sibling tree with SID-aligned cached
  comparison keys. The 2,048-root-stream measurement improves about 94%.
- Fresh XLS and PPT writers move generated stream buffers into `OleWriter`.
  PPT improves about 20%; XLS peak heap falls about 9.5%. DOC retains the
  exact-sized copy because moving its spare-capacity buffer regressed 58%.
- Directory writing allocates scratch structures proportional to every entry
  for each storage, and duplicate checks scan existing siblings.
- `HashMap`/`HashSet` iteration in fresh CFB directory construction requires a
  separate determinism audit; it is not treated as a performance result.

The substrate harness still does not measure semantic legacy DOC/XLS/PPT
open/edit/save. Any additional owned-stream experiment must start with those
end-to-end baselines; the previous spare-capacity DOC move remains rejected.

## RTF path

RTF currently has no performance-harness cases. Native `Document::from_bytes`
and borrowed body access exist, the first complete text result is cached, and
the native writer accepts a forward-only sink. The unified root path still
passes through owned UTF-8 strings and raw materializers, so it must not be
used as evidence for the native path.

The smallest source-audited allocation candidate is raw
`RtfDocument::text()`: it collects text slices into a temporary `Vec<&str>`
and then joins them. A direct pre-sized `String` plus `push_str` can remove that
temporary allocation shape, but only after matched public open/list/full-text/
stream-save/no-op/one-edit baselines exist.

## Source and detector path

`litchi-core::ReadAt` provides immutable positional reads and source versions.
Source-backed OPC and positional CFB now consume it; IWA also consumes it,
though current IWA physical ingress snapshots the complete source with one
full-range request.

Generic smart detection may scan one ZIP package through multiple format
owners and then discard the prepared parse before the selected owner opens it.
The focused iWork route has already disproved that this duplication is
architecturally necessary: `litchi-iwa-detect::PreparedSource` retains an
opaque classified physical catalog without exposing archive types in the root
facade. The generic detection path must be measured before adapting that
pattern elsewhere.

## Initial hypotheses

| # | Source-audit disposition | Measurement needed |
|---:|---|---|
| 1 | Refined: legacy OPC path and `Read` ingress slurp the source; source-backed ingress is positional. | Cold-filesystem bytes, syscalls, latency and RSS across both modes; deterministic range-source distributions now exist. |
| 2 | Confirmed: ordinary OPC open inflates every admitted Part. | Open/list/one-object scaling against total uncompressed bytes and member count. |
| 3 | Superseded for source-backed OPC: finite weighted LRU, per-entry single-flight and content-free diagnostics exist; legacy eager open does not use that cache. | Cache bytes are not yet charged to the hierarchical Budget; add contention and retention measurements. |
| 4 | Measured: ordinary OPC open is serial and explicit eager open has a local bounded session. Six large ZIP tasks reach 4.52x p50 at 12 CPUs; small tasks regress. | Broader real-package scaling and threshold tuning. |
| 5 | Confirmed: stored entries are CRC-checked then copied. | Stored-media one-Part read and package-open copied-byte/RSS deltas. |
| 6 | Refined by measurement: exact unchanged saves copy the source; owned same-topology mutations raw-copy unchanged entries; borrowed/topology-changing/unsupported sources rewrite fully. | Real media-heavy 1% updates, source-backed editable publication, and reducing the measured retained-source/payload-copy memory cost. |
| 7 | Confirmed structurally: duplicate indexes, boxed Parts, source-XML map, and linear fallback exist. | Allocation profiles, type sizes, cache counters and repeated noncanonical lookup. |
| 8 | Refined: source-backed XLSX structural open/list avoids timed reads; selected first/range reads physically overlap only the selected worksheet. | Broader source-backed selectors, edits and real workbook matrices. |
| 9 | Confirmed structurally: small XLSX edits scan/rebuild/reparse the complete touched sheet and repackage all Parts. | First/middle/last cell, 1% updates, and commit-versus-save separation. |
| 10 | Plausible but unmeasured: per-cell semantic ownership and transient parse duplication may dominate large stores. | Allocation count/bytes, type sizes, peak RSS and cache-miss profiles. |
| 11 | Refined by implementation and measurement: CFB has positional `SharedOleFile` and bounded bulk reads; MiniFAT parsing and sector reads no longer require the former temporary buffers, and child lookup descends the validated tree by cached exact keys. | Add deep-directory, MiniFAT-heavy, concurrent-read, and real DOC/XLS/PPT scenarios beyond the measured synthetic wide-root and writer corpora. |
| 12 | Confirmed for generic detection; disproved for focused prepared iWork detection. | Generic detect-then-open versus prepared-source handoff. |
| 13 | Measured for ODS snapshots: one package clone and duplicate package parse were removable; implemented without changing readback. | Broader ODF source-backed read and unchanged-member publication profiles. |
| 14 | Confirmed for DOCX direct-body batches: repeated full XML rebuild/parse work was removable while retaining ordinary durable operations and complete readback. | Real-producer/extension/security corpora and broader structural/bulk edit semantics. |
| 15 | Confirmed structurally for RTF text aggregation, but unmeasured: a temporary slice vector precedes final string construction. | Native RTF public semantic baseline before any implementation change. |

## Ranked work queue

The order below is provisional until baseline measurements are recorded.

| Rank | Candidate | Expected CRUD reach | Risk | ADR fit |
|---:|---|---|---|---|
| 1 | Extend source-backed OPC from selective reads to broad query/edit/patch coverage. | All OOXML selective read/query/edit paths; offsets the measured exact-source peak-memory cost. | High | Positional source/descriptors are implemented; cache Budget charging and CRUD coverage remain. |
| 2 | Extend targeted OPC preservation to source-backed editable packages and remove regenerated-payload copies. | Targeted OOXML save, especially media-heavy packages; reduces current peak-memory tradeoff. | High | Owned same-topology path is accepted; topology fallback and framing preservation are tested. |
| 3 | Tune explicit bounded-session thresholds and complete remaining I/O budget policy. | Large multi-Part open/save/validation. | Medium-high | 1/2/4/8/12 evidence exists; large tasks scale, small tasks regress; no hidden Rayon path remains. |
| 4 | Build one validated OPC publication plan and reuse its generated XML and Part order during emission. | Every rewritten OPC save. | Low-medium | Implemented; see `changes/0001-opc-publication-plan.md`. |
| 5 | Exact owned-source OPC no-op publication. | Owned DOCX/PPTX/XLSX open/read/no-op save. | Medium | Implemented; same-topology mutations now use targeted preservation. See changes 0004 and 0008. |
| 6 | Move already-owned XLS/PPT writer buffers into `OleWriter`. | Legacy fresh creation and some rebuilds. | Low | Implemented for XLS/PPT; DOC rejected by measurement. See `changes/0003-legacy-owned-stream-handoff.md`. |
| 7 | Use validated cached CFB sibling-tree descent and reusable sector buffers. | Legacy stream-heavy open/rebuild workflows. | Medium | Implemented; see `changes/0002-cfb-lookup-and-sector-buffers.md`. |
| 8 | Extend the accepted XLSX row-start index to broader selector and edit matrices. | Sparse range queries after sheet load. | Low-medium | Narrow ranges are accepted; preservation/readback gates and broad CRUD coverage remain unchanged. |
| 9 | Coalesce DOCX same-structure paragraph replacements and measure PPTX capture/fingerprint reuse. | 1% semantic document/presentation edits. | Medium-high | Implemented for canonical direct-body DOCX batches and PPTX selected-scene reuse; complete source validation and candidate readback remain. See changes 0010 and 0012. |
| 10 | Charge source-backed cache bytes to hierarchical budgets and measure contention. | Concurrent repeated Part reads. | Medium-high | Weighted bounded eviction and per-entry single-flight are implemented. |
| 11 | Extend ODF beyond the accepted ODS snapshot reuse: source-backed selectors and unchanged-member publication. | ODT/ODS/ODP open/query and changed save. | High | Public semantic baselines now exist; exact no-op and full readback must remain. |
| 12 | Add native RTF public semantic cases, then measure direct full-text aggregation. | RTF open/list/full-text/stream-save/no-op/one-edit. | Low-medium | No harness evidence exists yet; preserve cached text and native forward-only output contracts. |
| 13 | Add legacy DOC/XLS/PPT semantic open/edit/save baselines before further CFB ownership experiments. | OLE2 document CRUD rather than substrate-only insertion. | Medium | Positional CFB exists; the previous DOC move regression requires end-to-end guardrails. |
| 14 | Measure ODT snapshot handoff from existing shared transaction bytes. | ODT no-op and changed edit/save. | Low-medium | Preserve source lineage, limits, complete readback and exact no-op bytes. |
| 15 | SIMD or lock-free work. | Unknown. | High | Deferred until remaining hot loops/locks are measured after work elimination. |

## Evidence still missing

The deterministic harness now records warm latency distributions, confidence
intervals, corpus hashes, complete output validation, and sequential-write
call/byte counts. Targeted `heaptrack` runs also cover allocation count,
temporary allocation count, peak heap, and peak RSS for the implemented
changes. Remaining gaps are:

- Reproducible cold-cache distributions on a controlled host.
- Decompressed and recompressed byte observers. Positional range-request
  distributions now exist for OPC and XLSX, but not yet for every format/source.
- Broad hardware-counter evidence. A matched targeted-OPC run is committed now
  that the environment reports `perf_event_paranoid=1`; stage-1 remains without
  counters and no claim is generalized from the one measured save workload.
- Contention evidence beyond the committed explicit-context scaling curves.
- Format-semantic preservation evidence beyond the generated
  DOCX/PPTX/ODT/ODS/ODP slices and native targeted-OPC raw passthrough corpus.
