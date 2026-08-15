# Performance hotspot inventory

Status: source-audited; initial ZIP/OPC and CFB substrate measurements captured
Branch: `feat/office-format-completeness`
Evidence through: [`change 0144`](changes/0144-cfb-simulated-range-source-evidence.md)
(the newest accepted semantic-format optimization remains limited to four repeated
source-backed ODP full-text projections; the latest release CFB selective-range
evidence is the configured simulator result in
[`0144`](changes/0144-cfb-simulated-range-source-evidence.md), while
[`0094`](changes/0094-cfb-selective-read-evidence.md) retains the non-simulated
exact-range result, and the latest accepted
generic multi-format filesystem result remains
[`0089`](changes/0089-filesystem-release-repeated-evidence.md); Change 0143 is
the latest accepted before/after CFB filesystem result)

This document records facts established by source inspection. It is not a
performance-results report. A path is called a bottleneck only after the
process benchmark and profiler evidence in `BASELINE.md` confirms its effect
on a named corpus and scenario.

## Current resource observations (change 0115)

The [current-HEAD resource profile](results/resource-profile-current-head-0115.json)
adds process-total evidence for a narrow set of named paths.  It does not
promote any observation to a production bottleneck because heaptrack includes
startup and synthetic corpus construction, while `strace` covers the whole
process.

- The managed XLSX batch run recorded 6,130,956 allocation calls and
  1,026,348,498 allocated bytes in one heaptrack process profile.  This is a
  strong candidate for operation-attribution work, not proof that the timed
  edit itself owns all those allocations.
- The OPC source one-Part profile recorded 549 logical source reads and
  16,785,201 logical source bytes, while the CFB save profile recorded 1,825
  logical reads and 84,838,500 bytes before publishing 16,913,408 bytes.  The
  CFB read/output ratio is a concrete measurement target; it is not a physical
  disk-I/O claim.
- The existing bounded RTF stream retained zero output bytes and a 37-byte
  authoring window in the harness.  Heaptrack still observed 450,852 process
  allocation calls, so any further RTF work should separate authoring from
  corpus setup before changing the streaming path.
- Explicit 1/2/4/8/available execution-context runs on many-small OPC and CFB
  corpora were classified `nonideal_or_measurement_noise`: raw p50 showed no
  measured speedup and out-of-range Amdahl fractions are invalidated rather
  than treated as serial-fraction estimates.  The result supports investigating
  task granularity and serial work, but does not justify adding parallelism or
  changing the execution API.

The CFB read-amplification breakdown is now captured and its bounded
fingerprint-request hypothesis is accepted in Change 0143: logical bytes stay
84,838,500 while calls fall from 1,825 to 857, with both clean ABBA directions
improving warm and advisory-cold p50/p95/mean. The next evidence-oriented
priorities are operation-scoped allocation profiles for CFB and managed XLSX,
block-backed physical-cold/high-latency CFB evidence, and a CPU-pinned repeated
scaling run with uncertainty. None should be treated as an optimization
acceptance gate until matched controls and preservation gates exist.

## CFB fingerprint read coalescing (change 0143)

Complete CFB overlay fingerprints now use a right-sized request window capped
at 1 MiB; comparison and publication remain at 64 KiB, the buffers do not
overlap, and no fingerprint or stable-token validation stage is removed. A
clean CPU-2 `A1, B1, B2, A2` release run with 200 samples per warm and
advisory-cold state reduces exact logical requests 53.0411% (1,825 -> 857) with
unchanged logical bytes, output hash and one-span publication. Warm p50 improves
3.3327%/1.3163% and advisory-cold p50 10.7679%/9.4641%; p95 and mean agree in
both directions. A matched whole-process RSS boundary found no candidate
increase, but operation-only allocation/peak memory, physical I/O and proven
cold-storage claims remain open.

## ODP repeated full-text projection (change 0140)

The production threshold-two cache removes two of four complete semantic text
projections in the matched `SourceBackedPresentation` selector shape. A clean
CPU-2 `A1, B1, B2, A2` release run accepts p50 reductions of 45.80%/46.32% and
p95 reductions of 45.25%/45.83%; p99 and mean agree. Whole-process Heaptrack
allocation calls fall 14.31% and temporary allocations 17.25%, but peak heap
is unchanged at 89.22M and process VmHWM is near-neutral. The prepared-source
replay performs zero post-preparation reads, so this is parse/projection/cache
work rather than physical-I/O or decompression evidence. Broader slide-object,
single-call, open, edit/save, real-producer, and generic ODF work remains in the
ODF queue.

## Rejected XLSX publisher provenance reuse (change 0141)

A private lineage/version fast path was tested across calculation metadata,
defined names, page breaks, page margins, page setup, print options, and sheet
protection. It skipped the publication-time semantic reload but left the raw
ZIP publication path unchanged. Clean CPU-2 `A1, B1, B2, A2` evidence found a
1.04% regression in the pooled seven-case p50 geometric mean; calculation
metadata regressed 3.84% p50, and paired directions were mixed. Whole-process
allocation calls fell only 2.84%, temporary allocations 2.12%, peak heap was
unchanged, and VmHWM moved less than 1%. The production change was fully
reverted. Future XLSX publication work should target physical output work,
broad graph validation, whole-Part reconstruction, or a materially larger
semantic parse instead of reintroducing generic provenance fields to these
seven snapshots.

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
- The immutable source-backed package now also has one consuming low-level
  same-topology publisher. It accepts at most 64 unique selected existing
  Parts, validates/materializes that bounded set, regenerates only changed
  selected members, raw-copies every other physical member, and
  monitors source version through bounded sequential output. Signed real
  changes and unsupported layouts return typed zero-output refusals. DOCX now
  exposes a guarded exact-source main-document transaction over that publisher:
  raw-MCE identity and main-Part-only operations are required, transfers are
  refused. PPTX now exposes the analogous guarded exact-source selected-slide
  transaction: its raw package/presentation/slide relationship closure is
  bound into the snapshot, MCE-rewritten slides and more than one shape edit
  per selected slide are refused. A bounded outer batch now composes up to 32
  exact slide snapshots into one atomic multi-Part publication. XLSX now has a
  narrower guarded transaction for typed
  calculation properties/features or the direct defined-name catalog in
  `xl/workbook.xml`; cells, formulas, chains and topology remain outside those
  capabilities. Selected-worksheet
  variants now bind the workbook relationship and worksheet owner for direct
  typed page breaks, page margins, print options, relationship-free page
  setup, complete sheet-protection metadata, or typed core/Office 2010 data
  validations; all materialize two Parts and refuse wider worksheet/topology
  changes. Auto-filter and core conditional-formatting variants additionally
  bind styles/DXF state and materialize three Parts.
- `PackageWriter` previously reconstructed generated XML and Part order during
  emission. The measured `PublicationPlan` change now constructs, audits, and
  reuses that state once. It reduced allocation calls by 37.0% in the profiled
  256-Part save and mean latency by 5.49% in the 2,048-Part compressible save;
  full-Part recompression remains unchanged on the fallback path. Targeted raw
  publication separately improves p50 by 58-96% across the synthetic cells,
  while retained-source peak heap originally rose 37%; see change 0008. The
  changed Part now shares its existing immutable logical payload with the ZIP
  regeneration layer, removing one measured 4.19 MiB copy and reducing the
  matched peak by 3.42%; see change 0021. After validation, the ZIP layer also
  moves that entry's generated local span instead of cloning it, removing a
  second 4.20 MiB allocation and reducing matched peak heap another 3.20%; see
  change 0022. The source-backed one-Part path then removes three unselected
  Part materializations/recompressions on the four-Part corpus, reducing p50
  73.12%, instructions 65.42% and peak heap 3.20%; see change 0037. Complete
  physical archive input/output and the selected-Part compressor buffer remain.
  The DOCX facade integration then removes eager ownership and recompression of
  16 unselected Parts in the media-rich one-edit/save case: p50 falls 97.43%,
  instructions 74.91%, and semantic materializations 17 -> 1 while the eager
  DOCX guard remains neutral; see change 0039. The PPTX facade integration
  removes eager ownership and recompression of 227 unselected Parts in the
  fixed media-rich one-slide edit/save case: p50 falls 97.12%, instructions
  67.91%, and semantic materializations 229 -> 2 with byte-identical output;
  see change 0044. The atomic eight-slide follow-up regenerates eight selected
  slide members in one plan, cuts p50 95.78%, allocations 32.54%, peak heap
  8.94%, and materializations 229 -> 9 while preserving byte-identical output;
  see change 0077. The XLSX calculation-metadata integration removes eager
  ownership and recompression of 11 unselected Parts in its fixed media-rich
  edit/save case: p50 falls 99.2519% (133.67x), instructions 77.78%, and
  semantic materializations 12 -> 1 with byte-identical output; see change
  0046. The defined-name variant also materializes only the workbook Part and
  cuts p50 97.84% (46.32x), instructions 78.45% and materializations 12 -> 1;
  see change 0076. The selected-worksheet page-break, page-margin, print-options and
  relationship-free page-setup, sheet-protection and data-validation
  variants each materialize only the workbook catalog and target worksheet; on
  their matched media-rich controls, p50 falls 97.86%, 97.93%, 97.87%, 97.78%,
  97.75% and 97.75%, respectively; see changes 0061, 0067, 0070, 0073, 0078
  and 0079.

The committed managed source-cache change (`f8d417ac3`) charges exact
physical `InputBytes`, cumulative declared cold-load `Work`, retained
catalog/flight/payload `Objects`, and retained/in-flight payload `Memory` to the
caller's hierarchical `Budget`; compatibility constructors retain the finite
unmanaged `SourceCacheLimits` behavior. Focused correctness tests cover these
resource charges, retained-resource releases, flights, waiters, pinning, eviction,
cancellation, sibling competition, and contention invariants. Release
contention accepts no managed-versus-control speedup; allocation/peak-memory/
RSS, hardware, copied/decompressed-byte, CPU-utilization, and production-performance
evidence remain open.

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
     -> hand off the validated Store only below the cell/XML retention bounds
        and only when final Part and style/shared-string identities still match
  -> save
     -> recompress complete package
```

Confirmed source facts:

- `litchi-xlsx` now exposes bounded forward-only creation for one-sheet
  workbooks. Its correctness and resource-limit tests do not establish latency,
  allocation, peak-memory, or large-stream performance; matching harness and
  profiler evidence remains pending. This is distinct from the source-backed
  existing-cell publication result below.
- The legacy eager path still materializes all admitted Parts. The additive
  source-backed XLSX facade avoids timed source reads while listing after open;
  managed source-backed OPC caches charge exact physical `InputBytes`,
  cumulative declared cold-load `Work`, retained catalog/flight/payload
  `Objects`, and retained/in-flight payload `Memory` to a caller's hierarchical
  `Budget`, preserve externally pinned handles, and coordinate same-Part cold
  loads through one flight. Compatibility opens remain finite under the
  unmanaged `SourceCacheLimits` path. Correctness tests cover these resource
  charges, retained-resource releases, budget hierarchy, eviction, pinning, sibling
  competition, cancellation and failure; the release contention ABBA adds
  structural/distribution evidence but accepts no speedup. Allocation,
  peak-memory/RSS, hardware, copied/decompressed-byte, CPU, and
  production-performance evidence remain missing.
- The additive source-backed calculation-metadata editor loads only the
  workbook Part, stages existing typed `calcPr`/feature edits, reparses the
  complete candidate workbook XML, and consumes the commit into the accepted
  one-Part publisher. It recaptures owner/content-type/URI/XML/source-version
  identity before output; MCE projection and changed signed sources refuse.
  The fixed eight-media case improves p50 99.2519% and materializations 12 ->
  1. This does not authorize cell, formula, cached-result, relationship or
  calculation-chain edits.
- The additive source-backed defined-name editor likewise loads only the
  workbook Part. It binds the exact workbook owner/XML and ordered sheet
  catalog, validates global/local name scope, reparses the complete candidate,
  and refuses protected or MCE/unknown catalogs and changed signed sources.
  The media-rich control improves p50 97.84% and materializations 12 -> 1.
- The additive source-backed sheet-protection editor binds the exact workbook,
  selected worksheet and complete outbound worksheet-relationship set. It
  atomically replaces the complete direct core/Office 2010 protection state,
  reparses the result and refuses MCE-selected protection or changed closure.
  The media-rich control improves p50 97.75%, instructions 77.87%, and
  materializations 12 -> 2.
- The additive source-backed data-validation editor binds the same exact
  workbook, selected worksheet and complete outbound relationship closure. It
  atomically replaces complete typed direct core/Office 2010 collections,
  consumes checked post-write readback, and refuses MCE-selected collections
  or changed closure. The media-rich control improves p50 97.75%, instructions
  73.43%, and materializations 12 -> 2; allocation calls remain within policy.
- The additive source-backed auto-filter editor additionally binds the styles
  relationship and differential-format count so value/color/DXF filters and
  sorts cannot publish dangling style references. It replaces one direct
  worksheet filter/sort subtree, refuses MCE-selected or protected state, and
  raw-copies all unrelated Parts. The media-rich control improves p50 97.75%,
  instructions 73.57%, and materializations 12 -> 3.
- The additive source-backed conditional-formatting editor reuses that exact
  workbook/worksheet/relationship/styles closure and atomically replaces the
  complete direct core owner collection. Its matched selectable case uses the
  same typed values and worksheet rewriter in eager and source-backed paths,
  proves byte-identical publication, and records materializations 12 -> 3.
  Balanced ABBA evidence has not yet been retained, so this is not a latency,
  instruction, or allocation result.
- One first cell access parses the entire selected worksheet. The non-evicting
  `OnceLock` retains it for the snapshot lifetime.
- The sparse cell store is row-major and supports binary-search point lookup.
  A compact immutable row-start index now skips preceding rows for narrow
  ranges. The measured range query improves about 80%; full scan and first-cell
  guardrails remain near neutral.
- A targeted cell edit performs a semantic parse, an independent lossless
  layout scan, full replacement-byte construction, and a full changed-sheet
  semantic readback before publication.
- Source-backed scalar-cell publication now carries a tri-state source
  provenance proof from the checked snapshot into the publisher. Matched
  lineage/version avoids a second publication-time semantic worksheet reload;
  mismatched sources refuse and unavailable provenance retains the prior full
  reload/readback path. Balanced release ABBA across one-cell, `ceil(1%)` and
  exact-256 batches accepts p50 geomean improvements of 21.66%/22.65% and p95
  improvements of 21.38%/22.70%. Physical source reads and successful
  materializations are unchanged, so this is a semantic-reparse result rather
  than an I/O claim. See [`change 0096`](changes/0096-xlsx-source-provenance-publication.md).
- Plain worksheets previously ran a separate namespace-aware x14ac collection
  before every complete semantic parse even when no `dyDescent` token existed.
  Successful no-token reads now skip that pass; rejected inputs rerun it to
  preserve error precedence. Medium changed commits improve about 20% and cold
  reads about 35%; dense-wide 1% commit improves 19.62% p50, allocation calls
  fall 25.24%, and peak heap remains flat. Direct x14ac/MCE paths are unchanged.
- Eligible changed sheets now adopt that exact commit-validated store into the
  target snapshot. Medium commit plus first read improves 23.23% p50 and
  allocation calls fall 21.01%. The handoff is capped at 4,096 cells and 1 MiB
  XML; the unrestricted dense-wide prototype was rejected at +8.99% peak heap.
- Bulk cell actions are held in address order, then regrouped into nested
  row/cell `BTreeMap`s during worksheet emission. A direct owned-stream
  replacement removed that regrouping but improved formal 1% commit/save by
  at most 1.61% p50, so it was fully reverted in change 0030.
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

The additive source-backed PPTX editor instead snapshots one selected slide
and its exact package/presentation/slide relationship closure. One operation
may replace one shape or atomically replace up to 256 unique, nonoverlapping
shape texts in a single bounded scan/emission. It consumes the commit into the
source-backed one-Part publisher. The other 199 slides, all eight 2 MiB media
Parts, and every other unselected physical member remain on the raw-copy path.
MCE preprocessing that changes raw slide bytes, duplicate/overlapping batch
selectors, stale or foreign patches, topology changes, and changed signed
sources are refused before publication. The original one-shape case improves
97.12% p50; the matched eight-shape batch improves 97.45%, reduces allocation
calls 39.80%, and retains the 229 -> 2 materialization reduction.

These paths have strong preservation and atomicity tests plus generated-text
timing/allocation evidence. Real-producer, media/dependency, malformed,
security, copied-byte and cold-source matrices remain missing.

Change 0120 adds eight filesystem-isolated ordinary-root PPTX controls over
the 200-slide/eight-text-box/eight-2 MiB-media corpus. The source candidate
uses `litchi::Presentation::open(path)` and the eager control uses a prepared
byte root for query phases; `list_slides` materializes all owned slides and
`selected_slide` uses the selector-first `Presentation::slide(100)` API. A
separate untimed source replay classifies exact compressed ZIP payload-range
overlap: open/count are catalog-only, selected reads only slide 100, and list
reads all slides without media. This establishes a useful correctness and
logical-read guard for the unified facade, but it is not a performance result:
no latency, allocation, RSS, decompression, physical-I/O or cold-cache claim is
accepted before release ABBA. Eager controls explicitly have no source replay.

## ODF paths

ODT, ODS and ODP ordinary opens eagerly read and parse their ZIP packages.
The opt-in public semantic matrix now measures owned open, listing, one object,
full text, small creation, exact no-op and one supported edit/save across all
three owners. ODT indexed paragraph lookup still scans complete XML for
validation, but retains only the requested paragraph. ODP indexed slide lookup
likewise validates styles and content through EOF while retaining semantic text
and completed shapes only for the requested slide; repeated independent ODP
queries still rescan both XML inputs.
ODP content-only rich-object operations and ODT content-only paragraph
replacement, line-break, inline-run, hyperlink, insertion, and removal
operations now reuse checked raw preservation. On the fixed eight-by-2 MiB ODT
corpus, paragraph edit/save p50 falls 95.58%; the
matched line-break path falls 98.17% (54.59x), instructions fall 78.34%, and
allocation calls fall 6.90%. The matched inline-run path falls 98.39% (62.01x),
instructions fall 78.48%, and allocation calls fall 7.00%, with flat peak
heap/RSS. Structural insertion/removal fall 98.20%/98.27% p50 (55.55x/57.86x)
with exact member preservation. Oversized ODT content and resource-adding,
new-style, or richer structural ODF publication retain the established rebuild.
Generic packaged ODF chart-definition replacement now uses the same raw
publisher with an opt-in full payload preflight, retaining the former logical
writer's malformed-member rejection while preserving eligible unchanged ZIP
frames. Existing shared ODT/ODS/ODP raw paths remain lazy. This generic
integration is correctness-only pending matched release evidence; see
[`change 0101`](changes/0101-generic-odf-verified-raw-publication.md).

Existing ODT embedded-resource replacement now has matched selectable evidence
for 64 fixed existing package-backed image owners. The scalar control repeats
`replace_embedded_image` 64 times in one transaction; the bounded batch resolves
the same base-snapshot positions and publishes them through one
`edit_embedded_resources` call. Both reopen to the same complete
paragraph/image projection and retain exact frame names, paths, media types,
payload digests, retained media and untouched raw ZIP members. Case-specific
physical hashes are recorded without requiring scalar/batch byte identity. ODT
exposes no positional-source or logical-Part materialization diagnostics; the
record reports real bounded sink counters only and makes no performance claim
before frozen CPU-pinned balanced ABBA evidence. See
[`change 0085`](changes/0085-odt-embedded-resource-batch-evidence.md).

ODT one-paragraph lookup now has an additive public indexed selector. It keeps
the complete namespace-aware, resource-bounded EOF scan while retaining one
paragraph rather than the 10,000-paragraph collection. Large middle-paragraph
p50 falls 48.56%, allocation calls fall 27.05%, peak heap falls 24.74%, and
uninstrumented RSS falls 10.93%. The established list path remains neutral; a
shared-mode prototype that regressed it was removed. See
[`change 0047`](changes/0047-odt-indexed-paragraph-selector.md).

ODP one-slide lookup now uses a compile-time-specialized selector so the
established full-list parser does not carry a runtime mode. Large middle-slide
p50/mean/p95 improve 4.09%/4.20%/5.18%, whole-process allocation calls fall
3.86%, and the list, full-text, no-op, edit/save and media-save guards remain
within thresholds. Style inheritance, namespaces, shape/animation limits and
tail errors are still checked before return. See
[`change 0049`](changes/0049-odp-indexed-slide-selector.md).

ODP editing snapshots now pass their already validated slide projection into
private transaction staging instead of parsing every slide again from the same
immutable package bytes. Package/security reopening, settings, declarations,
page metadata, raw source-page coverage, isolated draft clones, changed
publication and complete final reopen/readback remain. Large exact no-op
edit/save improves 59.96% p50 and large changed edit/save improves 20.78%;
allocation calls fall 20.13% with flat peak heap/RSS. See
[`change 0060`](changes/0060-odp-snapshot-slide-projection-reuse.md).

Exact slide-only commits now keep the already mandatory parsed candidate until
final publication and move that projection into the immutable snapshot instead
of parsing the same bytes a second time. The independent final package reopen,
raw/compact XML audits and staged-media check still run; any RDF, chart, design,
annotation or rich-content operation retains the ordinary final parse. Large
one-slide edit/save improves 32.35% p50/32.92% mean, allocation calls fall
16.71%, and peak heap/RSS stay flat. See
[`change 0065`](changes/0065-odp-final-snapshot-handoff.md).

Direct ODT transaction snapshots now adopt the exact package allocation
created by validation and share it with staging rehydration. This removes two
complete archive copies while retaining both complete semantic parses. On the
same media-rich paragraph case, p50 falls 75.84% and peak heap/RSS remain flat;
the compactness audit formerly retained further archive-sized copies.

The changed-operation compactness audit now clones the validated predecessor's
private immutable package and borrows the validated candidate package. This
removes three complete archive copies (50.36 MB on the fixed media-rich case)
without removing archive/manifest parsing or compact XML/splice validation.
Edit/save p50 falls 30.44%, mean 31.36% and p95 32.41%; allocation calls fall
0.57% and peak heap/RSS remain flat. Final transaction materialization,
envelope classification and independent reopen/readback remain. See
[`change 0041`](changes/0041-odt-compact-audit-package-sharing.md).

Envelope classification now clones the immutable snapshot package handle
instead of allocating/copying another complete archive. ZIP validation and
manifest/signature/encryption inspection still run. Across two balanced ABBA
cycles on the same media-rich case, p50 falls 11.40%, mean 11.95%, and p95
12.19%; Heaptrack removes exactly two allocations per changed commit with flat
peak heap/RSS. The independent reopen/readback remains.
See [`change 0042`](changes/0042-odt-envelope-package-sharing.md).

Final changed-result publication now clones the already validated document's
private immutable package bytes into the byte-only snapshot. This removes one
16.79 MB copy and one redundant parse while retaining a fresh complete
`after.document()` reopen. Media-rich edit/save p50/mean/p95 improve
22.74%/22.56%/21.48%; allocation calls fall 3.46%, and peak heap/RSS remain
flat. The earlier parsed-final-document retention stays reverted; the guarded
medium one-paragraph path remains within 3% p50/mean and improves p95. See
[`change 0052`](changes/0052-odt-final-result-byte-handoff.md).

ODS unified snapshot construction previously cloned package bytes and parsed
the same ODS package twice: once for package/resource validation and again for
complete `Spreadsheet` readback. It now moves the one validated package into a
crate-private facade constructor. Large no-op edit/save p50 falls 11.78%; the
large changed case improves 2.06% because full spreadsheet rewrite/readback
dominates.

Eligible same-topology worksheet commits now reuse the bounded flat-ODS row
splicer: only changed modeled rows are serialized and untouched source spans
are copied exactly. Large/medium one-cell edit-save p50 falls 9.54% / 7.22%,
allocation calls fall 5.85%, and peak heap falls 27.18%. Structural changes
fall back to full-table replacement; an opaque untouched row is preserved
byte-for-byte, while touching it refuses publication. Compactness, package
reopen, snapshot parsing and complete typed-sheet readback remain mandatory.

The row-local editor now carries its exact checked source ranges through
package emission instead of flattening them and asking the package layer to
rediscover one maximal diff. On the fixed 2,048-cell plus 16 MiB-media case,
this avoids the full-package fallback that recompressed unchanged media.
Edit/save p50/mean/p95 improve 74.16%/74.17%/74.11%; instructions fall 69.04%
and matched peak heap/RSS remain flat. Foreign provenance and unexpected
assembled content refuse. Signatures, encryption-sensitive inputs,
unsupported ZIP layouts, structural edits and every unproved case retain the
established logical rebuild/signature policy. See
[`change 0057`](changes/0057-ods-row-splice-raw-publication.md).

The remaining unified-to-worksheet path formerly copied the exact archive at
each ownership boundary even after row publication stopped recompressing
media. Worksheet snapshots and patches now retain `Arc<Vec<u8>>`, the private
ODS package adopts that owner, and the unified worksheet handoff moves its
source and target allocations through validation with exact failure rollback.
On the same media-rich case, p50/mean/p95 improve
21.32%/21.30%/21.15%; peak heap falls 22.03% and uninstrumented RSS 20.57%.
The durable unified patch boundary and other semantic domains are unchanged.
See [`change 0068`](changes/0068-ods-shared-worksheet-archive-handoff.md).

A media-rich ODS publication case now adds eight deterministic 2 MiB opaque
resources. Eligible compact `content.xml` replacements raw-copy every other
validated ZIP member; exact local/central-member comparison skips unchanged
payload inflation only when the manifest is also exact. The media-rich
one-cell edit/save falls 4.73% p50, 5.73% mean and 7.65% p95, with peak heap
down 8.78%. The existing medium no-media p50 falls 0.77%. Encryption,
signatures, unsupported layouts and every unproved member retain established
logical rebuild/comparison. See
[`change 0031`](changes/0031-ods-unchanged-media-preservation.md).

ODS durable-patch construction formerly copied both exact package archives
into semantic blob bundles even though the outer patch already retained the
same immutable `Arc<[u8]>` owners, then hashed both packages again for
operation preconditions. The bundles now retain those existing allocations
and the preconditions reuse their content addresses. Media-rich one-cell
edit/save p50/mean/p95 improve 8.80%/9.07%/13.85%; the 33.58 MB payload-copy
site disappears and matched peak heap falls 1.92%. ZIP publication,
comparison, compact audit, final reopen and media verification remain. See
[`change 0054`](changes/0054-ods-shared-durable-patch-blobs.md).

ODS content-validation catalog CRUD is separately correctness-covered and
unmeasured. The clone-staged owner supports add/set/update/same-name
replace/remove/clear/rollback, exact no-op and source-checked reversible patch;
the unified document transaction publishes only `content.xml`, raw-preserves
untouched members and fully reopens the result. Referenced removal/clear,
unrepaired dangling references on changed commit, duplicate names, unsafe
rename, opaque/MCE/DTD owners, operation/output bounds and changed signed
packages refuse atomically. This exact closure does not establish a
performance hotspot or broader ODS cell/formula/style/structural capability.

The format-owned validation tranche now has bounded DOCX, PPTX, RTF and XLS
semantic reports in addition to the CFB, OPC and ODF reports. These paths are
finite correctness boundaries, not profiled hotspots. ODF repair remains one
typed non-destructive plan for removing a recognized local-header extra from a
first stored `mimetype` member. One opt-in selector now exercises its bounded
preflight, exact forward/inverse, refusal and zero-retained-output publication
contract, but supplies no latency or total-memory claim. Encrypted, signed,
macro, structural and semantic repairs refuse rather than widening the
preservation boundary.

A matching media-rich ODP case now adds one source-backed text box beside
eight deterministic 2 MiB opaque resources. Reusing the same accepted common
checked-splice/raw-copy primitive cuts edit/save p50 94.44%, mean 94.43%, and
p95 94.29%; allocation calls move +0.52% and peak heap/RSS stay flat. Exact
patch/inverse behavior, complete slide/rich-content/media readback, and every
common security/layout fallback remain. Resource-adding operations still use
the complete rebuild. See
[`change 0034`](changes/0034-odp-unchanged-media-preservation.md).

Existing ODP whole-model replacement now has matched selectable evidence for
eight fixed-name text boxes distributed across eight of 12 slides. The scalar
control repeats candidate staging eight times in one transaction; the bounded
batch resolves and publishes the same set once. Both reopen to the same full
slide/text/rich-content projection and retain exact auxiliary/media payloads.
The batch raw-preserves the manifest, while repeated scalar staging regenerates
it, so their physical output digests differ. ODP exposes no positional-source
or logical-Part materialization diagnostics; the record reports real bounded
sink counters only and makes no performance claim before frozen CPU-pinned
balanced ABBA evidence. See
[`change 0084`](changes/0084-odp-cross-slide-text-box-batch-evidence.md).

Change 0122 adds four opt-in matched selectors over the same 12-slide/eight-
2 MiB `Pictures/` ODP corpus: eager/source-backed open and eager/source-backed
one-middle-slide query. Source timing uses an uninstrumented `OwnedSource`,
while each measured sample has a separate `InstrumentedSource` replay for
exact calls, bytes, coalesced prior-range overlap, and compressed Pictures
overlap (`pictures_read_compressed_range_bytes`), distinct from prior-read
overlap (`source_read_range_overlap_bytes`). Open and one-slide query remain
distinct from a further explicit selected-media replay, which must cover one
complete selected compressed Pictures range and reports bytes outside
Pictures. The summary names compressed ZIP range totals separately from
uncompressed payload bytes/digests; the eager one-slide parity and selected
media checks run outside its timed query. A ZIP-tail catalog request may
physically touch the final Pictures range during open; that overlap is
retained as physical-range evidence and is not treated as media
materialization. Full eager/source semantic parity and deterministic media
digests remain outside timing. The selectors bring the matrix to 233 names
while leaving the default 36 cases / 198 records unchanged. This is
correctness/logical-read evidence only; no latency, decompression, allocation,
RSS, or release-ABBA claim is made.

Change 0123 adds four opt-in unified-root ODP filesystem selectors over the
same media-rich fixture: eager/source-backed open and eager/source-backed
middle-slide query. A temporary corpus is created and written before the
measurement; open timing covers only matching root owner construction, while
query timing covers only an already-open root query. Post-timing gates compare
full root semantics and metadata, source archive/member/hash identity, and
selected media payloads. Source controls pair each sample with a separate
direct typed `SourceBackedPresentation` instrumented replay, so catalog/query
media laziness and exact selected compressed-range coverage are evidence
domains distinct from root timing. This brings the matrix to 237 names while
leaving the default 36 cases / 198 records unchanged. Production routing
tests cover the filesystem handoff; no latency, physical-I/O, decompression,
allocation, RSS, or release-ABBA claim is made.

Change 0124 adds six opt-in ODS unified-root/source selectors over the existing
two-sheet media-rich ODS fixture: eager/source-backed root open, typed
selected-cell, and typed selected-media controls. Corpus and file publication,
eager cloning, and typed owner construction are outside the corresponding
timers. Each sample checks root names/count/text, complete cell and metadata
parity, exact source/archive/member/media identity, and typed ODS readback.
Independent `InstrumentedSource` replays report logical positional calls and
compressed-range overlap separately from uncompressed payload bytes; open
replays avoid unrelated media, selected-cell replay adds no reads after
content preparation, and selected-media evidence pairs an all-Pictures replay
with a selected-range-only replay, requiring both to cover exactly one
compressed member range and excluding other media. Eager source vectors are
empty. The six selectors bring the matrix to 243 names while leaving the
default 36 cases / 198 records unchanged. This is correctness/logical-range
evidence only, with no latency, physical-I/O, decompression, allocation, RSS
or release-ABBA claim.

Repeated `Spreadsheet::cell` scans previously linearly walked physical row and
cell runs for every coordinate. A new opt-in lookup-only sweep attributes the
cost, and the immutable facade now builds a sheet-aligned locator only after 64
successful queries. The large sweep falls 81.74% p50 and the existing
full-cell-text aggregate falls 52.65%; the dense locator requests 3,216 bytes,
is capped at 4 MiB, and peak heap/RSS stay flat. Repeated runs use cumulative
endpoints without expanding logical cells; point queries remain on the linear
path. See [`change 0027`](changes/0027-ods-adaptive-cell-locator.md).

Adopting an already parsed target package directly into the worksheet snapshot
was measured separately and fully reverted: large one-cell edit/save p50
improved only 0.44%, while p95 regressed 0.30%. Package/readback work remains a
hotspot, but that ownership handoff is not a material optimization.

ODT transaction snapshots created from an already validated `Document`
previously allocated and copied the complete package solely to establish the
snapshot owner. They now clone the package's private immutable `Arc` after the
same transaction size check. Large no-op edit/save p50 falls 18.51%, and
Heaptrack attributes exactly two fewer allocations and no package copy to each
snapshot; changed edit/save and unrelated open guardrails remain within 3%.
Direct snapshot byte ingress, full changed-package publication/readback, and
signed/encrypted envelope behavior are unchanged.

ODT full-text extraction now selects a private consuming parser mode: each
parser-created validated block string moves into its element, then into the
final text instead of being cloned at both boundaries. Repeated large-corpus
ABBA improves 3.25% p50 and 4.81% mean; process allocation calls fall 15.48%
and temporary allocations 45.52%, with peak heap and uninstrumented RSS flat.
Public structured block/list queries keep their original path and remain near
neutral. The unchanged open guard moves +3.94% p50/+4.17% mean; its +10.95%
p99 trigger is retained in change 0023. Repeated semantic scans, source-backed
reads and changed-member publication remain.

## Legacy CFB data path

```text
Read + Seek or positional `ReadAt`
  -> header and complete FAT
  -> complete directory bytes
     -> structural validation pass
     -> public entry decoding pass
  -> complete MiniFAT metadata
  -> validate every stream allocation chain
  -> semantic DOC / XLS / PPT owner
     -> lookup a child by cached validated sibling-tree keys
     -> materialize selected stream Vecs or a bounded caller-owned range
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
- Public `SharedOleFile::read_stream_range` now has pinned release ABBA evidence
  against legacy full-stream materialization. For the final 36-byte MiniFAT
  target, one physical source request falls from 261,184 to 36 bytes among 256
  siblings and from 2,096,192 to 36 bytes among 2,048 siblings. Read-stage p50
  improves 95.1%/94.8% and 99.2%/99.2% across the two ABBA directions; p95
  improves 94.4%/94.8% and 98.9%/99.1%. Total p50 moves 8.4%/14.2% and
  6.6%/11.9%. The 4 MiB FAT controls retain one request and one call; paired
  read and total p50 changes stay within 5% control drift. FAT p95/p99 and all
  MiniFAT p99 tails are not accepted. p99, cold
  filesystem, simulated high-latency range, allocation, and peak-RSS claims
  remain open. This is substrate evidence only, not DOC/XLS/PPT semantic
  adoption.
- Change 0125 adds a distinct 4095-byte MiniFAT boundary pair over the same
  256- and 2,048-sibling shapes. The target occupies 64 logical 64-byte
  mini-sectors (eight regular 512-byte sectors);
  the matched legacy/positional controls record separate open/read/total
  timing, exact source calls/bytes/range sizes, and payload hashes. The focused
  gate requires legacy source-byte amplification and one exact positional
  4095-byte request, exposing physical-run coalescing without
  making a latency or resource claim. Release ABBA, tails, cold/high-latency,
  allocation/RSS, and native semantic consumers remain open.
- Change 0126 adds eight ordinary-root DOCX filesystem selectors over the
  unchanged 200-paragraph/eight-incompressible-2 MiB-media corpus. The eager
  control times `fs::read` plus `Document::from_bytes`; the source control
  times `Document::open(path)`; prepared-root query selectors time only their
  exact query. Untimed parity covers semantic projections and metadata; exact
  source SHA plus logical OPC part/relationship/content-type/blob-hash gates
  cover package preservation, including media hashes and source immutability.
  A separate typed source replay classifies zero payload overlap at open,
  complete compressed main-document range coverage during query-selector preparation, and
  zero main/media/unselected/core overlap during the query, while recording
  calls, bytes, request sizes, coverage and materializations. This is
  correctness/logical-range evidence only; latency, physical-I/O,
  decompression, allocation, RSS, cold-cache, ABBA, broad-security and
  Markdown-performance claims remain open.
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
- A new source-backed same-length overlay substrate resolves selected existing
  streams through validated FAT/MiniFAT chains, derives bounded sorted physical
  spans, reopens the complete composed CFB and checks every selected stream
  before output. It rechecks source version and exact source/target fingerprints
  around 64 KiB sequential publication. Direct sinks receive typed partial
  progress; path publication uses synced sibling staging and atomic rename. The
  common wrapper retains signed/encrypted/DRM refusal and never falls back to a
  topology-changing render. No DOC/XLS/PPT end-to-end consumer or speed claim
  is adopted yet; generic CFB substrate correctness is not semantic format
  coverage.
- Atomic CFB overlay save now skips only the duplicate post-emission complete
  fingerprint scan. The saved path is mechanically `4N -> 3N`; direct
  `write_to` retains its post-emission scan. A pinned warm release ABBA run on
  the four-megabyte/five-entry corpus reduces exact logical source reads from
  101,751,908 bytes and 2,084 calls to 84,838,500 bytes and 1,825 calls:
  16,913,408 bytes (16.6222%) and 259 calls (12.4280%). All four legs publish
  the same 16,913,408-byte digest. The paired p50 directions are +3.7963% and
  -10.0141%, so no latency/speedup, RSS/allocation, physical-cold, or storage
  claim is accepted. Parent-wall and warm process-I/O counters remain
  descriptive only; see [change 0103](changes/0103-cfb-atomic-save-scan-evidence.md).
- Complete CFB overlay fingerprint scans now coalesce positional requests with
  a right-sized window capped at 1 MiB; comparison/emission stay at 64 KiB and
  no scan is removed. Clean balanced release evidence reduces calls from 1,825
  to 857 with unchanged 84,838,500 logical bytes and accepts p50/p95/mean in
  both directions for warm and advisory-cold states. The maximum code-local
  fingerprint buffer grows by 983,040 bytes; whole-process RSS is neutral in
  the matched boundary, while operation-only allocation and physical-I/O
  claims remain open. See [change 0143](changes/0143-cfb-fingerprint-read-coalescing.md).
- PPT root slide-order capture now passes its package-owned validated
  `OleFile` to independent live-document inspection instead of rebuilding the
  CFB index. Large root-open p50 improves 8.78% and allocation calls fall
  5.01%; the stream/current-user/live-persist and higher-level snapshot checks
  remain.
- Direct PPT text editing now holds its semantic selector result until the
  complete protection/editor preflight succeeds, then uses that editor for
  persisted-record resolution instead of opening the CFB editor a second time.
  Large direct edit/save p50 improves 14.12%; commit-time fresh-editor source
  comparison, publication and complete readback remain.
- The native PPT root now adopts a just-validated private text publication
  only after exact source and slide persist-ID checks. Default-limit root
  one-shape edit/save p50 improves 18.59%; custom limits and every structural
  path retain the complete root reopen.

The harness now measures native DOC/XLS/PPT open/list/one/full/no-op/one-edit
flows over deterministic writer artifacts. From the original baseline, large
one-edit/save p50 was 1.722 ms for XLS, 1.416 ms for DOC, and 0.357 ms for PPT.
XLS changed commit now reuses its already validated CFB editor instead of
discarding one BIFF parse and repeating the CFB open/capture; p50 improves
7.72%. Same-family fixed-width numeric commits now also certify that only the
requested Number/RK/MulRK fields changed, retain untouched worksheet
inventories, and clone only the edited sheet instead of rebuilding the complete
private offset inventory. The large 8,192-cell one-edit/save p50 improves a
further 7.83%; the complete independent public Workbook open/readback remains.
DOC publishes its ordinary WordDocument and table-stream replacements
as one failure-atomic object-editor batch instead of rendering/reopening the
CFB after each stream; p50 improves 10.52%. Both retain their final owner and
independent public-reader reopens. PPT root snapshot capture separately reuses
its first validated CFB open and improves p50 8.78%. Direct text-edit setup now
reuses its full editor preflight for record resolution and improves 14.12% p50.
Checked adoption of that result then improves root one-shape edit/save 18.59%
p50. The previous
spare-capacity DOC move remains rejected and must remain an independent writer
guardrail.

The source-backed XLS worksheet-visibility overlay landed in committed
production change `bac279116`. Change `0091` adds four opt-in eager/source-backed
scalar and bounded-batch selectors for one-owner and 64-owner visibility edits.
They verify complete worksheet/catalog/opaque-stream readback, exact overlay
bytes, patch/inverse, source fingerprints/spans, and cap/protection refusals.
Change `0095` makes the existing comment and visibility source-backed owners
submit only their exact NOTE/TXO or `BoundSheet8` byte ranges to the common CFB
splice planner. Replacement staging falls from 80,946 bytes to 109/27,904 for
one/256 comments and from 18,166 to 1/64 for one/64 visibility owners. Balanced
ABBA accepts no latency speedup: all source-backed p50 directions remain inside
1.5%, and each workload's largest absolute source-backed delta is below its
largest absolute eager-control delta. Allocation, RSS, peak-memory and physical I/O
remain open, and the complete candidate snapshot/readback remains.

Change `0136` now measures the fixed-width Number/RK/MulRK source-backed path
directly before further production work. On one pinned release process, the
source-backed Number p50 is 146.410 ms versus 31.492 ms eager (4.65x), and the
source-backed RK/MulRK p50 is 1.627 ms versus 0.100 ms eager (16.25x). Both
paths retain complete 16,995,840-byte or 202,752-byte targets and produce
byte-identical family outputs. Source-backed commit and publication phase
medians are 101.618/44.783 ms for Number and 1.117/0.509 ms for RK/MulRK.
This confirms complete target capture/publication as a high-value attribution
boundary; it does not yet isolate an allocation, reopen, hashing, or emission
substage and is not an accepted speedup/regression claim. Change `0137` now
adds two opt-in plan-only selectors that validate a composed target without
retaining a second complete CFB artifact. Their forward-only API deliberately
does not expose patch/inverse; full reopen, fingerprint, sink-failure,
no-op and exact source/target fingerprint preflights, topology and security
proofs remain, and complete bytes are still emitted at publication. Composed
semantic validation may allocate/read a candidate Workbook model, so zero
target-artifact bytes is not a bounded total-memory claim. No latency, memory,
allocation, RSS, or I/O claim is accepted until matched release ABBA evidence
is captured.

An XLS-only immediate handoff of the first validated terminal CFB rendering was
also measured and fully reverted. Tiny changed save improved 7.55%, but large
changed save was neutral at -0.39% p50 and four repeated large exact-no-op
cycles regressed 22.00% p50 / 16.69% mean. Allocation calls fell 0.33% and peak
heap stayed flat, which confirms that work was removed but not that the public
operation improved safely. See
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

## Native OLE2 semantic path

The native semantic matrix separates ordinary reader facades from exact-source
transaction owners. Open is timed explicitly; list/one/full operations start
from opened ordinary models; no-op and one-edit publication start from opened
DOC body-text, XLS cell-value, or PPT slide-order snapshots and include owned
output materialization. Complete semantic verification and patch/inverse
checks stay outside timing.

Measured large-corpus priorities:

1. XLS one-cell publication originally measured 1.722 ms p50. Reusing the
   rendered/reopened CFB editor removes a discarded BIFF parse and redundant
   package capture. Fixed-width numeric inventory carry-forward then reduces
   the current large path another 7.83% p50 while retaining exact byte-range
   proof and independent public readback. Complete Workbook validation, common
   CFB publication, patch construction and output materialization remain.
2. DOC one-paragraph publication originally measured 1.416 ms p50. Batching
   its ordinary two-stream replacement removes one intermediate CFB
   render/reopen, while complete revision, style/property and independent
   document readback remain. A later profile found repeated linear physical
   PieceTable scans at 36.89% of large-open self cycles. The accepted FC index
   reduces large open from 790.727 to 348.679 us p50 and the changed edit/save
   path from 1.379 to 0.950 ms p50 while retaining all FKP/readback validation.
   A subsequent profile found another 6.94% self cycles in repeated paragraph
   style resolution and validation. The accepted one-entry resolved-baseline
   cache reduces the current large open from 343.503 to 304.199 us p50 and
   allocation calls 18.61%, while every direct PAPX and style switch remains
   scalar and independently validated. A later CHPX profile attributed 7.56%
   of process self cycles to paragraph character-run extraction. The accepted
   monotonic range slice reduces the 512-paragraph list from 454.100 to 358.414
   us p50 and the frame to 1.23%, without adding storage or allocations. The
   next exact-source profile found two ordered containment tables restarted
   from the beginning for every paragraph terminator. Predecessor binary
   searches reduce the already-open 512-paragraph snapshot list from 206.644
   to 168.142 us p50 and the full one-edit/save path from 888.602 to 817.424 us
   p50; allocation calls and peak heap remain flat.
3. PPT one-shape publication (0.357 ms original p50) retains its complete
   text-owner commit and public readback. Root snapshot capture improves from
   37.522 to 34.227 us p50, the direct text-edit transaction improves from
   206.209 to 177.089 us, and checked root adoption reduces the full operation
   from 352.306 to 286.805 us p50.

Change 0117 adds pinned balanced release probes for native PPT lazy `Pictures`.
On the generated eight-slide/32-image corpus, independent untimed replay proves
that source-backed open reads 79,265 metadata/mandatory-stream bytes with zero
`Pictures` overlap, the cold all-images query reads the complete 8,389,408-byte
stream once, and a cached query adds no reads. A directly timed fresh
open-plus-all-images pair prevents misleading sums of phase medians. Both the
100-sample preflight and 200-sample/cooldown attempt failed the fixed
same-implementation drift gates, so no latency result is accepted. Allocation,
RSS attribution, cold-cache, producer-breadth, and save-path evidence remain
open.

An additive source-backed DOC owner now covers one ordinary Word97+ main-story
paragraph when its text and terminating mark are contained in one uncompressed
Unicode piece. Positional selection uses bounded chunks and same-width
`WordDocument` splicing, with exact no-op/source/fingerprint/stale checks,
candidate reopen/readback, inverse and typed partial-output coverage. Complete
artifact fingerprints and CFB validation/publication scans remain mandatory;
this is correctness/selector coverage only, with no end-to-end latency,
physical-I/O/range, allocation/RSS, cold/high-latency, real-producer or broad
DOC CRUD claim. See [`change 0105`](changes/0105-doc-source-backed-paragraph-splice.md).

See [`change 0015`](changes/0015-native-ole2-semantic-baseline.md),
[`change 0016`](changes/0016-xls-commit-editor-reuse.md), and
[`change 0017`](changes/0017-doc-batched-stream-publication.md), and
[`change 0050`](changes/0050-doc-piece-table-physical-index.md), and
[`change 0051`](changes/0051-doc-adjacent-style-baseline-cache.md), and
[`change 0053`](changes/0053-doc-chpx-range-index.md), and
[`change 0056`](changes/0056-doc-papx-containment-index.md), and
[`change 0105`](changes/0105-doc-source-backed-paragraph-splice.md), and
[`change 0024`](changes/0024-ppt-slide-order-open-reuse.md), and
[`change 0026`](changes/0026-ppt-text-edit-resolver-reuse.md), and
[`change 0062`](changes/0062-ppt-root-text-publication-adoption.md), and
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

The retained opaque-heavy common case now isolates editor open, candidate
publication, changed final rendering and the chained control at 1.382, 7.979,
5.473 and 26.086 ms p50. The stages are not additive: their sum is only 56.86%
of the end-to-end p50. A narrowly scoped inline recapture-allocation reuse
improved candidate publication 6.49% p50/5.95% mean but the complete operation
only 2.61%/2.30%, with p95 +0.54%; it was fully reverted. See
[`change 0036`](changes/0036-ole-common-stage-attribution.md).

## RTF path

`litchi-rtf` now also exposes bounded forward-only authoring. Escape-free
printable ASCII is emitted in direct spans capped at 32 bytes, without adding a
retained writer buffer and while retaining Work/Output reservations and
per-write cancellation checks. Balanced release ABBA accepts p50 geomean
improvements of 76.41%/76.47% and p95 improvements of 75.23%/75.76%; the large
case drops from 7,208,970 to 1,441,802 sink calls with exact bytes and hashes.
This is fresh creation only, and allocation, peak-memory/RSS and cold-I/O
evidence remain open. See
[`change 0097`](changes/0097-rtf-bounded-ascii-streaming.md).

Existing-document logical-tail append is now a separate, opt-in harness path
from streaming creation. Tiny/medium/large plain corpora append 4/64/256
bounded one-run paragraphs and verify candidate reopen, exact sequential bytes,
durable patch/inverse and foreign-source refusal. The fixed 16 KiB hashing-sink
window caps accepted bytes per write and retains zero output, but does not bound
the transaction's validated candidate snapshot. This is correctness/coverage
evidence only; no release latency, allocation, RSS, or speedup claim exists.
See [`change 0090`](changes/0090-rtf-logical-tail-append-evidence.md).

The standalone harness now records a fixed six-bucket distribution for every
serialized sink summary. It counts logical `Write::write` calls at the point
where bytes are accepted, includes zero-length calls, excludes rejected calls,
and checks that the bucket total equals `write_calls` at the exact inclusive
boundaries. This is reporting evidence only: it does not measure syscalls, disk
I/O, memory copies, compression, latency, allocation, RSS, or performance.
See [`change 0107`](changes/0107-output-write-size-evidence.md).

The existing seven native public cases cover owned open, lazy paragraph listing, one
paragraph, first complete text, exact stream save, exact empty-edit save, and
one checked paragraph edit/save over 24/200/10,000-paragraph corpora. The
unified root path remains intentionally outside this evidence.

`RtfDocument` now retains the total block-text byte length during its existing
owned-detach pass, so first full-text materialization performs one exact
allocation and one block pass instead of allocating and joining a temporary
fragment vector. The large full-text p50 improves 27.08%. Canonical text
emission now writes contiguous ASCII spans rather than formatting one character
per sink call, and text-only commits skip unused paragraph-property vectors and
scans. Together with early successful paragraph selection, large one-edit/save
p50 improves 25.79%. Full-text caching, forward-only sink errors, exact no-op
identity, opaque refusal, validation, and complete reopen/readback remain.

The next parser profile found a full `State::clone` on every ordinary body-text
flush at 8.53% exclusive samples. Ordinary flushes now borrow the state and
copy only effective encoding, formatting and paragraph properties; the full
state is retained only for insertion/deletion metadata. Large open p50 improves
20.09% and large one-edit/save improves 11.54%, with flat allocation count,
peak heap and RSS. Code-page selection, revision ranges and deletion behavior
have focused and complete-suite coverage.

The following profile attributed 15.37% of large-open and 14.46% of large
one-edit/save samples to extending parser transport buffers one byte at a time.
All-ASCII source tokens now enter those buffers in one extension; byte-valued
non-ASCII and invalid-Unicode input retain the checked per-character fallback.
Large open p50 improves 26.67% and one-edit/save 6.26%. Instructions fall
18.40%; allocation count, peak heap and RSS remain flat.

The next matched profile retained 17.36% exclusive large-open cycles in
`Lexer::tokenize_with_spans`: ordinary text decoded and advanced over every
UTF-8 scalar twice merely to find five ASCII delimiters. One checked byte scan
now finds those delimiters while retaining UTF-8 boundaries and exact source
spans. Large open p50 improves 17.23%, one-edit/save 14.65%, instructions fall
21.27%, and the lexer frame falls to 11.06%. Medium/large plain, raw CP-1252
and LZFu opens all improve; the prepared LZFu no-op microsegment exception is
disclosed in change 0040.

A follow-up attempted to move decoded block ownership directly into the final
document. The broad version removed 20.15% of process allocation calls and
improved raw CP-1252 open 3.08% p50, but moved ordinary ASCII allocation into
the parser loop and regressed plain large open 25.53% p50. Owned-only variants
measured -1.41% and +1.02% p50 across separate 4,000-sample/state runs. The
production parser was restored exactly; do not revisit this copy in isolation.
See change 0043.

The next changed-commit profile isolated a separate, short-lived owner: after
the initial complete parse, `ordinary_body_source_span` cloned the 540,051-byte
ASCII source, tokenized it again and scanned root depth before the required
candidate parse/readback. Direct uncompressed ASCII parses now retain a compact
range proven inside the parser's existing structural preflight. Ambiguous,
empty, binary, non-ASCII, compressed and over-32-bit ranges keep the established
locator/refusal. Large one-edit/save improves 10.72% p50 and 10.11% mean;
instructions fall 10.64%, and the 588 before-only locator allocation calls
over 20 edits disappear. Peak heap and uninstrumented RSS remain flat. See
change 0048.

The next large-open profile attributed 25.65% of cycles to `memmove`; the
10,000 retained `StyleBlock` values grew through 12 vector allocations to a
16,384-element / 16.12 MiB capacity. The existing structural preflight now
counts root text tokens and passes a bounded hint to the first retained block.
One 9.84 MiB exact reserve replaces those growth/copy steps. Large open p50
improves 21.17%, mean 21.00%, cycles 14.91%, and cache misses 32.20%; peak heap
falls 29.73%. Table/deletion-heavy and sub-64-KiB sources retain lazy growth.
Medium plain/CP-1252 p50 movements of +0.49%/+2.84% are disclosed. See change
0055.

Formatting/media, malformed/security, broader real-producer, cold-source,
broad edit and conversion matrices remain missing. Compressed LZFu and raw
CP-1252 open/read/no-op coverage is now measured but remains narrow.

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
| 1 | Refined: legacy OPC path and `Read` ingress slurp the source; source-backed ingress is positional. Five filesystem cases now record process-isolated warm/cold-requested counters and atomic-save hashes, including a repeated release tmpfs capture. | Repeat on a controlled block-backed filesystem/cache host. The release run's accepted cold advice and zero process `read_bytes` on tmpfs are counters/output evidence only, not physical cold-cache behavior. |
| 2 | Confirmed: ordinary OPC open inflates every admitted Part. | Open/list/one-object scaling against total uncompressed bytes and member count. |
| 3 | Implemented for managed source-backed OPC: finite weighted eviction, pinned-handle preservation, per-entry single-flight, exact physical `InputBytes`, cumulative declared cold-load `Work`, retained catalog/flight/payload `Objects`, retained/in-flight payload `Memory` Budget charging and content-free diagnostics exist; compatibility/unmanaged opens retain finite `SourceCacheLimits`, and legacy eager open does not use that managed cache. | Correctness tests cover all managed resource dimensions and charging/release invariants. Release contention ABBA covers structural/distribution counters but accepts no speedup. Add allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence. |
| 4 | Measured: ordinary OPC open is serial and explicit eager open has a local bounded session. Six large ZIP tasks reach 4.52x p50 at 12 CPUs; small tasks regress. | Broader real-package scaling and threshold tuning. |
| 5 | Confirmed: stored entries are CRC-checked then copied. | Stored-media one-Part read and package-open copied-byte/RSS deltas. |
| 6 | Refined by measurement: exact unchanged saves copy the source; owned same-topology mutations raw-copy unchanged entries; changed Parts share their immutable logical payload and validated generated local span without extra copies; the bounded source-backed publisher materializes only selected targets and raw-copies the rest; guarded DOCX, atomic same-slide and multi-slide PPTX shape-text batches, and XLSX calculation-metadata/defined-name/page-break/page-margin/print-options/page-setup/sheet-protection/data-validation/auto-filter facades consume it; borrowed/topology-changing paths rewrite fully, while unsupported source-backed layouts refuse. | Real-producer media-heavy multi-Part updates, broader semantic closures, signature/topology policies, and attribution of the remaining selected-Part/compressor-buffer memory cost. |
| 7 | Confirmed structurally: duplicate indexes, boxed Parts, source-XML map, and linear fallback exist. | Allocation profiles, type sizes, cache counters and repeated noncanonical lookup. |
| 8 | Refined: source-backed XLSX structural open/list avoids timed reads; selected first/range reads physically overlap only the selected worksheet; guarded calculation-metadata, defined-name, worksheet page-break, page-margin, print-options, relationship-free page-setup, sheet-protection, data-validation and auto-filter edits materialize only their one- to three-Part semantic closures. | Broader source-backed selectors, general cell/formula edits and real workbook matrices. |
| 9 | Refined by measurement: small XLSX edits scan/rebuild/reparse the complete touched sheet; bounded commits can reuse the validation store for first read, while large sheets fall back cold. Direct writer-local action regrouping was immaterial and reverted. | Attribute larger semantic-planning/emission/readback passes, first/middle/last cells, distinct bulk actions, structural edits, large-sheet retention and commit-versus-save separation without reviving direct regrouping alone. |
| 10 | Plausible but unmeasured: per-cell semantic ownership and transient parse duplication may dominate large stores. | Allocation count/bytes, type sizes, peak RSS and cache-miss profiles. |
| 11 | Refined by implementation and measurement: CFB has positional `SharedOleFile`, bounded bulk reads, exact-range reads and exact-range splice publication; MiniFAT parsing and sector reads no longer require the former temporary buffers; child lookup descends the validated tree; native DOC/XLS/PPT semantic baselines, XLS editor and inventory reuse, DOC batched publication and indexes, PPT root-open reuse, text-edit resolver reuse, and checked root text-publication adoption are accepted. Change 0094 accepts only the generic MiniFAT read-stage/source-byte and modest total-p50 evidence. Change 0095 adds native XLS comment/visibility semantic splice consumers with exact replacement-byte evidence but no latency speedup. Change 0102 range-resolves the native PPT one-shape selector but keeps full source fingerprint/publication checks and makes no end-to-end performance claim. Change 0105 adds a correctness-only Word97+ DOC main-story one-paragraph Unicode-piece splice with bounded positional selection, same-width replacement, candidate readback and source/inverse checks; complete CFB fingerprints and validation/publication scans remain. The XLS terminal-render handoff was neutral on large changed saves and regressed exact no-op. The opaque-heavy common case rejected direct shared writer payloads, an editor-wide validated-render cache, and inline recapture-allocation reuse; its open/publication/finish/end-to-end stage split is non-additive. | Attribute materially different final owner/public-reader work without reviving the rejected handoffs or recapture reuse; add deep-directory, MiniFAT-heavy, concurrent-read, real-producer, and security scenarios beyond generated corpora. Cold/high-latency sources and allocation/RSS evidence remain open; the DOC owner still needs matched source/read-range and real-producer breadth before any performance claim. |
| 12 | Confirmed for generic detection; disproved for focused prepared iWork detection. | Generic detect-then-open versus prepared-source handoff. |
| 13 | Measured for ODS snapshots: one package clone and duplicate package parse were removable. Same-topology ODS row-local publication retains exact range provenance through raw ZIP emission, its unified worksheet handoff now shares/moves the exact archive allocation through the nested snapshot and package validation, and compact ODS/ODP/ODT content publication avoids rebuilding untouched data; repeated ODS cell lookup uses a bounded lazy locator. ODT existing-document/direct-byte/final-result snapshots, changed-operation compact audits and envelope classification share exact validated package allocations, consuming full-text block strings and an indexed one-paragraph retention path are accepted, consecutive plain-text replacements publish one candidate, and scalar line-break/run/hyperlink plus plain paragraph insertion/removal use the accepted content-only publisher. A matched release ODT mixed model-content case now measures one staged publication against 49/193 scalar publications over 80/320 operations, with 96.8685%–96.8695% medium and 99.2289%–99.2381% large p50 reduction and equal per-shape hashes; this is a narrow repeated-publication result, not a general ODT/resource/I/O/memory claim. ODP one-slide lookup retains only its requested semantic projection while validating through EOF, ODP transaction staging reuses its snapshot-validated complete slide projection, and exact slide-only commits adopt that already validated candidate only after final package audits. Parsed final-document adoption remains reverted. All accepted paths retain readback and source lineage. | Broader ODF source-backed reads, repeated independent ODP semantic scans, formatted/non-text bulk edits, resource-adding/richer structural publication, real-producer media, and structural-edit profiles. |
| 14 | Confirmed for DOCX direct-body batches: repeated full XML rebuild/parse work was removable while retaining ordinary durable operations and complete readback. | Real-producer/extension/security corpora and broader structural/bulk edit semantics. |
| 15 | Measured and implemented for RTF full-text, text-only edit/save, ordinary parser and already-open story-query paths: temporary fragment/property vectors, per-character writer calls, unconditional full-state cloning, per-character ASCII transport-buffer extensions, twice-decoded ordinary-text delimiter traversal, the second ordinary-body source lexer and repeated full-block length scans were removable. Raw CP-1252, LZFu and a real-producer watermark have capability-bounded read/no-op coverage, and `relsize` has checked native semantic readback. | Extend the accepted native matrix to formatting/media, malformed/security, more real producers and broad edit scenarios; attribute a distinct remaining frame before another specialization. |

## Ranked work queue

The order below is provisional until baseline measurements are recorded.

| Rank | Candidate | Expected CRUD reach | Risk | ADR fit |
|---:|---|---|---|---|
| 1 | Extend source-backed OPC from selective reads and the bounded consuming publisher to broad query/edit/patch coverage. | All OOXML selective read/query/edit paths; offsets eager full-package work. | High | Positional source/descriptors, low-level one-Part/bounded multi-Part publication and managed cache charging across physical `InputBytes`, cumulative declared cold-load `Work`, retained `Objects`, and `Memory` are implemented and correctness-tested; broader semantic CRUD and controlled cache acceptance remain. |
| 2 | Broaden the accepted source-backed DOCX/PPTX and XLSX calculation-metadata/defined-name/page-break/page-margin/print-options/page-setup/sheet-protection/data-validation/auto-filter transactions only where complete semantic closures can be proved, with real media/signature/topology matrices. | Targeted OOXML save, especially media-heavy packages; avoids eager all-Part inflate/recompression where the same-topology proof applies. | High | DOCX is accepted in change 0039, guarded same-slide PPTX in 0044/0063 and bounded multi-slide PPTX in 0077, XLSX calculation metadata in 0046, page breaks in 0061, page margins in 0067, print options in 0070, relationship-free page setup in 0073, defined names in 0076, sheet protection in 0078, data validation in 0079, and auto filters in 0080. Change 0120 adds ordinary-root PPTX open/list/count/selected-slide logical-read controls and complete parity gates, but no speedup/resource claim. General XLSX cells/formulas/chains, table filters, printer settings and structural PPTX edits require wider closures; all accepted facades still need real-producer and broader topology/signature policy matrices. |
| 3 | Tune explicit bounded-session thresholds and complete remaining I/O budget policy. | Large multi-Part open/save/validation. | Medium-high | 1/2/4/8/12 evidence exists; large tasks scale, small tasks regress; no hidden Rayon path remains. |
| 4 | Build one validated OPC publication plan and reuse its generated XML and Part order during emission. | Every rewritten OPC save. | Low-medium | Implemented; see `changes/0001-opc-publication-plan.md`. |
| 5 | Exact owned-source OPC no-op publication. | Owned DOCX/PPTX/XLSX open/read/no-op save. | Medium | Implemented; same-topology mutations now use targeted preservation. See changes 0004 and 0008. |
| 6 | Move already-owned XLS/PPT writer buffers into `OleWriter`. | Legacy fresh creation and some rebuilds. | Low | Implemented for XLS/PPT; DOC rejected by measurement. See `changes/0003-legacy-owned-stream-handoff.md`. |
| 7 | Use validated cached CFB sibling-tree descent and reusable sector buffers. | Legacy stream-heavy open/rebuild workflows. | Medium | Implemented; see `changes/0002-cfb-lookup-and-sector-buffers.md`. |
| 8 | Extend the accepted XLSX row-start index and bounded validated-store handoff to broader selector and edit matrices. | Sparse range queries and first reads after eligible changed-sheet commits. | Low-medium | Narrow ranges and bounded commit/read reuse are accepted in changes 0006 and 0025; dense-wide handoff is intentionally excluded, and preservation/readback gates and broad CRUD coverage remain unchanged. |
| 9 | Coalesce DOCX same-structure paragraph replacements and measure PPTX capture/fingerprint reuse. | 1% semantic document/presentation edits. | Medium-high | Implemented for canonical direct-body DOCX batches and PPTX selected-scene reuse; complete source validation and candidate readback remain. See changes 0010 and 0012. |
| 10 | Measure and tune the managed source-backed cache under controlled contention. | Concurrent repeated Part reads. | Medium-high | Hierarchical charging across physical `InputBytes`, cumulative declared cold-load `Work`, retained `Objects`, and `Memory`, plus pinned-aware eviction and per-entry single-flight, are implemented and correctness-tested in change 0086; release ABBA in 0088 covers structural/distribution counters but accepts no speedup. Allocation, peak-memory/RSS, hardware, copied/decompressed-byte, CPU-utilization and production-performance evidence are open. |
| 11 | Extend ODF beyond accepted ODS snapshot, row-local provenance reuse/shared worksheet ownership, ODS/ODP/ODT unchanged-member publication, adaptive cell lookup, ODP indexed-slide retention/snapshot handoffs and ODT byte/full-text/indexed-query/audit/envelope/batch/final-byte ownership: positional source-backed reads, repeated independent ODP scans, richer non-text/bulk edits, resource-adding/richer structural publication and real-producer media. | ODT/ODS/ODP open/query and changed save. | High | Same-topology ODS row splicing now carries exact range proofs through raw ZIP emission and the adjacent nested worksheet/package owners share and move their archive allocation; compact ODS/ODP/ODT content raw preservation, bounded facade lookup, direct/existing/final-result ODT byte sharing, consuming full-text blocks, indexed paragraph/slide retention, ODP staging and final slide-only snapshot projection reuse, matched ODP text-box and ODT embedded-resource scalar/bounded evidence, compact-audit/envelope sharing, consecutive paragraph coalescing and scalar line-break/run/hyperlink plus plain paragraph insertion/removal publication are accepted. Change 0122 adds matched ODP eager/source-backed media-rich open and middle-slide logical-read selectors with explicit selected-Pictures replay; change 0123 adds matched unified-root eager/source-backed filesystem open and middle-slide controls with complete post-timing semantic/metadata/media/member/hash parity plus direct typed replay evidence; change 0124 adds matched unified-root ODS eager/source-backed open plus typed selected-cell/media controls with complete untimed root/typed/archive/member/hash parity and direct positional-read evidence. These are correctness/range evidence only. ODS content-validation catalog CRUD is correctness-covered but unmeasured. Parsed final-document adoption remains reverted for a read regression; other structural fallback, exact no-op and full readback remain. See changes 0011, 0014, 0018, 0019, 0020, 0023, 0027, 0031, 0034, 0035, 0038, 0041, 0042, 0045, 0047, 0049, 0052, 0057, 0060, 0065, 0068, 0071, 0072, 0074, 0075, 0084, 0085, 0122, 0123 and 0124. |
| 12 | Extend accepted native RTF work beyond the capability-bounded variant matrix after parser-state, transport batching, byte-delimiter scanning, retained ordinary-body ranges, retained story-length/cardinality handoffs and sparse paragraph selection. | RTF formatted/media, malformed/security, broader real-producer and broad edit paths. | Medium | Plain, raw CP-1252, LZFu and producer-watermark read/no-op inputs plus a narrow native shape-text chain are covered; plain generated paragraph queries and editing are timed, public paragraph cardinality is parser-retained, and explicit sparse `nth` no longer constructs discarded paragraph views. Cached full text, byte-valued fallback, revisions, candidate readback and native forward-only output contracts remain. See changes 0013, 0019, 0020, 0029, 0040, 0048, 0064, 0066 and 0069. |
| 13 | Remove the second complete target artifact from fixed-width native XLS publication, then continue attributing remaining OLE2 final-owner/public-reader work. | OLE2 spreadsheet/document/presentation edit publication rather than substrate-only insertion. | Medium-high | Changes 0136/0137 established the source-backed baseline and forward-only plan; change 0138 accepts complete-operation latency for Number and RK/MulRK after strict CPU-2 release A1/B1/B2/A2 (p50/p95/p99/mean agree in both directions). Number process VmHWM also agrees (-10.73%/-10.66%), while RK/MulRK RSS directions disagree and valid heaptrack A/B profiles show descriptive whole-process allocation reductions with identical peak heaps. The accepted result is limited to these deterministic fixed-width families; composed validation may allocate/read a candidate Workbook model, so zero target-artifact bytes is not a bounded total-memory claim. No physical-I/O, cold-cache, operation-only allocation or broad-producer claim is made. |
| 14 | Share existing ODT transaction bytes when a validated document creates a snapshot. | ODT no-op and changed edit/save. | Low-medium | Implemented with private `Arc` identity proof; no-op p50 -18.51% large, guardrails within 3%. See change 0014. |
| 15 | SIMD or lock-free work. | Unknown. | High | Deferred until remaining hot loops/locks are measured after work elimination. |

## Evidence still missing

The deterministic harness now records warm latency distributions, confidence
intervals, corpus hashes, complete output validation, and sequential-write
call/byte counts. Targeted `heaptrack` runs also cover allocation count,
temporary allocation count, peak heap, and peak RSS for the implemented
changes. Remaining gaps are:

- Reproducible physical cold-cache distributions on a controlled host. Change
  0087's one-sample debug warm/cold-requested run and change 0089's repeated
  tmpfs release run are correctness/counter and descriptive distributions only;
  neither proves a cold device or storage result.
- CFB selective-range acceptance is bounded to exact source-byte counters,
  MiniFAT read-stage p50/p95, and the modest total-p50 direction in change
  0094. Change 0144 additionally accepts p50/p95 only for the named configured
  simulated range source: both MiniFAT targets reduce to one exact request and
  improve in both ABBA directions, while the exact-work FAT control stays near
  neutral. Real cold/network/device range sources, FAT tail behavior, p99,
  allocation, and peak-RSS evidence remain open, and no DOC/XLS/PPT semantic
  consumer is covered.
- Decompressed and recompressed byte observers. Positional range-request
  distributions now exist for OPC and XLSX, but not yet for every format/source.
- Broad hardware-counter evidence. A matched targeted-OPC run is committed now
  that the environment reports `perf_event_paranoid=1`; stage-1 remains without
  counters and no claim is generalized from the one measured save workload.
- Cache contention acceptance: change 0088 has release structural/distribution
  ABBA evidence, but no accepted speedup, allocation, RSS, hardware,
  copied/decompressed-byte, or CPU-utilization result.
- XLS visibility overlay performance/resource evidence. Change 0091 is
  correctness/coverage only: it has no release ABBA, speedup, allocation, RSS,
  peak-memory, or physical-I/O claim, and its complete source-backed candidate
  snapshot is not bounded by the 64 KiB publication sink.
- Format-semantic preservation evidence beyond the generated
  DOC/XLS/PPT/DOCX/PPTX/RTF/ODT/ODS/ODP slices and native targeted-OPC raw
  passthrough corpus.
