# Performance hotspot inventory

Status: source-audited; initial ZIP/OPC and CFB substrate measurements captured
Branch: `feat/office-format-completeness`
Evidence through:
[`change 0039`](changes/0039-docx-source-backed-semantic-publication.md)

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
- The immutable source-backed package now also has one consuming low-level
  same-topology publisher. It validates/materializes only the selected existing
  Part, regenerates that member, raw-copies every other physical member, and
  monitors source version through bounded sequential output. Signed real
  changes and unsupported layouts return typed zero-output refusals. DOCX now
  exposes a guarded exact-source main-document transaction over that publisher:
  raw-MCE identity and main-Part-only operations are required, transfers are
  refused, and PPTX/XLSX facade integration remains absent.
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
  DOCX guard remains neutral; see change 0039.

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

These paths have strong preservation and atomicity tests plus generated-text
timing/allocation evidence. Real-producer, media/dependency, malformed,
security, copied-byte and cold-source matrices remain missing.

## ODF paths

ODT, ODS and ODP ordinary opens eagerly read and parse their ZIP packages.
The opt-in public semantic matrix now measures owned open, listing, one object,
full text, small creation, exact no-op and one supported edit/save across all
three owners. ODT/ODP repeated semantic queries still rescan complete XML.
ODP content-only rich-object operations and ODT content-only paragraph
replacement now reuse checked raw preservation. On the fixed eight-by-2 MiB
ODT corpus, paragraph edit/save p50 falls 95.58%, allocation calls fall 6.71%,
and peak heap is flat. Oversized ODT content and resource-adding or structural
ODF publication retain the established package rebuild.

Direct ODT transaction snapshots now adopt the exact package allocation
created by validation and share it with staging rehydration. This removes two
complete archive copies while retaining both complete semantic parses. On the
same media-rich paragraph case, p50 falls 75.84% and peak heap/RSS remain flat;
the predecessor copy used for reversible operation history remains.

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

A media-rich ODS publication case now adds eight deterministic 2 MiB opaque
resources. Eligible compact `content.xml` replacements raw-copy every other
validated ZIP member; exact local/central-member comparison skips unchanged
payload inflation only when the manifest is also exact. The media-rich
one-cell edit/save falls 4.73% p50, 5.73% mean and 7.65% p95, with peak heap
down 8.78%. The existing medium no-media p50 falls 0.77%. Encryption,
signatures, unsupported layouts and every unproved member retain established
logical rebuild/comparison. See
[`change 0031`](changes/0031-ods-unchanged-media-preservation.md).

A matching media-rich ODP case now adds one source-backed text box beside
eight deterministic 2 MiB opaque resources. Reusing the same accepted common
checked-splice/raw-copy primitive cuts edit/save p50 94.44%, mean 94.43%, and
p95 94.29%; allocation calls move +0.52% and peak heap/RSS stay flat. Exact
patch/inverse behavior, complete slide/rich-content/media readback, and every
common security/layout fallback remain. Resource-adding operations still use
the complete rebuild. See
[`change 0034`](changes/0034-odp-unchanged-media-preservation.md).

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

The harness now measures native DOC/XLS/PPT open/list/one/full/no-op/one-edit
flows over deterministic writer artifacts. From the original baseline, large
one-edit/save p50 was 1.722 ms for XLS, 1.416 ms for DOC, and 0.357 ms for PPT.
XLS changed commit now reuses its already validated CFB editor instead of
discarding one BIFF parse and repeating the CFB open/capture; p50 improves
7.72%. DOC publishes its ordinary WordDocument and table-stream replacements
as one failure-atomic object-editor batch instead of rendering/reopening the
CFB after each stream; p50 improves 10.52%. Both retain their final owner and
independent public-reader reopens. PPT root snapshot capture separately reuses
its first validated CFB open and improves p50 8.78%. Direct text-edit setup now
reuses its full editor preflight for record resolution and improves 14.12% p50;
the broader root one-shape edit/save improves 3.59%. The previous
spare-capacity DOC move remains rejected and must remain an independent writer
guardrail.

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
   package capture, but complete Workbook and exact cell-owner readback still
   dominate the accepted 1.639 ms path.
2. DOC one-paragraph publication originally measured 1.416 ms p50. Batching
   its ordinary two-stream replacement removes one intermediate CFB
   render/reopen, while complete revision, style/property and independent
   document readback remain in the accepted 1.348 ms path.
3. PPT one-shape publication (0.357 ms original p50) retains its complete
   commit and public readback. Root snapshot capture improves from 37.522 to
   34.227 us p50, while the direct text-edit transaction improves from 206.209
   to 177.089 us by removing its second editor open.

See [`change 0015`](changes/0015-native-ole2-semantic-baseline.md),
[`change 0016`](changes/0016-xls-commit-editor-reuse.md), and
[`change 0017`](changes/0017-doc-batched-stream-publication.md), and
[`change 0024`](changes/0024-ppt-slide-order-open-reuse.md), and
[`change 0026`](changes/0026-ppt-text-edit-resolver-reuse.md), and
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

The retained opaque-heavy common case now isolates editor open, candidate
publication, changed final rendering and the chained control at 1.382, 7.979,
5.473 and 26.086 ms p50. The stages are not additive: their sum is only 56.86%
of the end-to-end p50. A narrowly scoped inline recapture-allocation reuse
improved candidate publication 6.49% p50/5.95% mean but the complete operation
only 2.61%/2.30%, with p95 +0.54%; it was fully reverted. See
[`change 0036`](changes/0036-ole-common-stage-attribution.md).

## RTF path

Seven native public cases now cover owned open, lazy paragraph listing, one
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
| 1 | Refined: legacy OPC path and `Read` ingress slurp the source; source-backed ingress is positional. | Cold-filesystem bytes, syscalls, latency and RSS across both modes; deterministic range-source distributions now exist. |
| 2 | Confirmed: ordinary OPC open inflates every admitted Part. | Open/list/one-object scaling against total uncompressed bytes and member count. |
| 3 | Superseded for source-backed OPC: finite weighted LRU, per-entry single-flight and content-free diagnostics exist; legacy eager open does not use that cache. | Cache bytes are not yet charged to the hierarchical Budget; add contention and retention measurements. |
| 4 | Measured: ordinary OPC open is serial and explicit eager open has a local bounded session. Six large ZIP tasks reach 4.52x p50 at 12 CPUs; small tasks regress. | Broader real-package scaling and threshold tuning. |
| 5 | Confirmed: stored entries are CRC-checked then copied. | Stored-media one-Part read and package-open copied-byte/RSS deltas. |
| 6 | Refined by measurement: exact unchanged saves copy the source; owned same-topology mutations raw-copy unchanged entries; changed Parts share their immutable logical payload and validated generated local span without extra copies; the narrow source-backed publisher materializes only its target and raw-copies the rest; borrowed/topology-changing paths rewrite fully, while unsupported source-backed layouts refuse. | Real media-heavy 1% updates, semantic facade integration, signature/topology policies, and attribution of the remaining selected-Part/compressor-buffer memory cost. |
| 7 | Confirmed structurally: duplicate indexes, boxed Parts, source-XML map, and linear fallback exist. | Allocation profiles, type sizes, cache counters and repeated noncanonical lookup. |
| 8 | Refined: source-backed XLSX structural open/list avoids timed reads; selected first/range reads physically overlap only the selected worksheet. | Broader source-backed selectors, edits and real workbook matrices. |
| 9 | Refined by measurement: small XLSX edits scan/rebuild/reparse the complete touched sheet; bounded commits can reuse the validation store for first read, while large sheets fall back cold. Direct writer-local action regrouping was immaterial and reverted. | Attribute larger semantic-planning/emission/readback passes, first/middle/last cells, distinct bulk actions, structural edits, large-sheet retention and commit-versus-save separation without reviving direct regrouping alone. |
| 10 | Plausible but unmeasured: per-cell semantic ownership and transient parse duplication may dominate large stores. | Allocation count/bytes, type sizes, peak RSS and cache-miss profiles. |
| 11 | Refined by implementation and measurement: CFB has positional `SharedOleFile` and bounded bulk reads; MiniFAT parsing and sector reads no longer require the former temporary buffers; child lookup descends the validated tree; native DOC/XLS/PPT semantic baselines, XLS editor reuse, DOC batched publication, PPT root-open reuse and PPT text-edit resolver reuse are accepted. The XLS terminal-render handoff was neutral on large changed saves and regressed exact no-op. The opaque-heavy common case rejected direct shared writer payloads, an editor-wide validated-render cache, and inline recapture-allocation reuse; its new open/publication/finish/end-to-end stage split is non-additive. | Attribute materially different final owner/public-reader work without reviving the rejected handoffs or recapture reuse; add deep-directory, MiniFAT-heavy, concurrent-read, real-producer, and security scenarios beyond generated corpora. |
| 12 | Confirmed for generic detection; disproved for focused prepared iWork detection. | Generic detect-then-open versus prepared-source handoff. |
| 13 | Measured for ODS snapshots: one package clone and duplicate package parse were removable. Same-topology ODS row-local publication and compact ODS/ODP/ODT content raw preservation avoid rebuilding untouched data; repeated ODS cell lookup uses a bounded lazy locator. ODT existing-document and direct-byte snapshots share exact package allocations, and consuming full-text block strings are accepted. Direct final-document adoption remains reverted. All accepted paths retain readback and source lineage. | Broader ODF source-backed reads, repeated ODT/ODP semantic scans, resource-adding/structural publication, real-producer media, remaining predecessor-byte copies, and structural-edit profiles. |
| 14 | Confirmed for DOCX direct-body batches: repeated full XML rebuild/parse work was removable while retaining ordinary durable operations and complete readback. | Real-producer/extension/security corpora and broader structural/bulk edit semantics. |
| 15 | Measured and implemented for RTF full-text, text-only edit/save and ordinary parser paths: temporary fragment/property vectors, per-character writer calls, unconditional full-state cloning, per-character ASCII transport-buffer extensions and twice-decoded ordinary-text delimiter traversal were removable. Raw CP-1252, LZFu and a real-producer watermark have capability-bounded read/no-op coverage, and `relsize` has checked native semantic readback. | Extend the accepted native matrix to formatting/media, malformed/security, more real producers and broad edit scenarios; attribute a distinct remaining frame before another specialization. |

## Ranked work queue

The order below is provisional until baseline measurements are recorded.

| Rank | Candidate | Expected CRUD reach | Risk | ADR fit |
|---:|---|---|---|---|
| 1 | Extend source-backed OPC from selective reads and the narrow consuming publisher to broad query/edit/patch coverage. | All OOXML selective read/query/edit paths; offsets eager full-package work. | High | Positional source/descriptors and one low-level one-Part publication path are implemented; cache Budget charging and semantic CRUD coverage remain. |
| 2 | Integrate the accepted source-backed one-Part publisher into bounded DOCX/PPTX/XLSX transactions and real media/signature/topology matrices. | Targeted OOXML save, especially media-heavy packages; avoids eager all-Part inflate/recompression where the same-topology proof applies. | High | Change 0037 proves raw framing, source versions, signed/unsupported refusal and 4 -> 1 materializations for one low-level Part; semantic closure selection, explicit signature policy and topology handling remain. |
| 3 | Tune explicit bounded-session thresholds and complete remaining I/O budget policy. | Large multi-Part open/save/validation. | Medium-high | 1/2/4/8/12 evidence exists; large tasks scale, small tasks regress; no hidden Rayon path remains. |
| 4 | Build one validated OPC publication plan and reuse its generated XML and Part order during emission. | Every rewritten OPC save. | Low-medium | Implemented; see `changes/0001-opc-publication-plan.md`. |
| 5 | Exact owned-source OPC no-op publication. | Owned DOCX/PPTX/XLSX open/read/no-op save. | Medium | Implemented; same-topology mutations now use targeted preservation. See changes 0004 and 0008. |
| 6 | Move already-owned XLS/PPT writer buffers into `OleWriter`. | Legacy fresh creation and some rebuilds. | Low | Implemented for XLS/PPT; DOC rejected by measurement. See `changes/0003-legacy-owned-stream-handoff.md`. |
| 7 | Use validated cached CFB sibling-tree descent and reusable sector buffers. | Legacy stream-heavy open/rebuild workflows. | Medium | Implemented; see `changes/0002-cfb-lookup-and-sector-buffers.md`. |
| 8 | Extend the accepted XLSX row-start index and bounded validated-store handoff to broader selector and edit matrices. | Sparse range queries and first reads after eligible changed-sheet commits. | Low-medium | Narrow ranges and bounded commit/read reuse are accepted in changes 0006 and 0025; dense-wide handoff is intentionally excluded, and preservation/readback gates and broad CRUD coverage remain unchanged. |
| 9 | Coalesce DOCX same-structure paragraph replacements and measure PPTX capture/fingerprint reuse. | 1% semantic document/presentation edits. | Medium-high | Implemented for canonical direct-body DOCX batches and PPTX selected-scene reuse; complete source validation and candidate readback remain. See changes 0010 and 0012. |
| 10 | Charge source-backed cache bytes to hierarchical budgets and measure contention. | Concurrent repeated Part reads. | Medium-high | Weighted bounded eviction and per-entry single-flight are implemented. |
| 11 | Extend ODF beyond accepted ODS snapshot, row-local reuse, ODS/ODP/ODT unchanged-member publication, adaptive cell lookup and ODT byte/full-text ownership: source-backed selectors, repeated ODT/ODP scans, remaining predecessor-byte work, resource-adding/structural publication and real-producer media. | ODT/ODS/ODP open/query and changed save. | High | Same-topology ODS row splicing, compact ODS/ODP/ODT content raw preservation, bounded facade lookup, direct/existing-document ODT snapshot sharing and consuming full-text blocks are accepted; final-document adoption remains reverted for a read regression; structural fallback, exact no-op and full readback remain. See changes 0011, 0014, 0018, 0019, 0020, 0023, 0027, 0031, 0034, 0035 and 0038. |
| 12 | Extend accepted native RTF work beyond the capability-bounded variant matrix after parser-state, transport batching and byte-delimiter scanning. | RTF formatted/media, malformed/security, broader real-producer and broad edit paths. | Medium | Plain, raw CP-1252, LZFu and producer-watermark read/no-op inputs plus a narrow native shape-text chain are covered; only plain generated paragraph editing is timed. Cached text, byte-valued fallback, revisions and native forward-only output contracts remain. See changes 0013, 0019, 0020, 0029 and 0040. |
| 13 | Attribute and reduce remaining native XLS/DOC final-publication work. | OLE2 spreadsheet/document edit publication rather than substrate-only insertion. | Medium-high | Editor reuse and DOC stream batching are accepted in changes 0016/0017; PPT root-open reuse is accepted in 0024. XLS terminal-render, shared CFB writer payload, editor-wide validated-render and inline recapture-allocation prototypes are rejected in 0028/0033/0036. The 4x4 MiB common stage/control cases remain; exact patches, complete BIFF/CFB or DOC validation and independent public readback remain. |
| 14 | Share existing ODT transaction bytes when a validated document creates a snapshot. | ODT no-op and changed edit/save. | Low-medium | Implemented with private `Arc` identity proof; no-op p50 -18.51% large, guardrails within 3%. See change 0014. |
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
  DOC/XLS/PPT/DOCX/PPTX/RTF/ODT/ODS/ODP slices and native targeted-OPC raw
  passthrough corpus.
