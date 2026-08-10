# Performance program phase report

Date: 2026-08-11
Branch: `feat/office-format-completeness`
Production base for the latest semantic tranche: `56dfde4fdbe7433f66e2ac3e5a161cf8857c0c3f`

This report summarizes the measured implementation tranches to date. It is not a
claim that the end-to-end performance program or CRUD scenario matrix is
complete. The reproducible environment, original substrate baseline, corpus
definitions, commands, and profiler limitations are in
[`BASELINE.md`](BASELINE.md); raw reports are under [`results/`](results/).

## Current stable tranche

The original stage-1 results below remain historical evidence. The current
harness contains **88 selectable cases**: 36 default cases and 198 default
records, plus six opt-in simulated-range cases, two opt-in scaling cases, 16
opt-in DOCX/PPTX semantic cases, seven opt-in RTF semantic cases, and 21 opt-in
ODT/ODS/ODP semantic cases. It is still not broad program or CRUD coverage.

| Change | Current evidence | Scope / limitation |
|---|---|---|
| XLSX row-start index | ABBA p50 geomean **-80.499%**, mean geomean **-79.962%**; full scan **+0.03%** mean; first cell **-1.31%** mean | Heap allocations **+17**, RSS **+0.25%**; narrow-range query only |
| Targeted OPC raw publication | Four-cell ABBA p50 geomean **-84.98%**; few-large/incompressible **-71.70%**; matched cycles **-69.21%** | Peak heap **+37.18%**, one-shot RSS **+22.26%** from retained source/provenance and regenerated-payload copying |
| Positional CFB/ZIP and explicit execution | Large-task p50 scaling at 12 CPUs: OPC **4.52x**, CFB **5.93x**; no hidden global Rayon | Many-small tasks regress at high worker counts; default/legacy paths remain serial |
| Source-backed OPC and DOCX/XLSX/PPTX facades | EOCD structural-open source bytes **-73.6% to -98.5%**; ordinary payload overlap zero | No latency claim: later EntryId/cache-diagnostic changes confound comparison and some cells exceed 5% variance |
| Deterministic range simulation | XLSX listing has zero timed requests; selected reads have zero unselected-sheet overlap; full physical size distributions recorded | Synthetic latency model, not a cold filesystem or ambient network |
| DOCX/PPTX semantic selectors and edits | DOCX one paragraph **-4.72%** p50; PPTX 1% edit/save **-9.37%** p50 and mean; PPTX one-edit guardrail +0.28% p50 (neutral) | Generated text corpora; complete transaction capture dominates one edit; no ODF/iWork implication |
| Coalesced DOCX paragraph edits | Large 100-edit/save p50 **-94.99% (19.97x)** and mean **-95.02%**; medium two-edit/save p50 **-12.98%**; scalar one-edit guardrail neutral | Direct-body, strictly ordered paragraph text replacement; generated corpus; scalar API remains separate |
| ODF semantic baselines and ODS snapshot reuse | Medium/large ODS no-op edit-save p50 **-7.45% / -11.78%**; one-cell edit-save **-3.57% / -2.06%** | Generated ODT/ODS/ODP corpora; ODP is coverage-only and ODT has the focused follow-up below; changed ODS publication still rewrites the package |
| RTF semantic baseline and text paths | Medium/large full-text p50 **-38.39% / -27.08%**; one-edit/save **-33.40% / -25.79%** | Generated native RTF text corpus; open guard +0.96% / +3.41%; formatting/media/security matrices remain missing |
| ODT shared transaction bytes | Medium/large no-op edit-save p50 **-27.05% / -18.51%**; exactly two allocations and one archive copy removed per snapshot | Existing-document snapshot handoff only; changed edit/save and open guardrails remain within 3%; changed publication still rewrites the package |

Raw evidence: [`XLSX before A`](results/abba-xlsx-range-before-a.json),
[`after A`](results/abba-xlsx-range-after-a.json),
[`before B`](results/abba-xlsx-range-before-b.json),
[`after B`](results/abba-xlsx-range-after-b.json); [`EOCD before A`](results/abba-eocd-before-a.json),
[`after A`](results/abba-eocd-after-a.json), [`before B`](results/abba-eocd-before-b.json),
[`after B`](results/abba-eocd-after-b.json); and
[`source-versus-eager`](results/stage3-source-vs-eager-many-small.json). The
committed positional XLSX record is
[`xlsx-source-positional.json`](results/xlsx-source-positional.json): p50 open
is 33.881 us/56.493 us/139.897 us (tiny/medium/dense), listing after open has
zero timed source reads, and first/range reads physically overlap only the
selected worksheet member (zero unselected worksheet read calls). These are
physical-overlap counts, not materialization counts.

The semantic raw evidence is [`before A`](results/abba-semantic-before-a.json),
[`after A`](results/abba-semantic-after-a.json),
[`after B`](results/abba-semantic-after-b.json), and
[`before B`](results/abba-semantic-before-b.json). The dedicated 60-sample
one-edit guardrail is linked from
[`change 0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md).

The ODF/ODS raw evidence is [`before A`](results/abba-odf-before-a.json),
[`after A`](results/abba-odf-after-a.json),
[`after B`](results/abba-odf-after-b.json), and
[`before B`](results/abba-odf-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md).

The coalesced-DOCX raw evidence is
[`before A`](results/abba-docx-batch-before-a.json),
[`after A`](results/abba-docx-batch-after-a.json),
[`after B`](results/abba-docx-batch-after-b.json), and
[`before B`](results/abba-docx-batch-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0012`](changes/0012-docx-coalesced-paragraph-edits.md).
That record also links the dedicated four-leg large-corpus scalar one-edit
guardrail (p50 -1.28%, mean +0.79% with overlapping intervals), which is
treated as neutral.

The RTF raw evidence is
[`text before A`](results/abba-rtf-text-before-a.json),
[`text after A`](results/abba-rtf-text-after-a.json),
[`text after B`](results/abba-rtf-text-after-b.json), and
[`text before B`](results/abba-rtf-text-before-b.json). The independent open
guard, complete seven-case matrix, allocation/RSS evidence, and rejected first
candidate are in
[`change 0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md).

The ODT shared-snapshot raw evidence is
[`before A`](results/abba-odt-shared-before-a.json),
[`after A`](results/abba-odt-shared-after-a.json),
[`after B`](results/abba-odt-shared-after-b.json), and
[`before B`](results/abba-odt-shared-before-b.json). Allocation attribution,
pooled statistics, open/changed-publication guardrails, and RSS evidence are in
[`change 0014`](changes/0014-odt-shared-snapshot-bytes.md).

Source-backed cache bytes are bounded by `SourceCacheLimits` but are not yet
charged to hierarchical `Budget`. Raw ZIP preservation is now integrated for
owned same-topology OPC mutations; broader source-backed editing is pending.
See [`0005`](changes/0005-xlsx-row-start-index.md),
[`0006`](changes/0006-positional-containers-and-explicit-execution.md), and
[`0007`](changes/0007-source-backed-opc-and-facades.md),
[`0008`](changes/0008-targeted-opc-preservation.md), and
[`0009`](changes/0009-range-source-and-scaling.md), and
[`0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md), and
[`0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md), and
[`0012`](changes/0012-docx-coalesced-paragraph-edits.md), and
[`0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md), and
[`0014`](changes/0014-odt-shared-snapshot-bytes.md).

Consolidated changed-crate tests passed, along with focused changed-crate
warning-denied Clippy and formatter checks. An umbrella all-feature `litchi`
attempt exhausted local disk; it is not reported as a passing umbrella gate.

## Accepted results

All latency figures below are warm-memory release-build p50 results from
matched before/after binaries. Each linked change record contains raw-sample
counts, ABBA ordering, mean or interval context, hashes, and memory profiles.

| Workload group | Before | After | Result | Memory result |
|---|---:|---:|---:|---|
| Targeted OPC mutation, four synthetic cells | individual rows in record | individual rows in record | **-84.98% p50 geometric mean**; range -58.24% to -96.41% | Few-large/incompressible peak heap +37.18%; one-shot RSS +22.26% |
| Exact owned OPC no-op, 16.78 MB incompressible archive | 211.531 ms | 3.443 ms | -98.37% | Peak heap +22.6%; profiler RSS +25.5% because the compressed source is retained alongside eagerly inflated Parts |
| Exact owned OPC no-op, six named many-Part/large-Part cells | individual rows in record | individual rows in record | -99.93% p50 geometric mean | Many-small allocation calls -93.7%; large memory tradeoff above |
| CFB final-root-stream lookup, four 256/2,048-sibling cells | 1.067-7.596 us | 0.451-0.486 us | -84.70% p50 geometric mean | Wide-root peak heap +1.5%; profiler RSS +7.6% for retained exact comparison keys |
| CFB open, four 256/2,048-stream cells | 141.1-963.1 us | 136.8-974.9 us | -1.42% p50 geometric mean | Allocation calls -6.1% to -8.8%; temporary allocations -20.6% to -27.7% |
| OPC rewritten publication, eight named cells | individual rows in record | individual rows in record | -1.65% mean geometric mean; best intended cell -5.49% | Allocation calls -37.0%; peak heap -2.3% |
| Payload-heavy PPT fresh writer | 6.312 ms | 5.035 ms | -20.23% | Peak heap -12.4%; profiler RSS -12.9% |
| Payload-heavy XLS fresh writer | 4.126 ms | 4.065 ms | -1.48%, treated as latency-neutral | Peak heap -9.5%; profiler RSS -12.6% |
| DOCX one paragraph, 10,000-paragraph corpus | 2.945 ms | 2.805 ms | -4.72% p50 / -4.99% mean | 10 collection-growth allocations removed per selector invocation; process peak unchanged |
| DOCX 1% edit/save, 10,000 paragraphs / 100 edits | 487.542 ms | 24.418 ms | **-94.99% p50 (19.97x) / -95.02% mean**; scalar one-edit neutral | Allocation calls -94.11%; peak heap flat; uninstrumented RSS +0.37% (flat) |
| PPTX 1% edit/save, 10,000 text boxes | 399.320 ms | 361.915 ms | -9.37% p50 / -9.37% mean | Allocation calls -11.67%; peak heap flat; profiler RSS +1.28% |
| ODS no-op edit/save, 32,768 cells | 76.894 ms | 67.838 ms | -11.78% p50 / -12.08% mean | Peak heap flat; profiler RSS -0.13% |
| ODS one-cell edit/save, 32,768 cells | 384.150 ms | 376.237 ms | -2.06% p50 / -2.19% mean | Changed package rewrite/readback still dominates |
| RTF full text, 10,000 paragraphs | 33.095 us | 24.134 us | -27.08% p50 / -25.37% mean | One fragment-vector allocation removed per first materialization |
| RTF one paragraph edit/save, 10,000 paragraphs | 12.408 ms | 9.208 ms | -25.79% p50 / -25.53% mean | Allocation calls -707 over 100 samples; peak heap flat; RSS +0.32% (flat) |
| ODT no-op edit/save, 10,000 paragraphs | 3.950 us | 3.219 us | -18.51% p50 / -29.58% mean | Exactly two allocations and one 28.42 KiB archive copy removed per snapshot; peak heap/RSS flat |

The underlying records are:

- [`0001-opc-publication-plan.md`](changes/0001-opc-publication-plan.md)
- [`0002-cfb-lookup-and-sector-buffers.md`](changes/0002-cfb-lookup-and-sector-buffers.md)
- [`0003-legacy-owned-stream-handoff.md`](changes/0003-legacy-owned-stream-handoff.md)
- [`0004-opc-exact-owned-source.md`](changes/0004-opc-exact-owned-source.md)
- [`0008-targeted-opc-preservation.md`](changes/0008-targeted-opc-preservation.md)
- [`0009-range-source-and-scaling.md`](changes/0009-range-source-and-scaling.md)
- [`0010-docx-pptx-semantic-queries-and-edits.md`](changes/0010-docx-pptx-semantic-queries-and-edits.md)
- [`0011-odf-semantic-baseline-and-ods-snapshot.md`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md)
- [`0012-docx-coalesced-paragraph-edits.md`](changes/0012-docx-coalesced-paragraph-edits.md)
- [`0013-rtf-semantic-baseline-and-text-paths.md`](changes/0013-rtf-semantic-baseline-and-text-paths.md)
- [`0014-odt-shared-snapshot-bytes.md`](changes/0014-odt-shared-snapshot-bytes.md)

The DOC ownership-transfer variant was rejected and removed after a 58.42%
p50 regression. The earlier full-rewrite mutated-OPC guardrail was neutral on
incompressible data; targeted raw publication supersedes it only for the
strictly proved same-topology owned-source case. Fallback still uses that
validated full rewrite. Rejected, fallback and memory results are retained
rather than hidden in an aggregate.

## Work removed

- Exact unchanged owned OPC publication no longer regenerates manifests,
  reconstructs ZIP records, or recompresses logical Parts. It copies the
  complete validated source to the caller's sequential sink in writes bounded
  to 64 KiB and verifies complete output in the benchmark.
- Targeted same-topology OPC publication no longer recompresses unchanged
  Parts. It audits the ordinary publication plan, regenerates only changed
  payload/relationship/content-type closures, and raw-copies unchanged local
  spans and central records, including unknown non-part members.
- Rewritten OPC publication constructs and audits generated XML and stable
  Part order once before emission rather than once for validation and again
  for writing.
- CFB lookup follows the validated sibling-tree ordering with SID-aligned
  cached comparison keys rather than scanning the complete sibling tree.
- CFB FAT/DIFAT/MiniFAT parsing reuses a bounded sector buffer, MiniFAT decodes
  into its final table, and directory sectors read into their final buffer.
- Fresh XLS and PPT writers transfer already-owned generated stream buffers to
  CFB without a second payload copy. DOC deliberately retains its measured
  faster exact-sized copy.
- DOCX one-paragraph selection no longer constructs the complete paragraph
  collection, and source-backed paragraph counts no longer construct any
  paragraph views. Complete XML validation and limits still run.
- Canonical multi-paragraph DOCX text edits now plan disjoint replacements and
  emit/reparse one candidate XML document instead of rebuilding and reparsing
  the complete main document once per paragraph. Durable patches remain
  ordinary source-checked paragraph operations with complete final readback.
- Repeated PPTX shape-text edits no longer parse the selected slide scene a
  second time solely to map the already selected shape to its raw XML span.
- DOCX plaintext package output exposes the underlying forward-only OPC sink
  instead of imposing an unused `Seek` bound.
- ODS unified snapshot construction reuses its one validated package for full
  facade readback instead of cloning package bytes and parsing the package a
  second time.
- RTF first full-text materialization retains only a byte count during parse,
  then allocates the final string once and copies blocks in one pass instead of
  allocating and joining a temporary fragment vector.
- RTF canonical text emission writes ordinary ASCII spans in chunks instead of
  one formatted write per character. Text-only commits skip paragraph-property
  vectors/scans, and a successful paragraph selector stops at its target.
- ODT transaction snapshots created from an already validated `Document` clone
  its private immutable package handle instead of allocating and copying the
  complete archive. Direct snapshot byte ingress keeps independent validation.

No unsafe code, ambient I/O, dependency edge, public archive type, or global
synchronization primitive was introduced. Exact-source authorization is
revoked conservatively on every mutable OPC entry point, including failed and
semantic no-op calls. Borrowed ingress, topology-changing edits, and unsupported
ZIP layouts use the fully validated rewrite path before any sink output.

## Evidence and verification

The standalone harness provides 88 selectable cases and a 198-record default
matrix across deterministic ZIP/OPC, positional CFB/OPC, source-backed XLSX,
public DOC/XLS/PPT writer corpora, and DOCX/PPTX/RTF/ODT/ODS/ODP semantic
corpora.
It records
p50/p95/p99, raw samples, mean, sample deviation, Student's-t 95% mean interval,
corpus/output hashes, environment, bounded sequential-write behavior,
deterministic logical/physical range distributions, and exact execution
tasks/bytes. CI runs a non-gating deterministic smoke check and a
scheduled/manual release matrix.

The current local evidence includes consolidated changed-crate tests with
byte/hash checks, focused changed-crate warning-denied Clippy, formatter and
diff checks, YAML parsing, and JSON parsing. The umbrella all-feature `litchi`
attempt exhausted local disk, so it is not represented as a passing umbrella
gate. The historical stage-1 all-feature gate and its pre-existing Cargo
warning about DOCX/PPTX example output name `owner_native_smoke` remain scoped
to that earlier capture.

The repository-wide warning-denied rustdoc command remains blocked by existing
broken/private intra-doc links in unchanged OPC, DOC, XLS, and PPT files. The
dependency-direction checker unit suite passes, while the live policy check
reports existing unclassified edges (including `litchi-opc -> xml-minifier`
and several dev-only `-> soapberry-zip` edges); this tranche changes no Cargo
manifest or dependency edge. These pre-existing gate failures are not counted
as passing verification.

During the stage-1 capture, hardware counters were unavailable because that
host had `perf_event_paranoid=4`. The later targeted-OPC capture ran after the
environment reported `1`: matched process counters show cycles -69.21% and
instructions -69.85% for that one save cell. No counter claim is retroactively
made for stage 1 or generalized to other workloads; lock-wait evidence remains
missing.

## Remaining highest-impact work

The largest remaining limitation is the incomplete migration from eager OPC to
source-backed CRUD: selective open, source versions, finite cache and
single-flight now exist, but cache bytes are not yet charged to the hierarchical
budget and broad edit/patch coverage is incomplete. Raw ZIP preservation is
integrated for eager owned same-topology mutation, but not for source-backed
editable packages, and its measured retained-source/payload-copy memory cost
remains to be reduced.

Other high-priority gaps are cold-filesystem and real range-source matrices,
threshold tuning/contention work beyond the committed explicit scaling curves,
and broad format-semantic CRUD coverage beyond the new generated
DOCX/PPTX slice (bulk action distinctions, dependency-copy, merge/split,
patching, repair, security, malformed and real-producer corpora, plus broader
ODF and RTF coverage). Legacy DOC/XLS/PPT semantic open/query/edit/save
baselines are the next source-audited non-iWork prerequisite before further CFB
ownership experiments; broader ODT source-backed reads and changed publication
remain open.
iWork work is deliberately deferred while the `iwa-*` crates are changing
independently.
The scenario-by-scenario gap map and next case queue are in
[`CRUD_COVERAGE.md`](CRUD_COVERAGE.md).
The ranked source-level queue and path maps are maintained in
[`HOTSPOTS.md`](HOTSPOTS.md), and architectural gates are in
[`ADR_COMPLIANCE.md`](ADR_COMPLIANCE.md).
