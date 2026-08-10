# Performance program phase report

Date: 2026-08-10
Branch: `feat/office-format-completeness`
Production base for this tranche: `d6bd13c9043737885ad27fa9d6207f4c30c1addd`

This report summarizes the measured implementation tranches to date. It is not a
claim that the end-to-end performance program or CRUD scenario matrix is
complete. The reproducible environment, original substrate baseline, corpus
definitions, commands, and profiler limitations are in
[`BASELINE.md`](BASELINE.md); raw reports are under [`results/`](results/).

## Current stable tranche

The original stage-1 results below remain historical evidence. The current
harness contains **44 selectable cases**: 36 default cases and 198 default
records, plus six opt-in simulated-range cases and two opt-in scaling cases. It
is still not broad program or CRUD coverage.

| Change | Current evidence | Scope / limitation |
|---|---|---|
| XLSX row-start index | ABBA p50 geomean **-80.499%**, mean geomean **-79.962%**; full scan **+0.03%** mean; first cell **-1.31%** mean | Heap allocations **+17**, RSS **+0.25%**; narrow-range query only |
| Targeted OPC raw publication | Four-cell ABBA p50 geomean **-84.98%**; few-large/incompressible **-71.70%**; matched cycles **-69.21%** | Peak heap **+37.18%**, one-shot RSS **+22.26%** from retained source/provenance and regenerated-payload copying |
| Positional CFB/ZIP and explicit execution | Large-task p50 scaling at 12 CPUs: OPC **4.52x**, CFB **5.93x**; no hidden global Rayon | Many-small tasks regress at high worker counts; default/legacy paths remain serial |
| Source-backed OPC and DOCX/XLSX/PPTX facades | EOCD structural-open source bytes **-73.6% to -98.5%**; ordinary payload overlap zero | No latency claim: later EntryId/cache-diagnostic changes confound comparison and some cells exceed 5% variance |
| Deterministic range simulation | XLSX listing has zero timed requests; selected reads have zero unselected-sheet overlap; full physical size distributions recorded | Synthetic latency model, not a cold filesystem or ambient network |

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

Source-backed cache bytes are bounded by `SourceCacheLimits` but are not yet
charged to hierarchical `Budget`. Raw ZIP preservation is now integrated for
owned same-topology OPC mutations; broader source-backed editing is pending.
See [`0005`](changes/0005-xlsx-row-start-index.md),
[`0006`](changes/0006-positional-containers-and-explicit-execution.md), and
[`0007`](changes/0007-source-backed-opc-and-facades.md),
[`0008`](changes/0008-targeted-opc-preservation.md), and
[`0009`](changes/0009-range-source-and-scaling.md).

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

The underlying records are:

- [`0001-opc-publication-plan.md`](changes/0001-opc-publication-plan.md)
- [`0002-cfb-lookup-and-sector-buffers.md`](changes/0002-cfb-lookup-and-sector-buffers.md)
- [`0003-legacy-owned-stream-handoff.md`](changes/0003-legacy-owned-stream-handoff.md)
- [`0004-opc-exact-owned-source.md`](changes/0004-opc-exact-owned-source.md)
- [`0008-targeted-opc-preservation.md`](changes/0008-targeted-opc-preservation.md)
- [`0009-range-source-and-scaling.md`](changes/0009-range-source-and-scaling.md)

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

No unsafe code, ambient I/O, dependency edge, public archive type, or global
synchronization primitive was introduced. Exact-source authorization is
revoked conservatively on every mutable OPC entry point, including failed and
semantic no-op calls. Borrowed ingress, topology-changing edits, and unsupported
ZIP layouts use the fully validated rewrite path before any sink output.

## Evidence and verification

The standalone harness provides 44 selectable cases and a 198-record default
matrix across deterministic ZIP/OPC, positional CFB/OPC, source-backed XLSX,
and public DOC/XLS/PPT writer corpora. It records p50/p95/p99, raw samples,
mean, sample deviation, Student's-t 95% mean interval, corpus/output hashes,
environment, bounded sequential-write behavior, deterministic logical/physical
range distributions, and exact execution tasks/bytes. CI runs a non-gating
deterministic smoke check and a scheduled/manual release matrix.

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
and broad format-semantic CRUD coverage
(selective queries, 1% edits, dependency-copy, merge/split, patching, repair,
security, malformed and real-producer corpora).
The scenario-by-scenario gap map and next case queue are in
[`CRUD_COVERAGE.md`](CRUD_COVERAGE.md).
The ranked source-level queue and path maps are maintained in
[`HOTSPOTS.md`](HOTSPOTS.md), and architectural gates are in
[`ADR_COMPLIANCE.md`](ADR_COMPLIANCE.md).
