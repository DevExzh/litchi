# OPC, CFB, OLE2 Office, OOXML, RTF, and ODF performance baseline

`litchi-perf-baseline` is an isolated, reproducible measurement tool for the
ZIP/OPC and CFB/OLE2 substrates, fresh DOC/XLS/PPT writer packaging, and
public-API XLSX snapshot/edit/save flows, matched opt-in XLSX scalar-cell
eager/source-backed and managed source-backed publication controls, and opt-in DOC/XLS/PPT,
DOCX/PPTX/RTF/ODT/ODS/ODP semantic flows, including the opt-in RTF logical-tail
append transaction, bounded RTF/XLS/DOCX/PPTX/ODF validation reports, and a
source-backed DOCX section inventory. It creates every corpus in memory; it also exercises
source-backed XLSX catalog, worksheet reads, and guarded calculation-metadata,
defined-name, page-break/page-margin/page-setup/print-options/sheet-protection/data-validation/auto-filter
publication over positional I/O. The scalar-cell controls use deterministic
medium and dense/sparse four-sheet corpora with untouched media Parts. Their
timed interval covers open, selector planning, commit, and sequential
publication; reopen, semantic equality, exact hashes, raw media identity,
lifecycle gates, and source/materialization counter sampling remain outside the
reported timing. Source-backed JSON also records those stages separately,
including managed Budget/cache diagnostics and release-to-zero checks.
The reported duration is the sum of open/stage/commit and sequential
publication segments; source cache diagnostics are sampled between those
segments and are excluded. They are evidence for later release ABBA work, not
a speedup claim. It does not
depend on untracked office files, network state, or randomness. ODP builder
timestamps are replaced with fixed metadata before measurement. The JSON
report contains the generator parameters and SHA-256 hashes for the generated
container and target entry, so a result always identifies its exact input or
packaged output.

The tool is intentionally outside the root workspace and has no effect on
production dependency graphs.

The native OLE2, DOCX/PPTX, RTF, and ODF semantic matrices are deliberately
opt-in. They measure only current public APIs and therefore do not change the
default 36 cases / 198 records.

## Run

Run the complete default matrix (36 default cases; 198 result records: 144
substrate records, nine writer records, and 45 XLSX records). The six simulated
range cases, two execution-scaling cases, one low-level source-overlay save
case, one source-backed DOCX semantic publication case, one source-backed
media-rich PPTX semantic publication case, four matched same-slide/multi-slide
PPTX batch cases, two matched cross-slide ODP text-box publication cases, one
matched ODT embedded-resource publication pair, one matched ODT mixed
model-content publication pair, one XLSX commit/read attribution case,
two matched XLSX calculation-metadata publication cases, two matched XLSX
defined-name publication cases, two matched XLSX
page-break publication cases, two matched XLSX page-margin publication cases,
two matched XLSX page-setup publication cases,
two matched XLSX print-options publication cases,
two matched XLSX sheet-protection publication cases,
two matched XLSX data-validation publication cases,
two matched XLSX auto-filter/sort-state publication cases,
two matched XLSX conditional-formatting publication cases,
two XLSX merge/unmerge commit-plus-save cases, six matched eager/source-backed
XLSX scalar-cell publication cases (one cell, `ceil(1%)`, and the exact 256-cell
bound), one bounded unmanaged source-backed two-worksheet case, and four
managed source-backed scalar-cell cases (one cell, `ceil(1%)`, exact 256, and
two worksheets),
two bounded XLSX/RTF streaming-creation cases,
six matched CFB selective-read cases and six matched simulated high-latency
CFB selective-read cases,
four matched native XLS existing-comment publication cases,
six matched native XLS fixed-width numeric publication cases,
four matched native XLS worksheet-visibility publication cases,
four opaque-heavy common OLE2 stage/edit-save cases, 24 native OLE2 semantic cases, 16
DOCX/PPTX semantic cases, 15 RTF semantic cases (13 transport/read/edit
cases plus two logical-tail publication cases), 38 ODF semantic cases, and one
ODF `mimetype` repair-plan case are opt-in. Eight additional native PPT
`Pictures` selectors are available for matched eager/source-backed open,
cold all-images query, repeated all-images query, and fresh open-plus-all-images
phases on a deterministic picture-heavy corpus. Source-backed elapsed samples
for those native-PPT `Pictures` selectors use an uninstrumented
`litchi_core::OwnedSource`; independent untimed `InstrumentedSource` replays
provide their source-read counters. The current `Case` matrix exposes 271
selectable case names in total. Eight additional
PPTX ordinary-root filesystem selectors (`pptx_file_{eager,source}_{open,
list_slides,slide_count,selected_slide}`) are opt-in and do not alter the
default 36 cases / 198 records. Two repeated native-PPT selected-shape query
selectors (`ppt_semantic_repeated_shape_text` and
`ppt_source_backed_repeated_shape_text`) are also opt-in and use matched
eight-query eager/source-backed controls;
four matched ODP media-rich read selectors (`odp_media_eager_open`,
`odp_media_source_backed_open`, `odp_media_eager_one_slide`, and
`odp_media_source_backed_one_slide`) are also opt-in;
four unified-root ODP filesystem selectors (`odp_file_{eager,source}_{open,
selected_slide}`) are also opt-in and compare eager byte ownership with the
filesystem source-root handoff over the same media-rich corpus. Together they
bring the selectable matrix to 237 names while leaving the default 36 cases /
198 records unchanged;
six matched ODS unified-root/source selectors (`ods_file_{eager,source}_{open,
selected_cell,selected_media}`) are also opt-in over the deterministic
two-sheet media-rich ODS corpus. Root open timing covers only root construction;
typed selected-cell/media timing covers only the query after owner preparation.
Independent `SourceBackedSpreadsheet` replays report logical positional reads,
exact compressed-member range coverage, and zero-read selected-cell behavior;
compressed ranges and uncompressed payloads remain separate fields. Together
these additions bring the selectable matrix to 243 names while leaving the
default 36 cases / 198 records unchanged;
two additional matched CFB MiniFAT selectors (`cfb_selective_mini_4095_legacy_read`
and `cfb_selective_mini_4095_shared_read`) exercise a distinct 4095-byte
boundary target alongside the existing 36-byte control. Together these CFB
additions bring the selectable matrix to 245 names while leaving the default
36 cases / 198 records unchanged;
eight ordinary-root DOCX filesystem selectors (`docx_file_{eager,source}_{open,
paragraph_count,list_paragraphs,full_text}`) are also opt-in over the fixed
200-paragraph/eight-incompressible-2 MiB-media corpus. Root-open timing covers
`fs::read` plus eager `Document::from_bytes` versus source `Document::open`;
query roots are prepared outside timing and the timer covers only the named
root query. Independent untimed `litchi_docx::source_backed::Package` replays
record catalog/open reads and, for query selectors, complete coverage of the
compressed main-document range during document preparation plus zero
media/unselected/core overlap during the query. Request sizes, range coverage,
materializations, and an explicit classification are also recorded. Full
eager/source semantic parity, exact source hash, logical OPC
part/relationship/content-type/blob-hash gates, media hashes, and source
immutability remain verification outside timing. This is correctness and
logical compressed-range evidence only: it makes no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, ABBA, security, or Markdown claim.
Together these additions bring the selectable matrix to 253 names while
leaving the default 36 cases / 198 records unchanged;
two matched ODS repeated-cell sweep selectors
(`ods_file_eager_cell_sweep` and `ods_file_source_cell_sweep`) are also opt-in
over the same two-sheet media-rich corpus. Each owner is opened before timing,
four identical 2,048-cell sweeps cross the adaptive locator threshold, and
source replay counters are reset after preparation and must remain zero during
the sweep. Semantic digest/count and complete source/member/media preservation
gates are untimed. These additions bring the selectable matrix to 255 names
while leaving the default 36 cases / 198 records unchanged;
two matched ODS ordered cell-batch sweep selectors
(`ods_file_eager_cell_batch_sweep` and `ods_file_source_cell_batch_sweep`) are
also opt-in over that exact corpus. Owners and 2,048 borrowed selectors are
prepared before timing; the timer covers four bounded `cell_batch` calls (8,192
result slots) with results black-boxed. Independent source replay observes
exactly eight version checks and zero post-preparation payload reads per
four-call sweep; semantic digest/count and complete source/member/media
preservation gates remain untimed. These additions bring the selectable matrix
to 257 names while leaving the default 36 cases / 198 records unchanged;
the four opt-in fixed-width native XLS numeric selectors bring the current
selectable matrix to 261 names while leaving that default unchanged; two
additional plan-only native XLS numeric selectors bring it to 263 names;
two additional matched source-backed ODP repeated-text selectors
(`odp_source_backed_repeated_text_uncached` and
`odp_source_backed_repeated_text_cached`) are also opt-in over that same
12-slide, eight-2 MiB-media corpus. Each `SourceBackedPresentation` owner and
four output slots are prepared outside timing. The control reconstructs the
pre-cache public sequence (`slides()` plus `Slide::all_text()`, filtered and
joined with exact `\n\n`, followed by the trailing source check); the candidate
calls `SourceBackedPresentation::text()` four times, exercising the
threshold-two cache.
The timer contains only those four projections. An independent untimed
`InstrumentedSource` replay records preparation and post-preparation counters:
the four-call replay must have zero reads, bytes, compressed-range overlap,
and `Pictures` payload reads; freshness vectors are `[3, 3, 3, 3]` for the
control and `[3, 5, 2, 2]` for the candidate (12 observations total in each
case). Archive topology, eight-picture/16 MiB media identity, text parity and
hashes remain untimed. These additions bring the selectable matrix to 265
names while leaving the default 36 cases / 198 records unchanged. Change 0140
accepts the frozen CPU-2 selector-pair result for this exact four-call shape:
p50 improves 45.80%/46.32% and p95 improves 45.25%/45.83% in paired directions;
whole-process allocation-call counts improve 14.31%, while peak heap and RSS
remain neutral. No single-call, open, physical-I/O, decompression, cold-cache,
or generic ODF claim is made;
six additional matched simulated CFB selective-read cases
(`cfb_selective_simulated_{mini,mini_4095,fat}_{legacy,shared}_read`) reuse the
same deterministic final-position targets while applying the configured
bounded range delay. Legacy uses a harness-only delayed `Read + Seek` adapter;
the shared control uses the bounded `ReadAt` simulator. Each result records
open/read/total logical and physical request counts, bytes, sorted request
sizes, size buckets, and a deterministic simulated service floor alongside
the existing target hash and phase counters. The Mini controls require the
shared read stage to do exactly target work and the legacy stage to retain its
full-root amplification; the FAT pair remains a matched control. These cases
are selectable evidence only and do not by themselves imply a latency or
native semantic performance claim. Together these additions bring the
selectable matrix to 271 names while leaving the default 36 cases / 198 records
unchanged. Change 0144 documents the evidence boundary and the separate clean
release result.

The validation/section and scalar-cell selectors are opt-in and do not alter the default
36 cases / 198 records:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --json target/perf/container-baseline.json
```

For a short local smoke run:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 2 --shape tiny --payload compressible \
  --writer-shape tiny --xlsx-shape tiny --json -
```

Select comma-separated subsets when investigating a change:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --samples 30 --case cfb_open,cfb_list_streams,cfb_read_one --shape wide-root \
  --payload incompressible --json target/perf/wide-root.json
```

Measure the common OLE2 transactional publication path with four unchanged 4
MiB regular streams and one tiny edited MiniFAT stream:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case ole_common_open,ole_common_put_stream_publish,ole_common_finish_render,ole_common_one_edit_save \
  --shape few-large --payload incompressible --json target/perf/ole-common.json
```

The stage cases separately time public editor open, `put_stream` candidate
publication, and changed `finish` rendering. The end-to-end case times all
three. Stage preparation, exact deterministic output comparison and a complete
public CFB reopen of all five streams remain outside each timed interval.

Measure the current evidence-first OPC source overlay path on its fixed
few-large incompressible corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --case opc_source_overlay_one_part_save \
  --json target/perf/opc-source-overlay.json
```

This case times positional source open and the source-backed one-Part publisher,
which validates the selected payload, preserves the existing URI, content type,
relationships and topology, raw-copies every other member, and writes to a
sequential sink. Payload preparation and complete output verification stay
outside timing.

Measure the matched XLSX scalar-cell controls on their deterministic
media-rich four-sheet corpora:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case xlsx_eager_cell_values_one_edit_save,\
xlsx_source_backed_cell_values_one_edit_save,\
xlsx_eager_cell_values_one_percent_edit_save,\
xlsx_source_backed_cell_values_one_percent_edit_save,\
xlsx_eager_cell_values_batch_edit_save,\
xlsx_source_backed_cell_values_batch_edit_save,\
xlsx_source_backed_cell_values_multi_sheet_edit_save,\
xlsx_source_backed_managed_cell_values_one_edit_save,\
xlsx_source_backed_managed_cell_values_one_percent_edit_save,\
xlsx_source_backed_managed_cell_values_batch_edit_save,\
xlsx_source_backed_managed_cell_values_multi_sheet_edit_save \
  --xlsx-cell-crud-shape medium,dense-sparse \
  --json target/perf/xlsx-cell-values-crud.json
```

The one-cell, deterministic `ceil(1%)`, and exact-256 selectors are matched
eager/source-backed pairs. The additional source-backed two-worksheet control
is bounded to one existing cell per selected worksheet. Managed controls use
the same explicit finite cache policy as their unmanaged source-backed
controls, with an execution-context Budget for retained/in-flight OPC
`PartData` payloads. The source-backed path uses one selector-first
multi-sheet transaction and a sequential overlay publisher; the harness also
checks exact no-op, clear, and remove lifecycle behavior as untimed gates.
Every output is reopened and checked for semantic cell state, package topology,
relationships, deterministic hashes, and the bounded sink; eager outputs also
check untouched media payloads. Source-backed outputs additionally compare raw
local and central ZIP records for every unselected member, including media,
relationships, content types, workbook, and unselected worksheets. Source read
and successful materialization counters are sampled after the timed segments
(with diagnostics excluded from the reported sum). Balanced release ABBA in
[change 0096](../../docs/performance/changes/0096-xlsx-source-provenance-publication.md)
accepts removal of the redundant publication-time semantic worksheet reparse:
source-backed p50 geomean improves 21.66%/22.65% in the two directions, while
physical read/materialization counters remain unchanged. No allocation, RSS,
cold-I/O or decompression claim is attached to that result.

Change 0109's managed tranche has no controlled release ABBA comparison and
therefore makes no speedup or throughput claim. Its Budget covers only retained
and in-flight OPC `PartData` payload reservations; parsed stores, metadata,
staging, rewritten candidates, and output buffers are outside that accounting.
The tranche also does not claim allocations, RSS/peak memory, hardware/CPU
pinning, cold I/O, decompression, or real-producer breadth.

Measure the opt-in OPC source-cache Budget boundary and controlled contention
matrix on one fixed many-small incompressible corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --workers 1,2,4,8,available \
  --case opc_source_cache_budget_boundary,\
opc_source_cache_control_contention,opc_source_cache_managed_contention \
  --json target/perf/opc-source-cache.json
```

The boundary selector emits an exact-budget success and a one-byte-under
managed refusal. The refusal must report two reservation failures and perform
zero payload I/O. Each contention selector emits same-Part and fixed-work
disjoint-Part cells at `1/2x`, `1x`, and `2x` cache capacity for every capped,
deduplicated worker width. With five resolved widths, the three selectors emit
62 records: two boundary records plus 30 finite-control and 30 Budget-managed
records.

Every cell creates one persistent worker team and reuses it across warm-ups and
samples. Each iteration opens and prefills a fresh package, admits the initial
cohort through an explicit source gate, then times only post-admission service
completion. A distinct compressed payload incurs one fixed 10 ms source delay.
Returned `PartData` handles stay live until the whole wave completes, making
pin pressure observable. The report records exact cache counters, pre-release
flights/waiters, gate arrivals and concurrency, retained bytes/entries, Budget
use after handle/package drop, request throughput, and an Amdahl classification
only for the fixed-request disjoint cells. Same-Part widths change request
count, so they are explicitly throughput-only and do not receive a speedup or
serial-fraction estimate.

These deterministic delays and classifications are correctness and contention
evidence, not production-latency results. Do not make a performance claim from
them without clean release builds, CPU affinity, balanced control/managed ABBA
ordering, retained raw samples, allocation and peak-memory evidence, and stable
counter identities.

Measure the controlled filesystem tranche (five opt-in cases):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case opc_file_eager_open,opc_file_source_open,\
opc_file_eager_one_part_atomic_save,opc_file_source_one_part_atomic_save,\
cfb_file_same_length_overlay_atomic_save \
  --json target/perf/filesystem-crud.json
```

Use `--filesystem-cache warm` for a warm-only smoke, or
`--filesystem-cache warm,cold-requested` (the default) for both keyed states.
Use `--filesystem-root PATH` to place the source, destination, and sibling
temporary files under a caller-selected filesystem; the report records only
that a root was selected, not the path itself.

Each sample uses a fresh child process and reports child operation time plus
parent-observed wall time. A separate child primes the warm path immediately
before each warm sample; the cold sample requests Linux `posix_fadvise`
`DONTNEED` immediately before timing and records whether that advisory request
was accepted. `cold-requested` is a cache-state request, not a claim that the
kernel or storage device delivered a guaranteed cold cache. The additive
`filesystem_evidence` JSON section records per-sample `ReadAt` request and
return byte counts, request sizes, maximum in-flight reads, procfs I/O/fault/
RSS counters, deterministic output hashes and byte lengths, and semantic
reopen checks. The OPC corpus and expected one-Part output hashes, plus the
CFB corpus and exact 36-byte-overlay output hash, are pinned and checked before
samples. The CFB atomic-save case also divides its timed operation into three
non-overlapping intervals: positional source/CFB open, overlay planning and
complete candidate validation, and atomic publication. Each interval reports
elapsed time plus logical `ReadAt` call/request/return deltas. The harness
requires the three source-counter deltas to sum exactly to the whole operation
and phase elapsed time not to exceed it. These are logical source-work and
timing-attribution counters, not physical device I/O, copied bytes,
allocations, or decompression evidence.

When both OPC save cases are selected, their per-state sample hashes
must also match. Save cases seed a pre-existing destination before both warm and cold measurements
and publish through a same-filesystem sibling temporary file plus atomic
rename; the CFB case uses the checked same-length stream-overlay publisher and
records its changed-span and published-byte report fields. OPC materialization
counts are recorded when exposed by the public API. After every prime and
measured child, the parent re-reads the source and checks its pinned SHA-256
before proceeding to the next sample.

Measure the matched ordinary-root PPTX filesystem controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case pptx_file_eager_open,pptx_file_source_open,\
pptx_file_eager_list_slides,pptx_file_source_list_slides,\
pptx_file_eager_slide_count,pptx_file_source_slide_count,\
pptx_file_eager_selected_slide,pptx_file_source_selected_slide \
  --json target/perf/pptx-root-filesystem.json
```

These eight cases use the same 200-slide/eight-text-box/eight-2 MiB-media
PPTX corpus as the source-backed publication controls. Source candidates call
`litchi::Presentation::open(path)`; eager open times `fs::read` plus
`Presentation::from_bytes`, while query cases prepare the root before timing.
`slide_count` counts only, `list_slides` materializes the complete owned slide
vector, and `selected_slide` uses `Presentation::slide(100)`. Every sample is
isolated in a fresh warm/cold-requested child and verifies source hash, full
eager/source metadata, size, name, text, and semantic parity outside timing.
Source samples add one separate untimed `SourceBackedPresentation` replay with
exact compressed-range classification: open/count must overlap no slide or
media payload, selected must overlap only slide 100, and list must overlap all
slide payloads but no media. Eager samples explicitly have no `ReadAt` replay;
their generic filesystem counter scope is marked not applicable. These are
correctness/logical-read observations only, not latency, RSS, allocation,
decompression, physical-I/O, or cold-cache claims. See
[`change 0120`](../../docs/performance/changes/0120-pptx-root-source-path-evidence.md).

Measure the media-rich DOCX source-backed semantic-edit control:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --case docx_source_backed_one_edit_save \
  --json target/perf/docx-source-edit.json
```

The fixed corpus contains 200 paragraphs and eight deterministic
incompressible 2 MiB PNG payloads. The current control times the required
public migration through `SourceBackedPackage::into_opc_package`, one semantic
paragraph transaction, and sequential publication. Complete DOCX readback,
media and topology checks, patch/inverse/stale checks, and output hashing stay
outside timing.

Measure the media-rich PPTX source-backed semantic publication:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --case pptx_source_backed_one_edit_save \
  --json target/perf/pptx-source-edit.json
```

The fixed corpus has 200 slides, eight deterministic text boxes per slide, and
eight deterministic incompressible 2 MiB inert PNG media parts. This opt-in
case times positional open, one guarded source-backed slide shape-text edit,
and one-Part overlay publication into a bounded sequential sink. Full PPTX semantic
readback plus exact Part topology, relationships, content types, unselected
payload/media, source/output hashes, and source/sink checks remain untimed.

Measure the matched atomic eight-shape eager/source-backed controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case pptx_eager_batch_edit_save,pptx_source_backed_batch_edit_save \
  --json target/perf/pptx-source-batch.json
```

Both paths resolve and replace the same eight text boxes on slide 100 through
the public atomic batch API and emit byte-identical output. The eager control
materializes all 229 ordinary Parts; the source-backed path materializes only
the presentation root and selected slide. Corpus creation plus complete
semantic, topology, relationship, media, patch/inverse, output-hash, and
source/sink verification remain outside the timed interval.

Measure the matched eight-slide batch controls on the same corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case pptx_eager_multi_slide_batch_edit_save,pptx_source_backed_multi_slide_batch_edit_save \
  --json target/perf/pptx-multi-slide-batch.json
```

Both paths replace all eight text boxes on slides 0, 28, 57, 85, 114, 142,
171, and 199 and emit byte-identical output. The eager path materializes all
229 ordinary Parts; the source-backed batch materializes the presentation root
and eight selected slides, then regenerates only those eight slide members.
Complete semantic, topology, relationship, raw unselected-member, media,
patch/inverse, source-version, output-hash, and sink checks remain untimed.

Measure the matched XLSX calculation-metadata publication controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_calculation_metadata_edit_save,xlsx_source_backed_calculation_metadata_edit_save \
  --json target/perf/xlsx-calculation-edit.json
```

Both cases change only `calcPr` in `xl/workbook.xml` on the same workbook with
one worksheet, one DrawingML drawing, and eight referenced incompressible 2 MiB
PNG payloads. The eager control materializes all twelve ordinary Parts and
uses the owning XLSX writer. The source-backed case materializes only the
workbook Part and consumes the guarded one-Part OPC overlay. Complete package,
relationship, calculation-metadata, drawing/media, output-hash, and bounded
sequential-sink verification remains outside timing.

Measure the matched XLSX defined-name catalog publication controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_defined_names_edit_save,xlsx_source_backed_defined_names_edit_save \
  --json target/perf/xlsx-defined-names-edit.json
```

Both paths replace the direct workbook defined-name catalog with one global
range name and one hidden sheet-local name. The eager control materializes all
twelve ordinary Parts; the source-backed path materializes only the workbook
Part and raw-copies the other eleven. Complete name/scope and calculation-policy
readback, package topology, relationships, media, patch/inverse, output hash,
and bounded sequential-sink verification remain outside timing.

Measure the matched XLSX page-break publication controls on the same fixed
media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_page_break_edit_save,xlsx_source_backed_page_break_edit_save \
  --json target/perf/xlsx-page-break-edit.json
```

Both paths add one manual horizontal break to `Sheet1` and change only
`xl/worksheets/sheet1.xml`. The eager control materializes all twelve ordinary
Parts; the source-backed path materializes only the workbook catalog and the
selected worksheet, then raw-copies every other ZIP member. Full page-break
readback, calculation-metadata stability, package topology, relationships,
media payloads, output hash, and sequential-sink bounds are verified outside
timing.

Measure the matched XLSX page-margin publication controls on the same fixed
media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_page_margin_edit_save,xlsx_source_backed_page_margin_edit_save \
  --json target/perf/xlsx-page-margin-edit.json
```

Both paths add the same six typed margins to `Sheet1` and change only
`xl/worksheets/sheet1.xml`. The eager control materializes all twelve ordinary
Parts; the source-backed path materializes only the workbook catalog and the
selected worksheet, then raw-copies every other ZIP member. Full page-margin
readback, calculation-metadata stability, package topology, relationships,
media payloads, output hash, and sequential-sink bounds are verified outside
timing.

Measure the matched relationship-free XLSX page-setup publication controls on
the same fixed media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_page_setup_edit_save,xlsx_source_backed_page_setup_edit_save \
  --json target/perf/xlsx-page-setup-edit.json
```

Both paths add the same A4, landscape, 85%-scale typed settings to `Sheet1`
and change only `xl/worksheets/sheet1.xml`. The eager control materializes all
twelve ordinary Parts; the source-backed path materializes only the workbook
catalog and selected worksheet. It retains the complete selected-worksheet
relationship signature and refuses printer-settings references because those
require a wider multi-Part publication capability. Semantic readback,
calculation metadata, topology, relationships, media, output hash, and sink
bounds are verified outside timing.

Measure the matched XLSX print-options publication controls on the same fixed
media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_print_options_edit_save,xlsx_source_backed_print_options_edit_save \
  --json target/perf/xlsx-print-options-edit.json
```

Both paths enable the same typed horizontal-centering, headings, and gridline
flags on `Sheet1` and change only `xl/worksheets/sheet1.xml`. The eager control
materializes all twelve ordinary Parts; the source-backed path materializes
only the workbook catalog and selected worksheet, then raw-copies every other
ZIP member. Complete print-options readback, calculation metadata, topology,
relationships, media, output hash, and sequential-sink bounds are verified
outside timing.

Measure the matched XLSX sheet-protection publication controls on the same
fixed media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_sheet_protection_edit_save,xlsx_source_backed_sheet_protection_edit_save \
  --json target/perf/xlsx-sheet-protection-edit.json
```

Both paths create the same complete typed protection state on `Sheet1`,
including sheet locks, one core protected range and one Office 2010 protected
range with a strong verifier descriptor. The eager control materializes all
twelve ordinary Parts; the source-backed path materializes only the workbook
catalog and selected worksheet. Complete protection readback, calculation
metadata, topology, all worksheet relationships, media, output hash, and
sequential-sink bounds are verified outside timing.

Measure the matched XLSX data-validation publication controls on their fixed
media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_data_validation_edit_save,xlsx_source_backed_data_validation_edit_save \
  --json target/perf/xlsx-data-validation-edit.json
```

Both paths replace the same complete typed core and Office 2010 validation
collections on `Sheet1`. The eager control materializes all twelve ordinary
Parts; the source-backed path materializes only the workbook catalog and
selected worksheet. Complete validation readback, calculation metadata,
topology, worksheet relationships, media, output hash, and sink bounds are
verified outside timing.

Measure the matched XLSX auto-filter and sort-state publication controls on
their fixed media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_auto_filter_edit_save,xlsx_source_backed_auto_filter_edit_save \
  --json target/perf/xlsx-auto-filter-edit.json
```

Both paths replace the same typed value filter and sort state on `Sheet1`.
The eager control materializes all twelve ordinary Parts; the source-backed
path materializes the workbook catalog, selected worksheet, and styles Part.
Complete typed readback, calculation metadata, topology, worksheet
relationships, media, output hash, and sink bounds are verified outside
timing.

Measure the matched XLSX conditional-formatting publication controls on their
fixed media-rich archive:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_conditional_formatting_edit_save,xlsx_source_backed_conditional_formatting_edit_save \
  --json target/perf/xlsx-conditional-formatting-edit.json
```

Both paths replace the same complete ordered collection of three typed core
conditional-formatting owners on `Sheet1` through the same worksheet rewriter.
The eager control materializes all twelve ordinary Parts; the source-backed
path materializes the workbook catalog, selected worksheet, and styles Part.
Complete typed reopen, exact patch/inverse replay, calculation metadata,
topology, worksheet relationships, all unselected Part/media payloads, raw ZIP
members, output hash, source reads, and sequential-sink bounds are verified
outside timing. This selectable evidence makes no latency claim until a
balanced ABBA measurement is retained.

Measure the two opt-in XLSX merge/split lifecycle controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_merge_commit_save,xlsx_eager_unmerge_commit_save \
  --json target/perf/xlsx-merge-unmerge.json
```

Both cases use the same deterministic sparse `Sheet1` A1:B2 fixture with a
retained A1 anchor and unrelated C1 cell. Inputs and transaction edits are
prepared before timing; each sample times only one semantic commit and bounded
sequential `Workbook::write_to`. Reopen, merge membership, covered/uncovered
views, anchor/unrelated-cell retention, exact durable patch apply/inverse, and
stale-source refusal are verified outside timing. These cases add selectable
correctness evidence only and make no latency claim without controlled ABBA
evidence.

Measure the matched native XLS existing-comment publication controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xls_comments_eager_edit_save,xls_comments_source_backed_edit_save,\
xls_comments_eager_batch_edit_save,xls_comments_source_backed_batch_edit_save \
  --json target/perf/xls-comments-edit.json
```

All four cases use one deterministic BIFF8 workbook with 256 existing comments,
an untouched worksheet/comment, eight exact 2 MiB incompressible opaque streams,
and one opaque metadata stream. The one-edit cases replace the middle owner; the
batch cases replace exactly the supported 256-owner limit. Author/text lengths
and compressed encoding width stay unchanged. Each sample separately records
semantic staging/plan and publication time while total `elapsed_ns` is their
sum. The source-backed reports additionally retain changed comments, logical
streams, exact NOTE/TXO splice counts and replacement-byte totals, physical
spans, equal Workbook lengths, and exact source/target CFB fingerprints. The
one-owner control requires two splices (one NOTE and one TXO family), while
the 256-owner control requires 512; replacement bytes must be nonzero and
remain below the full Workbook stream. Generic source counters cover only the explicit owned-source
ingress because the public XLS comments API owns its source bytes internally;
sink counters cover complete bounded publication. Complete reopen, all 256
semantic values, worksheet/comment inventory, every untouched stream, explicit
eager fallback for a length-changing edit, and protected/refusal behavior stay
outside timing. Output hashes are deterministic per case, but eager rendering
and source-backed overlay are not required to have identical physical CFB
bytes. These cases add selectable evidence only and make no performance claim
without a frozen release-build, CPU-pinned ABBA run.

Measure the matched native XLS fixed-width numeric publication controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xls_numeric_eager_number_edit_save,xls_numeric_source_backed_number_edit_save,\
xls_numeric_eager_rk_mulrk_edit_save,xls_numeric_source_backed_rk_mulrk_edit_save,\
xls_numeric_plan_only_number_edit_save,xls_numeric_plan_only_rk_mulrk_edit_save \
  --json target/perf/xls-numeric-edit.json
```

The Number pair reuses the deterministic `Untouched!E21` `Number` cell in the
comments corpus (`42` -> `43`). The RK/MulRK pair uses a separate deterministic
native XLS corpus containing one standalone RK and one two-cell MulRK record;
one transaction updates all three cells with exact integer-RK replacements.
Each sample times transaction creation, `set_number`/`set_numeric`, the eager
`commit`, ordinary source-backed `commit_source_backed`, or forward-only
`commit_source_backed_plan`, and publication separately. The plan-only commit
includes composed semantic validation but retains no complete target artifact.
`total_ns` and the top-level `elapsed_ns` distribution are arithmetic sums of
those four separately timed phases, not a continuous wall-clock timer.
All publication paths write the complete target CFB to equivalently configured
preallocated bounded `CountingSink`s (64 KiB maximum write); ordinary
source-backed output avoids Workbook-stream reserialization but still retains a
complete target snapshot. Plan-only output retains no target snapshot or target
byte vector at commit and is explicitly forward-only without artifact
patch/inverse. Source ingress, expected outputs, sink capacity,
no-op/fingerprint, exact source/target fingerprint preflights,
security/unsupported refusal, full `Snapshot`/`Workbook` reopen, untouched CFB
topology/member bytes, and the untimed real-producer `54016.xls` forward gate
all run outside timing. Generic `source.read_calls`/`source.read_bytes` carry
the owned source-ingress counters; `source.xls_numeric` reports the separate
commit, publication, and total vectors, explicit target-retention/materialization
flags, complete target materialized bytes for retained-target paths, source-backed
splice/replacement/span/fingerprint evidence, sink bytes, write counts, digests,
and the explicit owned-input scope. These selectors are
correctness/coverage evidence only: no positional-I/O, allocation/RSS,
bounded-artifact-memory, speedup, or broad-producer claim is made. Composed
semantic validation may read and allocate a candidate `Workbook` model, so
zero target-artifact bytes does not imply bounded total memory.

Measure the matched native XLS worksheet-visibility publication controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xls_visibility_eager_edit_save,xls_visibility_source_backed_edit_save,\
xls_visibility_eager_batch_edit_save,xls_visibility_source_backed_batch_edit_save \
  --json target/perf/xls-visibility-edit.json
```

All four cases use one deterministic BIFF8 workbook with 66 worksheet owners,
eight exact 256 KiB incompressible opaque streams, and opaque metadata. The
one-edit cases change worksheet position 1's `BoundSheet8.hsState` byte; the
batch cases change exactly the supported 64-owner limit by hiding positions
1 through 64, leaving positions 0 and 65 visible. Each sample separately
records semantic staging/commit and sequential publication time through a
bounded 64 KiB sink; total `elapsed_ns`
is their sum. Source-backed reports additionally retain changed-owner/stream,
exact source-relative splice and replacement-byte diagnostics, physical-span,
equal Workbook-length, and exact source/target CFB fingerprint evidence. The
one-owner control requires 1 splice and 1 replacement byte; the 64-owner
control requires 64 and 64. Generic source counters cover only explicit owned-source ingress;
sink counters cover complete bounded publication. Complete worksheet/catalog
reopen, opaque-stream preservation, exact offset checks, eager patch
replay/inverse, no-op identity, cap-plus-one refusal, and protected-source
refusal stay outside timing. The source-backed API retains its complete
candidate snapshot, so these cases make no allocation, peak-memory, I/O,
materialization, or speedup claim without a frozen release-build, CPU-pinned
ABBA run.

For just the end-to-end legacy writer packaging runs:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to \
  --writer-shape tiny,large,payload-heavy --json target/perf/legacy-writers.json
```

Run the complete native DOC/XLS/PPT semantic matrix over the same deterministic
tiny and large writer artifacts (46 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --writer-shape tiny,large \
  --case doc_semantic_open,doc_semantic_list_paragraphs,doc_semantic_one_paragraph,doc_semantic_full_text,doc_semantic_noop_edit_save,doc_semantic_one_edit_save,xls_semantic_open,xls_semantic_list_worksheets,xls_semantic_one_cell,xls_semantic_full_cell_scan,xls_semantic_noop_edit_save,xls_semantic_one_edit_save,ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_one_shape_text,ppt_source_backed_one_shape_text,ppt_semantic_repeated_shape_text,ppt_source_backed_repeated_shape_text,ppt_semantic_fresh_open_one_shape_text,ppt_source_backed_fresh_open_one_shape_text,ppt_semantic_full_text,ppt_slide_order_snapshot_open,ppt_text_edit_one_edit_save,ppt_semantic_noop_edit_save,ppt_semantic_one_edit_save \
  --json target/perf/ole2-semantic-baseline.json
```

For the dense-wide XLSX range traversal that isolates selected-row scanning:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --case xlsx_narrow_column_range_scan \
  --xlsx-shape dense-wide --json target/perf/xlsx-narrow-range.json
```

For a short positional XLSX laziness and selected-range smoke run:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 \
  --case xlsx_source_open,xlsx_source_list_sheets,xlsx_source_first_cell,xlsx_source_narrow_column_range_scan \
  --xlsx-shape tiny --json -
```

Run the complete tiny semantic DOCX/PPTX smoke matrix (16 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case docx_semantic_open,docx_semantic_list_paragraphs,docx_semantic_one_paragraph,docx_semantic_full_text,docx_semantic_create_small,docx_semantic_noop_edit_save,docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,pptx_semantic_open,pptx_semantic_list_slides,pptx_semantic_one_slide,pptx_semantic_full_text,pptx_semantic_create_small,pptx_semantic_noop_edit_save,pptx_semantic_one_edit_save,pptx_semantic_one_percent_edit_save \
  --json target/perf/semantic-office-smoke.json
```

Measure the bounded validation reports and source-backed DOCX section inventory
over deterministic in-memory corpora:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --semantic-shape tiny,medium,large \
  --writer-shape tiny,large --rtf-variant plain \
  --case rtf_validation_report,xls_validation_report,docx_validation_report,\
docx_section_inventory,pptx_validation_report,odf_validation_report \
  --json target/perf/validation-sections.json
```

Corpus generation, source setup, canonical report/inventory summarization,
hashing, and gates are outside each timed interval. RTF validation and generic
ODF validation time only their borrowed-byte validator calls. XLS, DOCX, and
PPTX validation time the public source-backed validator, including its
positional `ReadAt` requests. The DOCX section case times source-backed package
open plus the section-inventory snapshot. Each result records deterministic
report hashes, check IDs/status classes, issue codes and counts, source
before/after hashes, and source-read counters where available. The section
record additionally retains every descriptor's ownership, paragraph range,
page/margin values, start marker, and header/footer relationship IDs. Warmups
are excluded from recorded source counters. These are selectable correctness
and baseline measurements only; make no speedup claim without a frozen release,
CPU-pinned balanced ABBA capture with retained raw samples.
The RTF validation selector accepts the plain, raw CP-1252, and LZFu variants;
the producer-watermark fixture is intentionally excluded because its opaque
drawing surface yields an unknown safety status rather than a complete terminal
validation report.

Run the complete tiny semantic ODF smoke matrix (24 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case odt_semantic_open,odt_semantic_list_paragraphs,odt_semantic_one_paragraph,odt_semantic_full_text,odt_semantic_create_small,odt_semantic_noop_edit_save,odt_semantic_one_edit_save,odt_semantic_one_percent_edit_save,ods_semantic_open,ods_semantic_list_sheets,ods_semantic_one_cell,ods_semantic_cell_sweep,ods_semantic_full_cell_text,ods_semantic_create_small,ods_semantic_noop_edit_save,ods_semantic_one_edit_save,ods_semantic_one_percent_edit_save,odp_semantic_open,odp_semantic_list_slides,odp_semantic_one_slide,odp_semantic_full_text,odp_semantic_create_small,odp_semantic_noop_edit_save,odp_semantic_one_edit_save \
  --json target/perf/semantic-odf-smoke.json
```

Run the fixed medium ODS publication case with eight 2 MiB incompressible
opaque resources:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --case ods_media_one_edit_save \
  --json target/perf/ods-media-publication.json
```

Run the fixed medium ODT paragraph publication case with eight 2 MiB
incompressible opaque resources:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --case odt_media_paragraph_edit_save \
  --json target/perf/odt-media-paragraph-publication.json
```

Run the matched ODT line-break publication case over the same fixed corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --case odt_media_line_break_edit_save \
  --json target/perf/odt-media-line-break-publication.json
```

Run the matched ODT inline-run publication case over the same fixed corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --case odt_media_append_run_edit_save \
  --json target/perf/odt-media-append-run-publication.json
```

Run the matched ODT hyperlink publication case over the same fixed corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --case odt_media_append_hyperlink_edit_save \
  --json target/perf/odt-media-append-hyperlink-publication.json
```

Run the matched ODT structural paragraph publication cases over the same fixed
corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 \
  --case odt_media_insert_paragraph_edit_save,odt_media_remove_paragraph_edit_save \
  --json target/perf/odt-media-structural-paragraph-publication.json
```

Run the matched ODT existing embedded-image replacement cases over a fixed
64-owner extension of the media-rich corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 \
  --case odt_embedded_resource_scalar_replace_save,odt_embedded_resource_batch_replace_save \
  --json target/perf/odt-embedded-resource-publication.json
```

Run the matched ODP cross-slide existing-text-box publication cases over a
fixed media-rich corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 \
  --case odp_media_textbox_scalar_replace_save,odp_media_textbox_batch_replace_save \
  --json target/perf/odp-cross-slide-textbox-publication.json
```

Run the plain tiny semantic RTF smoke matrix (13
records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case rtf_semantic_open,rtf_semantic_paragraph_count,rtf_semantic_list_paragraphs,rtf_semantic_collect_paragraphs,rtf_semantic_one_paragraph,rtf_semantic_full_text,rtf_semantic_text_to_sink,rtf_semantic_stream_save,rtf_semantic_noop_edit_save,rtf_semantic_one_edit_save,rtf_semantic_one_percent_edit_save,rtf_semantic_remove_paragraph_save,rtf_semantic_move_paragraph_save \
  --json target/perf/semantic-rtf-smoke.json
```

Select all transport and producer variants for the complete tiny RTF coverage
matrix (39 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --rtf-variant plain,byte1252,lzfu,watermark \
  --case rtf_semantic_open,rtf_semantic_paragraph_count,rtf_semantic_list_paragraphs,rtf_semantic_collect_paragraphs,rtf_semantic_one_paragraph,rtf_semantic_full_text,rtf_semantic_text_to_sink,rtf_semantic_stream_save,rtf_semantic_noop_edit_save,rtf_semantic_one_edit_save,rtf_semantic_one_percent_edit_save,rtf_semantic_remove_paragraph_save,rtf_semantic_move_paragraph_save \
  --json target/perf/semantic-rtf-variants-smoke.json
```

Measure logical append to an existing plain RTF separately from streaming
creation. The command below selects small, medium, and large existing
documents; the changed case appends 4, 64, or 256 bounded plain paragraphs,
while the no-op case commits an empty tail transaction. Both cases publish to
a fixed 16 KiB non-seek sink window. Every sample is checked against an
untimed output digest, and the preflight gates cover exact no-op identity,
in-memory patch replay/inverse, durable JSON decode/apply/inverse, complete
reopen, and foreign-source conflict refusal:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --semantic-shape tiny,medium,large \
  --case rtf_logical_tail_append,rtf_logical_tail_noop_save \
  --json target/perf/rtf-logical-tail-append.json
```

The timed interval includes append staging, candidate commit validation, and
sequential publication. Source parsing, expected-output construction, durable
wire work, reopening, and all correctness gates remain outside timing. The
`sink.rtf_tail_append` object records source/input/inserted/output bytes,
paragraph/run counts, the fixed sink window, and boolean gate results. This is
selectable baseline evidence only; it makes no speedup or latency claim until
a frozen release-build CPU-pinned ABBA comparison exists.

Exercise deterministic high-latency, range-bounded positional I/O without a
network. Every upstream logical read is split into physical requests no larger
than the selected maximum:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --shape tiny --payload compressible \
  --xlsx-shape tiny \
  --case opc_range_source_open,opc_range_source_open_main_read,xlsx_range_source_open,xlsx_range_source_list_sheets,xlsx_range_source_first_cell,xlsx_range_source_narrow_column_range_scan \
  --range-fixed-latency-us 100 --range-request-overhead-us 25 \
  --range-bandwidth-bytes-per-sec 52428800 \
  --range-max-physical-bytes 4096 --json target/perf/range-source.json
```

Collect explicit worker scaling points. Counts are capped to visible available
parallelism, deduplicated, and emitted in ascending order; this avoids
oversubscription in pinned CI containers:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --shape many-small --payload incompressible \
  --case opc_open_session_scaling,cfb_bulk_read_scaling \
  --workers 1,2,4,8,available --json target/perf/execution-scaling.json
```

## Corpus matrix

Each shape is generated twice, once with a repeated deterministic payload and
once with a deterministic xorshift payload. The latter is intended to be hard
to compress, not cryptographically random.

| Shape | Entries | Bytes per entry | Total logical payload | Primary CFB stressor |
|---|---:|---:|---:|---|
| `tiny` | 3 | 512 B | 1.5 KiB | Fast smoke input; CFB MiniFAT |
| `many-small` | 256 | 1 KiB | 256 KiB | CFB MiniFAT traversal |
| `few-large` | 4 | 4 MiB | 16 MiB | CFB regular FAT streams |
| `wide-root` | 2,048 | 64 B | 128 KiB | Wide CFB root sibling tree |

## Legacy writer matrix

The fresh writer cases are independent of the ZIP/CFB corpus shapes and
payload kinds. `--writer-shape` selects their deterministic bounded content
shape; omitting it runs all three.

| Writer shape | DOC paragraphs | XLS cells | PPT slides / text boxes | Bound |
|---|---:|---:|---:|---:|
| `tiny` | 3 | 16 | 1 / 2 | Fast public-API smoke input |
| `large` | 512 | 8,192 | 12 / 144 | Moderate end-to-end packaging input |
| `payload-heavy` | 128 × 20,000-byte text | 128 × 32,700-byte string cells across 128 sheets | 16 / 128 × 40,000-byte text | 4–8 MiB primary CFB stream |

DOC paragraphs and PPT text boxes contain fixed identifier text. XLS uses a
fixed sequence of finite numeric cells for `tiny`/`large`. `payload-heavy`
uses repeated deterministic text; its XLS string cells are distinct but each
worksheet contains exactly one cell, retaining stable shared-string ordering
through the public API. There is no random, clock, or filesystem-derived
content.

## XLSX corpus matrix

`--xlsx-shape` selects complete, in-memory workbooks with a fixed integer grid
on every sheet. The value at each coordinate is derived from its sheet, row,
and column. The deterministic ~1% update set is evenly distributed through the
complete logical cell order and rounds up to one cell where necessary.

| XLSX shape | Sheets | Rows × columns per sheet | Logical cells | ~1% updates | Primary use |
|---|---:|---:|---:|---:|---|
| `tiny` | 3 | 8 × 8 | 192 | 2 | Fast end-to-end smoke input |
| `medium` | 4 | 32 × 32 | 4,096 | 41 | Moderate snapshot/edit/save input |
| `dense-wide` | 2 | 256 × 256 | 131,072 | 1,311 | Stable narrow-column scan over dense selected rows |

`dense-wide` deliberately contains all 256 stored cells in each selected row.
Its narrow range is `B1:B256`, which returns 256 cells while making the
worksheet store examine the 65,536 stored cells in those rows. This preserves a
high-contrast public end-to-end case without relying on internal APIs.

## Opt-in bounded streaming creation

`xlsx_streaming_create` and `rtf_streaming_create` exercise the public
forward-only creation APIs. They use `--semantic-shape` only as a common shape
selector and are not part of the default matrix.

| Shape | XLSX rows / cells | RTF paragraphs / runs | Retained writer window |
|---|---:|---:|---:|
| `tiny` | 64 / 256 | 64 / 64 | XLSX 4 KiB row buffer; RTF 37 B encoder state |
| `medium` | 8,192 / 32,768 | 8,192 / 8,192 | Same |
| `large` | 131,072 / 524,288 | 131,072 / 131,072 | Same |

The XLSX case writes one numeric, inline-text, boolean, and explicit blank cell
per row. The RTF case writes one bounded UTF-8 run per paragraph. Text is
deterministic and includes non-ASCII and format-significant characters.

Timed publication uses a non-seek SHA-256 discard sink. It retains no output
bytes: only hash state, accepted bytes, write calls, and the largest write. A
complete artifact is generated separately before timing, reopened through the
ordinary `Workbook` or `Document` facade, and exhaustively checked against the
shape. That artifact is dropped before samples run. Every timed digest must
equal the reopened artifact digest.

The `sink` record additionally reports exact rows/cells or paragraphs/runs,
input bytes, authored worksheet/RTF bytes, `retained_output_bytes: 0`, and the
production writer's explicit `retained_authoring_window_bytes`. Increasing total output
therefore cannot be mistaken for an increasing retained authoring window.

The RTF writer batches only escape-free printable ASCII, with a hard 32-byte
sink-request ceiling and no additional retained writer buffer. Balanced
release ABBA in
[change 0097](../../docs/performance/changes/0097-rtf-bounded-ascii-streaming.md)
records p50 geomean improvements of 76.41%/76.47%; the large case falls from
7,208,970 to 1,441,802 sink calls while exact bytes and hashes remain stable.
This is a fresh-creation result, not an existing-document edit or memory-profile
claim.

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --semantic-shape tiny,large \
  --case xlsx_streaming_create,rtf_streaming_create \
  --json target/perf/streaming-create.json
```

## Opt-in DOCX/PPTX semantic corpus matrix

`--semantic-shape` creates complete public-API packages in memory. Text names,
edit selection, object order, and output verification are fixed. The one-percent
set is evenly spaced in the complete paragraph or text-box order and rounds up
to one object. `*_create_small` runs only for `tiny`, so its corpus metadata is
always truthful.

| Shape | DOCX paragraphs | PPTX slides × text boxes | DOCX ~1% edits | PPTX ~1% edits |
|---|---:|---:|---:|---:|
| `tiny` | 24 | 3 × 4 | 1 | 1 |
| `medium` | 200 | 12 × 8 | 2 | 1 |
| `large` | 10,000 | 100 × 100 | 100 | 100 |

The DOCX cases use `Package::from_reader`, `document().paragraph(index)`,
paragraph enumeration/text extraction, document transactions, and `to_stream`.
The one-percent case uses the canonical
`Edit::replace_body_paragraph_texts` transaction whenever the selection has
more than one paragraph, so validation, XML emission, candidate parsing, and
complete selected-paragraph readback are coalesced without changing the
ordinary durable patch operations. The one-edit case remains on
`replace_paragraph_text` as a scalar guardrail.
The seekable stream is instrumented with the existing `sink` schema field.
DOCX also accepts a forward-only `Write` sink; production correctness tests
cover that contract while the frozen before/after benchmark keeps the same
seekable counter implementation in both binaries.
PPTX uses `Package::from_bytes`/`from_vec`, presentation slide/text views,
opened-presentation transactions, and `to_bytes`. PPTX currently has no public
writer-sink API, so PPTX save records intentionally leave `sink` as `null`
rather than claiming unobservable write behavior.

## Opt-in RTF semantic corpus matrix

The RTF cases exercise only the ordinary native `litchi_rtf::Document` facade:
owned-byte open, lazy paragraph enumeration, one middle paragraph, first
complete-text materialization, bounded semantic-text output to a forward-only
sink, exact source streaming, exact empty-edit publication, and
capability-bounded one-paragraph and `ceil(1%)` paragraph edit/save. The two
`rtf_logical_tail_*` cases are a separate existing-document append tranche;
they do not reuse the streaming-creation path and are restricted to the
matched plain lifecycle corpus.
The two lifecycle cases use a matched default-formatted plain corpus because
the read/edit corpus's explicit font formatting is outside their changed
publication closure.
`--rtf-variant` defaults to `plain`.

| Variant | Source | Shapes | Supported cases |
|---|---|---|---|
| `plain` | Deterministic direct ASCII RTF | tiny, medium, large | All 15, including logical-tail append/no-op |
| `byte1252` | Deterministic raw CP-1252 bytes containing literal `0xe9` | tiny, medium, large | The original 13 open/read/text-to-sink/stream/no-op cases; changed splice and logical tail are excluded because candidate validation refuses this byte layout |
| `lzfu` | Deterministic LZFu compression of the plain bytes | tiny, medium, large | The original 13 open/read/text-to-sink/stream/no-op cases; changed transport and logical-tail rewrites are explicitly unsupported |
| `watermark` | Content-addressed real-producer `test-data/rtf/watermark.rtf` | tiny selector only | The original open/read/stream/no-op cases; semantic body-text output and logical-tail publication are excluded because the meaningful content is header drawing metadata rather than editable body text |

Every save uses the native forward-only `Write` contract and every output is
reopened and fully verified. The watermark verifier additionally requires the
three public header shapes and `gtextUNICODE=ASAP` metadata. LZFu generation
checks deterministic compression and exact decompression; byte-1252 generation
uses raw high-bit bytes rather than RTF hex escapes.

| Shape | Paragraphs | Source bytes | Text bytes |
|---|---:|---:|---:|
| `tiny` | 24 | 1,347 | 1,199 |
| `medium` | 200 | 10,851 | 9,999 |
| `large` | 10,000 | 540,051 | 499,999 |

The exact stream-save and no-op cases preserve every selected input byte for
byte and emit it as one sequential write. The one-edit case is limited to the
plain variant: it stages the middle paragraph through `replace_paragraph_text`,
commits with source and semantic readback checks, streams the changed snapshot,
and verifies every paragraph after reopen. Corpus creation, expected-output
construction, and input cloning remain outside the timed interval.

The logical-tail cases use the matched default-formatted lifecycle corpus,
which has 24/200/10,000 existing paragraphs for tiny/medium/large. They append
4/64/256 one-run plain paragraphs under explicit paragraph, run, input,
inserted-byte, output, and durable-patch limits. The resulting source/input /
inserted/output byte counters are 1,304/168/273/1,577 for tiny,
10,808/2,816/4,421/15,229 for medium, and
540,008/11,008/17,413/557,421 for large. Publication uses a 16 KiB
windowed non-seek sink that retains zero output bytes; the large case therefore
reports multiple bounded writes instead of one whole-document write. This
window is a sink accounting bound, not a claim that the append transaction's
validated candidate snapshot is memory-bounded.

## Opt-in ODF semantic corpus matrix

The ODF cases use the same `--semantic-shape` selection, but each format is
generated through its public builder and reopened from owned in-memory bytes.
Creation is timed only for `tiny`; every case validates its public semantic
projection and every edit/save case reopens the published bytes after timing.
No filesystem `save(path)` operation is included, so these records measure
in-memory publication rather than OS/filesystem behavior.

| Shape | ODT paragraphs | ODS sheets × rows × columns | ODP slides |
|---|---:|---:|---:|
| `tiny` | 24 | 1 × 8 × 8 | 3 |
| `medium` | 200 | 2 × 32 × 32 | 12 |
| `large` | 10,000 | 2 × 128 × 128 | 100 |

Each ODT batch uses `Builder`, `Document::from_bytes`, paragraph enumeration,
full-text extraction, and the source-bound `Document::edit` transaction.
`odt_semantic_one_paragraph` calls the public `Document::paragraph(index)`
selector. The selector validates the complete XML through EOF but retains only
the requested paragraph; it is not a positional or early-return XML read.
`odt_semantic_one_percent_edit_save` stages deterministic evenly spaced
replacements through the ordinary scalar transaction API, then commits once,
materializes the package, reopens it, verifies every paragraph and full text,
and checks that the operation result count equals the selected 1% closure.

Each ODS batch uses `Builder`, `Spreadsheet::from_bytes`, `sheets()`, the
public logical `cell()` view, a row-major cell sweep without cell-text
cloning, a deterministic row-major cell-text aggregate, and the unified
`document::Snapshot` transaction. ODS snapshot construction is
inside the timed edit/save interval so these cases expose the package-open cost
paid by this public editing entry point; the source-byte clone is outside the
interval. The timed work also includes staging, commit, and published-byte
observation. `ods_semantic_cell_sweep` isolates repeated public lookup cost
without cloning cell text, while `ods_semantic_full_cell_text` retains the
end-to-end aggregate because the facade exposes cells rather than a single
full-text method. `ods_semantic_one_percent_edit_save` selects a deterministic
evenly spaced `ceil(1%)` of row-major cells, partitions them by worksheet,
stages one bounded atomic `set_cells` batch per selected sheet, commits the
document once, then reopens and verifies the complete grid. It is a selectable
correctness and timing case; no comparative latency improvement is claimed by
its addition.

`ods_media_one_edit_save` is a separate fixed-medium corpus: 2,048 cells plus
eight deterministic 2 MiB resources under `Pictures/`. It times public unified
snapshot open, one middle-cell edit, commit and output materialization. Outside
timing it reopens the complete grid and verifies every resource path, manifest
media type, exact payload and deterministic output. This case does not vary
with `--semantic-shape` and is not part of the 24-record tiny ODF smoke matrix.

`odt_media_paragraph_edit_save`, `odt_media_line_break_edit_save`,
`odt_media_append_run_edit_save`, `odt_media_append_hyperlink_edit_save`,
`odt_media_insert_paragraph_edit_save`, and
`odt_media_remove_paragraph_edit_save` share a separate fixed-medium corpus:
200 paragraphs plus eight deterministic 2 MiB resources under `Pictures/`.
They time public snapshot open, respectively replace the middle paragraph,
append one line break, append one unstyled inline run, append one inert
hyperlink, insert one paragraph, or remove one paragraph, commit, and
materialize the output. Outside timing each case reopens every paragraph,
verifies hyperlink text and URL where applicable, verifies every resource path,
manifest media type and exact payload, checks patch replay, exact inverse and
stale-source refusal, and requires deterministic output.
The JSON's `output_sha256` makes that published artifact identity explicit.
Harness regressions additionally prove raw local/central record identity for
all untouched core and media members, including styled and unstyled run and
hyperlink publication. These cases do not vary with `--semantic-shape` and are
opt-in.

`odt_embedded_resource_scalar_replace_save` and
`odt_embedded_resource_batch_replace_save` share a deterministic extension of
that fixed-medium corpus: the same 200 paragraphs and eight 2 MiB opaque
resources plus 64 existing, uniquely named package-backed image owners with
4 KiB payloads. Replacement values and the immutable snapshot are prepared
outside timing. Both intervals create one transaction, replace the same 64
existing image owners onto corresponding fixed same-length target paths without
renaming frames, commit once,
and copy the complete published snapshot to one pre-reserved bounded sink. The
scalar control calls `replace_embedded_image` 64 times; the bounded case calls
`edit_embedded_resources` once with 64 base-snapshot selectors. Replacement
does not add, remove, or reorder owners, so scalar selector shifting is not a
factor. The displaced fixed source payloads remain packaged, as required by the
production replacement contract.

Outside timing, both paths reopen all 200 paragraphs and all 64 image owners,
verify every frame name/path/media type and source/replacement payload SHA-256,
verify all eight retained 2 MiB resources and raw identity for every untouched
ZIP member, and require the same complete semantic projection. Volatile and
deterministic-JSON durable forward replay, inverse restoration, and stale-source
refusal are checked for both paths. Case-specific deterministic output hashes
and exact one-write sink counters are retained. The owned ODT transaction API
exposes no positional source-read or logical-part materialization diagnostics,
so the report omits them. These cases are selectable evidence only: no latency,
instruction, allocation, memory, or materialization improvement is claimed
without a frozen CPU-pinned balanced ABBA capture.

Measure the matched ODT mixed model-content publication controls on the
deterministic semantic ODT corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 2 --semantic-shape tiny \
  --case odt_mixed_model_content_scalar_edit_save,odt_mixed_model_content_batch_edit_save \
  --json target/perf/odt-mixed-model-content-smoke.json
```

For a measurement matrix, select `--semantic-shape medium,large` and use a
release-build sample count appropriate for the host:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --semantic-shape medium,large \
  --case odt_mixed_model_content_scalar_edit_save,odt_mixed_model_content_batch_edit_save \
  --json target/perf/odt-mixed-model-content.json
```

The scalar control applies one deterministic five-operation group per selected
region (`remove`, `insert`, `replace`, styled `append_run`, and inert
`append_hyperlink`). The operation vector orders all count-preserving plain
operations first and its complete inline append tail second. The scalar
control publishes each plain operation separately, then publishes the entire
inline tail once so later publications cannot canonicalize away an earlier
styled span. The candidate stages that identical operation vector in one
transaction so the current model-content coalescing path is exercised. Tiny,
medium, and large use 1, 16, and 64 groups (5, 80, and 320 operations); the
paragraph count stays constant. Thus the scalar control reports three times
the group count plus one inline-tail publication (4, 49, and 193 for tiny,
medium, and large), while the candidate reports one publication per iteration.

The timed interval starts after a fresh source snapshot has been opened and
ends after the scalar publication sequence or single candidate commit. Corpus
construction, operation preparation, full reopen and semantic projection,
result counts, exact source/target member inventory and untouched raw-member
identity, volatile and durable replay/inverse/stale checks, barrier and
late-error atomicity, security-envelope classification, and text/operation
limits are outside timing. The JSON `source.odt_mixed` object records exact
operation, publication, result, output-byte, output-hash, and logical-result-hash
counters for every measured sample. This is matched selector evidence only;
it makes no latency or speedup claim until a frozen release-build balanced ABBA
capture is retained.

`odp_media_textbox_edit_save` is a separate fixed-medium source-backed
publication corpus: 12 slides plus eight deterministic 2 MiB resources under
`Pictures/`. It times public snapshot open, one `add_text_box` operation,
commit, and output materialization. Outside timing it checks every original
slide, the inserted text box through `rich_content`, exact patch/inverse and
stale-source behavior, deterministic output, and every resource payload and
manifest media type. It does not vary with `--semantic-shape` and is opt-in.

`odp_media_textbox_scalar_replace_save` and
`odp_media_textbox_batch_replace_save` share a second fixed-medium corpus: the
same 12 deterministic slides and eight 2 MiB resources plus eight existing,
globally unique rich-text boxes distributed over slide positions 0, 1, 3, 4,
6, 7, 9, and 11. Replacement models and the immutable editing snapshot are
prepared outside timing. Both intervals create one transaction, replace the
same complete models without renaming any owner, commit once, and copy the
published snapshot to one pre-reserved bounded sink. The scalar control calls
`replace_text_box_model` eight times, so it exercises repeated candidate
staging; the bounded case calls `replace_text_box_models` once.

Outside timing, both paths reopen the complete presentation and rich-content
inventory, require identical slide/full-text projections, verify all fixed
names, pages, paragraph counts and updated text, check volatile and durable
patch replay/inverse plus stale-source refusal, and retain deterministic input
and output hashes. `mimetype`, `styles.xml`, `meta.xml`, and all eight media ZIP
members remain physically identical. The batch also raw-preserves the manifest;
the repeated scalar staging regenerates it, so the two semantically identical
outputs intentionally have distinct bytes and digests. ODP's owned-byte editor
exposes neither positional source-read nor logical-Part materialization
diagnostics. The report therefore records real sink counters and omits
`source`/materialization fields instead of inventing an OPC-style count. These
cases add selectable evidence only: no latency, allocation, memory, or
materialization improvement is claimed without a frozen CPU-pinned balanced
ABBA capture.

`odp_media_eager_open` and `odp_media_source_backed_open` are matched
media-rich read controls over the same 12-slide/eight-2 MiB `Pictures/` corpus;
`odp_media_eager_one_slide` and `odp_media_source_backed_one_slide` keep the
same prepared owner and query the deterministic middle slide. Eager timing
uses an owned byte buffer and source timing uses an uninstrumented
`litchi_core::OwnedSource`; cloning/owner preparation and complete semantic
parity are outside the query-only interval. Every source sample has an
independent untimed `InstrumentedSource` replay recording exact calls, bytes,
coalesced prior-range overlap, and overlap with all compressed `Pictures/*`
ranges. The JSON names this aggregate `pictures_read_compressed_range_bytes`;
it is not the prior-read overlap counter (`source_read_range_overlap_bytes`).
The open replay may include a bounded ZIP-tail request that physically touches
the last Pictures range; that is recorded as compressed-range overlap, not
called a media materialization. One additional untimed replay reads exactly the
selected media member and gates its compressed-range overlap against the
selected member (and its non-Pictures bytes) before reporting the sample. The
summary distinguishes `pictures_compressed_range_bytes` (aggregate compressed
ZIP ranges) from `pictures_uncompressed_payload_bytes` and its uncompressed
payload digest; selected-media fields likewise distinguish compressed-range
bytes from uncompressed payload bytes/digest. The eager records retain the
same semantic/media digests but intentionally have no source vectors. These
selectors provide correctness/logical-read evidence only; they make no
latency, physical-I/O, decompression, allocation, RSS, or release ABBA claim.

`odp_file_eager_open` and `odp_file_source_open` are matched unified-root
controls over the existing 12-slide/eight-2 MiB `Pictures/` corpus;
`odp_file_eager_selected_slide` and `odp_file_source_selected_slide` prepare
the corresponding owner before timing and time only the deterministic middle
slide query. The eager control uses `litchi::Presentation::from_bytes` and the
source control uses `litchi::Presentation::open` on one prepared temporary
file. Corpus creation, writing, owner preparation for query controls, complete
semantic/metadata/media/member/hash parity, and selected-media verification
are outside timing. Source records pair each measured sample with an
independent direct `litchi_odp::SourceBackedPresentation` replay through an
instrumented positional source; open/query counters and selected-media range
overlap are kept separate from root timing. These selectors make no latency,
physical-I/O, decompression, allocation, RSS, or release-ABBA claim.

Each ODP batch uses `Builder`, `Presentation::from_bytes`, `slides()`,
`Presentation::text`, source snapshots, and public presentation transactions.
Opened source slides are preservation-only under the public rewrite contract,
so `odp_semantic_one_edit_save` performs the supported single-slide append and
validates every retained slide plus the addition. The current public ODP
builder writes wall-clock timestamps in `meta.xml`; the corpus generator
retains its authored content/style output but repackages it with fixed
benchmark metadata before measurement, so corpus SHA-256 values stay stable
across runs.

OPC parts retain their deterministic `benchmark/parts/00000.bin` names, and
the middle entry remains the fixed `zip_read_one` target. CFB streams are
fixed-width root siblings named `benchmark_stream_00000.bin`, etc. The final,
lexicographically greatest CFB stream is the `cfb_read_one` target. This makes
`wide-root` exercise a long successful validated-tree descent instead of
accidentally measuring a near-root hit, and keeps the original pre-change
full-tree traversal cost reproducible in comparison binaries.

Every synthetic OPC corpus has exactly one package-level Office Document
relationship, `rIdBenchmarkMain`, targeting that middle Part. These packages
use generator identifier `litchi-opc-synthetic-v2`; older v1 corpus hashes
remain distinguishable.

## Cases

- `zip_index`: parse the ZIP central directory and build Soapberry's index.
- `zip_read_one`: read and verify one target member from an already indexed ZIP.
- `opc_open`: parse a generated OPC ZIP into `litchi_opc::OpcPackage`.
- `opc_open_owned`: move an already-prepared archive allocation into
  `OpcPackage`; the input clone happens before timing.
- `opc_noop_save`: open the owned archive bytes once, then serialize the
  unchanged package through a bounded non-seekable counting sink. This measures
  Litchi's exact owned-source no-op path and verifies byte-identical output.
- `opc_mutated_save`: change one byte in the fixed target Part before timing,
  then publish through the full validated rewrite path and compare every output
  byte with a deterministic expected archive.
- `opc_source_open`: build the positional source's validated catalog with an
  explicit finite cache policy. A post-timing first payload access must still
  perform source I/O, proving that open did not materialize it.
- `opc_source_open_main_read`: open the same source and materialize its unique
  Office Document main Part in one timed operation.
- `opc_source_cached_main_read`: cold-load before timing, then require the timed
  access to add zero source reads and return the same pinned `Arc` allocation.
- `opc_source_concurrent_same_part`: start two cold main-Part reads together;
  both must receive the same pinned allocation, and a third access must be a
  zero-I/O cache hit.
- `opc_source_overlay_one_part_save`: on the fixed few-large incompressible
  corpus, time `InstrumentedSource` open and the source-backed one-Part
  publisher, replacing the existing main Part under the same URI and content
  type and writing sequential output. Source reads and sink writes are
  reported; exact bytes, every reopened Part, and output digest are verified
  after timing.
- `docx_source_backed_one_edit_save`: on a fixed 200-paragraph DOCX with eight
  incompressible 2 MiB images, time positional open, one source-backed
  main-document transaction, and its guarded one-Part overlay save. The path
  reports one semantic Part materialization per sample;
  complete paragraph/media/topology and reversible-patch checks remain
  untimed.
- `pptx_source_backed_one_edit_save`: on a fixed 200-slide PPTX with eight text
  boxes per slide and eight incompressible 2 MiB inert PNG parts, time
  positional open, one guarded source-backed shape transaction, and one-Part
  overlay publication through a bounded sequential sink. The path reports two
  semantic Part materializations per sample; full PPTX readback and exact topology,
  relationship, content-type, unselected-Part, and media checks remain untimed.
- `pptx_eager_batch_edit_save`: use the same media-rich corpus, eagerly own all
  229 ordinary Parts, atomically replace all eight text boxes on the selected
  slide, and publish through a bounded sequential sink.
- `pptx_source_backed_batch_edit_save`: perform the identical atomic eight-shape
  edit while materializing only the presentation root and selected slide. Both
  cases require the same deterministic output hash and complete untimed
  preservation/readback checks.
- `pptx_file_{eager,source}_{open,list_slides,slide_count,selected_slide}`:
  compare the unified ordinary-root filesystem path with an eager byte-root
  control on the fixed 200-slide/eight-text-box/eight-2 MiB-media corpus.
  Source samples use a separate untimed `SourceBackedPresentation` replay for
  exact slide/media range classification; eager samples have no `ReadAt`
  counters. These selectors are correctness/logical-read evidence only and
  do not make a latency or resource claim.
- `xlsx_eager_calculation_metadata_edit_save`: on the fixed media-rich XLSX
  corpus, time positional open, eager ownership of all twelve ordinary Parts,
  one calculation-properties transaction, and full sequential publication.
  It is the matched semantic control for the source-backed case.
- `xlsx_source_backed_calculation_metadata_edit_save`: perform the same
  `calcPr` edit while materializing only `xl/workbook.xml` and raw-copying all
  eleven unselected Parts. Both paths emit byte-identical output and run the
  same complete untimed semantic and preservation verification.
- `xlsx_eager_page_break_edit_save`: on the same fixed media-rich XLSX corpus,
  time positional open, eager ownership of all twelve ordinary Parts, one
  selected-worksheet page-break transaction, and full sequential publication.
- `xlsx_source_backed_page_break_edit_save`: perform the same page-break edit
  while materializing only `xl/workbook.xml` and
  `xl/worksheets/sheet1.xml`. The other ten ordinary Parts remain deferred and
  are physically raw-copied during one-Part overlay publication.
- `xlsx_eager_merge_commit_save` / `xlsx_eager_unmerge_commit_save`: on the
  deterministic sparse A1:B2 fixture, prepare the merge or unmerge transaction
  outside timing, then time only eager semantic commit plus bounded sequential
  save. Complete merge/split membership, anchor and covered/uncovered cell
  semantics, unrelated-cell retention, exact durable patch apply/inverse, and
  stale-source refusal are checked outside timing. These cases make no latency
  claim without controlled ABBA evidence.
- `cfb_open`: parse the complete generated container into `litchi_cfb::OleFile`.
- `cfb_list_streams`: enumerate and materialize all stream paths from an
  already-open CFB container.
- `cfb_read_one`: look up, materialize, and verify the final root stream from an
  already-open CFB container.
- `cfb_create_stream_borrowed`: insert one already-prepared target payload with
  `OleWriter::create_stream`, including its required payload copy in the timed
  interval.
- `cfb_create_stream_owned`: insert the same already-prepared target allocation
  with `OleWriter::create_stream_owned`. Comparing this directly with the
  borrowed case isolates the avoidable payload-copy cost.
- `cfb_shared_open`: validate and index a CFB container through
  `SharedOleFile` over the instrumented positional source.
- `cfb_shared_read_one`: open outside timing, reset counters, then read and
  verify the fixed target stream through positional I/O.
- `cfb_shared_concurrent_reads`: start reads of the first and final root streams
  together after open. The result records observed maximum in-flight reads
  without requiring overlap on very small corpora.
- `cfb_selective_mini_legacy_read` / `cfb_selective_mini_shared_read`: paired
  legacy full-stream materialization and positional exact-range reads of a
  deterministic 36-byte MiniFAT target at the final position among 256 or
  2,048 sibling streams. The positional case allocates an exact caller buffer
  inside the read stage and does not populate the root-mini-stream cache.
- `cfb_selective_mini_4095_legacy_read` /
  `cfb_selective_mini_4095_shared_read`: the same paired controls for a
  deterministic 4095-byte MiniFAT target at the final position. The target
  occupies 64 logical 64-byte mini-sectors and exposes physical-run request
  amplification: legacy materializes the complete root mini-stream, while the
  positional path records the exact source ranges used to fill the 4095-byte
  caller buffer. Source call/byte/range vectors and separate open/read/total
  timings are evidence only; no speed claim is implied.
- `cfb_selective_fat_legacy_read` / `cfb_selective_fat_shared_read`: the same
  paired control for a deterministic 4 MiB FAT target. These six selectors
  are opt-in and only emit the `many-small` and `wide-root` shapes; each result
  records separate open/read/total timings, stage-local instrumented read
  calls/bytes/range sizes, returned payload bytes, and hashes. They retain no
  sink. Change 0094 records the accepted pinned ABBA result and its explicit
  claim boundary; a one-sample invocation remains correctness evidence only.

Run the matched selective-read evidence explicitly (it is not in the default
matrix):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --shape many-small,wide-root \
  --case cfb_selective_mini_legacy_read,cfb_selective_mini_shared_read,\
cfb_selective_mini_4095_legacy_read,cfb_selective_mini_4095_shared_read,\
cfb_selective_fat_legacy_read,cfb_selective_fat_shared_read \
  --json target/perf/cfb-selective-read.json
```
The six simulated selectors (`cfb_selective_simulated_{mini,mini_4095,fat}_
{legacy,shared}_read`) use the same `many-small` and `wide-root` targets with
the configured fixed latency, request overhead, bandwidth, and physical-range
ceiling. Legacy uses a bounded delayed cursor; shared uses `ReadAt` range
simulation. Open/read/total logical and physical request counts, bytes, sorted
sizes, buckets, and deterministic service-floor nanoseconds are emitted beside
the existing phase counters and target hash. They are opt-in evidence only.
Change 0144 accepts p50/p95 only for its clean, configured-simulator ABBA;
no cold-I/O, physical-device, ambient-network, allocation, RSS, or semantic
native-Office claim is made.
- `doc_fresh_write_to`: construct a new `litchi_doc::writer::Writer`, add the
  selected fixed paragraphs through its public API, and package it with public
  `write_to`.
- `xls_fresh_write_to`: construct a new `litchi_xls::writer::Writer`, add the
  selected sheets and cells through its public API, and package it with public
  `write_to`.
- `ppt_fresh_write_to`: construct a new `litchi_ppt::writer::Writer`, add the
  selected slides and text boxes through its public API, and package it with
  public `write_to`.
- `doc_semantic_open`: open the generated DOC container and its document model
  through `litchi_doc::Package`.
- `doc_semantic_list_paragraphs` / `doc_semantic_one_paragraph` /
  `doc_semantic_full_text`: time ordinary document paragraph enumeration,
  middle-paragraph selection, or complete text extraction. The public
  one-paragraph path necessarily materializes the paragraph collection.
- `doc_semantic_noop_edit_save` / `doc_semantic_one_edit_save`: start from an
  already-open exact-source `body_text::Snapshot`, publish zero or one middle
  paragraph replacement, and materialize owned bytes. Exact no-op bytes,
  deterministic changed bytes, public reopen, forward patch application, and
  inverse restoration are checked outside timing.
- `doc_body_snapshot_list_paragraphs`: open one exact-source
  `body_text::Snapshot` before timing, then time only
  `paragraphs(Projection::All)`. Every returned position and paragraph text is
  verified after timing. This narrow opt-in case attributes the paragraph
  terminator/PAPX containment work used by exact-source DOC transactions.
- `xls_semantic_open`: open the generated native workbook through the ordinary
  `litchi_xls::Workbook` reader.
- `xls_semantic_list_worksheets` / `xls_semantic_one_cell` /
  `xls_semantic_full_cell_scan`: enumerate workbook tabs, resolve one middle
  numeric cell, or traverse every stored cell through the ordinary public
  worksheet API.
- `xls_semantic_noop_edit_save` / `xls_semantic_one_edit_save`: start from an
  already-open exact-source `cell_values::Snapshot`, publish zero or one
  middle-cell replacement, and materialize owned bytes with diagnostics,
  exact patch replay, inverse, and full-grid reopen checks outside timing.
- `ppt_semantic_open`: open the generated native presentation and its semantic
  presentation model through `litchi_ppt::Package`.
- `ppt_semantic_list_slides` / `ppt_semantic_one_shape_text` /
  `ppt_semantic_full_text`: enumerate slides, resolve one middle textbox, or
  extract all presentation text. The selected-shape path necessarily builds
  the public slide collection.
- `ppt_source_backed_one_shape_text` is the matched positional-source
  selected-shape query. It opens `text_edit::SourceSnapshot` before timing,
  then times only `read_text(Target)`; `ppt_semantic_fresh_open_one_shape_text`
  and `ppt_source_backed_fresh_open_one_shape_text` pair fresh eager/source
  open-plus-query phases. All four controls use the same deterministic PPT
  corpus and target. Source-backed elapsed samples use an uninstrumented
  immutable source; an independent `InstrumentedSource` replay reports exact
  logical read calls/bytes under `source.ppt_shape_text`. These controls make
  no allocation, physical-I/O, or speedup claim.
- `ppt_semantic_repeated_shape_text` / `ppt_source_backed_repeated_shape_text`
  keep one prepared eager/source-backed owner and issue eight identical
  selected-shape text queries. Setup is outside timing; source timing uses an
  uninstrumented `OwnedSource`, while an untimed replay records exact logical
  `ReadAt` calls, bytes, prior-covered bytes per later logical read, and one
  canonical semantic digest. Range overlap is coalesced within each current
  read; a byte may contribute again on a later query. These controls make no
  latency or resource claim.
- `ppt_slide_order_snapshot_open`: capture the public exact-source root
  `slide_order::Snapshot`, including its complete package, live-document,
  slide-order, and review-history validation. Generic public-reader semantic
  verification remains outside timing.
- `ppt_text_edit_one_edit_save`: start from an already-open exact-source
  `text_edit::Snapshot`, then time direct target resolution, one middle-shape
  replacement, and commit. Exact patch replay, inverse restoration, direct
  text-edit readback, and the complete generic semantic reopen remain outside
  timing.
- `ppt_semantic_noop_edit_save` / `ppt_semantic_one_edit_save`: start from an
  already-open exact-source `slide_order::Snapshot`, publish zero or one
  middle-shape text replacement, and materialize owned bytes. The complete
  patch and exact inverse are verified before reopening all slides and shapes.
  Native semantic cases use only `tiny` and `large`; `payload-heavy` remains a
  writer-throughput shape rather than a semantic-edit shape.
- `xlsx_open_owned`: move a prepared owned XLSX allocation into public
  `Workbook::from_bytes`; the input clone happens before timing.
- `xlsx_list_sheets`: enumerate sheet names after an already-open workbook,
  without asking any worksheet for cell semantics.
- `xlsx_first_cell`: access the first stored cell (`Sheet1!A1`) from an
  already-open workbook, measuring the first worksheet semantic access.
- `xlsx_full_cell_scan`: enumerate every stored cell in `Sheet1` through the
  public full-sheet range.
- `xlsx_narrow_column_range_scan`: preload `Sheet1` outside timing, then scan
  `B1:B<rows>`. On `dense-wide`, that returns 256 cells while traversing all
  65,536 stored cells in the selected rows.
- `xlsx_noop_commit`: commit a fresh public edit with no mutations and verify
  its patch is empty and its complete workbook semantics are unchanged.
- `xlsx_noop_commit_save`: do the no-op commit and public `Workbook::write_to`
  into the bounded in-memory sink, then require exact generated-corpus bytes.
- `xlsx_one_cell_commit` / `xlsx_one_cell_commit_save`: prepare one fixed cell
  update before timing, then time commit alone or commit plus public
  `write_to`; the resulting patch, workbook semantics, and save bytes are
  verified.
- `xlsx_one_cell_commit_first_read`: prepare the same fixed update before
  timing, then time commit plus the first public read of that cell. This
  attributes duplicate validation/materialization work without changing the
  36-case default matrix.
- `xlsx_one_percent_commit` / `xlsx_one_percent_commit_save`: do the same for
  the shape's deterministic ~1% update set (2 / 41 / 1,311 cells).
- `xlsx_source_open`: open `SourceBackedWorkbook` over the instrumented
  positional source. A fresh post-timing `Sheet1!A1` read must add selected
  worksheet member I/O, proving that open did not semantically materialize a
  worksheet.
- `xlsx_source_list_sheets`: enumerate the source-backed catalog after open,
  require zero timed source reads, and use the same fresh first-cell delta to
  prove that listing did not materialize a worksheet.
- `xlsx_source_first_cell`: cold-read and verify `Sheet1!A1`, then require a
  fresh second-sheet access to add unselected worksheet member I/O. This proves
  that the selected read did not semantically materialize another worksheet.
- `xlsx_source_narrow_column_range_scan`: cold-read and verify every address
  and value in `B1:B<rows>` through the source-backed public API, with the same
  second-sheet deferral proof.
- `opc_range_source_open` / `opc_range_source_open_main_read`: repeat the OPC
  structural-open and open-plus-main-Part flows through the deterministic
  latency/bandwidth/range simulator. The structural case requires a fresh main
  read to add physical requests after timing.
- `xlsx_range_source_open`, `xlsx_range_source_list_sheets`,
  `xlsx_range_source_first_cell`, and
  `xlsx_range_source_narrow_column_range_scan`: repeat the corresponding
  source-backed XLSX flows through the simulator. Listing must issue zero timed
  logical or physical requests; selected reads must leave every exact
  unselected worksheet compressed range untouched and pass a fresh second-sheet
  deferral probe.
- `opc_open_session_scaling`: eager-open every ZIP member with a caller-sized
  `OpenSession` local pool, then verify every generated OPC Part.
- `cfb_bulk_read_scaling`: use `SharedOleFile::bulk_read` with a caller-sized
  local pool, prewarmed outside timing, and verify every stream in input order.
  Neither scaling case uses the global Rayon pool.
- `docx_semantic_open`: open a complete in-memory DOCX through public
  `Package::from_reader`; the prepared owned input clone is outside timing.
- `docx_semantic_list_paragraphs` / `docx_semantic_one_paragraph` /
  `docx_semantic_full_text`: list every paragraph, inspect one middle paragraph
  through `document().paragraph(index)`, or extract complete document text.
- `docx_semantic_create_small`: author and serialize the tiny corpus entirely
  through public DOCX APIs, then reopen and fully verify it.
- `docx_semantic_noop_edit_save`, `docx_semantic_one_edit_save`, and
  `docx_semantic_one_percent_edit_save`: time document transaction capture,
  no-op/one/~1% paragraph replacement, publication, and seekable-stream save;
  reopen and verify every paragraph, full text, operation count, and recorded
  sink behavior. Multi-paragraph selection uses the canonical batch transaction
  while the one-edit case stays scalar.
- `pptx_semantic_open`: open a complete in-memory PPTX through public
  `Package::from_bytes`.
- `pptx_semantic_list_slides` / `pptx_semantic_one_slide` /
  `pptx_semantic_full_text`: enumerate the slide graph, inspect one middle
  slide's shape scene, or flatten all slide text.
- `pptx_semantic_create_small`: author and serialize the tiny corpus entirely
  through public PPTX APIs, then reopen and fully verify it.
- `pptx_semantic_noop_edit_save`, `pptx_semantic_one_edit_save`, and
  `pptx_semantic_one_percent_edit_save`: move owned input into `from_vec`,
  time opened-presentation transaction capture, no-op/one/~1% text-box edits,
  commit, publication, and `to_bytes`, then reopen and verify all slides,
  shapes, and text. The public API has no save-to-sink method.
- `rtf_semantic_open`: parse deterministic owned transport bytes through
  public `Document::from_bytes`.
- `rtf_semantic_paragraph_count`: query the public exact paragraph cardinality.
- `rtf_semantic_collect_paragraphs`: collect all lazy paragraph views so the
  allocation behavior is guarded separately from traversal.
- `rtf_semantic_list_paragraphs`, `rtf_semantic_one_paragraph`, and
  `rtf_semantic_full_text`: traverse lazy body paragraph views, resolve and
  flatten one middle paragraph, or materialize the snapshot's cached complete
  text for the first time.
- `rtf_semantic_text_to_sink`: write semantic UTF-8 body text to a pre-reserved,
  bounded, forward-only non-seek sink and verify the complete output outside
  the timed interval.
- `rtf_semantic_stream_save`: stream the immutable snapshot through public
  `Document::write_to` and require byte-exact source output.
- `rtf_semantic_noop_edit_save`: commit an empty edit, require shared snapshot
  identity and exact bytes, then stream through the same forward-only sink.
- `rtf_semantic_one_edit_save`: replace the middle paragraph through the
  checked native transaction, stream the changed snapshot, reopen it, and
  verify its complete semantic projection and sink counters.
- `rtf_semantic_one_percent_edit_save`: select deterministic evenly spaced
  `ceil(1%)` paragraph positions, stage one bounded atomic batch, commit once,
  stream once, and verify exact patch replay, inverse restoration, late-failure
  atomicity, and complete reopened text outside the timed interval.
- `rtf_semantic_remove_paragraph_save`: remove the exact middle paragraph from
  the matched generated default-formatted plain lifecycle corpus, commit once,
  and serialize into the same bounded
  sink. Complete reopen/full projection, volatile and durable forward/inverse,
  stale-source refusal, and the output hash are checked outside timing.
- `rtf_semantic_move_paragraph_save`: move the first paragraph to the final
  list position and perform the same untimed correctness gates. An
  additional untimed equal-position move proves exact snapshot/byte no-op.
  Both lifecycle cases deliberately reject changed CP-1252, LZFu, producer
  watermark, and opaque/formatted sources; their unchanged equal-position move
  remains exact. Native RTF has no logical-Part materialization counter, so the
  report records honest output hashes and bounded sink counters rather than a
  fabricated Part count.
- `rtf_logical_tail_append`: stage a bounded batch of borrowed plain paragraphs
  immediately before the exact root close of an existing lifecycle document,
  then commit and publish through the windowed non-seek sink. The untimed
  gates prove complete paragraph/text readback, exact source bytes, in-memory
  patch/inverse, durable JSON apply/inverse, and foreign-source refusal.
- `rtf_logical_tail_noop_save`: run the same existing-document tail transaction
  with an empty batch. It measures exact no-op commit plus sequential
  publication and reports shared snapshot identity separately from the changed
  append case.

The two lifecycle intervals include edit construction, the one staging call,
commit, a constant-size diagnostics assertion, one shared snapshot-handle
clone, and bounded sequential serialization. Complete byte/projection,
volatile/durable patch, stale-source, no-op, refusal, sink-summary and hash
oracles remain outside timing.

For both CFB stream-insertion cases, payload generation/cloning and writer
construction happen before timing, while writer and source destruction happen
afterward. They time stream insertion only, not complete CFB serialization. The
`few-large` shape is the useful ownership profile because its prepared target is
4 MiB; `tiny` stays small enough for a quick release-build smoke test.

The sink reserves a checked byte budget before timing and copies every output
byte into bounded memory. It records accepted bytes, write count, and the
largest write, rejects output beyond the budget or individual writes larger
than the case's configured ceiling, and verifies the deterministic no-op
output after timing. Most passthrough cases use a 64 KiB per-write ceiling;
whole-document RTF save cases permit one expected-output-sized write. This
keeps allocator growth out of the timed interval while preventing a
passthrough write from degenerating into bookkeeping that never reads bytes.

Each fresh writer iteration begins with a new writer and ends after its public
`write_to` call has completed into a seekable `Cursor<Vec<u8>>`. This exercises
the real DOC/XLS/PPT OLE packaging path used by the corresponding `save` API
without timing filesystem state. Construction and packaging are both included;
post-timing validation requires the package bytes to exactly match the
precomputed deterministic corpus. The manifest records the complete output
hash and the hash of its conventional primary stream (`WordDocument`,
`Workbook`, or `PowerPoint Document`).

Each XLSX save case uses public `Workbook::write_to` rather than filesystem
`save`, so it measures the same sequential serialization path without timing
filesystem state. No-op byte exactness is deliberately scoped to the harness's
generated XLSX corpus, whose owned-source serialization is deterministic; it
does not claim byte preservation for arbitrary third-party XLSX input.

## JSON contract

Reports use `schema_version: 1`. Each result contains sorted raw elapsed-ns
samples plus min, median p50, nearest-rank p95 and p99, max, mean, sample
standard deviation, and a two-sided Student's-t 95% confidence interval for
the mean. Untimed warm-up iterations are recorded in the configuration but are
never included in these statistics.

Each case embeds a corpus manifest with the format-specific generator version,
container format, shape, payload algorithm, logical and serialized sizes,
member count, target member, and SHA-256 hashes. Existing OPC manifest fields
are unchanged; CFB reports use `compression: "none"`
and count streams in `archive_member_count`. The report also records the git
revision and dirty state when available, `rustc` version, build profile and
target, visible logical CPUs, allocator, relevant Cargo/Rust flags, and Linux
`perf_event_paranoid` value. Filesystem-enabled reports additionally expose
best-effort non-path host evidence: OS/kernel, CPU model, total memory, page
size, filesystem type, current CPU affinity, and whether the measured
source/destination probe used one device. A storage identifier is deliberately
serialized as `null`; absolute paths and device identifiers are never emitted.
Metadata collection is best-effort and is complete before timed iterations.
The requested JSON parent directory is created automatically.

`configuration.writer_shapes` is an additive schema-v1 field identifying the
fresh writer shape selection. For writer records, `entry_count` is the logical
paragraph/cell/text-box count, `uncompressed_payload_bytes` is the generated
logical content byte count, and `entry_bytes` is zero because the content has
no uniform per-entry byte size. Existing substrate fields and generator
identifiers retain their previous meanings.

`configuration.xlsx_shapes` and the optional `corpus.xlsx` object are additive
schema-v1 fields. `corpus.xlsx` records sheet dimensions, the exact
deterministic ~1% update count, and `source_members`: the workbook, worksheet,
shared-string, and style ZIP member names whose exact compressed ranges feed
the positional counters. Existing non-XLSX corpus records keep it null.

`configuration.rtf_variants` and the optional `corpus.rtf_variant` field are
additive schema-v1 identifiers for the selected RTF input capabilities.
Non-RTF corpus records omit `rtf_variant`. Corpus names include the variant so
repeated case names remain unambiguous in multi-variant reports.

`configuration.range_simulation` records fixed latency, request overhead,
bandwidth, and maximum physical range. `configuration.execution_workers`
records the resolved, capped, deduplicated scaling points in deterministic
ascending order.

The twenty-three positional cases add a `source` object; older cases omit it. Its
arrays contain one value for every measured iteration and record `read_calls`,
`read_bytes`, compressed ordinary-OPC-payload range overlap, and
`max_in_flight_reads`. Applicable OPC cases also record a semantic per-sample
`ordinary_payload_materializations` count; it may be zero, one, or multiple
Parts depending on the timed operation. Payload-range overlap is a
physical request-amplification metric: bounded ZIP metadata reads may fetch
adjacent compressed payload bytes without decompressing or caching that Part.
Accordingly, `opc_source_open` may report overlap while still reporting zero
materializations; its post-timing cold access proves the distinction.
Native PPT selected-shape source cases additionally record the scalar canonical
text digest and per-sample logical read calls/bytes under
`source.ppt_shape_text`.

Simulated-range records additionally contain `source.simulation`: per-sample
logical read calls/bytes, physical request count/bytes, sorted physical request
sizes, and fixed request-size buckets. Request delays are computed only from
the recorded configuration; no ambient network or clock-derived input is used.

Filesystem records are additive under `filesystem_evidence`. Each evidence
sample is keyed by case, corpus manifest, sample index, and `cache_state`
(`warm` or `cold-requested`), and pairs child elapsed time with parent-observed
wall time. It includes logical `ReadAt` requested/returned bytes, fixed request
size buckets, maximum concurrent reads, procfs I/O/fault/RSS deltas (with
post-sample `VmHWM`), output SHA-256 and byte length for saves, and public API
publication counters where available. `opc_materialized_parts` is explicit
zero for the raw-copy source overlay path; CFB records changed spans and
published bytes. The configuration records selected cache states, process
isolation, fresh-child sampling, and whether a caller-selected filesystem root
was used. PPTX ordinary-root samples additionally record
`logical_read_counter_scope` and, for source candidates only, an untimed
`pptx_source_replay` object with source hash, total request sizes, exact slide
and media payload-range overlap, selected/unselected slide counters, union
coverage bytes, full-range coverage counts, semantic hash, and classification.
Selected replay classification requires the complete target compressed range;
list replay classification requires every slide range. Eager PPTX samples set
the replay object to null and mark the generic counter scope
`not_applicable_eager_pptx`; their zero generic fields must not be interpreted
as source-read measurements.
Eager OPC filesystem samples use the same explicit boundary,
`not_applicable_eager_opc`, because their timed `fs::read` path has no
`ReadAt` counter; source-backed OPC and CFB overlay samples retain
`timed_read_at` when their positional counter is active.

Each filesystem `CaseResult` additionally carries an additive
`operation_metrics` envelope. Its `sample_count` and `alignment` identify the
sorted `elapsed_ns.samples` vector, and every measured numeric vector has that
same cardinality and order. The envelope separates logical source-read
vectors, procfs process vectors, post-operation output length, publication
counters, and OPC materialization counts. A vector with a measured zero is
serialized as a numeric zero; unsupported or unavailable values omit the
numeric vector and retain an explicit `status` (`not_applicable` or
`unavailable`). Procfs CPU/fault/context-switch/RSS values are operation
deltas. `peak_rss_bytes` is the process-lifetime after-sample `VmHWM`, not an
operation peak, and `rss_delta_bytes` is not a peak. The envelope makes no
allocation, copied-byte, decompressed-byte, or recompressed-byte claim because
the filesystem child does not instrument those quantities. The existing raw
`filesystem_evidence` samples remain unchanged.

ODP media-rich selectors additionally record `source.odp_media`: the exact
phase/timing scope, selected middle slide and media member, canonical full
semantic digest, uncompressed media payload digest, total source calls/bytes
and prior-range overlap, compressed `Pictures/*` overlap, and independent
selected-media calls/bytes/prior-range-overlap vectors. The
`pictures_compressed_range_bytes` and `selected_media_compressed_range_bytes`
fields are compressed ZIP ranges; the corresponding `*_uncompressed_*` fields
are payload bytes or digests. The selected-media replay proves one full
Pictures range and reports bytes outside Pictures separately; its prior-read
and compressed-range overlap vectors are named
`selected_media_read_prior_range_overlap_bytes` and
`selected_media_read_compressed_range_overlap_bytes`. Source open and
one-slide query vectors remain separate from that explicit media-read replay.

Positional XLSX records additionally contain `source.xlsx` arrays for physical
overlap with the workbook, selected worksheet, all unselected worksheets,
shared strings, and styles compressed member ranges. These overlap counters
are intentionally truthful about ZIP read amplification and therefore are not
semantic materialization counters. No XLSX materialization count is emitted,
because the production API does not directly expose one. Instead, each case
enforces its semantic deferral claim with a fresh post-timing worksheet access
that must add I/O for that worksheet's exact compressed member range.

Scaling records contain `execution.worker_count`, `logical_tasks`, and
`logical_bytes`. One result is emitted per resolved worker count, in ascending
order. Those fields and the elapsed samples are sufficient to compute
throughput, speedup, scaling efficiency, and an Amdahl serial-fraction estimate
outside the harness.

Publication cases may additionally emit `output_sha256`, independently
identifying the deterministic changed archive without changing schema v1. For
`opc_source_overlay_one_part_save`, its
`ordinary_payload_materializations` value is exactly one per sample: the
selected original Part is validated, while every unselected member is copied
physically without semantic materialization.

Streaming-creation cases also emit `output_sha256`. Their `sink` extensions are
content-free scalar evidence; `retained_output_bytes` is always zero and
`retained_authoring_window_bytes` is the row/text encoder bound, not process RSS or a
claim about allocator internals.

Every serialized `sink` summary also contains the fixed
`write_size_buckets` object: `bytes_0`, `bytes_1_to_512`,
`bytes_513_to_4096`, `bytes_4097_to_16384`, `bytes_16385_to_65536`, and
`bytes_over_65536`. These are counts of logical bytes accepted by the sink's
`Write::write` calls, including accepted zero-length calls; rejected calls do
not increment any counter. The six bucket counts always sum to `write_calls`,
including at the exact boundary values. They do not measure syscalls, disk
I/O, memory copies, compression, or performance.

Logical-tail RTF cases also emit `output_sha256`. Their `sink.rtf_tail_append`
object records the operation (`append` or `exact_noop`), source, caller-input,
inserted, and published output bytes, appended paragraph/run counts, the 16 KiB
sink window, and boolean exact-no-op, in-memory patch, durable patch, reopen,
and source-conflict gates. `retained_output_bytes: 0` describes only the timed
sink; the append API intentionally retains its validated candidate snapshot,
so the window is not a process-RSS or transaction-memory claim.

Native XLS comment cases emit `output_sha256` and a `source.xls_comments`
object with the explicit owned-source counter scope, source-backed flag,
update count, separate semantic/publication distributions, changed
comment/stream counts, exact NOTE/TXO `splice_count` and `replacement_bytes`
(source-backed cases), source and Workbook lengths, and changed physical spans
and exact source/target fingerprints. Native XLS worksheet-visibility cases
emit the same shape under `source.xls_visibility`, with changed
worksheet/stream counts and exact `BoundSheet8.hsState` splice/replacement
diagnostics. These diagnostics are content-free and do not imply a bounded
candidate allocation or a performance advantage.

Native XLS fixed-width numeric cases emit `output_sha256` and a
`source.xls_numeric` object. It records the Number or RK/MulRK family, update
count, separate edit/set/commit/publication/total vectors, input/output CFB and
Workbook sizes, source-backed splice/replacement/span/fingerprint vectors, and
sink bytes/write-count/digest vectors. The eager/source-backed selectors retain
and report complete target materialization; the two plan-only selectors report
`target_artifact_retained_at_commit: false`,
`target_artifact_materialized_at_commit: false`, and zero
`complete_target_materialized_bytes` while their sink vectors still prove full
publication bytes. Their forward-only contract reports
`patch_or_inverse_supported: false`. Generic `source.read_calls` and
`source.read_bytes` carry the owned source-ingress counters. Its
`owned_input_scope` explicitly
describes complete in-memory CFB ingress; source-backed publication is not a
bounded-artifact-memory or positional-I/O claim. Plan-only composed semantic
validation may read and allocate a candidate `Workbook` model, so zero
commit-boundary target-artifact bytes do not imply zero target-semantic
allocation or bounded total memory.

`docx_source_backed_one_edit_save` also emits `output_sha256`, source/sink
distributions, and `ordinary_payload_materializations`. Its value is exactly
one per sample: the raw main document is loaded for semantic selection while
the eight media payloads and every other unselected member remain physically
source-backed through publication.

`pptx_source_backed_one_edit_save` also emits deterministic `output_sha256`,
source/sink distributions, and `ordinary_payload_materializations`. Its value
is exactly two per sample: the mandatory presentation root and selected slide
are loaded for semantic validation, while every other slide and all eight media
payloads remain source-backed through physical raw-copy publication.

The paired PPTX batch controls emit the same evidence and byte-identical
output. Their `ordinary_payload_materializations` values are exactly 229 and
two per sample, respectively.

The two XLSX calculation-metadata publication cases also emit deterministic
`output_sha256`, complete source/sink distributions, and semantic
materialization counts. The eager control reports twelve Parts per sample; the
source-backed path reports one. Their output hashes are required to match.

The two XLSX defined-name publication cases use the same twelve-Part media-rich
archive. Their eager/source-backed materialization counts are twelve and one,
respectively, and their output hashes are required to match.

The two XLSX page-break publication cases emit the same evidence. Their eager
control reports twelve semantic materializations per sample and the
source-backed path reports two; their output hashes are also required to
match.

The two XLSX page-margin publication cases use the same twelve-Part media-rich
archive and evidence contract. The eager control reports twelve semantic
materializations per sample and the source-backed path reports two; their
output hashes are required to match.

The two XLSX print-options publication cases use that same matched evidence
contract. Their eager/source-backed materialization counts are twelve and two,
respectively, and their output hashes are required to match.

The two XLSX page-setup publication cases use the same twelve-Part archive and
evidence contract for relationship-free settings. Their eager/source-backed
materialization counts are twelve and two, respectively, and their output
hashes are required to match.

The two XLSX sheet-protection publication cases use the same archive and bind
the complete selected-worksheet relationship closure. Their eager/source-backed
materialization counts are twelve and two, respectively, and their complete
typed readback and output hashes are required to match.

The two XLSX data-validation publication cases use an equivalent twelve-Part
media-rich archive seeded with core and Office 2010 collections. Their
eager/source-backed materialization counts are twelve and two, respectively,
and their complete typed readback and output hashes are required to match.

The two XLSX auto-filter publication cases use an equivalent twelve-Part
media-rich archive seeded with a value filter and sort state. Their
eager/source-backed materialization counts are twelve and three, respectively,
because both controls validate the styles relationship needed by authored
color/DXF filters and sorts; their complete typed readback and output hashes
are required to match.

## External profiling

Build the binary once before collecting counters, then invoke it directly so
Cargo is not part of the profile:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
perf stat -d tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --case cfb_read_one --shape wide-root \
  --payload incompressible --json target/perf/perf-cfb-read.json
```

For a sampled profile, use whichever installed tool is appropriate:

```sh
cargo flamegraph --release --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --case opc_noop_save --shape few-large \
  --payload compressible --json target/perf/flamegraph-run.json

samply record tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --case opc_noop_save --shape few-large \
  --payload compressible --json target/perf/samply-run.json
```

External profilers are optional. On Linux, `perf` can be installed yet denied
by the kernel's `perf_event_paranoid` policy or container permissions. The
harness records that policy when readable and still runs normally without any
profiler, providing the wall-clock JSON baseline. If `perf` is unavailable or
denied, run the binary directly (or use `samply`/`cargo flamegraph` if present)
rather than changing system policy merely to complete a smoke run.

Timing is only a first baseline. It intentionally does not claim peak RSS,
allocation counts, CPU utilization, lock contention, or cache misses; those
need dedicated instrumentation and a controlled runner.
