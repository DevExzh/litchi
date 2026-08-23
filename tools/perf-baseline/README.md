# OPC, CFB, OLE2 Office, OOXML, RTF, and ODF performance baseline

`litchi-perf-baseline` is an isolated, reproducible measurement tool for the
ZIP/OPC and CFB/OLE2 substrates, fresh DOC/XLS/PPT writer packaging, and
public-API XLSX snapshot/edit/save flows, matched opt-in XLSX scalar-cell
eager/source-backed and managed source-backed publication controls, one-cell
eager/source-backed clear/remove lifecycle controls, and opt-in DOC/XLS/PPT,
DOCX/PPTX/RTF/ODT/ODS/ODP semantic flows, including matched ODS
owned/source-backed existing-cell publication controls and the opt-in RTF logical-tail
append and ordinary-paragraph split/adjacent-merge transactions, bounded RTF/XLS/DOCX/PPTX/ODF validation reports, and a
source-backed DOCX section inventory. It creates every corpus in memory; it also exercises
source-backed XLSX catalog, worksheet reads, and guarded calculation-metadata,
defined-name, page-break/page-margin/page-setup/print-options/sheet-protection/data-validation/auto-filter
publication over positional I/O. Four additional XLSX composition selectors
exercise disjoint `Edit::join`, overlapping join refusal, disjoint three-way
planning, and explicit three-way conflict resolution over the same deterministic
cell CRUD corpus. The scalar-cell controls use deterministic
medium and dense/sparse four-sheet corpora with untouched media Parts. Their
timed interval covers open, selector planning, commit, and sequential
publication; reopen, semantic equality, exact hashes, raw media identity,
lifecycle gates, and source/materialization counter sampling remain outside the
reported timing. Source-backed JSON also records those stages separately,
including managed Budget/cache diagnostics and release-to-zero checks. Managed
cell-value records retain pre-publication `InputBytes`/`OutputBytes`/`Work`/
`Objects` limits and usage, cache/catalog object reservations, and an
immediate post-publication shared-budget snapshot so direct output charging is
visible without extending the timed interval.
The reported duration is the sum of open/stage/commit and sequential
publication segments; source cache diagnostics are sampled between those
segments and immediately after managed publication; all such diagnostics are
excluded. They are evidence for later release ABBA work, not a speedup claim.
It does not
depend on untracked office files, network state, or randomness. ODP builder
timestamps are replaced with fixed metadata before measurement. The JSON
report contains the generator parameters and SHA-256 hashes for the generated
container and target entry, so a result always identifies its exact input or
packaged output.

Every newly emitted report also carries an additive top-level
`binary_identity` descriptor for the actual running executable. It follows the
resource-profile descriptor vocabulary: `path`, `binary_sha256`,
`binary_bytes`, `mode_bits`, `executable`, and `profile`. The executable is
hashed with a bounded streaming reader after all timed cases have completed;
the harness does not hash the binary inside warmups, samples, or child
operation timers. Unix reports record permission bits, while platforms without
a portable permission-bit equivalent emit `mode_bits: null`; `profile` remains
explicit on every platform. A changed file identity or size during the read
fails report generation rather than publishing unverifiable provenance.

The tool is intentionally outside the root workspace and has no effect on
production dependency graphs.

Balanced A1/B1/B2/A2 captures can be validated and summarized with the
standard-library-only helper at `tools/perf_abba_summary.py`:

```sh
python3 tools/perf_abba_summary.py \
  control-a.json candidate-a.json candidate-b.json control-b.json \
  --json-out summary.json
```

The helper recomputes p50, mean, p95, and p99 from retained samples and fails
closed when harness configuration, stable environment facts, executable
identity, corpus identity,
source metrics, sink metrics, or result coverage differ between legs. Its
default same-implementation drift ceilings are 5% for p50 and mean, 10% for
p95, and 15% for p99; use `--drift-ceilings` only when a recorded calibration
justifies different values. Reports must be complete schema-v1 JSON from the
`litchi-perf-baseline` harness, with typed tool/configuration/environment
fields, a clean worktree, positive warm-up count, and at least 15 retained
samples per row. A1/A2 must carry one non-empty control revision, B1/B2 one
different non-empty candidate revision; each leg's environment and canonical
report hash is retained in the summary. Configured shape arrays are checked
against result rows; filesystem evidence is the schema-specific exception for
format-specific filesystem shapes such as `media-rich`.

The ABBA binary descriptor is an exact identity within each implementation:
A1 must equal A2 and B1 must equal B2, including path, size, mode, and profile.
The control and candidate `binary_sha256` values must differ. Missing or
malformed descriptors fail closed before any statistic is accepted.

The helper verifies every integer statistic, Welford sample standard deviation,
and the two-sided Student's-t 95% confidence interval using the harness
formula. p50 is the Rust u64 floor midpoint and p95/p99 use integer
nearest-rank. A source or sink that is present must be byte-for-byte/canonical
JSON equal on all four legs; an absent or null value is reported only as
`consistently_absent`, never as `verified_equal`. When present,
`output_sha256` must be a lowercase SHA-256 identity equal on all legs. This
strict source/sink identity policy means the ABBA summary is not suitable for
optimizations that intentionally change I/O, source reads, or sink writes;
use a matching evidence protocol for those changes. It never infers a speedup
claim from an accepted statistic.

The four raw reports and their strict summary can be packaged for archival
with the standard-library-only `tools/perf_abba_package.py` helper. It invokes
the external `zstd` executable in a deterministic single-threaded mode,
retains raw JSON bytes, recomputes the complete summary with the canonical
summary implementation, and refuses any summary tamper, duplicate ABBA role,
output overwrite, or path escape. Output names are flat basenames:

```sh
python3 tools/perf_abba_package.py \
  --change 0238-perf-package \
  --output-dir docs/performance/results/0238-perf-package \
  --summary target/perf/summary.json \
  --artifact a1=target/perf/control-a.json \
  --artifact b1=target/perf/candidate-a.json \
  --artifact b2=target/perf/candidate-b.json \
  --artifact a2=target/perf/control-b.json
```

The deterministic manifest records each artifact's role, compressed and
uncompressed byte counts and SHA-256 digests, plus the summary's raw and
canonical identity and report bindings. It also records the canonical `zstd`
executable path, version string, file size, and executable SHA-256 used for
compression. All files are first written under a private staging directory
inside the output directory and then published with exclusive hard links; a
write, directory, or publication failure removes the staging directory and
every file published by that invocation. Existing files are never replaced;
use a fresh output directory for a rerun. For safety, the selected output
directory itself may not be a symlink. The directory is held open with
`O_DIRECTORY|O_NOFOLLOW` for the complete transaction, and all publication and
cleanup use descriptor-relative operations, so swapping its pathname cannot
redirect the package.

The native OLE2, DOCX/PPTX, RTF, and ODF semantic matrices are deliberately
opt-in. They measure only current public APIs and therefore do not change the
default 36 cases / 198 records.

The XLSB semantic CRUD baseline is a separate opt-in binary so that its
format-specific dependency does not alter the default matrix. It uses the
public XLSB owner and umbrella facade over the deterministic POI
`testVarious.xlsb` fixture (including its inert opaque/VBA members), and
reports warm raw samples with p50/mean/p95/p99, source/output SHA-256
identities, worksheet/cell coordinates, sizes, and untimed reopen,
semantic-readback, exact no-op patch/unrelated-Part semantic-digest,
malformed-input, package/cell resource-limit, and sparse-iteration gates. The stored-cell scan and the
`ceil(1%)` edit are explicitly scoped to the selected worksheet. The binary
fails closed when any gate does not pass:

```sh
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 \
  cargo run --manifest-path tools/perf-baseline/Cargo.toml --release \
  --features xlsb-crud --bin xlsb_crud -- \
  --case all --warmup 3 --samples 30 \
  --json target/perf/xlsb-crud.json
```

The default XLSB fixture SHA-256 is
`8c600e97d719b0266dcfb49c1872feb8d10c6ed12bc768ff16ace7dae555ebfc`.
Use `--fixture` to select another public fixture; capability-dependent
operations refuse unsupported corpora rather than fabricating an edit.

## Run

Run the opt-in DOCX story-hyperlink planning case over its deterministic
49-story/1,152-relationship corpus:

```sh
cargo run --manifest-path tools/perf-baseline/Cargo.toml --release -- \
  --case docx_story_hyperlink_plan --warmup 3 --samples 30 \
  --json target/perf/docx-story-hyperlink-plan.json
```

The corpus, source-backed story inventory, selector guards, one untimed
semantic publication/readback, and one repeat publication for deterministic
output validation are prepared outside the measured interval. Each retained
sample times eight repeated `Snapshot::plan_target_urls` calls on the prepared
immutable snapshot. The case is opt-in and does not change the default 36-case
matrix.

Run the complete default matrix (36 default cases; 198 result records: 144
substrate records, nine writer records, and 45 XLSX records). The six simulated
range cases, two execution-scaling cases, one low-level source-overlay save
case, one source-backed DOCX semantic publication case, one source-backed
media-rich PPTX semantic publication case, four matched same-slide/multi-slide
PPTX batch cases, two owned-source PPTX cross-presentation slide-copy evidence
selectors plus one opt-in source-backed plain cross-copy selector, two matched
cross-slide ODP text-box publication cases, one
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
two XLSX merge/unmerge commit-plus-save cases, four opt-in XLSX edit-composition
selectors, four matched eager/source-backed
XLSX scalar-cell clear/remove lifecycle cases, six matched eager/source-backed
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
four opaque-heavy common OLE2 stage/edit-save cases, 24 native OLE2 semantic cases, 17
DOCX/PPTX semantic cases, 21 RTF semantic cases (15 transport/read/edit
cases plus six logical-tail publication cases), 38 ODF semantic cases, one
ODF `mimetype` repair-plan case, and six matched ODF content-COW publication
cases are opt-in. Eight additional native PPT
`Pictures` selectors are available for matched eager/source-backed open,
cold all-images query, repeated all-images query, and fresh open-plus-all-images
phases on a deterministic picture-heavy corpus. Source-backed elapsed samples
for those native-PPT `Pictures` selectors use an uninstrumented
`litchi_core::OwnedSource`; independent untimed `InstrumentedSource` replays
provide their source-read counters. Before the repeated-edit selector described
below, the `Case` matrix exposed 340 selectable case names, including two opt-in RTF standalone-picture
CRUD selectors, two opt-in RTF ordinary-paragraph split/adjacent-merge
selectors, four opt-in XLSX scalar-cell clear/remove lifecycle selectors, and
four opt-in XLSX existing-row visibility lifecycle selectors, four opt-in ODS
existing-cell publication selectors, and four opt-in XLSX edit-composition
selectors.
One additional matched source-backed ODS repeated-edit selector
(`ods_source_backed_repeated_edit`) is also opt-in over the same two-sheet
media-rich corpus. One `SourceBackedSpreadsheet` owner and four fixed-window
hashing sinks are prepared outside the timer; the timer covers exactly four
sequential one-cell edit/commit/publish transactions, each verified
byte-exactly against the eager owned oracle, with per-sample
stage/commit/publication phase sums and a single untimed `InstrumentedSource`
replay. This brings the selectable matrix to 341 names while leaving the
default 36 cases / 198 records unchanged.
One additional opt-in DOCX semantic selector
(`docx_semantic_one_paragraph_text`) locates the middle paragraph before timing
and times only `Paragraph::text()` for the tiny, medium, and large shapes; exact
text and error guards remain outside the timed interval. This brings the
selectable matrix to 342 names while leaving the default 36 cases / 198 records
unchanged.
Four additional opt-in XLSB lifecycle selectors (`xlsb_semantic_open`,
`xlsb_semantic_list_worksheets`, `xlsb_semantic_one_cell`, and
`xlsb_semantic_full_cell_scan`) use deterministic tiny, medium, large, and
sparse BIFF12 corpora. Archive cloning, workbook/worksheet preparation, and
verification stay outside the timer; the full scan times only boxed
`worksheet.cells()` consumption and verifies an exact canonical cell digest.
These selectors bring the matrix to 346 names while leaving the default 36
cases / 198 records unchanged.
Two additional high-level XLSX filesystem selectors (`xlsx_file_open` and
`xlsx_file_open_lifecycle`) use the deterministic medium cell-CRUD XLSX corpus
and a temporary source file. The first times exactly
`litchi::Workbook::open(Path)`;
the lifecycle selector times that call plus worksheet names, count, and text.
Both selectors compare those projections and metadata with an untimed
`litchi::Workbook::from_bytes(Vec<u8>)` oracle and verify the source archive
SHA-256. They leave the default 36 cases / 198 records unchanged.
Two additional high-level XLSX bytes selectors (`xlsx_bytes_open` and
`xlsx_bytes_open_lifecycle`) exercise `litchi::Workbook::from_bytes(Vec<u8>)`
over the deterministic medium cell-CRUD corpus. The first times exactly the
bytes facade construction; the lifecycle selector times that construction plus
worksheet names, count, and full text. An untimed typed eager
`litchi_xlsx::Workbook` semantic projection (including names, count, and full
text), archive SHA-256, and independently opened eager OPC/property metadata
digest guard each case. These selectors bring the
selectable matrix to 344 names while leaving the default 36 cases / 198
records unchanged.
Four additional opt-in OOXML wire selectors (`omml_formula_range_scan`,
`omml_formula_extract`, `pptx_drawingml_extract`, and
`pptx_drawingml_range_scan`) use one fixed, namespace-adversarial XML corpus.
Their independent `quick_xml::NsReader` projections are prepared and checked
before timing; each timed iteration only executes the selected production
scanner or public extractor and verifies its exact ranges, formula strings, or
text after the clock stops. The formula-string digest is length-delimited and
computed before timing. These selectors bring the
selectable matrix to 348 names while leaving the default 36 cases / 198 records
unchanged.
Eight additional
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
`fs::read` plus a harness-local recreation of the historical eager facade
preparation versus source `Document::open`; the eager control uses smart
detection, the typed DOCX owner, full-text validation, and facade-equivalent
metadata conversion because normal `Document::from_bytes` is source-backed as
of change 0257. Query roots are prepared outside timing and the timer covers
only the named root query. Independent untimed
`litchi_docx::source_backed::Package` replays
record catalog/open reads and, for query selectors, complete coverage of the
compressed main-document range during document preparation plus zero
media/unselected/core overlap during the query. Completed return sizes, range
coverage, materializations, and an explicit classification are also recorded.
Full eager/source semantic parity, exact source hash, logical OPC
part/relationship/content-type/blob-hash gates, media hashes, and source
immutability remain verification outside timing. This is correctness and
logical compressed-range evidence only: it makes no latency, physical-I/O,
decompression, allocation, RSS, cold-cache, ABBA, security, or Markdown claim.
Together these additions bring the selectable matrix to 253 names while
leaving the default 36 cases / 198 records unchanged;
eight matched DOCX/PPTX ordinary-root lifecycle selectors additionally time a
fresh historical eager byte-owner control or source-backed filesystem open
plus paragraph count, full text, slide count, or one selected slide. They reuse
the same fixed media-rich corpora, fresh-child protocol, complete untimed
semantic/archive gates, and independent positional replays. Change 0188
retains the selectors
and raw warm release ABBA, but accepts no latency statistic because every
p50/mean pair misses at least one same-implementation drift gate and all tails
are conservatively withheld. These selectors bring the current matrix to 332
names without changing the default 36 cases / 198 records;
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
two additional matched source-backed ODT repeated-text selectors
(`odt_source_backed_repeated_text_uncached` and
`odt_source_backed_repeated_text_cached`) are also opt-in over one fixed
large 10,000-paragraph ODT with eight deterministic 2 MiB `Pictures/*`
members. Each source owner and four output slots are prepared outside timing.
The uncached control reproduces the current `SourceBackedDocument::text`
semantics from retained `content.xml` through public
`TextElements::extract_text`, including exactly two source-freshness
observations per call; the candidate calls public `SourceBackedDocument::text()`
four times. The timer contains exactly those four full-text projections.
Independent untimed `InstrumentedSource` replays require the control's eight
freshness observations (`[2, 2, 2, 2]` per call), zero post-preparation
`ReadAt` reads, zero `Pictures/*` reads after preparation, and deterministic
content/media range counters. The candidate accepts either the pre-cache
`[2, 2, 2, 2]` shape or the cache-enabled `[2, 4, 2, 2]` publication-window
shape and records the observed shape and total. Exact semantic text, archive
topology, source XML hash, eight-media payload identity, and projection hashes
remain outside timing.
These additions bring the selectable matrix to 322 names while leaving the
default 36 cases / 198 records unchanged; they provide correctness and phase
evidence only and make no latency claim;
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
Two additional opt-in PPTX cross-presentation slide-copy selectors
(`pptx_cross_copy_plain` and `pptx_cross_copy_media_rich`) use deterministic
three-slide source/two-slide destination packages. They report plan, commit,
and OPC sequential publication timings separately, with reopen timing retained
as a non-publication diagnostic. Complete semantic/package/dependency-closure,
collision-remap, source-immutability, durable-patch, and stale/foreign/refusal
gates are untimed. The selectors brought the selectable matrix to 273 names at
that point while leaving the default 36 cases / 198 records unchanged. They make no
speedup, allocation, RSS, release-ABBA, or physical-I/O claim at the 0145
revision. Change 0158 now accepts the later owned-source additive-topology
publisher on these exact selectors at the bounded prepared-operation scope.

Change 0159 adds the independent opt-in
`pptx_source_backed_cross_copy_plain` selector over the exact same plain
three-slide-source/two-slide-destination bytes. It calls the public
`SourceBackedPresentationEditor::plan_cross_slide_copy` and
`publish_cross_slide_copy_to_stream` APIs; source-backed publication's
internal rerun/verification/topology work therefore remains part of the
publication phase. Only plan and public publication durations are reported;
setup, reopen, raw/topology/semantic gates, typed stale/foreign refusals, and
source-read checks are outside timing. The schema records separate source and
destination logical `ReadAt` call/byte counters, with no cache, allocation,
RSS, physical-I/O, eager/source speedup, ABBA, or media-rich claim. This
selector brought the selectable matrix to 302 names at change 0159 while leaving the
default 36 cases / 198 records unchanged.

Change 0146 adds 12 shared CFB MiniFAT `open_stream` evidence selectors:
36-byte and 4095-byte targets across `many-small` and `wide-root`, with
one-shot, repeat-3, and sequential repeat-8 operations plus matched
deterministic-delay simulator controls. They call `SharedOleFile::open_stream`
directly and record per-invocation output hashes, logical read events, source
version/refusal checks, root Mini Stream identity, and inferred direct/cache
counter evidence. Repeat-8 is bounded sequential repetition; it is not
`bulk_read`, concurrent traversal, or a full-workload claim. The selectors are
opt-in and bring the current selectable matrix to 285 names while leaving the
default 36 cases / 198 records unchanged. Change 0146 makes no performance
claim until identical-harness clean release ABBA evidence is available.

Change 0147 supplies that clean release ABBA. Under the exact 100 us fixed +
25 us/request, 50 MiB/s, 4 KiB-range simulator, all four named one-shot cells
improve total p50/p95/p99/mean by about 62-64% in both directions, while exact
one-shot source work falls from the complete root Mini Stream to one 36- or
4095-byte target range. Repeated many-small cells retain a smaller modeled
cost from the extra first direct request, so no generic repeat improvement is
claimed. Local wall-clock tails, allocation/RSS, physical-I/O, native-format,
and cross-format claims remain open.

Change 0148 extends the same shared-baseline harness with six production-only
selectors for different-SID A-B-A, public bulk A-B-A, and overlapping
same-target calls at the 36-byte and 4095-byte targets:
`cfb_open_stream_mini_shared_{different_sid,bulk,concurrent}` and
`cfb_open_stream_mini_4095_shared_{different_sid,bulk,concurrent}`. The
selectable matrix was 291 names at that revision; change 0154 later made it
301, and change 0159 now makes it 302.
The default remains 36 cases / 198 records. The runner accepts the control root-only vector, the prior
direct-then-root vector, and the target-aware same-SID repeat vector, while
recording ordered workload names, exact positional ranges, output hashes,
source-version stability, and typed missing-stream refusal. Concurrent
workers use only a harness-side overlap gate for deterministic entry; the
bulk selector exercises the public `bulk_read` API. This tranche is
correctness/evidence-only and makes no latency, allocation, RSS, physical-I/O,
release, or generic performance claim. Failure/retry, ineligible-root, FAT,
native semantic, resource, and performance acceptance for those extended
selectors remain open. See
[`0148-cfb-same-target-repeat-policy.md`](../../docs/performance/changes/0148-cfb-same-target-repeat-policy.md).

Change 0149 retains the clean release comparison for that target-aware policy:
four strict `A1 control, B1 candidate, B2 candidate, A2 control` CPU-2 legs,
20 warmups, 200 samples, and 36 records per leg (28,800 samples total). The
control and candidate use the same runner and deterministic corpora. Under the
exact 100 us fixed + 25 us/request, 50 MiB/s, 4 KiB-ceiling simulator, all
eight aggregate repeat-3/repeat-8 total cells improve by roughly 56-64% at
p50/p95/p99/mean in both adjacent directions. Same-target work changes from
`[D,C,0...]` to `[D,D,...]`; later calls remain target-sized reads rather than
zero-source cache hits. One-shot model controls are neutral. Local,
per-invocation, bulk, concurrent, allocation/RSS, physical-I/O, native-format,
and generic claims are withheld after explicit local tail/drift review
triggers. See the
[`0149` release record](../../docs/performance/changes/0149-cfb-same-target-repeat-release-abba.md)
and [summary](../../docs/performance/results/cfb-repeat-abba-0149-summary.json).

Change 0152 supplies the final clean release ABBA for same-target MiniFAT
single-flight, introduced by `c270c8f3b` and finalized in `f46381c6f`, against
clean control `e486e4b1` on CPU 2. The four legs used 20 warmups and 500 samples
across 24 records per leg (48,000 retained samples); all correctness/source-
event invariants passed. Existing concurrent scenarios recorded 6,473
candidate versus 8,000 control logical source calls, 19.09% fewer. This is
accepted only as source-event/correctness evidence. At the 0152 revision the
291-name selector matrix was unchanged; change 0153 adds four RTF selectors
measured at the pre-staged publication-call interval, making that matrix 295.
Change 0154 adds six ODT/ODS/ODP content-COW publication selectors, making the
matrix 301 at that revision; change 0159 made it 302, change 0160 made it 303,
change 0162 made it 305, change 0163 made it 309, and change 0164 makes it 311.
Change 0165 extends the existing native-DOC phase selector without adding a
case, so that revision remained at 311 names. Change 0166 adds four row-
visibility selectors, making the current matrix 315 names; the default remains
36 cases / 198 records.
No runtime selector was added to 0152; only `cfg(test)` source-event acceptance
and tests changed. Root
MiniStream cache and
resource-accounting boundaries and broader performance gaps remain. Local or
generic latency, allocation/RSS/peak memory, physical I/O/syscalls,
cold-cache/device/network, decompression, native semantic, OOXML, ODF, RTF,
and iWork claims are withheld. See the
[`0152` release record](../../docs/performance/changes/0152-cfb-same-target-singleflight-release-abba.md)
and [summary](../../docs/performance/results/cfb-singleflight-abba-0152-summary.json).

Change 0166 adds four opt-in XLSX existing-row visibility selectors, taking
the current selectable matrix from 311 to 315 names while leaving the default
36 cases / 198 records unchanged. The matched eager/source-backed controls
cover one-row hide from a visible source and exact-256-row unhide from a
source whose first 256 rows are hidden. Each selector runs on the same
single-sheet media-rich corpus in `medium` (512 × 16) and `large` (2,048 ×
32) shapes. Open, stage/plan, commit, sequential publication, and lifecycle
vectors are separate; source-backed records add generic logical `ReadAt`
counters and pre-publication cache diagnostics only. Measured publication is
bound by exact length/SHA-256 to an untimed semantically reopened expected
artifact, and raw untouched-member identity is checked separately. Exact
no-op, foreign/stale, signed/protected/formula/MCE/macro/relationship, and
partial/zero-sink refusal fields are source-backed-only and omitted from eager
records. This is correctness/phase evidence only and makes no
latency, speedup, allocation/RSS, physical-I/O, cold-cache, decompression, or
real-producer claim. See
[`0166`](../../docs/performance/changes/0166-xlsx-row-visibility-evidence.md).

Change 0167 changes production publication only and adds no selector. Matched
source-backed row-visibility publication reuses the existing cell-values
lineage/version proof, removing one semantic worksheet reload, cell parse and
row scan while retaining the mandatory OPC selected-member read. Clean CPU-2
A/B/B/A records observe descriptively lower publication vectors, but the 5%
same-implementation drift gate fails: maximum absolute drift is 34.80% for
control large/unhide publication p99 and 10.23% for candidate medium/hide
complete-workflow p50; first-pair medium hide/unhide complete-workflow p99
regresses 6.95%/2.69%. No acceptance-grade latency or resource claim is made.
See
[`0167`](../../docs/performance/changes/0167-xlsx-row-visibility-provenance-reuse.md).

Change 0168 changes production validation only and adds no selector. Native
XLS Number/RK/MulRK plan-only commit now runs BIFF owner validation on the exact
composed CFB view inside the existing pre/post fingerprint fence. This removes
two redundant post-plan complete source scans per effective edit while
retaining CFB reopen, selected-range checks, source preconditions, final
source/target fingerprints, security policy, semantic readback, and publication
fences. Clean CPU-2 A/B/B/A records observe descriptively lower complete-
workflow and semantic-commit p50/mean/p95/p99 values in both paired directions,
but the 5% same-implementation drift gate fails. No acceptance-grade latency,
tail, allocation/RSS, physical-I/O, or producer claim is made. See
[`0168`](../../docs/performance/changes/0168-xls-numeric-validation-fusion.md).

Change 0169 changes the shared hierarchical budget charge path and adds no
selector. The existing `xlsx_streaming_create` large shape exposed a transient
allocation in every cumulative row/cell charge. `Budget::consume` now walks the
immutable ancestor chain by reference, while releasable reservations retain up
to four charged nodes inline and spill for deeper caller-defined hierarchies.
Clean CPU-2 release A/B/B/A runs use 20 warmups and 200 samples per shape.
Medium and large p50/mean/p95/p99 improve in both paired directions by
1.05%-9.76%; tiny p50/mean/p95 also improve, while tiny p99 regresses
1.81%/2.75% and is withheld. Matched whole-process Heaptrack captures record
48.81% fewer allocation calls and 69.77% fewer temporary allocations with
unchanged 225.45M peak heap. RSS directions disagree. The result is limited to
warm in-memory, one-sheet inline-scalar forward-only XLSX creation; it is not a
total-memory, physical-I/O, cold-cache, producer, or broad `Budget` claim. See
[`0169`](../../docs/performance/changes/0169-xlsx-streaming-budget-charge.md).

Change 0170 changes only the XLSX streaming encoder and adds no selector.
Measured ordinary text is validated in source order but appended in contiguous
UTF-8 runs between XML entities; byte length avoids a redundant scalar-count
pass when it already proves Excel's character bound, and one row-number lexical
form is reused across that row. Clean CPU-2 release A/B/B/A runs use 20 warmups
and 300 samples per shape. Large p50/mean/p95/p99 improve in both paired
directions by 5.02%-6.99%; medium p50/mean/p95 improve by 4.45%-5.52%; tiny p50
improves by 5.03%/7.74%. Tiny mean/tails and medium p99 are withheld because
paired directions disagree. Exact archive/worksheet hashes, sink topology,
zero retained output, and the 4 KiB row window remain fixed. Whole-process
instructions/branches fall in matched counter runs, but branch misses regress;
no allocation, RSS, total-memory, physical-I/O, cold-cache, richer-authoring,
or producer claim is made. See
[`0170`](../../docs/performance/changes/0170-xlsx-streaming-escape-runs.md).

Change 0154 adds matched owned-rebuild and source-positional content-only
publication selectors for ODT, ODS, and ODP. Each selector prepares the real
semantic edit and owner outside timing, then measures one public publication
call plus the same fixed 16 KiB non-seek hashing sink. Exact content, semantic
reopen, package inventory, positional untouched-member raw identity, physical
and central order, no-op, limits, cancellation, source immutability, and
logical `ReadAt` counters remain untimed gates. A clean CPU-2 release ABBA with
20 warmups and 100 samples per record accepts p50 improvements of
96.35%-96.63% in both pair directions for this prepared in-memory publication
boundary. It makes no end-to-end, allocation/RSS, physical-I/O, decompression,
cold-cache, filesystem, real-producer, or iWork claim. See the
[`0154` record](../../docs/performance/changes/0154-odf-content-cow-publication-evidence.md)
and [summary](../../docs/performance/results/odf-content-cow-abba-0154-summary.json).

Change 0158 adds no selector. It compares clean release control `e8a67b19e`
and candidate `d900ae633` on the existing plain and media-rich PPTX cross-copy
cases using 20 warmups, 200 samples, and strict CPU-2 A/B/B/A order. Total p50
improves 29.643%/26.196% and 43.294%/43.604%; media-rich publication p50
improves 49.321%/49.680%. Plain publication tails are withheld after declared
same-implementation drift triggers. The accepted boundary is canonical
generated owned-source prepared slide copy; source-backed/cold-I/O,
decompression, real-producer, generic OPC/PPTX, and iWork claims remain open.
See the
[`0158` record](../../docs/performance/changes/0158-pptx-additive-topology-release-abba.md)
and [summary](../../docs/performance/results/pptx-additive-topology-abba-0158-summary.json).

Change 0159 adds one source-backed plain PPTX cross-copy selector. It is
correctness/counter evidence only over the matched existing plain corpus:
`plan_ns` and `publication_ns` are separated, setup/reopen/gates are untimed,
and source/destination logical `ReadAt` call/byte counters are reported
separately. The public source-backed publication includes its preparation
rerun and topology publication work; no eager/source speedup, cache,
allocation/RSS, physical-I/O, media-rich, real-producer, or release-ABBA claim
is made. See the
[`0159` record](../../docs/performance/changes/0159-pptx-source-backed-cross-copy-evidence.md).

The validation/section and scalar-cell selectors are opt-in and do not alter the default
36 cases / 198 records:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --json target/perf/container-baseline.json \
  --corpus-manifest target/perf/container-baseline.corpus-v2.json
```

`--corpus-manifest` writes the additive schema-2 deterministic corpus catalog
and places a reference under `corpus_catalog` in the schema-1 report.  It does
not change the existing case/corpus identity keys or their comparator digest.
See [`docs/performance/CORPUS_MANIFEST_V2.md`](../../docs/performance/CORPUS_MANIFEST_V2.md).

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

[Change 0150](../../docs/performance/changes/0150-xlsx-managed-cell-values-budget-evidence.md)'s
managed tranche has no controlled release ABBA comparison and therefore makes
no speedup or throughput claim. Its Budget covers only retained
and in-flight OPC `PartData` payload reservations plus the managed publisher's
accepted `OutputBytes`; parsed stores, metadata, staging, rewritten
candidates, and output buffers are outside that accounting. The current
tranche also performs one untimed one-byte-under first-publication-request
replay per managed selector: the typed `OutputBytes` refusal accepts zero
output and preserves the source version. Declared `Work` remains separate
from decompressed/read bytes. The tranche does not claim allocations, RSS/
peak memory, hardware/CPU pinning, cold I/O, decompression, or real-producer
breadth.

Measure the matched one-cell XLSX clear/remove lifecycle controls on the same
media-rich four-sheet corpora:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case xlsx_eager_cell_clear_edit_save,\
xlsx_source_backed_cell_clear_edit_save,\
xlsx_eager_cell_remove_edit_save,\
xlsx_source_backed_cell_remove_edit_save \
  --xlsx-cell-crud-shape medium,dense-sparse \
  --json target/perf/xlsx-cell-lifecycle-crud.json
```

Each selector targets one existing numeric owner from the deterministic
medium or dense/sparse inventory. `clear` retains an empty owner while
`remove` deletes that owner; neither selector claims formula, rich metadata,
missing-cell, or third-party-producer parity. Eager controls use the public
`WorksheetEdit` API; source-backed controls use the positional value editor.
Open, selector planning/staging, commit, and publication are reported as
separate phase vectors. Timed publication uses a fixed-memory SHA-256 sink
with zero retained output. Complete semantic/reopen, package/raw-member,
exact-hash, and sink gates run outside the timed samples. The shared
source-backed lifecycle preflight supplies exact no-op, clear/remove owner,
volatile-patch, and stale-source gates; eager controls retain their own
semantic/package checks. Source-backed results report generic positional
`ReadAt` calls/bytes and successful cached OPC payload materializations; the
source-backed clear/remove pair is correctness/counter evidence only, with no
durable-patch, allocation, RSS, physical-I/O, cold-cache, decompression, or
speedup claim.

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

Measure the controlled filesystem tranche (six opt-in cases):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case opc_file_eager_open,opc_file_source_open,\
opc_file_eager_one_part_atomic_save,opc_file_source_one_part_atomic_save,\
cfb_file_same_length_overlay_atomic_save,\
cfb_file_owned_same_length_overlay_atomic_save \
  --json target/perf/filesystem-crud.json
```

Use `--filesystem-cache warm` for a warm-only smoke, or
`--filesystem-cache warm,cold-requested` (the default) for both keyed states.
Use `--filesystem-root PATH` to place the source, destination, and sibling
temporary files under a caller-selected filesystem; the report records only
that a root was selected, not the path itself.

`cold-requested` remains the existing advisory state.  It records an accepted
Linux `posix_fadvise(DONTNEED)` request and does not imply that the source was
evicted.  The opt-in `--filesystem-cache cold-verified` state is stricter and
is admitted only on 64-bit Linux; 32-bit Linux is explicitly reported as
`ineligible_linux_non64_bit`.  It runs only source-touching
open/lifecycle/save controls (the prepared PPTX/DOCX query controls are
explicitly ineligible), and leaves the default cache selection unchanged.

Cold-verified samples are admitted only when all of the following hold:

- the source is a regular, non-empty, page-aligned file opened read-write on an
  allowlisted block-backed filesystem identified from the opened FD's numeric
  `statfs` magic (`0xef53` ext2/3/4, `0x58465342` XFS, `0x9123683e` Btrfs,
  `0xf2f52010` F2FS, or `0x2fc12fc1` ZFS);
- the source has been `fsync`ed and accepted `posix_fadvise(DONTNEED)` advice;
- the canonical, hashed, versioned external `fincore` binary emits one strict
  JSON record immediately before the timed operation with zero resident,
  dirty, and writeback bytes; only its basename, executable hash/version, and
  stderr digest/length plus method/fallback evidence are retained, and any
  unrecognized fallback is ineligible; and
- the child’s `/proc/self/io` `read_bytes` delta is positive during the
  source-touching interval.

Use, for example:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 --filesystem-cache cold-verified \
  --filesystem-root /path/on/ext4 \
  --case opc_file_source_open --json target/perf/filesystem-cold-verified.json
```

The verifier creates a private page-aligned source copy (ZIP alignment uses
the EOCD comment field and does not change logical package members).  It does
not add source paths, absolute tool paths, stderr contents, or device
identifiers to the report.  Each eligible proof records the aligned source
SHA-256 and size, numeric filesystem magic, and canonical fincore basename,
SHA-256, and version.  `filesystem_evidence` records an explicit
`cold_verified_status` for every requested case and keeps ineligible cases out
of timed results; a status such as `ineligible_filesystem_unsupported`,
`ineligible_fincore_invalid_json`, or `ineligible_read_bytes_zero` is a
successful fail-closed outcome, not a cold result.  A verified result proves
page-cache residency/dirty/writeback state plus process `read_bytes` only.  It
does not prove physical-media temperature, device-cache state, or durable
storage latency.

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
records its changed-span and published-byte report fields. The owned CFB
selector reads the source once before the nested CFB phase timers, seals it as
`Arc<[u8]>`, and calls `SharedOleFile::open_owned`; its total operation time
includes that filesystem ingress, while its provenance, source-byte hash, and
separate open/plan/atomic-publication timings are recorded under `cfb_owned`.
Its generic logical `ReadAt` vectors are explicitly not applicable rather than
invented for immutable slice ownership. OPC materialization counts are
recorded when exposed by the public API. After every prime and
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

Measure the opt-in PPTX cross-presentation slide-copy evidence selectors:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case pptx_cross_copy_plain,pptx_cross_copy_media_rich \
  --json target/perf/pptx-cross-slide-copy.json
```

The plain selector copies source slide 3 into destination slide 2 at zero-based
position 1 without media; the media-rich selector adds eight deterministic 2
MiB PNG resources to the source and destination packages so the copied
dependency closure exercises deterministic collision-avoidance remapping. Both selectors use the public
`Snapshot::plan_cross_slide_copy`, `Package::apply_cross_slide_copy_plan`, and
`OpcPackage::to_stream` APIs. The timer reports plan, commit, and sequential
publication phases separately; setup, reopen, semantic/package/closure,
source-immutability, collision, durable-patch, stale/foreign/refusal, and
output-hash checks remain untimed. This is correctness/evidence only and makes
no performance claim.

Measure the independent source-backed plain cross-copy evidence selector:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case pptx_source_backed_cross_copy_plain \
  --json target/perf/pptx-source-backed-cross-slide-copy.json
```

This selector uses the exact plain corpus bytes from `pptx_cross_copy_plain`
and copies source slide 3 into destination slide 2 at zero-based position 1.
The public source-backed publisher reruns preparation and topology checks
before writing the sequential output, so the reported phases are only
`plan_ns` and `publication_ns`; there is no synthetic commit phase. Reopen,
raw ZIP member/order/comment preservation, semantic/dependency gates, typed
source/destination revision refusal, and foreign-editor refusal are untimed.
The output is correctness/counter evidence only and does not compare eager
and source-backed speed or claim cache, physical-I/O, allocation, RSS,
media-rich, or real-producer behavior.

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

Measure the four opt-in XLSX edit-composition controls:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_join_disjoint_commit_save,xlsx_join_conflict_plan,\
xlsx_three_way_disjoint_commit_save,xlsx_three_way_conflict_resolve_save \
  --xlsx-cell-crud-shape medium,dense-sparse \
  --json target/perf/xlsx-edit-composition.json
```

These selectors reuse the deterministic media-rich four-sheet scalar corpus.
Both branches are prepared from one immutable workbook lineage outside timing.
The disjoint join case times `join`, `commit`, and 64 KiB write-call-bounded
sequential `write_to` into an output-retaining sink; the join-conflict case
times only the typed overlap refusal. The disjoint three-way case times plan,
finish, commit, and that publication, while the conflict case also times
explicit `MergeChoice::Left` resolution. Exact no-op, empty-join identity, and
`MergeChoice::Neither` gates are outside every timed interval. Reopen, exact
output, package/media identity, durable JSON determinism, forward/inverse
replay, and stale/foreign refusal additionally gate the three save-bearing
selectors. These are correctness and phase evidence only; no latency or
speedup claim is made.

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

Run the complete tiny semantic DOCX/PPTX smoke matrix (17 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case docx_semantic_open,docx_semantic_list_paragraphs,docx_semantic_one_paragraph,docx_semantic_one_paragraph_text,docx_semantic_full_text,docx_semantic_create_small,docx_semantic_noop_edit_save,docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,pptx_semantic_open,pptx_semantic_list_slides,pptx_semantic_one_slide,pptx_semantic_full_text,pptx_semantic_create_small,pptx_semantic_noop_edit_save,pptx_semantic_one_edit_save,pptx_semantic_one_percent_edit_save \
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

Run the matched owned/source-backed existing-cell ODS controls over that same
two-sheet, 2,048-cell, eight-resource media-rich corpus:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 \
  --case ods_source_eager_one_edit_save,ods_source_backed_one_edit_save,\
ods_source_eager_one_percent_edit_save,ods_source_backed_one_percent_edit_save \
  --json target/perf/ods-source-cell-crud.json
```

The one-cell selector edits the existing middle cell used by the fixed ODS
media case; the one-percent selector uses the deterministic evenly spaced
`semantic_update_indices` closure (21 existing cells). The eager controls use
the owned ODS transaction and write the complete committed snapshot through
the fixed 16 KiB non-seek hashing sink. The source-backed controls use
`SourceBackedSpreadsheet::from_read_at`, `edit_cells`, `set_cell`/`set_cells`,
`commit`, and `write_to` without `materialize()`. Open, staging, commit/plan,
and sequential publication are timed. Reopen, full semantic digest, media
payload hashes, and source/output hashes are untimed. Source-backed records
add aligned lifecycle/open/stage/commit/publication vectors and a separate
untimed `InstrumentedSource` replay. Source-only raw-untouched-member,
patch/inverse, exact-noop, foreign-source, replacement-limit, partial-sink and
immutability gates are nullable and omitted from eager records. Stale/version,
cancellation, signed/protected, formula/unknown/repeated-row, and transaction-
bound contracts remain production-test evidence outside this selector. The
fixed sink retains zero output bytes. The selector makes no physical-I/O,
decompression, allocation/RSS, cold-cache, producer, or broad CRUD claim;
accepted release latency is scoped by the linked change record.

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

The matched RTF tail publication-call controls are separate opt-in selectors:
`rtf_logical_tail_commit_append`, `rtf_logical_tail_plan_append`,
`rtf_logical_tail_commit_noop_save`, and `rtf_logical_tail_plan_noop_save`.
They use the same tiny/medium/large plain uncompressed lifecycle corpus and
the same changed-append/exact-no-op inputs. Commit and publication-plan
construction is pre-staged before the timed interval; `elapsed_ns` is exactly
the pre-staged publication-call interval around `TailAppendCommit::write_to`
or `TailAppendPublicationPlan::write_to` to the fixed 16 KiB non-seek sink.
These calls intentionally perform different validation, digest/source-version,
budget, window, and final-verification work; this is not a symmetric-work
comparison. The
`source.rtf_tail_publication` object keeps planning, publication, reopen, and
lifecycle vectors separate and explicitly reports source-retained,
complete-candidate-retained, and publication-window bytes. Exact output bytes,
digest, semantic paragraph projection, no-op identity, durable patch
apply/inverse/stale/foreign checks, cancellation, sink failure/partial
progress, limits, and source-version gates are untimed correctness checks.
This tranche makes no end-to-end, rich-format, allocation/RSS, physical-I/O,
or ABBA latency claim.

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --semantic-shape tiny,medium,large \
  --case rtf_logical_tail_commit_append,rtf_logical_tail_plan_append,rtf_logical_tail_commit_noop_save,rtf_logical_tail_plan_noop_save \
  --json target/perf/rtf-tail-publication-plan.json
```

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

The row-visibility selectors use a separate one-sheet corpus so worksheet
selection and row-state work remain explicit:

| Row-visibility shape | Rows × columns | Logical cells | Fixed media | Batch bound |
|---|---:|---:|---:|---:|
| `medium` | 512 × 16 | 8,192 | 8 × 512 KiB | 256 rows |
| `large` | 2,048 × 32 | 65,536 | 8 × 512 KiB | 256 rows |

The corpus contains deterministic numeric cells and untouched media members;
the row-visibility evidence does not reuse the multi-sheet scalar-cell CRUD
shape or make a claim about broad worksheet/row structural editing.

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

RTF streaming results also publish aligned, per-retained-sample
`operation_metrics`. Process counters are best-effort same-process procfs
deltas whose scope includes the after-snapshot probe overhead; unsupported
platforms report them as unavailable. The allocator target reports checked
operation regions, while the ordinary binary leaves allocation metrics absent.
Allocator live/high-water values are absolute before/after process counters,
and `peak_rss_bytes` is the process-lifetime high-water mark, not an
operation-local peak. Resource probes, digest finalization, and correctness
checks remain outside the elapsed-time interval.

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

The XLSX case now has a clean release A/B/B/A result in
[change 0169](../../docs/performance/changes/0169-xlsx-streaming-budget-charge.md).
After removing transient hierarchical-budget charge allocations, medium and
large p50/mean/p95/p99 improve in both paired directions by 1.05%-9.76%; tiny
p50/mean/p95 also improve, while its p99 regresses 1.81%/2.75% and is withheld.
Matched whole-process Heaptrack captures record 38,672,384 -> 19,794,608
allocation calls and 22,545,902 -> 6,815,902 temporary allocations, with
unchanged 225.45M peak heap. The exact archive/worksheet hashes, 4 KiB retained
row window, and zero retained output remain fixed. The allocation counts cover
the whole benchmark process, not only the timed writer; RSS directions disagree
and no total-memory, physical-I/O, cold-cache, multi-sheet, shared-string/style/
formula/date, real-producer, or broad `Budget` claim is made.

[Change 0170](../../docs/performance/changes/0170-xlsx-streaming-escape-runs.md)
keeps the same selector and exact corpus while batching ordinary text between
XML entity boundaries. Large p50/mean/p95/p99, medium p50/mean/p95, and tiny
p50 improve in both clean paired directions; tiny mean/tails and medium p99 are
withheld. The output hashes, logical sink counters, zero retained output, and
4 KiB row window remain unchanged. Matched process-wide counters record
6.15%-6.19% fewer instructions and 10.54%-10.57% fewer branches, while branch
misses regress 8.99%-14.37%. These counters are descriptive whole-process
evidence, not operation-local CPU, allocation, memory, or I/O attribution.

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
PPTX semantic save controls use `Package::from_bytes`/`from_vec`, presentation
slide/text views, opened-presentation transactions, and `to_bytes`; those
legacy controls intentionally leave `sink` as `null`. The cross-presentation
slide-copy evidence selectors use the public `OpcPackage::to_stream` writer
through a bounded sequential sink and report its counters explicitly, without
extending the claim to the other PPTX save paths.

## Opt-in RTF semantic corpus matrix

The RTF cases exercise only the ordinary native `litchi_rtf::Document` facade:
owned-byte open, lazy paragraph enumeration, one middle paragraph, first
complete-text materialization, bounded semantic-text output to a forward-only
sink, exact source streaming, exact empty-edit publication, and
capability-bounded one-paragraph and `ceil(1%)` paragraph edit/save. The two
historical `rtf_logical_tail_*` cases and four matched Commit/PublicationPlan
controls are a separate existing-document append tranche; they do not reuse
the streaming-creation path and are restricted to the matched plain lifecycle
corpus.
The two lifecycle cases use a matched default-formatted plain corpus because
the read/edit corpus's explicit font formatting is outside their changed
publication closure.
The split/merge cases use that same literal-ASCII ordinary-body corpus across
the 24/200/10,000 paragraph shapes. Split targets paragraph `count/2` at its
interior ASCII midpoint; merge targets that paragraph and its immediate
successor. Independent raw splices establish the exact `+5`/`-5` byte output
delta and unchanged surrounding bytes before timing. Each selector reports
separate open, stage, commit, publication, and lifecycle vectors and uses a
16 KiB windowed non-seek sink that retains zero output bytes. Untimed gates
cover semantic reopen, volatile and deterministic durable forward/inverse
replay, forged result-artifact, stale/foreign refusal, strict limits and
unsupported/protected refusal, partial/zero sinks, and source/output hashes.
These selectors provide correctness and phase evidence only; they make no
performance or transaction-memory claim.
`--rtf-variant` defaults to `plain`.

| Variant | Source | Shapes | Supported cases |
|---|---|---|---|
| `plain` | Deterministic direct ASCII RTF | tiny, medium, large | All 21, including six logical-tail and two ordinary-paragraph split/merge selectors |
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

The standalone-picture CRUD selectors `rtf_picture_payload_batch_replace` and
`rtf_picture_batch_remove` use a deterministic uncompressed ASCII RTF corpus
with root-level standalone PNG/JPEG groups. Tiny, medium, and large shapes
contain 2, 8, and 64 alternating 16-byte decoded payloads; replacement leaves
one picture unselected and both batches stay within the public 64-operation
ceiling. The corpus uses mixed-case hexadecimal digits with deterministic
spaces/newlines, so the independent expected-output splice replaces only
digit slots while preserving every source whitespace byte and each slot's
case. Each selector reports separate open, stage, commit, publication, and
lifecycle vectors; publication is `commit.snapshot().write_to` into a bounded
hashing sink. Untimed gates cover semantic reopen, raw unselected preservation,
exact no-op identity, volatile patch apply/inverse, deterministic durable JSON
serialization plus forward/inverse/stale/foreign checks, refusal cases,
partial/zero sinks, and source/output hashes. These selectors provide
correctness and phase evidence only; they make no performance claim.

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

The matched ODT repeated-text selectors use a separate large media-rich
corpus. Run both phases with:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case odt_source_backed_repeated_text_uncached,odt_source_backed_repeated_text_cached \
  --json target/perf/odt-repeated-text.json
```

The corpus contains 10,000 deterministic paragraphs and eight incompressible
2 MiB `Pictures/*` members. Owner construction and output-slot reservation are
outside the timed interval; exactly four full-text projections are timed per
sample. The uncached control obtains retained `content.xml` with
`SourceBackedDocument::content_xml()` and calls public
`TextElements::extract_text`, preserving the current two-freshness-check
semantics per call. The candidate invokes public `SourceBackedDocument::text()`
four times, so the same harness remains valid when the production text cache is
enabled. Untimed replays verify exact semantic text, archive topology, media
payload hashes, content/media compressed-range classification, the control's
eight freshness observations (`[2, 2, 2, 2]`), and zero source reads after
owner preparation. The candidate records either that pre-cache shape or the
cache-enabled `[2, 4, 2, 2]` shape (ten total observations). This is
phase/correctness evidence only; no latency claim is made.

Four unified-root ODT filesystem selectors (`odt_file_eager_open`,
`odt_file_source_open`, `odt_file_eager_open_full_text_lifecycle`, and
`odt_file_source_open_full_text_lifecycle`) are also opt-in over this same
large 10,000-paragraph/eight-2 MiB-picture corpus. Eager timing includes the
filesystem `fs::read` and `litchi::Document::from_bytes`; source timing calls
`litchi::Document::open(path)`. The two lifecycle selectors additionally keep
`Document::text()` inside the timer. Corpus creation and temporary-file setup,
eager/source paragraph, text, table, element, and metadata parity, archive
member/hash identity, and compressed-range/media payload identity remain
outside timing. Each source selector has an independent untimed
`InstrumentedSource` replay proving that preparation and the full-text query
overlap zero `Pictures/*` payload bytes. These selectors add filesystem/root
correctness and source-range evidence directly. The separately frozen
[change 0191](../../docs/performance/changes/0191-odt-unified-source-ingress.md)
A1/B1/B2/A2 run retains the correctness and logical-range evidence but accepts
no open-only latency statistic because same-implementation drift fails every
tier. The full-text lifecycle accepts p50/mean/p95/p99 reductions of 30.02% to
35.36% after both paired directions and implementation drifts pass. It makes no
allocation, RSS, physical-I/O, decompression, cold-cache, producer, edit/save,
or broad ODF claim.

Run the four ingress phases with:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case odt_file_eager_open,odt_file_source_open,\
odt_file_eager_open_full_text_lifecycle,odt_file_source_open_full_text_lifecycle \
  --json target/perf/odt-root-filesystem.json
```

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

Measure native DOC owner/public phase attribution explicitly (it is not in
the default matrix):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 --writer-shape tiny,large,payload-heavy \
  --case doc_owner_public_phases \
  --json target/perf/doc-owner-public-phases.json
```

The three selected writer shapes reuse the exact deterministic DOC bytes from
the legacy writer matrix. The result stores one vector entry per retained
sample for each named phase and for checked attributed/unattributed totals;
`elapsed_ns.samples` is the sorted view of the same measured lifecycle totals.
The complete semantic, patch, refusal, hash, and preservation gates run before
the sample loop or after each sample's timer stops. Successful event-order and
cardinality validation runs after each named outer interval but before the
complete lifecycle timer stops, so its recorder work is visible in checked
unattributed time. Separate format tests bind balanced error events. Observer
callbacks are synchronous; their dispatch and recorder work are present in
outer/lifecycle measurements even though individual phase intervals stop
before final recorder bookkeeping. `Snapshot::finish` is timed as output
materialization after commit; it is not a file save/publication operation.
Change 0160 adds this one opt-in selector, taking the current matrix from 302
to 303 names without changing the default 36 cases / 198 records.

Change 0165 extends that existing selector with a private native-DOC lazy/fused
fingerprint proof and a bounded descriptive comparison. It does not add a selector or change the
historical `measured_total_ns` lifecycle boundary. Per-snapshot diagnostic
fingerprints are lazy; same-allocation patch replay is measured as a separate
post-lifecycle workflow extension, and the first source/target fingerprint
demand is measured separately as deferred work. The extension also records
independently computed expected source/target FNV-1a fingerprints and four
additional gates. The final post-rebase comparison is clean at control
`d6818e290` and candidate `5dd813b1e`: CPU-2 A1/B1/B2/A2 uses 20 warmups and
500 retained samples per shape (6,000 primary samples) plus 24,000 guard
samples. Lifecycle p50 positive-faster deltas are +33.77%/+33.21% tiny, +12.28%/+13.81%
large, and +17.33%/+17.82% payload-heavy; immediate fingerprint-demand
workflow p50 positive-faster deltas are +14.56%/+13.89%, +4.50%/+5.83%, and
+6.55%/+7.08%. Final DOC guards are noop +78.84%/+79.89% tiny and
+71.08%/+70.40% large, one-edit +37.23%/+40.81% and +20.45%/+19.79%, while
DOC open is -3.52%/+0.13% and +0.55%/-1.80%. Neighboring XLS one-edit/open
guards are mostly neutral or improved; its nanosecond noop remains noisy.
The three-sample, preflight-inclusive whole-process Heaptrack probe records
50,677 allocation calls and a 128.28 MiB peak heap on both sides; it is not
operation-scoped, and RSS is descriptive only. The former `const` fingerprint accessors are
a capability change, not a deprecation. No physical-I/O, cold-cache,
real-producer, or generic total-memory result is claimed. See
[`0165`](../../docs/performance/changes/0165-doc-lazy-fingerprint.md), the
[summary](../../docs/performance/results/doc-lazy-fingerprint-0165-summary.json),
and the [release manifest](../../docs/performance/results/doc-lazy-fingerprint-0165-manifest.json).

## Cases

- `zip_index`: parse the ZIP central directory and build Soapberry's index;
  each iteration verifies the member count against the deterministic corpus
  manifest and retains measured `source.zip_index.observed_member_counts`
  (warm-ups excluded) beside its `expected_member_count`.
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
- `xlsx_join_disjoint_commit_save` and `xlsx_join_conflict_plan`: prepare two
  one-cell branches from the same immutable media-rich cell CRUD workbook before
  timing. The disjoint case times typed join, commit, and 64 KiB
  write-call-bounded sequential `write_to` into an output-retaining sink; the
  overlap case times only the recoverable conflict refusal and retains the
  rejected branch.
- `xlsx_three_way_disjoint_commit_save` and
  `xlsx_three_way_conflict_resolve_save`: prepare the same one-cell branches
  outside timing, then time non-applying three-way planning, finish, commit, and
  the same publication; the conflict case additionally times explicit
  left-branch resolution. Exact no-op, empty-join identity, and `Neither`
  resolution gate all four cases outside timing; durable replay/inverse,
  stale/foreign refusal, reopen, and media preservation additionally gate the
  save-bearing cases. All four selectors are correctness/phase evidence only
  and make no latency claim.
- `xlsx_{eager,source_backed}_cell_{clear,remove}_edit_save`: on the fixed
  media-rich four-sheet scalar corpus, clear or remove one existing numeric
  owner. Eager controls use `WorksheetEdit`; source-backed controls use the
  positional value editor. The timed interval records open, selector
  planning/staging, commit, and publication separately, with a fixed-memory
  zero-retained SHA-256 sink. Reopen, semantic owner/value state, package and
  source raw-member preservation, exact hashes, and sink gates are preflight or
  postflight checks. The source-backed lifecycle preflight additionally checks
  exact no-op, volatile patch, and stale-source behavior. These are
  correctness/counter selectors only and do not claim durable source patches,
  formula/metadata/third-party-producer parity, allocation/RSS, physical I/O,
  cold-cache behavior, decompression, or speedup.
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
- `doc_owner_public_phases`: opt-in native DOC owner/public-reader phase
  attribution over the exact `doc_fresh_write_to` bytes. The case runs the
  deterministic `tiny`, `large`, and `payload-heavy` writer shapes when
  selected, and reports per-sample strict-owner/public-reader/source-retain
  open phases, `Edit::new`, replacement staging, in-memory `Edit::finish`,
  final strict-owner/public-reader/source-retain/patch phases, and separate
  `Snapshot::finish` output materialization. The production
  `performance-diagnostics` feature emits only ordered content-free events;
  the harness owns `Instant` and uses a bounded preallocated recorder. Event
  order/cardinality is validated after each named outer interval but inside
  the lifecycle timer and checked unattributed remainder. Semantic
  source/output reopen, exact no-op/changed state, forward/inverse/stale patch
  behavior, typed refusal, malformed input, output hashes, and untouched CFB
  stream preservation are untimed gates.
  `Finish` here means in-memory owner rendering, not file publication or save.
  Attribution is per-sample checked arithmetic with explicit unattributed time;
  synchronous observer overhead and non-additive boundary work remain visible.
  This is correctness/attribution evidence only, with no speedup, physical-I/O,
  allocation, RSS, cold-cache, or real-producer claim.
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
- `xlsx_file_open`: open the deterministic medium cell-CRUD XLSX corpus through
  the high-level `litchi::Workbook::open(Path)` filesystem route. The timed
  scope is exactly that public path call; the untimed oracle compares metadata,
  worksheet names/count, full text, and the source archive SHA-256 with
  `Workbook::from_bytes(Vec<u8>)`.
- `xlsx_file_open_lifecycle`: use the same temporary source file and oracle,
  but time the high-level path open followed by worksheet names, worksheet
  count, and full text. The selector is opt-in and does not claim a source
  speedup or physical-I/O result.
- `xlsx_bytes_open`: move a prepared owned XLSX allocation into
  `litchi::Workbook::from_bytes(Vec<u8>)`; the input clone and typed eager
  semantic/hash guards are outside timing, so the measured scope is exactly
  the facade bytes construction.
- `xlsx_bytes_open_lifecycle`: use the same owned bytes path and guards, but
  time facade construction followed by worksheet names, worksheet count, and
  full text. These selectors are opt-in and do not claim a source-backed or
  physical-I/O result.
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
  through `document().paragraph(index)` with scanner-only timing, or extract
  complete document text.
- `docx_semantic_one_paragraph_text`: locate the middle paragraph before timing
  and time only its `Paragraph::text()` call; exact text/error checks and the
  complete semantic verification remain outside the timed interval.
- `omml_formula_range_scan`: scan the fixed OMML fragment corpus through
  `litchi_ooxml_common::xml::scan_omml_formula_ranges`; exact byte ranges are
  checked against an untimed `NsReader` oracle.
- `omml_formula_extract`: extract the same corpus through the public
  `litchi_ooxml_common::xml::extract_omml_formulas`; exact formula XML strings
  are independently reconstructed from the oracle's byte ranges and hashed
  outside the timed interval.
- `pptx_drawingml_extract`: extract text from the fixed DrawingML fragment
  corpus; exact text is checked against an untimed `NsReader` oracle.
- `pptx_drawingml_range_scan`: scan `p` element ranges from the same fixed
  DrawingML corpus; exact ranges are checked against an untimed `NsReader`
  oracle.
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
- `rtf_picture_payload_batch_replace`: replace a bounded batch of same-length
  standalone PNG/JPEG payloads while preserving the source's mixed-case hex
  and whitespace transport; one trailing picture remains unselected for the
  raw-preservation gate.
- `rtf_picture_batch_remove`: remove a bounded batch of alternating standalone
  PNG/JPEG groups from the same deterministic root-level corpus. Both picture
  selectors report open/stage/commit/publication/lifecycle phase vectors and
  keep correctness, patch, refusal, sink, and hash gates outside timing.

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
the mean. `elapsed_ns.sample_order` is the retained-iteration index for each
sorted sample, so additive per-sample evidence can be reordered without
guessing from duplicate elapsed values. Untimed warm-up iterations are
recorded in the configuration but are never included in these statistics.

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

The twenty-seven positional cases add a `source` object; older cases omit it. Its
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
(`warm`, `cold-requested`, or `cold-verified`), and pairs child elapsed time
with parent-observed wall time. It includes logical `ReadAt`
requested/returned bytes, fixed request size buckets, largest
requested/returned range sizes, a range-order pattern, maximum concurrent
reads, procfs I/O/fault/RSS deltas (with post-sample `VmHWM`), output SHA-256 and
byte length for saves, and public API publication counters where available.
`opc_materialized_parts` is explicit
zero for the raw-copy source overlay path; CFB records changed spans and
published bytes. The configuration records selected cache states, process
isolation, fresh-child sampling, and whether a caller-selected filesystem root
was used. PPTX ordinary-root samples additionally record
`logical_read_counter_scope` and, for source candidates only, an untimed
`pptx_source_replay` object with source hash, completed return sizes, exact slide
and media payload-range overlap, selected/unselected slide counters, union
coverage bytes, full-range coverage counts, semantic hash, and classification.
Its `read_return_sizes` vector records bytes returned by each logical replay
read, not requested buffer lengths.
Selected replay classification requires the complete target compressed range;
list replay classification requires every slide range. Eager PPTX samples set
the replay object to null and mark the generic counter scope
`not_applicable_eager_pptx`; their zero generic fields must not be interpreted
as source-read measurements.
Eager OPC filesystem samples use the same explicit boundary,
`not_applicable_eager_opc`, because their timed `fs::read` path has no
`ReadAt` counter; source-backed OPC and CFB overlay samples retain
`timed_read_at` when their positional counter is active.

The positional wrapper's `sequential`/`random` label is categorical evidence
about completed logical ranges only: `sequential` requires at least two full,
contiguous observations in completion order, while a non-contiguous transition
is `random`. Empty, short, concurrent, or insufficient observations are
`unknown`; invalid range arithmetic fails closed. The label says nothing about
kernel readahead, page-cache behavior, device/network requests, or physical I/O.
Largest range sizes and requested/returned byte totals have the same
logical-source boundary. An attempted `ReadAt` increments calls and requested
bytes before delegation; an underlying error leaves returned bytes unchanged,
and a source that returns more than the requested buffer is rejected before
returned-byte accounting. Counter overflow, invalid range arithmetic, and
poisoned metric state fail the sample rather than emitting fabricated values.
The generic operation envelope publishes returned-byte totals and the largest
returned range only; it does not fabricate a per-call returned-size
distribution from the requested-size vector. The untimed PPTX/DOCX replay
vectors are the exception: their `*_return_sizes` fields are exact completed
return lengths from the replay source.
The timed generic `ReadAt` and atomic-save callbacks expose no exact compressed
member, decompressed-byte, or recompressed-byte boundary, so the corresponding
`operation_metrics.source` vectors are explicitly `unavailable` (or
`not_applicable` for an uninstrumented selector); raw source/output lengths are
not substituted.

Each filesystem `CaseResult` additionally carries an additive
`operation_metrics` envelope. Its `sample_count` and `alignment` identify the
sorted `elapsed_ns.samples` vector; `sample_indices` records the stable child
identity used to order ties. Every measured numeric vector has that same
cardinality and order. `latency_claim` is explicitly
`evidence_only_filesystem_selector`, so these elapsed vectors are not a
comparator or eager/source latency claim. The envelope separates logical source-read
vectors (including largest range sizes); the aligned categorical
`logical_read_pattern` vector is descriptive and is not a numeric policy
metric. It also separates procfs process vectors, post-operation output length,
publication counters, and OPC materialization counts. A vector with a measured
zero is
serialized as a numeric zero; unsupported or unavailable values omit the
numeric vector and retain an explicit `status` (`not_applicable` or
`unavailable`). The `/proc/self/io` vectors (`rchar`, `wchar`, `read_bytes`,
`write_bytes`, `cancelled_write_bytes`, `syscr`, and `syscw`) use the explicit
scope `child_process_interval_delta_including_procfs_probe_overhead`. The
after-snapshot procfs read can itself add `rchar` and `syscr`; these vectors
therefore provide no operation-only or physical-storage attribution. Other
procfs vectors include faults, context switches, and RSS. `peak_rss_bytes` is
the process-lifetime after-sample `VmHWM`, not an operation peak, and
`rss_delta_bytes` is not a peak. The envelope makes no allocation,
copied-byte, decompressed-byte, or recompressed-byte claim because the
filesystem child does not instrument those quantities. The existing raw
`filesystem_evidence` fields retain their prior semantics; the new counters are
additive.

The default `litchi-perf-baseline` binary makes no allocation, copied-byte,
decompressed-byte, or recompressed-byte claim. The separate
`litchi-perf-baseline-alloc` target enables the benchmark-only
`allocator-metrics` feature and adds `operation_metrics.allocation` for
filesystem child operations. It wraps `std::alloc::System` with global atomic
counters, starts a non-overlapping region immediately before the timed child
operation, and publishes checked per-sample differences aligned to
`elapsed_ns.samples`. Totals include allocations from worker threads. Absolute
live/high-water counters are sampled before and after; they are never reset and
are not presented as an operation peak. A counter overflow or region-acquire
failure retains an explicit `overflow`/`unavailable` status and omits numeric
vectors. The tool identity records `tool.instrumentation` as
`system_allocator_operation_scoped`; allocator-instrumented elapsed samples
are never used for latency claims. `perf_abba_summary.py` rejects that identity
for latency ABBA, while `perf_compare.py` with a matching policy identity
withholds elapsed comparisons. The checked allocator policy selects only the
`opc_file_eager_open` warm/cold filesystem rows and requires every measured
allocation vector; its case/corpus/cache manifest is
`docs/performance/results/perf-regression-allocator-manifest-v1.json`. Use
`docs/performance/perf-regression-policy-allocator-v1.json` for that
allocator-only comparator; the checked normal policy continues to reject this
binary identity. The existing raw `filesystem_evidence` samples remain
unchanged.

In-process sink-only envelopes use `latency_claim: comparable_timed_operation`;
filesystem envelopes use `evidence_only_filesystem_selector`. The Python
comparator validates the claim on both reports, excludes evidence-only elapsed
vectors from latency comparisons, and includes `cache_state` in the result key
when present so warm and cold-requested rows cannot be paired accidentally.

Cases with a top-level `sink` summary also carry the same envelope. Their
`sink.accepted_bytes`, `sink.write_calls`, `sink.largest_write`, and
`sink.write_size_buckets` vectors are aligned views of the already-validated
deterministic summary, repeated once per retained elapsed sample. These are
logical lengths accepted by the harness sink's `Write::write` boundary; they
are not requested lengths, rejected calls, operating-system syscalls, disk
I/O, memory-copy counts, or writer-internal buffering. Requested-versus-accepted
lengths are therefore not inferred, especially for short-write sinks.
`sink.output_bytes` remains a separate final-output-length metric and is never
derived from accepted sink bytes. `sink.write_status` reports applicability for
the write vectors
independently of the existing `sink.status` for `output_bytes`. Filesystem
selectors do not expose a logical sink summary, so their sink-write vectors
remain explicitly `not_applicable`.

When `cold-verified` is selected, the evidence object additionally records
`cold_verified_status`, optional `cold_verified_samples`,
`cold_verified_claim_scope`, and `cold_verified_fincore_command`. Each
`cold_verified_samples` entry records the explicit status, numeric
`filesystem_magic`, page/source and aligned-source byte counts, aligned-source
SHA-256, fsync/advice state, fincore size/resident/dirty/writeback counts,
`read_bytes_before`/`after`/`delta`, and method/fallback evidence. Fincore
provenance is privacy-preserving: `fincore_tool` is only the canonical
basename, `fincore_sha256` and `fincore_version` identify the executable, and
`fincore_stderr_sha256`/`fincore_stderr_bytes` plus their version-stderr
counterparts record digests and lengths without serializing stderr contents.
Ineligible entries remain in `cold_verified_samples`/status evidence but do
not produce a timed `CaseResult`; prepared query controls use
`ineligible_prepared_query_control`.

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

Positional XLSX query records additionally contain `source.xlsx` arrays for physical
overlap with the workbook, selected worksheet, all unselected worksheets,
shared strings, and styles compressed member ranges. These overlap counters
are intentionally truthful about ZIP read amplification and therefore are not
semantic materialization counters. Those query cases emit no XLSX
materialization count because their production APIs do not directly expose
one. Instead, each case
enforces its semantic deferral claim with a fresh post-timing worksheet access
that must add I/O for that worksheet's exact compressed member range.

Scaling records contain `execution.worker_count`, `logical_tasks`, and
`logical_bytes`. One result is emitted per resolved worker count, in ascending
order. Those fields and the elapsed samples are sufficient to compute
throughput, speedup, scaling efficiency, and an Amdahl serial-fraction estimate
outside the harness.

The top-level `parallel_metrics` envelope is explicitly marked
`claim: "descriptive"` and repeats only sound parallelism evidence;
the comparator validates its shape and metadata but does not compare it as a
latency/resource regression metric. `configured_worker_budget` is the sorted
`configuration.execution_workers` selection, and each scaling result reports
its configured `execution.worker_count` and deterministic
`execution.logical_tasks`; every reported worker width must be a member of
that configured budget. OPC cache contention results additionally report
`observed_local_worker_count` only when one explicitly created worker team is
present; a serial zero-team result is `not_applicable`, and multiple teams are
`unavailable` because the local boundary is ambiguous. This is a
harness-created local team width, not a process thread count.
`deterministic_chunk_count` is unavailable until a producer exposes an exact
chunk counter. Range-simulation results may additionally report their exact
per-sample `source.simulation.physical_request_count` vector, reordered by
`elapsed_ns.sample_order`, as deterministic request/chunk evidence. The CFB
selective path reports the `read` phase only, and CFB `open_stream` sums its
`per_operation` phases only; both deliberately exclude the timed-open phase.
Other results leave `deterministic_chunk_count` unavailable rather than infer
it from bytes.
`lock_wait_ns` is unavailable because waiter counts are not lock-time
measurements. The envelope never reads process-global thread lists, converts
CPU utilization or waiter counts into worker/lock metrics, or infers chunks
from bytes. Every unavailable value carries a scope and reason. The comparator
cross-checks measured scalar and sample-vector values against their result and
configuration fields, including cache-state identity and the exact sorted
`sample_order`; the ABBA summary and package tools apply the same validation
when a schema-v1 report emits this envelope. None treats the descriptive
parallel evidence as a latency or resource regression metric.

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
I/O, memory copies, compression, or performance. When the summary is promoted
into `operation_metrics`, each scalar and bucket count uses the same accepted
boundary and is repeated only because the harness proved the retained sink
summary deterministic; no per-sample requested length is fabricated.

Logical-tail RTF cases also emit `output_sha256`. Their `sink.rtf_tail_append`
object records the implementation, operation (`append` or `exact_noop`),
source/source-retained, candidate-retained, caller-input, inserted, and
published output bytes, appended paragraph/run counts, the 16 KiB sink window,
and boolean exact-no-op, in-memory patch, durable patch, reopen, and
source-conflict gates. The historical pair retains its historical timing
boundary (staging/commit remain inside `elapsed_ns`) and uses
`WindowedHashingSink`; it is not publication-only evidence. All four new
Commit/PublicationPlan selectors are measured at the pre-staged
publication-call interval using `WindowedCountingSink` and additionally emit
`source.rtf_tail_publication` with separate planning/publication/reopen/lifecycle
vectors and an explicit `performance_claim` withholding end-to-end,
rich-format, allocation/RSS, physical-I/O, and ABBA latency claims.
`planning_ns` and `publication_ns` have one entry per retained sample;
`reopen_ns` and `lifecycle_ns` are one-element preflight-only vectors because
the expensive correctness gates run once outside the sample loop rather than
repeating for every sample. Their cardinality therefore intentionally differs
from the per-sample timing vectors.
`retained_output_bytes: 0` describes only the timed sink. The changed Commit
selector retains a complete candidate; the PublicationPlan retains source plus
bounded insertion; exact-no-op selectors report zero distinct target-candidate
retention. None of these fields is a process-RSS or transaction-memory claim.

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

The four XLSX scalar-cell clear/remove lifecycle selectors expose the same
`source.xlsx_cell_values` phase vectors as the existing scalar-cell controls:
`open_ns`, `plan_ns`, `commit_ns`, `publication_ns`, and `reopen_ns`, plus
exact output and semantic hashes. Eager records intentionally report zero
source reads/materializations; source-backed records report generic positional
`ReadAt` calls/bytes and successful cached OPC payload materializations. The clear
semantic hash requires the target owner to remain with an empty payload, while
the remove hash requires that owner to be absent and all other numeric owners
unchanged. Their publication sink reports `retained_output_bytes: 0`.

### XLSX existing-row visibility lifecycle controls

The four opt-in selectors are:

- `xlsx_eager_row_visibility_edit_save`
- `xlsx_source_backed_row_visibility_edit_save`
- `xlsx_eager_row_visibility_batch_edit_save`
- `xlsx_source_backed_row_visibility_batch_edit_save`

`--xlsx-row-visibility-shape medium,large` selects a deterministic one-sheet,
media-rich corpus with eight fixed 512-KiB media entries. `medium` contains
512 × 16 cells and `large` contains 2,048 × 32 cells. The scalar pair hides
one existing visible row; the batch pair unhides exactly 256 existing rows
whose initial state is hidden. Eager and source-backed outputs are checked
against their own deterministic expected bytes, while semantic row-state
digests must agree across implementations.
Each retained sample reports separate `open_ns`, `plan_or_stage_ns`,
`commit_ns`, `publication_ns`, and `lifecycle_ns` vectors plus its semantic
SHA-256. `lifecycle_ns` reopens the untimed expected artifact; exact measured
sink length/SHA-256 binds that artifact to the zero-retention publication.
Source-backed records additionally expose logical owned-source `ReadAt`
calls/bytes, selected versus unselected worksheet reads, source-version
checks, and cache diagnostics. These are ingress and pre-publication
diagnostics only; they do not describe physical I/O, allocations, total
memory, or post-publication cache state. The fixed windowed hashing sink
retains zero output bytes.

Common untimed gates cover exact output, semantic reopen, and raw untouched
ZIP-member identity. Exact no-op, foreign/stale package and source revisions,
signed, protected, formula, markup-compatibility, macro, relationship,
partial-sink, zero-output, and source-counter gates exercise the source-backed
transaction only and are omitted from eager records. The controls are
correctness/phase evidence only; they make no release speedup or latency claim.

Change 0167 measures a production-only source-backed publisher refinement
without changing these selectors or matrix counts. Matched row commits now
reuse the existing cell-values lineage/version provenance and skip one
publication-time semantic worksheet reload, cell parse and row scan; the
mandatory OPC selected-member read and complete sequential publication remain.
Clean CPU-2 A/B/B/A records use 20 warmups and 500 samples. Publication
p50/mean/p95/p99 is descriptively 50.42%-68.23% lower in both pair directions,
but same-implementation drift exceeds the 5% gate and two medium complete-
workflow p99 directions regress. No acceptance-grade end-to-end latency,
allocation/RSS, physical-I/O, or producer claim is made. See
[`0167`](../../docs/performance/changes/0167-xlsx-row-visibility-provenance-reuse.md)
and the
[summary](../../docs/performance/results/xlsx-row-visibility-provenance-0167-summary.json).

Change 0168 refines the existing native XLS plan-only numeric selectors without
changing their inputs or matrix counts. The CFB planner now offers an additive
owner-validation callback on the same composed positional view that it reopens
and verifies. XLS uses that callback for Workbook coverage, protection, macro,
and numeric readback checks before CFB's final fingerprint fence. The former
two post-plan `composed_source()` calls are gone: Number avoids 33,991,680
logical source bytes and 34 one-MiB `ReadAt` calls per sample; RK/MulRK avoids
405,504 bytes and two calls. Those are code-derived in-memory scan counts, not
physical-I/O measurements. Clean 20-warmup/500-sample CPU-2 A/B/B/A records
observe 19.22%-28.16% lower complete-workflow and 37.58%-48.04% lower semantic-
commit p50/mean/p95/p99 values in both paired directions. The stability gate
fails at 10.56% maximum control and 9.81% maximum candidate drift, so the
production work elimination is retained without an acceptance-grade latency,
tail, allocation/RSS, physical-I/O, or real-producer claim. See
[`0168`](../../docs/performance/changes/0168-xls-numeric-validation-fusion.md)
and the
[summary](../../docs/performance/results/xls-numeric-validation-fusion-0168-summary.json).

### Native DOC owner/public phase evidence

`doc_owner_public_phases` adds `source.doc_owner_public_phases` without
changing schema version 1 or the default case matrix. Its per-sample vectors
are `open_owner_ns`, `open_public_ns`, `open_retain_ns`, `edit_new_ns`,
`edit_replacement_ns`, `edit_authoring_ns`, `edit_finish_ns`,
`edit_final_owner_ns`, `edit_final_public_ns`, `edit_final_retain_ns`,
`edit_patch_ns`, `edit_commit_outer_ns`, and
`edit_output_materialization_ns`. Change 0165 additionally records the
independent `expected_source_fingerprint` and `expected_target_fingerprint`
values, per-sample `source_fingerprints` and `target_fingerprints`, and the
post-lifecycle vectors `same_lineage_apply_ns`, `deferred_fingerprint_ns`,
`workflow_no_diagnostic_ns`, and `workflow_with_fingerprint_demand_ns`.
`deferred_fingerprint_ns` covers the first source/target diagnostic demand;
the two workflow vectors are checked arithmetic extensions and are not folded
into `measured_total_ns`. The four corresponding boolean gates are
`same_lineage_apply_verified`, `reopened_source_apply_verified`,
`independent_fingerprints_verified`, and `workflow_arithmetic_verified`.
The summary also records open/edit outer totals, measured lifecycle totals,
attributed totals, checked unattributed time, source/candidate sizes and
hashes, one output hash per sample, and the existing semantic/patch/refusal/
preservation gates. The vectors retain sample iteration order;
`elapsed_ns.samples` is sorted by the existing statistics helper, so
comparisons use the same measured-total multiset rather than assuming
positional alignment with the evidence vectors.

The feature-gated production observer is clock-free and content-free. The
harness alone timestamps events with `Instant` using a bounded preallocated
recorder. Observer dispatch and recorder validation are present in outer or
lifecycle measurements; final per-event bookkeeping occurs after an individual
phase duration is sampled. The explicit unattributed remainder makes that
non-additive work visible. No result is evidence of speedup, physical I/O,
allocation, RSS, cold-cache behavior, or real-producer coverage.

## External profiling

Build the binary once before collecting counters, then invoke it directly so
Cargo is not part of the profile:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
perf stat -d tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --case cfb_read_one --shape wide-root \
  --payload incompressible --json target/perf/perf-cfb-read.json
```

Allocation evidence uses the separate benchmark-only target. Build it with
the feature and invoke `litchi-perf-baseline-alloc`; its report identity is
distinct from the normal binary, and its elapsed samples are not latency
evidence:

```sh
cargo build --release --locked --features allocator-metrics \
  --bin litchi-perf-baseline-alloc \
  --manifest-path tools/perf-baseline/Cargo.toml
tools/perf-baseline/target/release/litchi-perf-baseline-alloc \
  --warmup 3 --samples 15 --case opc_file_eager_open \
  --filesystem-cache warm,cold-requested \
  --json target/perf/perf-opc-file-alloc.json
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
