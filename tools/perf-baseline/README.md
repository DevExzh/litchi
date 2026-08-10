# OPC, CFB, legacy-writer, OOXML, RTF, and ODF performance baseline

`litchi-perf-baseline` is an isolated, reproducible measurement tool for the
ZIP/OPC and CFB/OLE2 substrates, fresh DOC/XLS/PPT writer packaging, and
public-API XLSX snapshot/edit/save flows, and opt-in DOCX/PPTX/RTF/ODT/ODS/ODP
semantic flows. It creates every corpus in memory; it also exercises
source-backed XLSX catalog and worksheet reads over positional I/O. It does not
depend on untracked office files, network state, or randomness. ODP builder
timestamps are replaced with fixed metadata before measurement. The JSON
report contains the generator parameters and SHA-256 hashes for the generated
container and target entry, so a result always identifies its exact input or
packaged output.

The tool is intentionally outside the root workspace and has no effect on
production dependency graphs.

The DOCX/PPTX/RTF/ODF semantic matrices are deliberately opt-in. They measure
only current public APIs and therefore do not change the default 36 cases / 198
records.

## Run

Run the complete default matrix (36 default cases; 198 result records: 144
substrate records, nine writer records, and 45 XLSX records). The six simulated
range cases, two execution-scaling cases, 16 DOCX/PPTX semantic cases, seven
RTF semantic cases, and 21 ODF semantic cases are opt-in, for 88 selectable
cases in total:

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

For just the end-to-end legacy writer packaging runs:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to \
  --writer-shape tiny,large,payload-heavy --json target/perf/legacy-writers.json
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

Run the complete tiny semantic ODF smoke matrix (21 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case odt_semantic_open,odt_semantic_list_paragraphs,odt_semantic_one_paragraph,odt_semantic_full_text,odt_semantic_create_small,odt_semantic_noop_edit_save,odt_semantic_one_edit_save,ods_semantic_open,ods_semantic_list_sheets,ods_semantic_one_cell,ods_semantic_full_cell_text,ods_semantic_create_small,ods_semantic_noop_edit_save,ods_semantic_one_edit_save,odp_semantic_open,odp_semantic_list_slides,odp_semantic_one_slide,odp_semantic_full_text,odp_semantic_create_small,odp_semantic_noop_edit_save,odp_semantic_one_edit_save \
  --json target/perf/semantic-odf-smoke.json
```

Run the complete tiny semantic RTF smoke matrix (seven records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case rtf_semantic_open,rtf_semantic_list_paragraphs,rtf_semantic_one_paragraph,rtf_semantic_full_text,rtf_semantic_stream_save,rtf_semantic_noop_edit_save,rtf_semantic_one_edit_save \
  --json target/perf/semantic-rtf-smoke.json
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

The RTF cases use deterministic direct ASCII RTF source with the same
paragraph counts as DOCX. They exercise only the ordinary native
`litchi_rtf::Document` facade: owned-byte open, lazy paragraph enumeration,
one middle paragraph, first complete-text materialization, exact source
streaming, exact empty-edit publication, and one checked paragraph edit/save.
Every save uses the native forward-only `Write` contract and every output is
reopened and fully verified.

| Shape | Paragraphs | Source bytes | Text bytes |
|---|---:|---:|---:|
| `tiny` | 24 | 1,347 | 1,199 |
| `medium` | 200 | 10,851 | 9,999 |
| `large` | 10,000 | 540,051 | 499,999 |

The exact stream-save and no-op cases preserve the generated input byte for
byte and emit it as one sequential write. The one-edit case stages the middle
paragraph through `replace_paragraph_text`, commits with source and semantic
readback checks, streams the changed snapshot, and verifies every paragraph
after reopen. Corpus creation, expected-output construction, and input cloning
remain outside the timed interval.

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
`odt_semantic_one_paragraph` necessarily calls the public `paragraphs()` API
and then selects the middle value because ODT has no public indexed paragraph
query; it is deliberately not described as a lazy lookup.

Each ODS batch uses `Builder`, `Spreadsheet::from_bytes`, `sheets()`, the
public logical `cell()` view, a deterministic row-major cell-text aggregate,
and the unified `document::Snapshot` transaction. ODS snapshot construction is
inside the timed edit/save interval so these cases expose the package-open cost
paid by this public editing entry point; the source-byte clone is outside the
interval. The timed work also includes staging, commit, and published-byte
observation. `ods_semantic_full_cell_text` is named explicitly because the
facade exposes cells rather than a single full-text method.

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
- `doc_fresh_write_to`: construct a new `litchi_doc::writer::Writer`, add the
  selected fixed paragraphs through its public API, and package it with public
  `write_to`.
- `xls_fresh_write_to`: construct a new `litchi_xls::writer::Writer`, add the
  selected sheets and cells through its public API, and package it with public
  `write_to`.
- `ppt_fresh_write_to`: construct a new `litchi_ppt::writer::Writer`, add the
  selected slides and text boxes through its public API, and package it with
  public `write_to`.
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
- `rtf_semantic_list_paragraphs`, `rtf_semantic_one_paragraph`, and
  `rtf_semantic_full_text`: enumerate lazy body paragraph views, resolve and
  flatten one middle paragraph, or materialize the snapshot's cached complete
  text for the first time.
- `rtf_semantic_stream_save`: stream the immutable snapshot through public
  `Document::write_to` and require byte-exact source output.
- `rtf_semantic_noop_edit_save`: commit an empty edit, require shared snapshot
  identity and exact bytes, then stream through the same forward-only sink.
- `rtf_semantic_one_edit_save`: replace the middle paragraph through the
  checked native transaction, stream the changed snapshot, reopen it, and
  verify its complete semantic projection and sink counters.

For both CFB stream-insertion cases, payload generation/cloning and writer
construction happen before timing, while writer and source destruction happen
afterward. They time stream insertion only, not complete CFB serialization. The
`few-large` shape is the useful ownership profile because its prepared target is
4 MiB; `tiny` stays small enough for a quick release-build smoke test.

The sink reserves a checked byte budget before timing and copies every output
byte into bounded memory. It records accepted bytes, write count, and the
largest write, rejects output beyond the budget or individual writes larger
than 64 KiB, and verifies the deterministic no-op output after timing. This
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
`perf_event_paranoid` value. Metadata collection is best-effort and is complete
before timed iterations. The requested JSON parent directory is created
automatically.

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

`configuration.range_simulation` records fixed latency, request overhead,
bandwidth, and maximum physical range. `configuration.execution_workers`
records the resolved, capped, deduplicated scaling points in deterministic
ascending order.

The seventeen positional cases add a `source` object; older cases omit it. Its
arrays contain one value for every measured iteration and record `read_calls`,
`read_bytes`, compressed ordinary-OPC-payload range overlap, and
`max_in_flight_reads`. OPC cases also record semantic
`ordinary_payload_materializations` (0 or 1). Payload-range overlap is a
physical request-amplification metric: bounded ZIP metadata reads may fetch
adjacent compressed payload bytes without decompressing or caching that Part.
Accordingly, `opc_source_open` may report overlap while still reporting zero
materializations; its post-timing cold access proves the distinction.

Simulated-range records additionally contain `source.simulation`: per-sample
logical read calls/bytes, physical request count/bytes, sorted physical request
sizes, and fixed request-size buckets. Request delays are computed only from
the recorded configuration; no ambient network or clock-derived input is used.

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
