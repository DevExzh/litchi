# OPC, CFB, OLE2 Office, OOXML, RTF, and ODF performance baseline

`litchi-perf-baseline` is an isolated, reproducible measurement tool for the
ZIP/OPC and CFB/OLE2 substrates, fresh DOC/XLS/PPT writer packaging, and
public-API XLSX snapshot/edit/save flows, and opt-in DOC/XLS/PPT,
DOCX/PPTX/RTF/ODT/ODS/ODP semantic flows. It creates every corpus in memory; it also exercises
source-backed XLSX catalog, worksheet reads, and guarded calculation-metadata
and page-break/page-margin/print-options publication over positional I/O. It does not
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
media-rich PPTX semantic publication case, one XLSX commit/read attribution case,
two matched XLSX calculation-metadata publication cases, two matched XLSX
page-break publication cases, two matched XLSX page-margin publication cases,
two matched XLSX print-options publication cases,
four opaque-heavy common OLE2 stage/edit-save cases, 21 native OLE2 semantic cases, 16
DOCX/PPTX semantic cases, nine RTF semantic cases, and 26 ODF semantic cases
are opt-in, for 132 selectable cases in total:

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

For just the end-to-end legacy writer packaging runs:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to \
  --writer-shape tiny,large,payload-heavy --json target/perf/legacy-writers.json
```

Run the complete native DOC/XLS/PPT semantic matrix over the same deterministic
tiny and large writer artifacts (40 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --writer-shape tiny,large \
  --case doc_semantic_open,doc_semantic_list_paragraphs,doc_semantic_one_paragraph,doc_semantic_full_text,doc_semantic_noop_edit_save,doc_semantic_one_edit_save,xls_semantic_open,xls_semantic_list_worksheets,xls_semantic_one_cell,xls_semantic_full_cell_scan,xls_semantic_noop_edit_save,xls_semantic_one_edit_save,ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_one_shape_text,ppt_semantic_full_text,ppt_slide_order_snapshot_open,ppt_text_edit_one_edit_save,ppt_semantic_noop_edit_save,ppt_semantic_one_edit_save \
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

Run the complete tiny semantic ODF smoke matrix (23 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case odt_semantic_open,odt_semantic_list_paragraphs,odt_semantic_one_paragraph,odt_semantic_full_text,odt_semantic_create_small,odt_semantic_noop_edit_save,odt_semantic_one_edit_save,odt_semantic_one_percent_edit_save,ods_semantic_open,ods_semantic_list_sheets,ods_semantic_one_cell,ods_semantic_cell_sweep,ods_semantic_full_cell_text,ods_semantic_create_small,ods_semantic_noop_edit_save,ods_semantic_one_edit_save,odp_semantic_open,odp_semantic_list_slides,odp_semantic_one_slide,odp_semantic_full_text,odp_semantic_create_small,odp_semantic_noop_edit_save,odp_semantic_one_edit_save \
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

Run the backward-compatible plain tiny semantic RTF smoke matrix (nine
records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --case rtf_semantic_open,rtf_semantic_paragraph_count,rtf_semantic_list_paragraphs,rtf_semantic_collect_paragraphs,rtf_semantic_one_paragraph,rtf_semantic_full_text,rtf_semantic_stream_save,rtf_semantic_noop_edit_save,rtf_semantic_one_edit_save \
  --json target/perf/semantic-rtf-smoke.json
```

Select all transport and producer variants for the complete tiny RTF coverage
matrix (33 records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny \
  --rtf-variant plain,byte1252,lzfu,watermark \
  --case rtf_semantic_open,rtf_semantic_paragraph_count,rtf_semantic_list_paragraphs,rtf_semantic_collect_paragraphs,rtf_semantic_one_paragraph,rtf_semantic_full_text,rtf_semantic_stream_save,rtf_semantic_noop_edit_save,rtf_semantic_one_edit_save \
  --json target/perf/semantic-rtf-variants-smoke.json
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

The RTF cases exercise only the ordinary native `litchi_rtf::Document` facade:
owned-byte open, lazy paragraph enumeration, one middle paragraph, first
complete-text materialization, exact source streaming, exact empty-edit
publication, and capability-bounded paragraph edit/save. `--rtf-variant`
defaults to `plain`, preserving the historical seven-row commands.

| Variant | Source | Shapes | Supported cases |
|---|---|---|---|
| `plain` | Deterministic direct ASCII RTF | tiny, medium, large | All seven |
| `byte1252` | Deterministic raw CP-1252 bytes containing literal `0xe9` | tiny, medium, large | Open/read/stream/no-op; changed splice is excluded because candidate validation refuses this byte layout |
| `lzfu` | Deterministic LZFu compression of the plain bytes | tiny, medium, large | Open/read/stream/no-op; changed transport rewrites are explicitly unsupported |
| `watermark` | Content-addressed real-producer `test-data/rtf/watermark.rtf` | tiny selector only | Open/read/stream/no-op; its meaningful content is header drawing metadata rather than editable body text |

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
full-text method.

`ods_media_one_edit_save` is a separate fixed-medium corpus: 2,048 cells plus
eight deterministic 2 MiB resources under `Pictures/`. It times public unified
snapshot open, one middle-cell edit, commit and output materialization. Outside
timing it reopens the complete grid and verifies every resource path, manifest
media type, exact payload and deterministic output. This case does not vary
with `--semantic-shape` and is not part of the 23-record tiny ODF smoke matrix.

`odt_media_paragraph_edit_save` is a separate fixed-medium corpus: 200
paragraphs plus eight deterministic 2 MiB resources under `Pictures/`. It
times public snapshot open, replacement of the middle paragraph, commit, and
output materialization. Outside timing it reopens every paragraph, verifies
every resource path, manifest media type and exact payload, checks patch
replay, exact inverse and stale-source refusal, and requires deterministic
output. The harness regression additionally proves raw local/central record
identity for all untouched core and media members. This case does not vary
with `--semantic-shape` and is opt-in.

`odp_media_textbox_edit_save` is a separate fixed-medium source-backed
publication corpus: 12 slides plus eight deterministic 2 MiB resources under
`Pictures/`. It times public snapshot open, one `add_text_box` operation,
commit, and output materialization. Outside timing it checks every original
slide, the inserted text box through `rich_content`, exact patch/inverse and
stale-source behavior, deterministic output, and every resource payload and
manifest media type. It does not vary with `--semantic-shape` and is opt-in.

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

`configuration.rtf_variants` and the optional `corpus.rtf_variant` field are
additive schema-v1 identifiers for the selected RTF input capabilities.
Non-RTF corpus records omit `rtf_variant`. Corpus names include the variant so
repeated case names remain unambiguous in multi-variant reports.

`configuration.range_simulation` records fixed latency, request overhead,
bandwidth, and maximum physical range. `configuration.execution_workers`
records the resolved, capped, deduplicated scaling points in deterministic
ascending order.

The twenty-one positional cases add a `source` object; older cases omit it. Its
arrays contain one value for every measured iteration and record `read_calls`,
`read_bytes`, compressed ordinary-OPC-payload range overlap, and
`max_in_flight_reads`. Applicable OPC cases also record a semantic per-sample
`ordinary_payload_materializations` count; it may be zero, one, or multiple
Parts depending on the timed operation. Payload-range overlap is a
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

Publication cases may additionally emit `output_sha256`, independently
identifying the deterministic changed archive without changing schema v1. For
`opc_source_overlay_one_part_save`, its
`ordinary_payload_materializations` value is exactly one per sample: the
selected original Part is validated, while every unselected member is copied
physically without semantic materialization.

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
