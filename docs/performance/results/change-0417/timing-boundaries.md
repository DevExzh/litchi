# Change 0417 selector timing boundaries

This is the source-backed scope map for the current-revision representative
CRUD matrix. It records what the harness timer contains. The index remains 30
selectors across 14 categories, with dynamic calculation/refresh unsupported.
The 0417 one-row-per-selector preflight passed. The retained 0417 normal and
allocator captures are governed by [`protocol.json`](protocol.json) and are
separate from this boundary map. `M` retains the index's default/full-run
baseline acceptance contract. `C` has no default/full-run baseline acceptance;
an opt-in 0417 report can still be accepted as separately retained descriptive
timing evidence without promoting its selector or changing its index status.

`U` has no selector. “Lifecycle” below means the named case's
open/stage/commit/publication interval; it does not mean the complete process
or the complete correctness workflow. Corpus construction, expected-output
preflight, reopen/readback, refusal gates, and most drops are outside the
reported interval unless the row says otherwise. `CountingSink` retains the
full output; its 64 KiB value limits individual write calls, not total memory.
Hashing/windowed sinks retain no output but can still have expected artifacts
prepared outside the timer. The 0417 `workers=1` setting is runner
configuration, not evidence of parallel scaling. The 37-byte RTF scratch
window and 4 KiB XLSX row window are retained-authoring windows, not process
heap bounds (`lib.rs:205-208`).

| Category | Status | Selector | Corpus and flags | Timed interval | Allocation attribution |
| --- | --- | --- | --- | --- | --- |
| Content reading/extraction | M | `xlsx_full_cell_scan` | Checked `xlsx-opc-zip`; `--xlsx-shape tiny,medium,dense-wide` | `Workbook::from_bytes` and `Sheet1` selection precede the timer; full stored-cell range scan only (`lib.rs:38291-38313`) | None (`result` sets `operation_metrics: None`, `lib.rs:54551-54562`) |
| Structural queries/analysis | M | `xlsx_list_sheets` | Same checked XLSX catalog and shapes | Workbook open precedes the timer; sheet-name collection only (`lib.rs:38240-38264`) | None (`lib.rs:38264`) |
| Conversion/export | C | `rtf_semantic_text_to_sink` | Generated schema-2; `--semantic-shape tiny,medium,large --rtf-variant plain` | Parsed RTF and sink setup precede the timer; `Document::write_text_to` only (`lib.rs:27334-27514`) | None |
| Conversion/export | C | `odt_semantic_text_to_sink` | Generated schema-2; `--semantic-shape tiny,medium,large` | Parsed document/oracle precede the timer; `Document::write_text_to` only (`lib.rs:30353-30407`) | None |
| Conversion/export | C | `ods_semantic_text_to_sink` | Generated schema-2; `--semantic-shape tiny,medium,large` | Parsed spreadsheet/oracle precede the timer; `Spreadsheet::write_text_to` only (`lib.rs:32121-32172`) | None |
| Conversion/export | C | `odp_semantic_text_to_sink` | Generated schema-2; `--semantic-shape tiny,medium,large` | Parsed presentation/oracle precede the timer; `Presentation::write_text_to` only (`lib.rs:33341-33402`) | None |
| Creation from scratch | M | `doc_fresh_write_to` | Checked `doc-cfb`; `--writer-shape tiny,large,payload-heavy` | `write_fresh_doc`: authoring plus serialization (`lib.rs:54513-54548`) | None |
| Creation from scratch | M | `xls_fresh_write_to` | Checked `xls-cfb`; `--writer-shape tiny,large,payload-heavy` | `write_fresh_xls`: authoring plus serialization (`lib.rs:54513-54548`) | None |
| Creation from scratch | M | `ppt_fresh_write_to` | Checked `ppt-cfb`; `--writer-shape tiny,large,payload-heavy` | `write_fresh_ppt`: authoring plus serialization (`lib.rs:54513-54548`) | None |
| Template/replacement | M | `xlsx_one_cell_commit` | Checked `xlsx-opc-zip`; `--xlsx-shape tiny,medium,dense-wide` | Workbook open and edit staging precede the timer; `Edit::commit()` only, with no save (`lib.rs:38996-39036`) | None |
| Template/replacement | M | `xlsx_one_percent_commit` | Same checked XLSX catalog and shapes | Same commit-only boundary, using 2/41/1,311 deterministic updates (`lib.rs:38996-39036`) | None |
| Append/incremental generation | C | `xlsx_streaming_create` | Generated schema-2; `--semantic-shape tiny,medium,large` | Hashing sink setup precedes the timer; `write_streaming_xlsx` only (`lib.rs:54402-54466`) | Deliberately none: resource collection is RTF-only (`lib.rs:54414-54431`) |
| Append/incremental generation | C | `rtf_streaming_create` | Generated schema-2; `--semantic-shape tiny,medium,large --rtf-variant plain` | Hashing sink setup precedes the timer; `write_streaming_rtf` only (`lib.rs:54402-54466`) | Dedicated allocator/process snapshots surround the writer call (`lib.rs:54414-54439,54481-54499`); process values are process-level deltas |
| Structural editing | C | `xlsx_eager_defined_names_edit_save` | Fixed generated `media-rich` corpus | The timed helper performs source open/materialization, defined-name edit/commit, and sequential publication; expected-output checks are outside (`lib.rs:46864-46901,46905-46974`) | None (`lib.rs:46984-46994`) |
| Structural editing | C | `xlsx_eager_row_visibility_edit_save` | Generated fixed corpus; `--xlsx-row-visibility-shape medium,large` | Whole case-local lifecycle: eager open, row staging, commit, and `write_to`; semantic hash and preservation gates are separate (`lib.rs:39547-39706,39712-39827`) | None (`lib.rs:39836-39845`) |
| Content deletion/sanitization | C | `xlsx_eager_cell_remove_edit_save` | Fixed multi-sheet media corpus; `--xlsx-cell-crud-shape medium,dense-sparse` | Whole case-local lifecycle: eager open, remove staging, commit, and publication; reopen/output checks are excluded (`lib.rs:39849-40200`) | None (`lib.rs:40206-40216`) |
| Content deletion/sanitization | C | `docx_story_hyperlink_redaction_save` | Fixed generated tiny seven-story corpus (`7-story-kinds-14-links-7-media-1-opaque`; index shape label `media-rich`) | Fresh source/sink setup precedes the phase timer; elapsed is `open + plan/apply + publication`; digest, cache, refusal, and output gates are outside (`docx_story_hyperlink_publication.rs:30-60,1135-1304,1453-1456`) | Dedicated allocation region starts before open and finishes after publication, so it includes package cleanup/drop after the elapsed phase timer stops (`docx_story_hyperlink_publication.rs:1282-1304,1444-1521`); normal target is unavailable, allocator target may report it |
| Cross-document copying | C | `pptx_source_backed_cross_copy_plain` | Fixed generated `plain` source/destination corpus | Source/destination open and semantic setup precede timing; elapsed is source-backed plan plus public publication; reopen/topology/refusal checks are outside (`lib.rs:46572-46723,46750-46754`) | None (`lib.rs:46781-46791`); sink retains output and 64 KiB is only a write-call bound |
| Cross-document copying | C | `pptx_cross_copy_media_rich` | Fixed generated `media-rich` source/destination corpus | Package/snapshot setup precedes timing; elapsed is separate plan + commit + OPC publication phases; reopen is recorded separately and excluded (`lib.rs:46365-46520,46530-46533`) | None (`lib.rs:46559-46569`); not an end-to-end or bounded-total-memory measurement |
| Merging/splitting | C | `xlsx_eager_merge_commit_save` | Fixed generated `sparse-a1-b2` corpus | Source open and edit preparation precede timer; commit plus sequential `write_to` only (`lib.rs:40316-40367`) | None (`lib.rs:40369-40380`) |
| Merging/splitting | C | `rtf_semantic_split_paragraph_save` | Generated schema-2; `--semantic-shape tiny,medium,large --rtf-variant plain` | One opened document, stage, commit, and windowed sequential write; correctness gates are preflight-only (`lib.rs:28037-28104,28115-28122`) | None (`lib.rs:28141-28153`) |
| Comparison/patching | C | `xlsx_join_disjoint_commit_save` | Generated XLSX `medium`; `--xlsx-cell-crud-shape medium` | Branch preparation uses scoped threads and precedes timer; timed join + commit + sequential publication; output reopen is excluded (`lib.rs:40427-40451,40560-40876,40814-40824,41121-41125`) | None (`lib.rs:41109-41117`); full output is retained by `CountingSink` |
| Comparison/patching | C | `xlsx_three_way_disjoint_commit_save` | Generated XLSX `medium`; `--xlsx-cell-crud-shape medium` | Branch preparation uses scoped threads and precedes timer; timed three-way plan + finish + commit + sequential publication; output reopen is excluded (`lib.rs:40427-40451,40560-40945,40814-40824,41129-41134`) | None (`lib.rs:41109-41117`); 64 KiB limits writes, not total retained output |
| Validation/repair | C | `odf_validation_report` | Generated schema-2; `--semantic-shape tiny,medium,large` | Borrowed in-memory bytes; `validate_package` call only (`lib.rs:29446-29504,29618-29630`) | None |
| Validation/repair | C | `odf_mimetype_repair_plan` | Generated schema-2; `--semantic-shape tiny,medium,large` | Sink setup and canonical preflight precede timer; validation + repair-plan construction + `plan.write_to` are inside (`lib.rs:29754-29793`) | None (`lib.rs:29799-29812`) |
| Dynamic calculation/refresh | U | No selector | Explicit capability required | No formula evaluation or external refresh workload is admitted | None |
| Security/protection | C | `xls_validation_report` | Generated legacy XLS; actual `--writer-shape tiny,large` only | `InstrumentedSource` setup precedes timer; `validate_source` plus its positional reads are inside (`lib.rs:29446-29504,29507-29518`) | None |
| Security/protection | C | `xlsx_eager_sheet_protection_edit_save` | Fixed generated `media-rich` corpus | Timed helper performs source open/materialization, protection edit, and sequential publication; output checks are outside (`lib.rs:49214-49286`) | None (`lib.rs:49296-49306`) |
| Low-level package operations | M | `cfb_list_streams` | Checked `cfb-ole2`; `--shape tiny,many-small,few-large,wide-root --payload compressible,incompressible` | `OleFile` open precedes timer; `list_streams` only (`lib.rs:52102-52119`) | None |
| Low-level package operations | M | `cfb_read_one` | Same checked CFB 4-shape × 2-payload catalog | `OleFile` open precedes timer; final target `open_stream` only (`lib.rs:52122-52140`) | None |
| Low-level package operations | M | `opc_open` | Checked `opc-zip`; same 4-shape × 2-payload flags | `OpcPackage::from_bytes` and part-count check are inside timer (`lib.rs:51056-51077`); archive bytes are already in memory | None |

The checked catalog and generated-shape contracts are defined in
[`crud-coverage-index-v1.json`](../../crud-coverage-index-v1.json) and
[`perf-corpus-manifest-v2.json`](../perf-corpus-manifest-v2.json). The index's
measured rows still require the default report contract (`container-baseline`
with at least 15 retained samples); this opt-in 0417 matrix does not alter
`Case::DEFAULT` or promote the 20 correctness-only selectors.
