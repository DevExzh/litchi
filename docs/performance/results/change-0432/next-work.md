# Change 0432 next work: bounded ODS semantic authoring

This is a source and coverage audit at revision `cf212b9a8`. It records a
candidate workstream; it contains no new implementation or performance result.
The overall non-iWork goal remains open.

## Recommended coverage gap

The recommended next creation/append workstream is a bounded semantic ODS
row-authoring path. This recommendation follows the missing coverage; its
performance impact must be measured before selecting an optimization. The goal requires separate evidence for streaming creation,
logical append, adding a package Part, and arbitrary modification followed by
repackaging ([GOAL.md](../../../../docs/GOAL.md#L272-L305)). It also requires
large row, paragraph, or slide streams whose authoring memory is proportional to
an explicit window ([GOAL.md](../../../../docs/GOAL.md#L804-L810)). ODS currently
has semantic builders and append methods, but no semantic writer that consumes a
row stream directly into a bounded sequential sink.

The current paths have these distinct scopes:

| Required semantic | Current path and evidence | What it does not establish |
|---|---|---|
| Streaming creation from scratch | [`Builder::add_sheet`](../../../../crates/litchi-ods/src/authoring/builder.rs#L102-L110) parses all existing sheets into a `Vec<Sheet>`; [`Builder::build`](../../../../crates/litchi-ods/src/authoring/builder.rs#L237-L263) retains `content_xml` and returns `Vec<u8>`. | Bounded row authoring or bounded package output. |
| Logical append to an existing structure | [`Edit::append_row` and `Edit::append_sheet`](../../../../crates/litchi-ods/src/document.rs#L1017-L1128) call [`advanced::append_row`](../../../../crates/litchi-ods/src/advanced.rs#L624-L645) or `append_sheet`, then stage a full candidate. [`splice_content`](../../../../crates/litchi-ods/src/advanced.rs#L5426-L5447) copies and rebuilds the package as `Vec<u8>`. | A bounded append transaction, source-window bound, or bounded candidate snapshot. |
| Adding a package Part | [`PackageWriter::add_file_reader_with_media_type_and_compression`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1265-L1315) streams opaque non-XML members; [`finish_to_writer`](../../../../crates/litchi-odf-common/src/core/writer.rs#L2028-L2044) streams final package output. | Semantic worksheet creation. The reader API explicitly rejects XML members ([writer.rs](../../../../crates/litchi-odf-common/src/core/writer.rs#L1317-L1347)). |
| Arbitrary modification and repackaging | ODS splice/edit paths preserve narrow source fragments and enforce output limits, but still materialize the candidate package through `splice_content`. | A general bounded-memory editing or repackaging claim. |

The coverage record confirms the gap: ODS sequential text output is conversion
evidence only, and broader streaming authoring remains missing
([CRUD_COVERAGE.md](../../../../docs/performance/CRUD_COVERAGE.md#L383-L385)).
The append/incremental index currently has only `xlsx_streaming_create` and
`rtf_streaming_create` ([crud-coverage-index-v1.json](../../../../docs/performance/crud-coverage-index-v1.json#L267-L298)).

## Recommended bounded batch

Add an explicitly narrow, opt-in ODS writer, tentatively
`StreamingSpreadsheetWriter<W: Write>`, for one ordinary worksheet and scalar
cells. The caller would submit strictly ordered rows incrementally. The writer
would retain only a reusable row/XML buffer and explicit counters, with finite
limits for rows, cells, cell text, row bytes, generated content XML, total
output, work, and cancellation. It should report typed sink/limit failures and
accepted partial-output bytes, matching the failure discipline of the existing
bounded writers.

The first contract should deliberately exclude formulas, styles, merged cells,
repeated-row expansion, foreign extensions, and arbitrary model edits. For
generated tiny/medium/large one-sheet scalar corpora, the harness should check
deterministic output, exact row/cell/text counts, manifest and mimetype, compact
XML, independent reopen/readback, sink write limits, cancellation, zero/partial
writes, and one-under/equal/over-limit behavior. Timing should cover only the
writer call; package construction, reopen, and exhaustive semantic oracles stay
outside it. Allocator-region, process-RSS, and sink-window observations must be
reported separately, with no speedup or total-memory claim until retained
ABBA/allocator evidence supports one.

This is a format-completion slice rather than another XLSX or RTF measurement.
`StreamingWorkbookWriter::write_row` already provides ordered sparse-row
creation with reusable row scratch ([streaming.rs](../../../../crates/litchi-xlsx/src/streaming.rs#L345-L390)),
and `StreamingRtfWriter` already documents fresh forward-only paragraph/run
creation without a document-wide model ([streaming.rs](../../../../crates/litchi-rtf/src/streaming.rs#L1-L8),
[streaming.rs](../../../../crates/litchi-rtf/src/streaming.rs#L252-L313)). The
performance harness recognizes streaming creation only for those two formats
([perf-baseline/src/lib.rs](../../../../tools/perf-baseline/src/lib.rs#L2407-L2409),
[perf-baseline/src/lib.rs](../../../../tools/perf-baseline/src/lib.rs#L55916-L55943)).
ODS `ods_semantic_create_small` instead times `semantic_ods_bytes` and performs
an untimed reopen, while `ods_semantic_text_to_sink` times
`Spreadsheet::write_text_to` on a document opened before the timer
([perf-baseline/src/lib.rs](../../../../tools/perf-baseline/src/lib.rs#L32848-L32880),
[perf-baseline/src/lib.rs](../../../../tools/perf-baseline/src/lib.rs#L33041-L33090)).
Neither is bounded semantic ODS creation.

## XML/package ownership review

`PackageWriter::with_writer` and `finish_to_writer` are suitable transport
building blocks, but the existing generic reader publication seam intentionally
accepts opaque non-XML members only. A generated `content.xml` path therefore
needs an ownership decision between `litchi-ods` authoring and
`litchi-odf-common` package writing. The next implementation should add only a
typed, generated-ODS XML seam with the same manifest, XML validation, output,
limit, cancellation, and partial-publication guarantees; it should not turn the
opaque `Read` API into an arbitrary XML authoring escape hatch. The seam review
must also decide where compact XML encoding and final `content.xml` closure are
owned before any public API is frozen.

A later, separate batch may add a source-backed tail-row publication plan for
logical append. That plan would need bounded `ReadAt` source windows, a bounded
row fragment, source-version/provenance checks, and typed partial output. It must
not inherit a bounded-memory claim from the current append methods while
[`XmlSplicePublication::assemble`](../../../../crates/litchi-odf-common/src/core/xml_splice.rs)
and package rebuild still materialize complete XML/package candidates. Adding an
opaque Part and arbitrary repackaging remain separate CRUD measurements.
