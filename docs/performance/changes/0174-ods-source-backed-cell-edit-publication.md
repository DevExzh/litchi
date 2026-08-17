# Change 0174: ODS source-backed existing-cell publication

Date: 2026-08-17

Status: correctness and matched harness coverage; release performance claim
pending

## Scope

`litchi-ods::SourceBackedSpreadsheet` now exposes a bounded transaction for
editing existing ordinary scalar cells without first materializing the complete
ODS candidate artifact. The source owner already retains validated `content.xml`,
`styles.xml`, and worksheet projections; the new transaction clones only
touched worksheet models, rewrites eligible physical rows, validates a complete
semantic readback, and delegates sequential ZIP publication to the existing
bounded ODF source publisher. Untouched package members are raw-copied.

The public `SourceCellSnapshot`, `SourceCellEdit`, `SourceCellCommit`, and
`SourceCellPatch` types preserve exact source lineage and failure atomicity.
The patch is deliberately an in-process semantic `content.xml` patch, not a
durable or byte-exact ZIP patch. Atomic filesystem save is outside this API.

## Closure and refusals

One transaction accepts at most 4,096 unique existing cells. It supports
repeated cell-run splitting inside non-repeated physical rows and preserves
standard direct table metadata, including `table:table-column`, outside the
rewritten row spans. It refuses missing cells, duplicate coordinates, repeated
physical rows, formulas, merges, style retargeting, protected owners, unknown
values anywhere in a rewritten row, changed signatures, encryption, and row
markup that cannot round-trip through the ordinary cell model.

The row audit is a linear contiguous-subtree walk. It permits only direct
`table-cell`/`covered-table-cell` owners with at most one direct plain
`text:p`; nested elements, multiple paragraphs, non-whitespace text outside a
paragraph, CDATA, general references, comments, processing instructions, and
DTDs fail closed before publication.

Exact semantic no-ops use the common byte-exact source-copy path, including
signed and protected sources. Changed publication retains typed source-change,
limit, cancellation, execution-budget, and partial-sink progress reporting.

## Harness coverage

Four opt-in selectors raise the selectable matrix from 315 to 319 without
changing the historical default 36 cases / 198 records:

- `ods_source_eager_one_edit_save`
- `ods_source_backed_one_edit_save`
- `ods_source_eager_one_percent_edit_save`
- `ods_source_backed_one_percent_edit_save`

They use the same deterministic two-sheet, 2,048-cell, eight-resource
media-rich corpus and the same 16 KiB non-seek hashing sink, which retains no
output bytes. The one-percent cases update 21 evenly spaced existing cells.
Open, staging, commit, and sequential publication are timed; semantic reopen,
complete cell digest, media payload hashes, source/output hashes, raw untouched
member identity, exact no-op, foreign patch, replacement-limit, and
partial-sink checks are untimed gates.

These selectors establish correctness and a matched timing boundary only. No
latency improvement, allocation/RSS, physical-I/O, decompression, cold-cache,
real-producer, atomic-save, or broader ODS CRUD claim is accepted without a
clean release A/B/B/A record.

## Verification

- `cargo test --locked -p litchi-ods --all-targets`
- `RUSTFLAGS='-D warnings -D deprecated' cargo clippy --locked -p litchi-ods --lib --test source_cell_transactions -- -D warnings -D deprecated`
- focused ODS source-cell tests, including exact 4,096/4,097 bounds, standard
  table metadata, row-grammar preservation, unknown-neighbor refusal,
  signatures/protection, stale source, limits, cancellation, and partial sink
- focused harness selector and enum-count tests
- strict harness Clippy and deterministic one-sample selector smoke
