# Change 0124: ODS unified-root filesystem and positional-read evidence

The performance harness now exposes six opt-in ODS controls over the existing
deterministic media-rich ODS corpus:

- `ods_file_eager_open` and `ods_file_source_open` compare eager byte-backed
  and filesystem source-backed `litchi::Workbook` construction;
- `ods_file_eager_selected_cell` and `ods_file_source_selected_cell` compare
  typed `Spreadsheet` and `SourceBackedSpreadsheet` cell queries after owner
  preparation; and
- `ods_file_eager_selected_media` and `ods_file_source_selected_media` compare
  typed selected-member reads over one `Pictures/*` member.

The six names are opt-in. They do not alter the default 36 cases / 198 result
records. The selectable case matrix now contains 243 names.

Corpus construction and temporary-file publication are outside timing. Eager
byte cloning is outside the open timer. Root open timing covers only root
owner construction. Typed selected-cell and selected-media timing opens the
owner before the timer and measures only the selected query. Every sample then
checks root worksheet names/count/text, complete cell parity, metadata, exact
source-file bytes and hash, archive member count, manifest/media payloads, and
typed ODS parity.

Source evidence is collected by separate instrumented
`SourceBackedSpreadsheet` replays for each measured sample. The open replay
records logical positional reads while proving that unrelated `Pictures/*`
compressed payload ranges are untouched. After the retained content owner is
prepared, selected-cell replay must add zero source reads. Selected-media
evidence pairs an all-Pictures replay, whose overlap must equal the selected
member range, with a second replay instrumented with only that selected
compressed range; both must cover exactly the selected range, excluding other
media ranges. Compressed source-range bytes and uncompressed payload bytes are
reported in separate fields. Eager controls intentionally emit empty source
vectors.

These are correctness and logical-range observations only. They do not claim
latency improvement, physical disk I/O, decompression work, allocation count,
peak memory/RSS, release behavior, or a release-ABBA result.

## Verification

The focused harness test is
`media_rich_ods_unified_root_file_selectors_are_matched_and_lazy`. A one-sample
smoke run can select all six names:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --case ods_file_eager_open,ods_file_source_open,\
ods_file_eager_selected_cell,ods_file_source_selected_cell,\
ods_file_eager_selected_media,ods_file_source_selected_media \
  --samples 1 --warmup 0 --json target/perf/ods-root-source-smoke.json
```

The smoke output is evidence that the selectors execute and preserve the
declared invariants; it is not a benchmark result.
