# Public compact source-proof guards

This directory is an isolated, baseline-compatible source bundle for the
0552 XLSX compact source-cell proof experiment. It is deliberately outside
the live Rust tree: copy the files to their matching repository-relative
paths before running the test target. The copied parent test adds one module
declaration for `compact_source_proof.rs`; the existing planning and
post-EOF child modules are copied as well so the bundle is self-contained.

The new module uses only the public `SourceBackedEditor`, `MultiSourceEdit`,
`MultiCommit`, `MultiPatch`, `Snapshot`, `OpcPackage`, and `Error` surfaces,
plus the existing parent fixture helpers. It has five focused tests:

* `compact_public_guard_preserves_late_unused_attribute_decode_errors`
  changes `B1` while an unchanged row `spans`, row `xmlns:q`, or cell
  `xmlns:q` contains `&missing;`. The source validator and ordinary parser
  admit those allowed attributes; the commit-time layout scan must return
  exactly `Error::Invalid("at 1..8: unrecognized entity \`missing\`")`.
  Two retries must produce the same error, preserve source bytes, and leave a
  `Sheet2` no-op usable.
* `compact_public_guard_keeps_row_order_in_planning_and_cell_order_in_commit`
  keeps the existing boundary explicit. A row-order violation returns from
  `edit_sheets` as `Error::Invalid("worksheet row 1 appears after row 2")`.
  The ordinary semantic parser admits B1/A1 wire order, but a requested
  rewrite returns from `commit` as
  `Error::Invalid("cell edits require strictly increasing cell references within each row")`.
  Both error locations and strings are exact; the test does not accept a
  fallback-stage alternative.
* `compact_public_guard_handles_empty_sheet_data_inverse_and_recommit`
  checks an empty-`sheetData` no-op (`changed() == false`, empty patch, and
  byte-identical stream publication), inserts `A1`, clones the immutable
  snapshot, applies and inverts the patch, verifies the original worksheet
  bytes are restored, and reopens the published result for an independent
  second edit.
* `compact_public_guard_preserves_empty_rows_and_cells` changes `C3` in a
  sheet containing an empty row, an empty cell with an explicit address, and
  an inferred empty cell. It checks semantic ownership and the emitted
  empty forms remain present alongside ordered output.
* `compact_public_guard_preserves_typed_prefixed_scalar_forms` changes a
  boolean scalar in transitional, strict-prefixed, and transitional
  prefix-rebinding worksheets. It checks semantic `Value::Bool(false)`, the
  preserved qualified element names, the `t="b"` type, `<v>0</v>` payload,
  and owning-package readback.

The exact public error formatting comes from the current `litchi_xlsx::Error`
definition: `Error::Invalid` stores the inner string and displays it with the
`invalid XLSX structure: ` prefix. The expected entity error therefore has
the display form `invalid XLSX structure: at 1..8: unrecognized entity
\`missing\``. Source-revision and patch-lineage behavior remains covered by
the copied parent tests, including typed
`Error::Package(OpcError::SourceChanged { .. })` publication failures and
`Error::PatchConflict` replay failures.

The relevant existing baseline commands are:

```text
cargo test -p litchi-xlsx --test source_backed_cell_values
cargo test -p litchi-xlsx --test source_backed_cell_values compact_public_guard
```

The coordinator owns copying this bundle, running the baseline and candidate
commands, and deciding whether the candidate is admitted. No build or test
was run while preparing these files.
