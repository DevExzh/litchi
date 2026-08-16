# Change 0163: XLSX scalar-cell clear/remove evidence

Date: 2026-08-17

Status: opt-in correctness, phase, logical-source-counter, and sequential-sink
evidence only. No latency, speedup, allocation/RSS, physical-I/O, cold-cache,
decompression, durable-source-patch, or real-producer claim is accepted.

## Scope

The standalone harness adds four selectors:

- `xlsx_eager_cell_clear_edit_save`
- `xlsx_source_backed_cell_clear_edit_save`
- `xlsx_eager_cell_remove_edit_save`
- `xlsx_source_backed_cell_remove_edit_save`

Each selector edits one existing numeric owner, `Sheet1!A1`, through the
public eager `WorksheetEdit` or positional source-backed cell-values editor.
Clear retains the empty `<c>` owner; remove deletes that owner. The selectors
reuse the existing `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`
corpus in `medium` and `dense-sparse` shapes: four worksheets, deterministic
numeric cells, and untouched media Parts. This contributes eight shape records
and raises the selectable matrix from 305 to 309; the default remains 36 cases
and 198 records.

## Timing and evidence boundary

Each retained sample reports separate `open_ns`, `plan_ns`, `commit_ns`,
`publication_ns`, and complete lifecycle vectors. The timer covers opening,
selector planning/staging, commit, and sequential publication. Publication
uses a fixed 64-KiB windowed hashing sink retaining zero output bytes. The
report records output digest, accepted bytes, write-size bounds, and generic
logical source/materialization counters; eager source counters are explicitly
zero/not applicable. Reopen, semantic verification, package identity, source
raw-member checks, and lifecycle gates are outside the timed interval.

## Gates

The preflight and per-sample checks require deterministic source/output hashes,
semantic reopen, unchanged unselected cell values, and the precise clear/remove
owner-count distinction. Source-backed publication additionally proves exact
raw preservation of unselected ZIP members/media and uses volatile source-bound
patch forward/inverse, exact no-op, stale-source, and foreign-source checks.
The existing focused XLSX cell-values tests remain the authority for the wider
protected, signed, formula, MCE, metadata, relationship, malformed-input,
limit, and failure-atomic refusal matrix; this benchmark tranche does not
turn those tests into a new timed security matrix. It adds no durable patch
wire or durable replay gate because the source-backed cell-values patch has no
durable serialization contract.

## Reproduction

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  xlsx_cell_lifecycle_clear_and_remove_controls_are_matched_and_bounded

cargo run --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --xlsx-cell-crud-shape medium,dense-sparse \
  --case xlsx_eager_cell_clear_edit_save,xlsx_source_backed_cell_clear_edit_save,\
xlsx_eager_cell_remove_edit_save,xlsx_source_backed_cell_remove_edit_save \
  --json target/perf/xlsx-cell-clear-remove-0163-smoke.json
```

The debug selector smoke is correctness evidence only. A future latency claim
would require clean release CPU-pinned matched ABBA runs with retained raw
reports, identical corpus/output hashes, and separately justified resource and
I/O measurements.

## Remaining gaps

The closure is limited to generated ordinary numeric scalar owners. It does
not cover formulas, shared strings, dates, metadata-rich cells, merged or
validated cells, extension/MCE owners, row/column structural deletion,
dependency recalculation, signed or real-producer workbooks, durable source
patches, allocation/peak-memory behavior, physical or cold I/O, or general
XLSX deletion semantics.
