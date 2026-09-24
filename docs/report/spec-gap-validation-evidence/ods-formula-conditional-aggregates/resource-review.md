# Conditional aggregate resource review

This review covers the value-evaluator seams in `value.rs`, the shared
criterion matcher, and `value/conditional.rs` for `SUMIF`, `SUMIFS`,
`COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and `AVERAGEIFS`. It applies ADR 0005's
finite work and storage accounting and ADR 0006's typed-failure and
failure-atomicity requirements, together with the conditional aggregate
contract in `contract.md`. No production code was changed by this review.

The current focused evidence is:

* `cargo check -p litchi-ods` passes.
* `cargo test -p litchi-ods --test ods_formula_conditional_evaluation --test ods_formula_conditional_limits -- --quiet` passes: 22 evaluation tests and 11 resource tests.
* `cargo clippy -p litchi-ods --lib -- -D warnings` passes.
* The isolated all-target gate passes `cargo test --locked -p litchi-ods`,
  `cargo clippy --locked -p litchi-ods --all-targets -- -D warnings`,
  rustdoc, format, and the batch rustfmt check; the gate logs and hashes are
  recorded under `gates/`.

## Closed resource checks

`read_area_cell_read` charges cell work before checked coordinate arithmetic
and before `read_reference_cell`. `read_reference_cell` enforces the
cumulative `max_reference_cells` counter and checks execution both before and
after the provider operation. The outer value evaluation retains its
source-version fence and final cancellation check. The focused limits tests
cover zero-cell, work, storage, cancellation, cumulative-read, unsupported
provider, and source-change failures.

Conditional range and matcher vectors now declare their reservation before
the vector. Reverse local drop order therefore destroys each vector before
releasing its reservation on both success and `?` exits. Compiled criteria
retain borrowed matcher text. Provider `CellRead::Text` becomes a borrowed
`TextValue`; candidate-cell wrappers do not copy or reserve text per cell.

The range scan is streaming. It retains range metadata, compiled matchers,
and fixed-size numeric state, and it reads an optional destination only after
all criteria match. The current shape checks include the selected target for
`IFS` functions. Checked additions guard cell indexes, coordinates, generated
planes, and destination geometry; the focused invalid-geometry tests did not
find an indexing panic. The criterion cache classifier conservatively rejects
position-dependent scalar expressions and excludes `MUNIT`'s scalar size
argument; the position-sensitive MUNIT regression passes.

## Resource findings closed in the current revision

1. **Geometry work is charged before resolver metadata.**
   `validate_projected_geometry` now charges its bounded plane unit before
   resolving the generated sheet and extent. `projected_cell` charges the
   selected cell before any generated-sheet or extent lookup. The repeated
   `sheet_name_at`/`sheet_extent` calls in `projected_cell` remain a bounded
   optimization opportunity; hoisting a per-plane plan can be deferred until
   profiling shows that metadata lookup is material.

2. **`compatible_shape` traversal is charged.**
   The shape helper now receives the evaluator and charges before each range
   and physical-area inspection while retaining formula `#VALUE!` results.

3. **Direct array-error inspection is charged.**
   Conditional dispatch now mirrors the value aggregate helper: scalar error
   inspection charges scalar work, each materialized array cell is charged
   before inspection, and `Areas` remain lazy and untraversed.

4. **Match-count overflow preserves a formula Number error.**
   `ConditionalState::observe_match` keeps checked addition and records
   `ScalarError::Number` as a generated formula error on overflow. Typed
   provider and resource failures still propagate immediately, and selected
   formula errors retain precedence at finish.

These findings are independent of the earlier text-allocation concern: the
provider text path is borrowed and allocation-free in the current evaluator.
