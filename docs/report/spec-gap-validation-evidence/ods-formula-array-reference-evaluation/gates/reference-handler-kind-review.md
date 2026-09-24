# Reference-handler kind review

This is a source-only review of the lazy-handler path used while planning the
shape of `:`, `!`, and `~` operands. No Cargo command or production edit was run
for this review.

## Reviewed snapshot

| Source | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `4c0ce9917d23482910a1c7f8b111253430608ee36407ff3d3d4ffc04fec172e0` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `680a3fd9c3b80951408f3210f5c5373dcea7dd00c1b561b0902043438b79d907` |

## Call chain reviewed

`shape_hint_demand` schedules a reference infix as a `ShapeFrame::Reference`,
which calls `reference_operator_shape`, then `reference_shape_value`. The latter
uses one `ReferenceShapeScratch` stack. For a lazy handler it calls
`schedule_reference_handler`; that function uses `evaluate_scalar_at` for a
condition or non-reference handler value and schedules a selected reference
branch back onto the same reference-shape stack.

The old `run_isolated_matrix` implementation is not called in this snapshot
(`rg` finds only its definition). The active handler path therefore does not
re-enter `matrix_branch_shape` recursively: nested reference handlers are
scheduled on the existing explicit stack. `evaluate_scalar_at` runs in
`Mode::Scalar`, so its `run_from` cannot enter matrix shape planning. The prior
native-recursion concern is cleared for the actual caller, although the dead
helper should be removed to prevent accidental future use.

## Confirmed semantic blockers

### RHK-1: array-valued IF conditions are reduced to the top-left cell

`schedule_reference_handler` always calls
`evaluate_scalar_at(condition, Shape::new(1, 1)?, 0)`. In the ordinary value VM,
when `finish_if` is in `Mode::Matrix` and the condition is array-like, it calls
`start_matrix_if`, which produces a `RuntimeValue::Array` and applies selection
per projected cell. The shape probe instead projects only the first condition
cell and schedules one branch as though the whole handler returned a reference.

For example:

```text
IF({TRUE();FALSE()};[.C1:.D1];[.E1:.F1])
```

has an array condition under matrix semantics. It must retain array kind (and
per-cell branch selection) or reach the reference operator's typed
`Unsupported(ReferenceOperator)` path as an array operand. The current probe can
select `[.C1:.D1]` from the top-left
`TRUE()` and return `Areas`, so an enclosing `:`/`!`/`~` planner can infer
reference geometry that the ordinary matrix value path does not have.

### RHK-2: array-valued IFERROR/IFNA input is also reduced to the top-left cell

For a non-reference handler input, the helper again uses a `1x1` scalar probe.
The ordinary matrix path sees an array-like value in `finish_if_error`, calls
`start_matrix_if_error`, materializes the array, and applies catching per cell.
For example:

```text
IFERROR({#DIV/0!;1};[.C1:.D1])
```

must retain an array result with only the error element selecting the alternative.
The shape probe sees the first `#DIV/0!`, catches it, and returns the reference
alternative as `Areas`, incorrectly treating the whole handler as one reference.
The same top-left error-selection problem applies to `IFNA` and to a non-error
first cell followed by an error later in the array.

The issue also includes a reference-valued handler input. `reference_candidate`
classifies `[.C1:.D1]` as a candidate and schedules its `Areas` value, so
`IFERROR([.C1:.D1];fallback)` is retained as `Areas` by the probe. In ordinary
`Mode::Matrix`, `finish_if_error` treats `Areas` as array-like, materializes the
reference, and returns an `Array` (with per-cell error handling), which a
reference operator must reject under the typed capability contract. The helper
therefore needs to distinguish scalar-demand projection from matrix-demand
array behavior for reference-valued handler inputs as well.

`reference_operator_shape` also currently maps `RuntimeValue::Array(_)` to
`1x1`. Once array kind is preserved, that arm must return unknown/typed refusal
for a reference operator; fabricating `1x1` would retain the same under-shape.

These are blockers for the advertised evaluated-reference handler profile. The
implementation must model array kind and selected per-cell behavior, or make a
clearly documented `Unsupported(ReferenceOperator)` refusal before reference
geometry is published. The existing `coerce_reference_areas` contract returns
that typed capability failure for scalar/array operands; this review does not
claim an additional formula-level `VALUE` mapping.

### RHK-3: one-argument IF has an incorrect true path

`finish_if` implements the existing scalar contract: `IF(condition)` returns
`Logical(condition)`. In `schedule_reference_handler`, a true one-argument IF
falls through to `node.child(1)`, which is absent, and returns
`InvalidExpression`; the false path happens to return `Logical(false)`. Thus
`IF(TRUE())` is inconsistent even before it is used as a reference-operator
operand. In an outer operator the ordinary runtime would eventually produce the
typed `Unsupported(ReferenceOperator)` for this scalar operand, whereas the
shape helper fails earlier with `InvalidExpression`. The handler should preserve
the ordinary one-argument result and let the consuming operator apply its normal
typed refusal.

## Required regressions

- Use an array IF condition with differing booleans and differing reference
  branch shapes as a `:`, `!`, and `~` operand. Verify the result is typed
  `Unsupported(ReferenceOperator)` or the explicitly supported per-cell
  semantics, never top-left branch geometry.
- Use mixed error/non-error arrays for both `IFERROR` and `IFNA`; verify catches
  are per element and no reference alternative is promoted from only the first
  element; an array operand must retain the typed
  `Unsupported(ReferenceOperator)` contract.
- Exercise `IF(TRUE())` and `IF(FALSE())` through the handler shape path and
  verify the same logical values as ordinary evaluation.
- Keep an unselected missing-sheet reference in each case and verify shape
  planning does not request its metadata.

## Disposition

The active nested-handler traversal is iterative and its shared stack/work/
storage/cancellation accounting remains bounded. RHK-1 through RHK-3 prevent
acceptance of the current handler-kind implementation until array kind and
one-argument semantics are corrected and the outer diagnostic behavior is
verified. The reviewed snapshot is historical; later source changes require a
new hash and re-review.
