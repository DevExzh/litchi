# Runtime reference-kind review

This is a source-only review of the runtime kind walker used while planning
lazy `IF`, `IFERROR`, and `IFNA` operands of `:`, `!`, and `~`. The reviewer did
not edit production code, run Cargo, or create a build/worktree. The root
agent's baseline-40 directed diagnostic confirmed the list-classification
counterexample recorded below.

## Reviewed snapshot

| Source | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `e351dcd1589c6c97dca115a1c665468ca35b9f25c57ab03226e79029c4ce3dd4` |

## Scope and cleared paths

`reference_handler_operand_kind` now uses `ReferenceKindScratch` and an
explicit `Visit`/`Unary`/`Infix`/function/lazy-continuation stack. The current
kind walker does not call `reference_shape_value`; the earlier kind-to-shape
native recursion is therefore absent in this snapshot. Reference operators
are applied to retained runtime operands on the same kind stack.

`ReferenceKindScratch` declares each vector before its matching reservation.
The same ordering is used by `RuntimeAreaSet` and its nested records, so
normal returns and `?` exits drop retained vector elements and buffers before
their budget reservations. I found no reservation leak in this path.

Lazy continuations visit a condition/value first and schedule only the selected
branch. One-argument `IF` and the direct nested-`Null` typed-refusal case are
covered by the later directed regressions; those earlier findings are cleared
for this snapshot. `AND` and `OR` remain scalar sequence aggregates. Ordinary
arrays and non-list references are classified as array-like for `XOR` and the
other eager matrix functions, matching the corresponding `apply_function`
branches.

## Blocking finding

### RRK-1: `ReferenceList` is collapsed into `Reference`

At `reference_operand_kind_from_runtime` (currently lines 3508–3515), every
`RuntimeValue::Areas(_)` becomes `ReferenceOperandKind::Reference`; the
`RuntimeAreaSet::is_list` bit is ignored. This is inconsistent with the value
VM, where `ReferenceList` is a distinct value. `materialize_for_array` and the
list conversion in `apply_function` produce a scalar `Error(Value)` for lists
except for the special `AND`/`OR` and reference-operator paths.

Consequently, the `Unary`, non-reference `Infix`, and generic `Function`
branches of the kind walker can report `Array` for a list. A lazy fallback can
then be refused as `Unsupported(ReferenceOperator)` before the ordinary VM has
converted the list to `Error(Value)` and allowed `IFERROR` to select the
fallback.

A reachable matrix-demanding example is:

```text
=IF({TRUE()};([.A1]:IFERROR(XOR(([.A1]~[.B1])![.A1]);[.C1:.D1]));0)
```

The inner `!` returns a reference list. The ordinary matrix VM converts that
list when evaluating `XOR`, yielding `Error(Value)`; `IFERROR` catches it and
returns `[.C1:.D1]`, which the enclosing range can consume. The current kind
walker classifies the list as `Reference`, then `XOR` as `Array`; the enclosing
handler is consequently rejected during shape planning. Root's baseline-40
diagnostic reproduces `Unsupported(ReferenceOperator)`. Existing direct
`NOT(list)` coverage does not exercise this lazy reference-operand path.

The same mismatch is reachable with unary or binary scalar functions, for
example `IFERROR(NOT(([.A1]~[.B1])![.A1]);[.C1:.D1])` used as a range operand.
The fix must preserve a distinct list kind (or retain the runtime area value)
through unary, infix, function, and lazy continuations, then apply the actual
list conversion rules. A syntax-only whitelist or treating every area as an
array would retain the wrong type boundary.

## Non-blocking follow-up

The reference-infix fast path at lines 3782–3785 propagates every
`ReferenceOperandKind::Error` before invoking `coerce_reference_areas`, while
the runtime operator specially propagates only `ScalarError::Reference`.
That is broader than the runtime contract and merits a regression for a
non-reference error flowing into a second reference operator. The directed
nested-`Null` case now returns the expected typed refusal, so this review does
not retain it as an acceptance blocker for the tested expression.

The kind walk resolves runtime geometry and then the later shape walk resolves
it again; scalar lazy values can likewise be probed before ordinary evaluation.
This is bounded and its temporary reservations are released, but it duplicates
work/admission checks and should remain visible in resource/performance
documentation. It is not a leak finding.

## Disposition

The explicit stack, lazy branch exclusion, and reservation lifetime are sound
within the reviewed snapshot. RRK-1 remains a production correctness blocker
for the advertised evaluated-reference/list profile until a distinct
`ReferenceList` state is carried through the kind walker and the lazy fallback
regression above passes. This artifact does not claim full VM acceptance.
