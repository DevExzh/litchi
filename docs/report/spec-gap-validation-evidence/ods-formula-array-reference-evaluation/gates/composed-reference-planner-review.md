# Composed reference shape planner review

Status: source-only review; no production edits, Cargo runs, or runtime acceptance claim.

This review covers the generalized reference-operator shape path in `value.rs` and its
`:`, `!`, and `~` implementations in `references.rs`. It is limited to explicit stack
geometry evaluation, list/plane handling, budget and cancellation behavior, and
provider/materialization boundaries. Cache and fixed-point behavior is outside this
review.

## Reviewed source

| Source | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `c76e3ddb2546b151a3189a578c5c08208e500c5264de4ba9fa92dff9ea0c464e` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `680a3fd9c3b80951408f3210f5c5373dcea7dd00c1b561b0902043438b79d907` |

## Confirmed sound within the syntactic-reference scope

`reference_shape_value` uses an explicit `Visit`/`Apply` stack for parenthesized
references and nested `:`, `!`, and `~` nodes. It does not recurse through the AST.
The stack and value vector grow through `ensure_capacity`, which checks the scalar
stack limit, the storage budget, and the retained execution context before each
allocation. Each frame is charged through the shared scalar evaluator, so work and
cancellation remain cumulative with the surrounding shape/value evaluation.

The helper delegates area construction and operator semantics to the runtime reference
module. Range construction preserves the inclusive sheet span and bounding rectangle;
intersection preserves record/list identity and only retains intersecting records;
union preserves ordered areas, lists, and duplicates. The shape adapter accepts only a
single non-list area on one plane. A list or multi-plane result remains unknown rather
than being flattened into a misleading two-dimensional shape.

These paths inspect sheet metadata and finite extents but do not read provider cells or
materialize selected arrays merely to infer geometry. The lazy shape planner selects a
branch before descending into it, so the existing selected-branch-only behavior does
not resolve an unselected branch. Temporary runtime area values, vectors, and their
reservations are local to the shape probe and are dropped on success, `Ok(None)`, and
error paths. I found no reservation or parent VM-stack leak in this path.

The planner does perform the reference operator once during shape planning and the
ordinary value path performs it again. That duplicates bounded metadata work and
admission checks, but the work/cancellation/storage checks use the shared evaluator
and the temporary allocations are released. It should remain visible in performance
and limit documentation; it is not by itself a correctness blocker.

### CRP-2: temporary shape buffers release their charge before their storage

The declarations in `reference_shape_value` currently place each `Vec` before its
matching `Option<Reservation>` (`frames` then `frame_reservation`, and `values` then
`value_reservation`). Local Rust bindings are dropped in reverse declaration order,
so an early `?`, `Ok(None)`, or normal return releases the reservation before dropping
the vector buffer. This is the opposite of the retained-storage invariant used by
the owned result, where storage is declared before its reservation so the bytes are
freed before the budget charge is released.

This is not a permanent leak and does not change the calculated geometry, but it can
temporarily under-report shared `Resource::Memory` usage while the shape buffers (and
any `RuntimeValue` area storage they own) are still being deallocated. That matters
when evaluations share a budget concurrently and is a resource-accounting defect
for a strict retained-memory guarantee. It affects all exits, including unsupported
operands and limit/cancellation errors, because those exits use automatic local
destruction.

The smallest safe repair is a private holder whose fields are declared as storage
followed by its reservation, for example a holder containing
`frames`, `frame_reservation`, `values`, and `value_reservation`. Dropping the holder
then drops both vectors (including their elements) before either outer reservation,
and the same order applies on every early return. An explicit cleanup closure can
also enforce the order, but it must wrap every `?`/`Ok(None)` path; a declaration-only
reordering of the four current locals is easy to regress.

## Blocking finding

### CRP-1: evaluated reference-valued lazy functions are silently under-shaped

`reference_shape_value` handles `Parenthesized`, `Reference`, and reference infix
nodes only. Its catch-all arm returns `Ok(None)` for a `Function` node. The outer
`reference_operator_shape` then returns `None`, and `combine_planned_children` uses
that result as the final shape for every `:`, `!`, or `~` node, even when its child
shape traversal already learned more geometry.

The value VM can produce a reference from a lazy function. In matrix mode,
`finish_if` projects an array-like condition but otherwise calls `push_branch`; a
selected direct reference remains a `RuntimeValue::Areas`. The runtime `combine_range`
path accepts that value as an operand. Therefore this is an evaluated-reference case,
not merely a syntactic spelling issue.

For example, a matrix-demanding expression of this form selects `C1:D1` while keeping
the missing reference unselected:

```text
=IF({TRUE()};([.A1]:IF(TRUE();[.C1:.D1];[Missing.A1:.Z100]));0)
```

The selected inner function returns a two-cell reference, so the outer range's
inclusive geometry is `A1:D1` (one row by four columns). The general shape walk can
infer the inner lazy branch, but the outer reference-specific helper encounters the
`IF` function and returns unknown. The generic shape fallback can consequently retain
a `1x1` demand; the later matrix evaluation then has no demand for the remaining
cells. The same risk applies to `IFERROR` and `IFNA`, including a selected
reference-valued alternative.

This is a blocker for a profile that advertises reference operators over evaluated
`Reference`/`ReferenceList` values. The local specification review describes the
operands semantically as references/lists, and the value VM already preserves a
reference as a first-class matrix result. The implementation must either:

1. extend the bounded shape evaluator to handle reference-valued lazy functions,
   evaluating only the selected branch and retaining the same work/stack/cancellation
   accounting; or
2. explicitly type-refuse function-produced references in these operators before
   shape planning, with a documented narrower profile and a regression proving that
   no silent `1x1` fallback occurs.

A regression is required for the expression above and for a wider selected reference
returned by `IFERROR` or `IFNA`; it should verify output shape and selected-cell
metadata/reads, and verify that `Missing.A1:Z100` is never resolved. Add list and
multi-plane variants to ensure a returned `ReferenceList` remains unsupported for a
2-D matrix demand rather than being flattened.

## Additional bounded regressions to retain

- Deep nested syntactic `:`, `!`, and `~` expressions should refuse through the
  typed stack/work limit without Rust stack growth.
- A low storage or cancellation budget should refuse while the temporary shape areas
  and reservations are still released.
- A multi-sheet range should charge and validate every physical sheet in the span,
  including intermediate sheets, while preserving duplicate logical sheet names by
  physical index.

## Disposition and scope

The explicit stack, syntactic operator geometry, list distinction, provider read
boundary, and local reservation cleanup are cleared by source inspection. CRP-1 keeps
the generalized VP15-3 planner from being accepted as complete: the current helper is
sound only for syntactic reference/operator trees whose operands resolve directly to
references. CRP-2 is a separate bounded resource-lifetime follow-up: all temporary
storage is eventually released, but the release order does not yet meet the shared
budget's storage-before-reservation invariant. No full VM, cache/fixed-point, or
integration-test acceptance is claimed here.
