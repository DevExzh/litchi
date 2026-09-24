# ODF value-inspection and conversion functions: integration review

Status: design review only. This report contains no production or test edits. The
normative contract is still being extracted by the semantic review; the integration
plan below deliberately leaves the contract-dependent array, reference, and error
choices explicit.

The batch covers the sixteen pure §6.13 functions:

`ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`,
`ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `NUMBERVALUE`, `TYPE`,
and `VALUE`.

## Current seams

The value VM already has the right representation for preserving the distinctions
these functions need. `Value` and `RuntimeValue` retain `Empty`, `Text`, `Logical`,
`Number`, `Error`, arrays, references, and reference lists. `RuntimeElement` retains
`Empty` and `Present(WorkingValue)`. `read_to_element` also creates a borrowed
`TextValue` after checking the text limit. A reference cell can therefore reach a
function without being cloned or silently changed into a scalar zero.

The scalar bridge is the dangerous boundary. `Slot` has `Empty` and `Value`, but
`scalar::argument` currently coerces `Slot::Empty` to a target value (normally
number zero) before the scalar evaluator sees it. That is correct for ordinary
arithmetic argument conversion and is incorrect for raw identity functions such as
`ISBLANK`, `ISNUMBER`, `ISTEXT`, `TYPE`, and `ERROR.TYPE`. Adding an `Empty` variant
to the global `WorkingValue` would widen the scalar ABI and would change existing
coercion paths. The bounded integration should keep `WorkingValue` unchanged and
add an inspection input/apply path beside `scalar::argument`, consuming `Slot`
before empty coercion.

The scalar evaluator's ordinary function admission and dispatch are in
`evaluation.rs` (`visit_function`, eager scheduling, and `apply_function`). The
scalar bridge's `eager` allow-list is a second admission point. All three must
recognise the new family, with per-function arity and input policy, rather than
falling through to the generic invalid-arity/error propagation path.

At the value layer, generic `ValueEvaluator::map_function` first turns
`RuntimeValue::Areas` into an array by `materialize_for_array`. That allocates a
cell vector and reads the entire area. `map_text_function` is the existing safe
pattern: it computes the broadcast shape, reserves a bounded output vector, selects
one `RuntimeElement` per output coordinate, and lets `select_area_element` stream a
single reference cell with work, cancellation, read-count, and text-limit checks.
Inspection functions that can receive references must use that pattern (or a
generalized mapper), not the generic materializer.

The generic `apply_function` path also converts a reference list to a formula
`#VALUE` before ordinary matrix dispatch. A contract-defined shape or pseudotype
refusal must run before that conversion and before `select_area_element`; a refusal
must perform zero resolver reads. Conversely, a contract that permits a descriptor
to be classified without reading cells must keep that operation metadata-only.

`shape_hint_demand` and `combine_planned_children` currently know the reducer and
matrix families. They do not have a value-inspection family with an explicit
output-shape policy. The new functions cannot be added as one blanket “scalar” or
one blanket “matrix” family: `TYPE` array semantics and the matrix lifting policy
for the conversion/predicate functions are contract decisions.

Finally, the demand-cache classifier and apply-time cache branches currently know
the existing sequence, reducer, text, and conditional families. A projected lazy
`IF` can therefore reuse a scalar result unless the new position-sensitive
functions are classified deliberately. A reference/range value selected at two
different output coordinates must not share a cached result merely because the
function itself is pure.

## Bounded implementation plan

### 1. Use one function catalog, with raw and converting input modes

Add a small shared catalog (or equivalent predicates in the existing function
modules) describing, for each name:

* arity and whether a missing argument is legal;
* raw identity input versus explicit numeric/text conversion;
* whether formula errors are inspected as values or propagated;
* accepted scalar/array/reference pseudotypes;
* scalar result versus elementwise matrix result; and
* cache and reference-coordinate sensitivity.

The catalog should be used by both scalar dispatch and value-layer scheduling so a
function cannot be admitted by one layer and rejected by another. Keep exact
normative codes and coercion rules in the contract table, rather than duplicating
ad-hoc name lists in `evaluation.rs`, `value/scalar.rs`, and `value.rs`.

The scalar-side input can remain private and small:

```text
InspectionInput<'a> = Empty | Value(WorkingValue<'a>)
```

It can be constructed from `Slot` before `scalar::argument` is called. Raw
predicates inspect this input directly. `N`, `VALUE`, `NUMBERVALUE`, `ISEVEN`, and
`ISODD` use a separate, explicit conversion operation whose treatment of `Empty`,
text, logicals, errors, non-finite numbers, and complex values comes from the
contract. `NA` has no input and returns the specified formula error value.

This keeps the existing arithmetic/logical coercion ABI intact. It also prevents an
empty cell from becoming number zero before `ISBLANK` or `TYPE` sees it. A scalar
function cannot classify an array or a reference by inventing a `WorkingValue`;
descriptor-level handling belongs to the value VM described below.

The raw/converted distinction is necessary for error behavior. Formula errors
stored in a cell are values that `ISERR`, `ISERROR`, `ISNA`, and, where specified,
`ERROR.TYPE` may inspect. Provider/resource/cancellation/source failures are
evaluation failures, not formula values; they must never be caught or converted by
these functions or by `IFERROR` around them.

### 2. Add a streaming inspection mapper

Add `map_inspection_function`, or generalize `map_text_function` into a mapper that
takes a scalar input adapter. Its reference path should:

1. perform arity and contract pseudotype checks before selecting or reading a cell;
2. compute the contract-defined shape without materializing a `RuntimeAreaSet`;
3. reserve only the bounded output array and reusable scalar argument storage;
4. charge cell work before each selected cell, then use
   `read_reference_cell` and `read_to_element`;
5. pass `RuntimeElement::Empty` as raw `Slot::Empty`, preserving borrowed text and
   formula-error identity; and
6. publish only after the normal source-version and cancellation fences.

This path must be selected before the generic `RuntimeValue::Areas` list conversion
and before `materialize_for_array`. It should share `runtime_shape`,
`select_matrix_value`, `select_area_element`, and the existing storage/work
reservation helpers where their contract allows. In particular, it must retain the
existing order of work charging, max-reference checks, cancellation checks, and
successful-read accounting. Typed resolver failures (`Unsupported`, resource
limit, cancellation, or source change) must bubble directly while retained formula
errors remain ordinary input values.

For a known shape/type refusal, the mapper must return the contract's formula error
or typed refusal before any resolver call. This includes the contract's decision on
reference lists/unions, multi-area descriptors, and three-dimensional references.
Computed reference expressions may need to evaluate their own upstream expression;
the zero-read guarantee applies once the resulting descriptor is known to be a
refusal. Do not use the existing whole-area materializer as a shortcut for a
computed expression.

### 3. Schedule and dispatch the family before generic scalar mapping

In `visit_function`, recognise the catalog before the generic eager/matrix argument
classification. Preserve a complete reference descriptor when the function's
contract operates on a reference; do not apply implicit intersection while merely
building a projected lazy branch. If a function is elementwise in matrix mode,
schedule its argument as a matrix argument and let the mapper select the coordinate.
If it is scalar-only, keep the argument descriptor intact until the explicit
scalar/reference policy is applied.

In `ValueEvaluator::apply_function`, dispatch the family before the existing
`Areas.is_list` conversion and before the generic `map_function` materialization.
The scalar-mode branch should use the raw `Slot` adapter for ordinary scalar values;
the matrix/reference branch should use the streaming mapper. Apply strict arity
checks and the catalog's per-function error policy at the dispatch boundary.

The scalar `apply_function` and `value/scalar::eager` allow-list need corresponding
branches. Unsupported names must not reach the ordinary invalid-arity path after
their children have been evaluated, because that path assumes the ordinary
coercion and formula-error rules.

### 4. Make shape and reference classification contract-driven

Introduce a small shape policy for this family and use it in
`shape_hint_demand`, `combine_planned_children`, and the reference-kind/function
classification. At minimum it must distinguish:

* scalar-result inspection/conversion functions whose input may be a complete
  reference descriptor;
* elementwise functions that lift an array/reference to one result per selected
  coordinate; and
* `TYPE` (and any other function whose contract classifies an array/reference as a
  value without iterating it).

Do not let a scalar-result function accidentally inherit the ordinary child shape
and produce an array. Do not force an elementwise function through the reducer
shape of `AVERAGE` either. The exact entries are blocked on the normative contract,
especially array `TYPE` codes, reference-list acceptance, and whether matrix mode
is an implicit-intersection or lifting operation for each predicate/converter.

The same policy must be used by reference-kind classification. A complete reference
argument must not be silently reduced to the current coordinate when the function
is defined over the descriptor, and a function that is scalar-only must reject an
unsupported list/area shape before reads.

### 5. Protect lazy-IF demand caches

Classify the new family in the early projected lookup and apply-time cache paths.
Add a dedicated `cacheable_inspection_branch` rather than treating every pure
function as position invariant. A conservative first rule is:

* literal-only arguments, and references proven to be one invariant cell, may use
  scalar payload/error-only cache entries;
* range references, implicit intersections, projected `RuntimeAreaSet` values,
  position-sensitive descendants, and unresolved/computed descriptors do not use a
  shared scalar cache entry; and
* a cache entry never stores a typed evaluation failure and never turns a cached
  formula error into a provider/resource/cancellation result.

If the contract gives an operation complete matrix arguments and proves that the
result is coordinate invariant, that operation can opt into the existing complete
argument cache path. Otherwise, keep the conservative per-coordinate evaluation.
The classifier must remain separate from the function's purity: `ISNUMBER` of a
range selected under a projected `IF` is pure but position-sensitive.

A minimum regression should use two cells with different types, for example a
projected lazy branch equivalent to:

```text
IF({TRUE;TRUE}; ISNUMBER([.A1:.A2]); FALSE)
```

and assert the two outputs differ and the two reference cells are read once each.
Equivalent cases are needed for `ISBLANK`, `ISTEXT`, `TYPE`, and the conversion
functions once their matrix/reference contract is fixed. A cache hit for a literal
`ISNUMBER(1)` should remain possible; the test should not force all inspection
functions out of the cache.

## Resource and failure requirements

The implementation should retain the following existing guarantees:

* zero resolver reads for a known descriptor/shape refusal;
* one streaming read per admitted projected cell, with work charged before the
  read and cancellation/max-reference checks before and after the provider call;
* borrowed text through `TextValue::borrowed`, with text-byte charging before any
  parser or conversion work and no whole-haystack/whole-reference clone;
* checked output/reservation limits and refund/drop order for arrays, temporary
  arguments, and conversion buffers;
* source-version and cancellation fences before and after evaluation and before
  publication; and
* typed failures bubbling through `IFERROR`, `ISERROR`, `NA`, and all new dispatch
  paths. Only formula errors already represented as cell values may be inspected or
  returned according to the function's contract.

`NUMBERVALUE` and `VALUE` need particular care around separator/parser scratch
storage. They should charge the borrowed input text and reserve only bounded
temporary/output state. `N` may preserve an admitted formula error as a value if
the contract says so; it must not use a broad `Result` catch that also catches
resource or provider failures. `ERROR.TYPE` and the `IS*` family should inspect a
formula-error value without invoking numeric/text coercion.

## Contract gates before implementation

The following choices must be recorded in the semantic contract and then copied
into the shared catalog before production coding:

1. Which functions accept `Any`, `Scalar`, a complete `Reference`, an array, or a
   reference list; and which known shapes are formula `#VALUE` versus typed
   unsupported. This determines the zero-read preflight.
2. Whether a formula error is inspected or propagated for every function, including
   `ERROR.TYPE`, `N`, `TYPE`, `VALUE`, and the parity predicates. Empty cells and
   formula `""` text must remain distinguishable.
3. `TYPE` codes and its array/reference behavior, including whether it classifies a
   descriptor without iteration, lifts elementwise, or uses implicit intersection.
4. Matrix lifting versus scalar-result behavior for all predicates and converters,
   including nested calls under lazy `IF`.
5. `ISEVEN`/`ISODD` numeric conversion, sign, fractional, zero, and non-finite
   rules.
6. `NUMBERVALUE` separator defaults, validation, grouping, whitespace/sign rules,
   and output/error behavior; and the accepted grammar/date behavior of `VALUE`.
7. `NA` arity and whether its returned `#N/A` is a formula value catchable by
   `IFERROR`.

Until these gates are closed, a semantic implementation cannot be called complete.
The integration boundary is otherwise bounded: retain `WorkingValue` as-is, add a
raw `Slot` inspection path, stream reference operands through a mapper generalized
from the text path, and keep position-sensitive inspection results out of shared
demand-cache entries.

## Validation plan

After the contract is frozen, add focused tests covering scalar values, `Empty`,
empty text, numbers, logicals, complex values, every formula error, borrowed text,
wrong arity/missing arguments, arrays and broadcasts, single-cell references,
reference lists/multi-area descriptors, and computed reference expressions. Add
resource-limit cases for zero-read refusal, reference cells, text bytes, output
cells, work, cancellation, source change, and typed resolver failures after an
earlier formula error.

Run cache cases with two positions whose referenced cells differ in both type and
UTF-8 text width. Verify that a projected branch does not reuse a position-sensitive
inspection result, while literal-only results remain cacheable. Run the focused
semantic and limits targets, then the locked/offline package check and the isolated
gate suite. This report makes no PASS claim for those tests because production
integration has not started and the normative contract is still open.
