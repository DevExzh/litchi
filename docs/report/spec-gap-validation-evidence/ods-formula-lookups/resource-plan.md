# Lookup and reference-function resource plan

Status: discovery and design review only. This file makes no production-support
claim and contains no implementation or test result. The reviewed baseline is
`635fd2e1348b621426b50909cbd5765c91837306`; the normative scope is
`ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`, `LOOKUP`, `MATCH`,
`OFFSET`, and `VLOOKUP`.

## Current evaluator seam

The function catalog names all nine functions, but the value VM has no lookup
module or dispatch. `ValueEvaluator::visit_function` in
`crates/litchi-ods/src/codec/formula/evaluation/value.rs` classifies only the
existing matrix, aggregate, statistical, descriptive, paired, order,
metadata, conditional, database, and logical families. `apply_function` has no
lookup branch, so these names reach generic argument projection and the scalar
bridge, which reports an unsupported function.

That generic path is not a safe fallback for this batch. It can project or
materialize an area before a ForceArray consumer sees it, and it visits every
ordinary argument before application. In particular, it would evaluate all
`CHOOSE` branches, lose descriptor identity for `OFFSET`/`INDIRECT`, and infer
the wrong shape or reference kind for `INDEX`.

The existing value VM already supplies the resource primitives required by a
lookup implementation:

- `RuntimeAreaSet` retains ordered physical areas and logical records,
  including `is_list`, 3-D planes, duplicate records, and checked cell/area
  reservations.
- `read_reference_cell` checks cumulative `max_reference_cells` and
  cancellation before and after `Resolver::read_cell`, then counts successful
  reads. `read_to_element` preserves resolver text as borrowed `TextValue`.
- `charge_cell_work` provides per-cell work and periodic cancellation checks;
  `new_element_vec` and `ensure_capacity` bound result and evaluator storage.
- `evaluate` fences source identity before and after execution and checks
  cancellation before publishing a retained result.
- `RuntimeAreaSet::derived` creates a first-class reference record with
  `reference: None`; this is the correct lifetime-neutral representation for a
  computed reference, provided consumers use its areas and records rather than
  assuming a lexical `Reference` is present.

`Resolver` exposes one-cell reads plus bounded sheet extent/order/name queries,
not a range iterator. Correct lookup code should therefore stream through the
existing one-cell API. An optional provider optimization can be considered
later, but the evaluator must not depend on direct worksheet storage.

## Argument contexts and result kinds

The lookup module should expose function classification, argument-context
selection, application, cache classification, shape planning, and reference
kind planning. The VM hooks must be added before the module can be reachable.

| Function | Required argument context | Result/resource behavior |
| --- | --- | --- |
| `ADDRESS` | Scalar row, column, absolute-mode, A1-mode, and sheet arguments | Bounded text result; no resolver reads. |
| `CHOOSE` | Visit index first; visit only the selected value branch in the caller's mode | Selected scalar, array, or reference is preserved; unselected branches are untouched. |
| `INDEX` | Data source is `ReferenceList\|Array`, retained complete as `Any`; row, column, and area number are scalar | Select one logical reference record, preserving its 3-D physical planes. Row/column zero or omission may produce a bounded row, column, or area result according to the contract. |
| `MATCH` | Scalar search and match type; complete `Reference\|Array` vector | Scalar one-based position; validate one-row/one-column shape before reads. |
| `LOOKUP` | Scalar search; complete searched source and optional complete result source | Scalar selected result; validate orientation and vector/result extension before reads. |
| `VLOOKUP` | Scalar lookup, column, and range flag; complete table `Reference\|Array` | Scalar selected cell; validate table height/column before reads. |
| `HLOOKUP` | Scalar lookup, row, and range flag; complete table `Reference\|Array` | Scalar selected cell; validate table width/row before reads. |
| `OFFSET` | Complete single `Reference`; scalar row/column offsets and optional dimensions | Derived reference only; no cell reads. Reference lists are refused before scanning. |
| `INDIRECT` | Scalar text and optional A1 flag | Derived reference only; parse and resolve geometry without cell reads. |

`INDEX` is the important correction to a generic ForceArray rule: its first
parameter is `ReferenceList|Array`, so complete `Any` context must preserve an
ordered list and selected descriptors. `OFFSET` rejects a `ReferenceList`.
`INDEX` area selection selects one logical record and clears list identity,
while retaining all physical planes of that selected 3-D record.

`CHOOSE` needs a dedicated lazy continuation analogous to the existing IF
frames. The index must be converted and range-checked before a branch frame is
created. A selected branch must inherit the active scalar/matrix projection;
the branch may still be a complete reference or array when its consumer is a
ForceArray or reference consumer. A generic `Apply` frame cannot provide this
behavior because its argument loop has already evaluated every child.

## Reference, shape, and source-policy integration

The target functions require explicit cases in all four planning paths:

1. `visit_function` must choose complete source/reference arguments and the
   lazy `CHOOSE` schedule. Projected IF branches must retain full source
   descriptors, while scalar lookup keys and row/column parameters remain
   position-sensitive where the contract permits arrays or computed values.
2. `shape_base` must report scalar shape for `ADDRESS`, `MATCH`, `LOOKUP`,
   `VLOOKUP`, and `HLOOKUP`; derive `OFFSET` shape from its input and explicit
   dimensions; derive `INDEX` shape from source geometry and row/column
   parameters; and defer dynamic `CHOOSE`/`INDIRECT` shape when it cannot be
   determined without evaluating a value. Do not combine the shapes of all
   CHOOSE alternatives.
3. `reference_kind` must classify `OFFSET` and valid `INDIRECT` as references,
   preserve selected `CHOOSE` kinds, and distinguish `INDEX`'s selected record,
   scalar cell, and array outputs. The search functions and `ADDRESS` are
   scalar results. Unknown generic-function handling must not be allowed to
   turn a reference result into an implicit intersection.
4. `reference_shape_value` must use the same selected-branch and source-policy
   rules as runtime evaluation. Internal deferred/refused descriptor probes
   may be mapped to planning outcomes, but public resolver failures must remain
   typed `EvaluationFailure`s and must not become formula errors.

`OFFSET` and `INDIRECT` should construct `RuntimeAreaSet::derived`, retaining
   canonical `SheetRef` names from the resolver. `INDIRECT`'s temporary parsed
   reference owns syntax text, so it must not be stored in the runtime area
   lifetime. Resolve a named sheet with `sheet_index` and then borrow its stable
   `sheet_name_at` result; use `Current` for a current-sheet address. This
   avoids adding an owned `SheetRef` lifetime or leaking parser strings. The
   resulting public `ReferenceView.reference()` may be `None` by design, while
   `ISREF`, `AREAS`, `ROWS`, `COLUMNS`, `SHEET`, and `SHEETS` continue to use
   retained geometry.

Source-qualified references must remain inert or typed unsupported according to
the frozen profile. They must not be resolved by `INDIRECT`, and provider
failures from sheet lookup, extent, or name queries must not be caught by
`IFERROR` during shape or reference-kind discovery.

## Read, work, cancellation, and storage rules

Every function must preflight arity, pseudotype, list/3-D/vector/table shape,
integer conversion, dimensions, and source policy before its first `read_cell`.
The following are zero-cell-read refusals:

- invalid `CHOOSE` index or missing selected branch;
- invalid lookup table/vector dimensions, list admission, row/column/index,
  area number, range flag, or result-vector length;
- `OFFSET` list input, signed-coordinate overflow, out-of-extent output, or
  non-positive explicit height/width;
- malformed or source-qualified `INDIRECT` text; and
- any `ADDRESS` conversion or output-size refusal.

Resolver metadata calls needed to establish sheet order, extent, or canonical
names are still work and cancellation operations; “zero reads” means no cell
value access. Geometry must be checked against both `max_reference_cells` and
`max_reference_areas` before retaining an area set. This applies to a derived
`OFFSET`/`INDIRECT` reference even when its eventual consumer will read only one
cell.

Reference searches must stream in logical row/column and record order. For each
candidate, charge work before calling `read_reference_cell`, convert with
`read_to_element`, and keep only fixed-size comparison state and the selected
coordinate. Exact searches may stop at their first match. Approximate searches
may use a bounded binary or upper-bound search, but the algorithm must retain
the specified duplicate tie and its read trace must be tested. Do not build a
candidate-cell vector or materialize an entire input reference.

For a selected result, read only the returned cell after the search unless the
candidate and result coordinate are the same already-read cell. For arrays,
borrow existing elements and clone only the scalar value needed by a result;
for reference row/column/area outputs, allocate only the selected output and
charge it against `max_array_cells`. Every output allocation must use checked
capacity and drop storage before releasing its reservation.

Comparison and address parsing must charge input text bytes without cloning
borrowed resolver text. `INDIRECT` parsing needs a bounded text/work budget;
`ADDRESS` must reserve and bound generated text. Long search and array loops
must call the same work/cancellation path even when they do not read a cell.

Formula-error cells are values governed by the lookup contract. Typed resolver,
resource, cancellation, source-version, and unsupported-cell failures are
`EvaluationFailure`s and must bubble through the lookup module. They cannot be
converted to `ScalarError`, hidden by a search fallback, or made catchable by
`IFERROR`/`IFNA`. Source and cancellation fences remain outside the complete
lookup evaluation and publication.

## Demand-cache rules

The current demand cache stores only scalar Empty, Missing, Number, Logical,
Error, and Complex payloads. It cannot safely retain a reference descriptor,
array, or owned text result.

Cache scalar `MATCH`, `LOOKUP`, `VLOOKUP`, `HLOOKUP`, and scalar `INDEX` only
when the source is complete and all search/index/flag arguments are proven
coordinate-independent. Direct descriptors may be cacheable within one
source-fenced evaluation; computed projected source expressions must remain
uncached. The lookup classifier must inspect nested descendants conservatively
and retain the existing position-sensitive exclusion for `MUNIT` and similar
scalar descendants.

`CHOOSE` may cache a selected scalar branch only after the index and selected
branch are independently invariant. It must never visit or classify an
unselected branch as evaluated. `OFFSET` and `INDIRECT` return descriptors and
must not enter the scalar payload cache. `ADDRESS` should remain uncached until
there is a bounded text cache representation; otherwise a projected text result
could be incorrectly reused or force an unbounded clone.

The early projected cache lookup must happen before scheduling source arguments,
as it does for existing reducers. The apply path must use the same classifier
for get/put. Cache hits must not bypass the outer source-version and
cancellation fences.

## Contract and test prerequisites

The contract must settle these points before implementation and oracle capture:

- which `ReferenceList` and 3-D forms each consumer admits, and how INDEX
  `AreaNumber` counts logical records while preserving selected 3-D planes;
- INDEX omitted/empty/zero row and column result kind and shape, including
  scalar publication of a selected row or column;
- CHOOSE array-index and branch-array behavior, fractional index conversion,
  explicit empty versus missing branch slots, and lazy branch errors;
- integer, logical, text, Empty, missing, and formula-error conversion for all
  indices, flags, offsets, and dimensions;
- MATCH/LOOKUP/HLOOKUP/VLOOKUP approximate ordering, duplicate direction,
  Number/Text mismatch, case folding, error cells, and unsorted input;
- LOOKUP two- versus three-argument orientation, result-vector length and
  extension rules;
- OFFSET empty dimension slots, negative offsets, list refusal, 3-D identity,
  and boundary behavior; and
- INDIRECT A1/R1C1 syntax, `.`/`!` separators, quoted sheets, relative formula
  position, external/source text, whole-axis text, and malformed names.

Focused limits tests should assert zero cell reads for every refusal, bounded
ordered reads for exact and approximate searches, no reads for reference
construction, no reads in unselected CHOOSE branches, typed-failure
precedence, borrowed text, source/cancellation fences, and bounded state under
large vectors or 3-D records. Matrix tests must cover projected IF/IFERROR,
selected references, computed arrays, cache hits, and position-sensitive nested
expressions.

## Suggested implementation order

1. Freeze the normative contract and conversion/comparison profile.
2. Add the shared lookup catalog, checked scalar converters, comparison kernel,
   bounded A1/R1C1 parser, and ADDRESS formatter.
3. Add descriptor selection and streaming search helpers, then VM scheduling,
   apply, shape, reference-kind, reference-shape, and cache hooks.
4. Implement lazy CHOOSE, derived OFFSET/INDIRECT, and INDEX record selection;
   then implement MATCH and the LOOKUP/HLOOKUP/VLOOKUP search wrappers.
5. Run focused semantic/resource/cache tests and only then freeze source for
   independent oracle, native, and isolated-gate evidence.

## Historical implementation review notes

The following notes retain findings from an early, unfrozen implementation
snapshot. References to the current implementation below refer to that snapshot,
not the final candidate. They are not a final support disposition; the final
resource review and frozen tests must establish which findings were resolved.

### INDIRECT parser reservation

`evaluation/lookup/indirect.rs` currently reserves
`body.len() * 4 + 128` for `reference::parse_body` and sets
`max_components` from the input length. That envelope does not demonstrate
coverage of the parser's owned `String` components, endpoint structures, and
the growing `Vec<Subtable>`. The parser grows that vector up to its component
limit; a component-count-derived `size_of::<Subtable>()` term and checked
fixed-structure terms are required. A long quoted-sheet/subtable input should
be a focused max-storage test at the exact envelope boundary.

The normalization `String` is also still live while `parse_body` runs. Both
the A1 and R1C1 normalizers currently release their normalization reservation
before returning an owned `Cow`/`String`, so the parser reservation does not
cover the simultaneous normalization scratch. Keep that reservation alive
until the normalized string is dropped, or reserve one checked combined
envelope and release the scratch only after parsing. The parser's internal
`try_reserve` calls cannot enforce the evaluator budget by themselves.

`map_reference_error` maps every unlisted `litchi_core::Error` variant to
formula `#REF!`. Only syntax/format errors should become a formula reference
error. Resource and allocation errors are already preserved; unexpected
provider, execution, or future typed variants must remain typed evaluator
failures rather than being erased by the wildcard arm. Add a focused injected
error test at this boundary.

The R1C1 path performs several potentially long lexical scans and reaches a
single `charge_work(0)` after normalization. The initial byte charge checks
execution only once. Periodic work/cancellation checkpoints are needed in the
sheet split, top-level separator, endpoint, and output scans, as already done
for the main A1 normalization loop.

### Search and ADDRESS adapters

The search kernel correctly keeps comparison state fixed-size, retains text by
borrowed `&str`, and propagates `read_reference_cell` failures. The following
resource details still need explicit closure:

- The `LOOKUP` result-extension path calls `sheet_extent` and constructs an
  enlarged `RuntimeArea` without the ordinary complete geometry/area admission
  path. Check both reference limits, charge metadata work, and check
  cancellation before publishing the extended view.
- A searched reference candidate and the selected result can be the same cell.
  Decide whether the read trace intentionally counts a second selected-cell
  read or retain the last candidate cell as fixed state and avoid rereading it;
  pin that choice in focused read-count tests.
- `SearchGrid` currently refuses lists and multi-plane areas before cell reads.
  This is acceptable only after the contract freezes that policy for every
  search function. A 3-D/list refusal must stay a zero-cell-read path.
- `ADDRESS` reserves output bytes and charges input/output text, but output
  sizing helpers use saturating arithmetic for quoted sheet names. Replace
  saturation with checked sizing so an overflow cannot be hidden by a later
  length mismatch or allocation attempt. Long sheet/output loops also need a
  cancellation checkpoint policy.

No production or test edit was made by this review. The in-progress source
remains outside any PASS claim until these points and the later full evaluator
review are complete.
